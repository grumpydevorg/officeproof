using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Spreadsheet;

/// <summary>
/// The write pass of <see cref="CellEditor"/>, on XDocuments rather than the SDK's typed model.
/// 2026-10-15, bats risk xlsx: re-serialising a worksheet through the typed model rewrote Excel's
/// default-namespace markup as <c>x:worksheet</c>/<c>x:c</c> prefixes. Valid, but every line of the part
/// changes and naive readers that match unprefixed names break. XLinq keeps the part's own prefixes,
/// attribute order and namespace declarations, so only the edited elements differ.
/// </summary>
internal static partial class WorkbookWriter
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // CT_Workbook children that follow calcPr, in schema order (ECMA-376 Part 1, 18.2.27).
    private static readonly string[] AfterCalcPr =
        ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
         "fileRecoveryPr", "webPublishObjects", "extLst"];

    /// <summary>Applies the plan; returns whether a calcChain was deleted.</summary>
    public static bool Write(SpreadsheetDocument doc, List<Planned> plan)
    {
        var wb = doc.WorkbookPart ?? throw new InvalidDataException("no workbook part");
        // sharedStrings is rewritten only when a string is written or a shared-string cell is
        // overwritten (its reference leaves the shared-string count, review 2026-10-15); otherwise it keeps its bytes.
        var strings = plan.Exists(p => p.Kind == CellValueKind.Text || p.ReplacesSharedString) ? new SharedStrings(wb) : null;

        foreach (var group in plan.GroupBy(p => p.RelId))
        {
            var part = wb.GetPartById(group.Key);
            var xml = Docx.Load(part);
            var worksheet = xml.Root ?? throw new InvalidDataException($"{part.Uri} is empty");
            var data = worksheet.Element(S + "sheetData") ?? throw new InvalidDataException($"{part.Uri} has no sheetData");
            foreach (var edit in group)
                Store(GetOrCreateCell(worksheet, data, edit.At), edit, strings);

            if (worksheet.Element(S + "dimension")?.Attribute("ref") is { } dimension && CellRange.TryParse(dimension.Value, out var range))
            {
                foreach (var edit in group) range = range.Including(edit.At);
                dimension.Value = range.ToString();
            }
            Docx.Save(part, xml);
        }
        strings?.Save();

        // Edited values invalidate the cached results of every dependent formula, and a calcChain
        // that disagrees with the cells' formulas is one of the things Excel "repairs" on open.
        // Deleting it and forcing a full calculation on load is what Excel accepts silently; the
        // cached <v> values stay stale until then (DuckDB read_xlsx reads them, 2026-10-15).
        // DeletePart also drops the workbook relationship and the content-type override.
        var removed = false;
        if (wb.CalculationChainPart is { } chain)
        {
            wb.DeletePart(chain);
            removed = true;
        }
        var book = Docx.Load(wb);
        var root = book.Root ?? throw new InvalidDataException("empty workbook part");
        var calcPr = root.Element(S + "calcPr");
        if (calcPr is null)
        {
            calcPr = new XElement(S + "calcPr");
            var next = root.Elements().FirstOrDefault(e => e.Name.Namespace == S && AfterCalcPr.Contains(e.Name.LocalName));
            if (next is null) root.Add(calcPr);
            else next.AddBeforeSelf(calcPr);
        }
        calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        Docx.Save(wb, book);
        return removed;
    }

    private static void Store(XElement cell, Planned edit, SharedStrings? strings)
    {
        if ((string?)cell.Attribute("t") == "s") strings?.Release();
        cell.Elements().Where(e => e.Name == S + "f" || e.Name == S + "v" || e.Name == S + "is").Remove();
        // s (style) stays. cm/vm point at cell/value metadata (dynamic arrays, rich values) that
        // described the old content.
        cell.Attribute("t")?.Remove();
        cell.Attribute("cm")?.Remove();
        cell.Attribute("vm")?.Remove();
        // Only extLst can remain, and f/v come before it.
        switch (edit.Kind)
        {
            case CellValueKind.Formula:
                cell.AddFirst(new XElement(S + "f", edit.Stored));
                break;
            case CellValueKind.Number:
                cell.AddFirst(new XElement(S + "v", edit.Stored));
                break;
            case CellValueKind.Boolean:
                cell.SetAttributeValue("t", "b");
                cell.AddFirst(new XElement(S + "v", edit.Stored));
                break;
            default:
                cell.SetAttributeValue("t", "s");
                cell.AddFirst(new XElement(S + "v", strings!.Add(edit.Stored).ToString(CultureInfo.InvariantCulture)));
                break;
        }
    }

    private static XElement GetOrCreateCell(XElement worksheet, XElement data, CellAddress at)
    {
        XElement? row = null, nextRow = null;
        uint index = 0;
        foreach (var r in data.Elements(S + "row"))
        {
            index = UInt(r, "r") ?? index + 1;
            if (index == at.Row) { row = r; break; }
            if (index > at.Row) { nextRow = r; break; }
        }
        if (row is null)
        {
            row = new XElement(S + "row", new XAttribute("r", at.Row.ToString(CultureInfo.InvariantCulture)));
            if (nextRow is null) data.Add(row);
            else nextRow.AddBeforeSelf(row);
        }

        XElement? last = null;
        var column = 0;
        foreach (var c in row.Elements(S + "c"))
        {
            column = CellAddress.TryParse((string?)c.Attribute("r"), out var a) ? a.Column : column + 1;
            if (column == at.Column) return c;
            if (column > at.Column)
            {
                var before = NewCell(worksheet, row, at);
                c.AddBeforeSelf(before);
                return before;
            }
            last = c;
        }
        var cell = NewCell(worksheet, row, at);
        // A row may end in extLst, so append after the last cell rather than at the end.
        if (last is null) row.AddFirst(cell);
        else last.AddAfterSelf(cell);
        return cell;
    }

    /// <summary>A new cell styled as Excel styles a typed-into empty cell: row format, else column format.</summary>
    private static XElement NewCell(XElement worksheet, XElement row, CellAddress at)
    {
        var cell = new XElement(S + "c", new XAttribute("r", at.ToString()));
        var style = (string?)row.Attribute("customFormat") is "1" or "true"
            ? UInt(row, "s")
            : worksheet.Elements(S + "cols").Elements(S + "col")
                .Where(c => UInt(c, "min") <= at.Column && at.Column <= UInt(c, "max"))
                .Select(c => UInt(c, "style"))
                .FirstOrDefault();
        if (style is > 0) cell.SetAttributeValue("s", style.Value.ToString(CultureInfo.InvariantCulture));

        // spans="1:2" is an optional load hint; keep a single span covering the new cell.
        if (row.Attribute("spans") is { } spans && spans.Value.Split(':') is [var a, var b]
            && int.TryParse(a, NumberStyles.None, CultureInfo.InvariantCulture, out var lo)
            && int.TryParse(b, NumberStyles.None, CultureInfo.InvariantCulture, out var hi)
            && (at.Column < lo || at.Column > hi))
        {
            spans.Value = string.Create(CultureInfo.InvariantCulture, $"{Math.Min(lo, at.Column)}:{Math.Max(hi, at.Column)}");
        }
        return cell;
    }

    private static uint? UInt(XElement e, string attribute) =>
        uint.TryParse((string?)e.Attribute(attribute), NumberStyles.None, CultureInfo.InvariantCulture, out var v) ? v : null;

    /// <summary>
    /// ST_Xstring escaping (ECMA-376 Part 1, 22.9.2.19): characters XML cannot carry are written as
    /// _xHHHH_, and a literal "_xHHHH_" in the text is protected as _x005F_xHHHH_ so Excel does not decode it.
    /// </summary>
    internal static string EscapeXstring(string text)
    {
        var protectedText = EscapePattern().Replace(text, "_x005F_$1");
        var sb = new StringBuilder(protectedText.Length);
        foreach (var ch in protectedText)
        {
            if (ch < 0x20 && ch is not ('\t' or '\n' or '\r') || ch is '\uFFFE' or '\uFFFF')
                sb.Append(CultureInfo.InvariantCulture, $"_x{(int)ch:X4}_");
            else
                sb.Append(ch);
        }
        return sb.ToString();
    }

    // Only the leading '_' of a literal escape-shaped run is encoded: "_x0041_" → "_x005F_x0041_".
    [GeneratedRegex("_(x[0-9A-Fa-f]{4}_)")]
    private static partial Regex EscapePattern();

    /// <summary>xl/sharedStrings.xml as an append-only table that reuses existing plain entries.</summary>
    private sealed class SharedStrings
    {
        private readonly OpenXmlPart _part;
        private readonly XDocument _xml;
        private readonly XElement _sst;
        private readonly Dictionary<string, int> _plain = new(StringComparer.Ordinal);
        private int _unique;
        private int _referenceDelta;

        public SharedStrings(WorkbookPart wb)
        {
            if (wb.SharedStringTablePart is { } part)
            {
                _part = part;
                _xml = Docx.Load(part);
            }
            else
            {
                _part = wb.AddNewPart<SharedStringTablePart>();
                _xml = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(S + "sst", new XAttribute("count", "0"), new XAttribute("uniqueCount", "0")));
            }
            _sst = _xml.Root ?? throw new InvalidDataException("empty sharedStrings part");
            foreach (var si in _sst.Elements(S + "si"))
            {
                // Only an si that is exactly one <t> is the plain string; rich or phonetic entries
                // carry formatting a plain edit must not inherit.
                if (si.Elements().ToList() is [var t] && t.Name == S + "t") _plain.TryAdd(t.Value, _unique);
                _unique++;
            }
        }

        public int Add(string text)
        {
            _referenceDelta++;
            var stored = EscapeXstring(text);
            if (_plain.TryGetValue(stored, out var existing)) return existing;
            var t = new XElement(S + "t", stored);
            if (stored.Length > 0 && (char.IsWhiteSpace(stored[0]) || char.IsWhiteSpace(stored[^1])))
                t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            var si = new XElement(S + "si", t);
            if (_sst.Element(S + "extLst") is { } ext) ext.AddBeforeSelf(si);
            else _sst.Add(si);
            _plain[stored] = _unique;
            return _unique++;
        }

        public void Release() => _referenceDelta--;

        /// <summary>count (references) and uniqueCount (entries) are optional hints; keep present ones right.</summary>
        public void Save()
        {
            if (long.TryParse((string?)_sst.Attribute("count"), NumberStyles.None, CultureInfo.InvariantCulture, out var count))
                _sst.SetAttributeValue("count", Math.Max(0, count + _referenceDelta).ToString(CultureInfo.InvariantCulture));
            if (_sst.Attribute("uniqueCount") is not null)
                _sst.SetAttributeValue("uniqueCount", _unique.ToString(CultureInfo.InvariantCulture));
            Docx.Save(_part, _xml);
        }
    }
}
