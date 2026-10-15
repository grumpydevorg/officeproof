using System.Globalization;
using OfficeProof.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeProof.Spreadsheet;

/// <summary>
/// Writes cell values into a copy of a workbook, touching only the parts an edit needs: the edited
/// worksheets, xl/sharedStrings.xml (only when a string is written or a shared-string cell overwritten),
/// xl/workbook.xml (calcPr) and the deleted xl/calcChain.xml with its relationship and content-type
/// override. Every other part keeps its bytes.
/// </summary>
public static class CellEditor
{
    /// <summary>
    /// Copies <paramref name="input"/> to <paramref name="output"/> and applies <paramref name="edits"/>.
    /// All edits are checked against the input first; any error leaves <paramref name="output"/> unwritten.
    /// </summary>
    public static EditResult Apply(string input, string output, IReadOnlyList<CellEdit> edits)
    {
        ArgumentNullException.ThrowIfNull(edits);
        if (Path.GetFullPath(input) == Path.GetFullPath(output))
            throw new ArgumentException("output must differ from input; officeproof never edits in place", nameof(output));

        var (plan, errors) = Plan(input, edits);
        if (errors.Count > 0) return new EditResult([], errors, CalcChainRemoved: false);

        File.Copy(input, output, overwrite: true);
        var removed = false;
        if (plan.Count > 0)
        {
            using var doc = SpreadsheetDocument.Open(output, isEditable: true);
            removed = WorkbookWriter.Write(doc, plan);
        }
        return new EditResult([.. plan.Select(p => p.Change)], [], removed);
    }

    // ---- Read-only pass over the input: resolve every edit or report why it cannot be made.

    private static (List<Planned> Plan, List<Finding> Errors) Plan(string input, IReadOnlyList<CellEdit> edits)
    {
        var plan = new List<Planned>();
        var errors = new List<Finding>();
        using var doc = SpreadsheetDocument.Open(input, isEditable: false);
        var wb = doc.WorkbookPart ?? throw new InvalidDataException($"{input} has no workbook part");
        var sheets = wb.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var indexes = new Dictionary<string, SheetIndex>(StringComparer.Ordinal);
        var seen = new HashSet<(string, CellAddress)>();
        List<string>? strings = null;

        foreach (var edit in edits)
        {
            var where = $"{edit.Sheet}!{edit.Cell}";
            var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, edit.Sheet, StringComparison.OrdinalIgnoreCase));
            if (sheet?.Id?.Value is not { } relId)
            {
                errors.Add(Error("xlsx-edit-sheet", $"no sheet named '{edit.Sheet}'", null, where));
                continue;
            }
            if (wb.GetPartById(relId) is not WorksheetPart wsPart)
            {
                errors.Add(Error("xlsx-edit-sheet", $"'{edit.Sheet}' is not a worksheet", null, where));
                continue;
            }
            var part = wsPart.Uri.ToString();
            if (!CellAddress.TryParse(edit.Cell, out var at))
            {
                errors.Add(Error("xlsx-edit-cell", $"'{edit.Cell}' is not an A1 cell reference within A1:XFD1048576", part, where));
                continue;
            }
            where = $"{sheet.Name?.Value}!{at}";
            if (!seen.Add((relId, at)))
            {
                errors.Add(Error("xlsx-edit-cell", $"{where} is edited more than once", part, where));
                continue;
            }
            if (Resolve(edit) is not { } resolved)
            {
                errors.Add(Error("xlsx-edit-value", $"'{edit.Value}' is not a valid {edit.Kind} value", part, where));
                continue;
            }
            var (kind, stored, shown) = resolved;

            if (!indexes.TryGetValue(relId, out var index))
                indexes[relId] = index = SheetIndex.Build(wsPart.Worksheet?.GetFirstChild<SheetData>());
            index.Cells.TryGetValue(at, out var cell);

            if (cell?.CellFormula is { } f && f.FormulaType?.Value == CellFormulaValues.Shared)
            {
                var si = f.SharedIndex?.Value;
                if (f.Reference is not null)
                {
                    // The master carries the only copy of the formula text; followers (same si, no text)
                    // would be left pointing at nothing. Excel itself re-homes the formula on the next
                    // follower with shifted references, which needs a formula rewriter officeproof lacks.
                    var followers = index.Cells.Values.Count(c => !ReferenceEquals(c, cell) && IsFollower(c, si));
                    if (followers > 0)
                    {
                        errors.Add(Error("xlsx-shared-formula",
                            $"{where} is the master of shared formula si={si}, reused by {followers} other cell(s); overwriting it would break them",
                            part, where));
                        continue;
                    }
                }
                else if (kind != CellValueKind.Formula)
                {
                    errors.Add(Error("xlsx-shared-formula",
                        $"{where} holds shared formula si={si}; refusing to replace a formula with a value (edit it with a formula)",
                        part, where));
                    continue;
                }
            }

            // 2026-10-15 review finding 10: Excel refuses "change part of an array" from the UI; written
            // into the XML, a value inside an array or data-table range contradicts the master's ref.
            // Even a formula on the master would shrink the range under the other cells, so every
            // cell of a multi-cell range is refused.
            if (index.Arrays.FirstOrDefault(a => a.Range.Contains(at)) is { Type: not null } array)
            {
                errors.Add(Error("xlsx-array-formula",
                    $"{where} is inside the {array.Type} formula range {array.Range} (master {array.Master}); refusing to change part of it",
                    part, where));
                continue;
            }

            if (cell?.DataType?.Value == CellValues.SharedString)
                strings ??= SharedStringTexts(wb);
            var previous = Display(cell, strings);
            plan.Add(new Planned(relId, at, kind, stored, cell?.DataType?.Value == CellValues.SharedString,
                new CellChange(sheet.Name?.Value ?? edit.Sheet, at.ToString(), kind, shown, previous)));
        }
        return (plan, errors);
    }

    private static bool IsFollower(Cell c, uint? si) =>
        c.CellFormula is { Reference: null } f && f.FormulaType?.Value == CellFormulaValues.Shared && f.SharedIndex?.Value == si;

    /// <summary>The kind, the text stored in the XML, and the value as reported back.</summary>
    private static (CellValueKind Kind, string Stored, string Shown)? Resolve(CellEdit edit)
    {
        var v = edit.Value ?? "";
        switch (edit.Kind)
        {
            case CellValueKind.Text:
                return (CellValueKind.Text, v, v);
            case CellValueKind.Formula:
                return Formula(v.StartsWith('=') ? v[1..] : v);
            case CellValueKind.Number:
                return Number(v);
            case CellValueKind.Boolean:
                return Boolean(v, allowDigits: true);
            default:
                if (v.Length == 0) return null;
                if (v.StartsWith('=')) return Formula(v[1..]);
                return Number(v) ?? Boolean(v, allowDigits: false) ?? (CellValueKind.Text, v, v);
        }

        static (CellValueKind, string, string)? Formula(string f) =>
            string.IsNullOrWhiteSpace(f) ? null : (CellValueKind.Formula, f, "=" + f);

        static (CellValueKind, string, string)? Number(string s)
        {
            const NumberStyles style = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent;
            if (!double.TryParse(s, style, CultureInfo.InvariantCulture, out var d) || !double.IsFinite(d)) return null;
            var text = d.ToString(CultureInfo.InvariantCulture); // shortest round-trip form
            return (CellValueKind.Number, text, text);
        }

        static (CellValueKind, string, string)? Boolean(string s, bool allowDigits)
        {
            if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || (allowDigits && s == "1"))
                return (CellValueKind.Boolean, "1", "TRUE");
            if (s.Equals("FALSE", StringComparison.OrdinalIgnoreCase) || (allowDigits && s == "0"))
                return (CellValueKind.Boolean, "0", "FALSE");
            return null;
        }
    }

    private static string? Display(Cell? cell, List<string>? strings)
    {
        if (cell is null) return null;
        if (cell.CellFormula is { } f)
            return string.IsNullOrEmpty(f.Text) ? $"=(shared formula si={f.SharedIndex?.Value})" : "=" + f.Text;
        var v = cell.CellValue?.Text;
        var t = cell.DataType?.Value;
        if (t == CellValues.SharedString)
            return int.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out var i) && strings is not null && i < strings.Count
                ? strings[i] : null;
        if (t == CellValues.Boolean) return v == "1" ? "TRUE" : "FALSE";
        if (t == CellValues.InlineString) return cell.InlineString?.InnerText;
        return string.IsNullOrEmpty(v) ? null : v;
    }

    private static List<string> SharedStringTexts(WorkbookPart wb) =>
        [.. wb.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(ItemText) ?? []];

    /// <summary>Plain text of an si: its t, or its rich-text runs joined (phonetic runs excluded).</summary>
    private static string ItemText(SharedStringItem si) =>
        si.Text?.Text ?? string.Concat(si.Elements<Run>().Select(r => r.Text?.Text));

    /// <summary>Cells with their addresses; rows and cells may omit r, which then means "previous + 1".</summary>
    private static IEnumerable<(CellAddress At, Cell Cell)> Cells(SheetData? data)
    {
        if (data is null) yield break;
        uint row = 0;
        foreach (var r in data.Elements<Row>())
        {
            row = r.RowIndex?.Value ?? row + 1;
            var column = 0;
            foreach (var c in r.Elements<Cell>())
            {
                column = CellAddress.TryParse(c.CellReference?.Value, out var a) ? a.Column : column + 1;
                yield return (new CellAddress(column, (int)row), c);
            }
        }
    }

    /// <summary>A worksheet's cells by address, and the multi-cell ranges owned by array and data-table formulas.</summary>
    private sealed record SheetIndex(Dictionary<CellAddress, Cell> Cells, List<(CellRange Range, CellAddress Master, string Type)> Arrays)
    {
        public static SheetIndex Build(SheetData? data)
        {
            var cells = new Dictionary<CellAddress, Cell>();
            var arrays = new List<(CellRange, CellAddress, string)>();
            foreach (var (at, cell) in CellEditor.Cells(data))
            {
                cells[at] = cell;
                // Only the master holds <f t="array|dataTable" ref="…">; the other cells of the range
                // carry cached values and no f. A single-cell array is replaced whole by any edit.
                if (cell.CellFormula is { Reference.Value: { } r } f
                    && (f.FormulaType?.Value == CellFormulaValues.Array || f.FormulaType?.Value == CellFormulaValues.DataTable)
                    && CellRange.TryParse(r, out var range) && range.First != range.Last)
                {
                    arrays.Add((range, at, f.FormulaType.Value == CellFormulaValues.Array ? "array" : "dataTable"));
                }
            }
            return new SheetIndex(cells, arrays);
        }
    }

    private static Finding Error(string rule, string message, string? part, string location) =>
        new(rule, Severity.Error, message, part, location);
}

/// <summary>
/// An edit checked against the input: where, the resolved kind, the text stored in the XML, and whether
/// the cell it overwrites held a shared string (whose reference count the write pass must release).
/// </summary>
internal sealed record Planned(string RelId, CellAddress At, CellValueKind Kind, string Stored, bool ReplacesSharedString, CellChange Change);
