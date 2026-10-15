using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeProof.Spreadsheet;

/// <summary>Read-only survey of an .xlsx package.</summary>
public static class WorkbookInspector
{
    private static readonly XNamespace Clbl = "http://schemas.microsoft.com/office/2020/mipLabelMetadata";

    /// <summary>Opens <paramref name="path"/> read-only and reports its sheets, names and package parts.</summary>
    public static WorkbookReport Inspect(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, isEditable: false);
        var wb = doc.WorkbookPart ?? throw new InvalidDataException($"{path} has no workbook part");
        var workbook = wb.Workbook ?? throw new InvalidDataException($"{path} has an empty workbook part");
        var sheetEntries = workbook.Sheets?.Elements<Sheet>().ToList() ?? [];

        var sheets = sheetEntries.Select(s => InspectSheet(wb, s)).ToList();
        var names = workbook.DefinedNames?.Elements<DefinedName>()
            .Select(n => new DefinedNameReport(
                n.Name?.Value ?? "",
                n.LocalSheetId?.Value is { } i && i < sheetEntries.Count ? sheetEntries[(int)i].Name?.Value : null,
                n.Hidden?.Value ?? false,
                n.Text))
            .ToList() ?? [];

        var all = doc.GetAllParts().ToList();
        var calcChain = wb.CalculationChainPart;
        var label = doc.LabelInfoPart;
        var package = new PackageReport(
            CalcChain: calcChain is not null,
            CalcChainEntries: calcChain?.CalculationChain?.Elements<CalculationCell>().Count() ?? 0,
            FullCalcOnLoad: workbook.CalculationProperties?.FullCalculationOnLoad?.Value ?? false,
            CustomXmlItems: [.. all.OfType<CustomXmlPart>().Select(p => p.Uri.ToString()).Order(StringComparer.Ordinal)],
            SensitivityLabelPart: label?.Uri.ToString(),
            SensitivityLabels: label is null ? [] : Labels(label),
            VbaProject: wb.VbaProjectPart is not null,
            ExternalLinks: [.. wb.ExternalWorkbookParts.Select(p => p.ExternalRelationships.FirstOrDefault()?.Uri.OriginalString ?? p.Uri.ToString())],
            PrinterSettings: all.OfType<SpreadsheetPrinterSettingsPart>().Count(),
            PivotCaches: wb.PivotTableCacheDefinitionParts.Count());

        return new WorkbookReport(
            sheets,
            names,
            package,
            wb.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Count() ?? 0,
            wb.WorkbookStylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>().Count() ?? 0);
    }

    private static SheetReport InspectSheet(WorkbookPart wb, Sheet sheet)
    {
        var name = sheet.Name?.Value ?? "";
        var state = sheet.State?.InnerText ?? "visible";
        var part = sheet.Id?.Value is { } id ? wb.GetPartById(id) : null;
        if (part is not WorksheetPart wsPart)
        {
            var kind = part switch
            {
                ChartsheetPart => "chartsheet",
                DialogsheetPart => "dialogsheet",
                _ => "unknown",
            };
            return new SheetReport(name, kind, state, part?.Uri.ToString() ?? "", null, 0,
                new FormulaCounts(0, 0, 0, 0), [], [], [], [], null);
        }

        var ws = wsPart.Worksheet ?? throw new InvalidDataException($"{wsPart.Uri} is empty");
        int cells = 0, masters = 0, followers = 0, plain = 0, array = 0;
        foreach (var cell in ws.Descendants<Cell>())
        {
            cells++;
            switch (cell.CellFormula)
            {
                case null:
                    break;
                case { FormulaType.Value: var t } f when t == CellFormulaValues.Shared:
                    if (f.Reference is null) followers++;
                    else masters++;
                    break;
                case { FormulaType.Value: var t } when t == CellFormulaValues.Array || t == CellFormulaValues.DataTable:
                    array++;
                    break;
                default:
                    plain++;
                    break;
            }
        }

        List<ConditionalFormatReport> cf =
        [
            .. ws.Elements<ConditionalFormatting>().Select(c => new ConditionalFormatReport(
                c.SequenceOfReferences?.InnerText ?? "",
                [.. c.Elements<ConditionalFormattingRule>().Select(r => r.Type?.InnerText ?? "")],
                Extended: false)),
            .. ws.Descendants<X14.ConditionalFormatting>().Select(c => new ConditionalFormatReport(
                c.GetFirstChild<Xm.ReferenceSequence>()?.Text ?? "",
                [.. c.Elements<X14.ConditionalFormattingRule>().Select(r => r.Type?.InnerText ?? "")],
                Extended: true)),
        ];
        List<DataValidationReport> dv =
        [
            .. ws.Descendants<DataValidation>().Select(v => new DataValidationReport(
                v.SequenceOfReferences?.InnerText ?? "", v.Type?.InnerText ?? "none", Extended: false)),
            .. ws.Descendants<X14.DataValidation>().Select(v => new DataValidationReport(
                v.GetFirstChild<Xm.ReferenceSequence>()?.Text ?? "", v.Type?.InnerText ?? "none", Extended: true)),
        ];

        return new SheetReport(
            name, "worksheet", state, wsPart.Uri.ToString(),
            ws.SheetDimension?.Reference?.Value,
            cells,
            new FormulaCounts(masters, followers, plain, array),
            [.. ws.Descendants<MergeCell>().Select(m => m.Reference?.Value ?? "")],
            cf,
            dv,
            [.. wsPart.TableDefinitionParts.Select(t => new TableReport(
                t.Table?.DisplayName?.Value ?? t.Table?.Name?.Value ?? "",
                t.Table?.Reference?.Value ?? "",
                t.Uri.ToString()))],
            ws.GetFirstChild<AutoFilter>()?.Reference?.Value);
    }

    private static List<SensitivityLabel> Labels(OpenXmlPart part) =>
        [.. Docx.Load(part).Descendants(Clbl + "label").Select(l => new SensitivityLabel(
            (string?)l.Attribute("id") ?? "",
            (string?)l.Attribute("method"),
            (string?)l.Attribute("enabled") is "1" or "true",
            (string?)l.Attribute("siteId")))];
}
