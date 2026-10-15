using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeProof.Spreadsheet.Tests;

/// <summary>
/// Builds a small workbook shaped like the real ones officeproof edits (2026-10-15 bats risk table):
/// shared formulas with a calcChain, conditional formatting, autoFilter, merged cells, a table,
/// defined names and styles.
/// </summary>
internal static class XlsxFixture
{
    /// <summary>Shared strings in the fixture, by index.</summary>
    public static readonly string[] Strings = ["alpha", "beta", "gamma"];

    /// <summary>
    /// Sheet "Data" (sheet1):
    /// <code>
    ///      A        B                          C                 D..E      
    ///   1  10 s=1   =A1*2 shared master B1:B5  "alpha"           merged "beta"
    ///   2  11       follower                   =SUM(A1:A3)
    ///   3  12       follower                                     E3 = 99
    ///   5  14       follower   (row 5 customFormat s=1, spans 1:2)
    /// </code>
    /// Sheet "Notes" (sheet2): A1 "beta", A2 "gamma", table Notes A1:A2; array formula D1:D2 (master D1),
    /// single-cell array E1, data table F1:F2 (master F1).
    /// </summary>
    public static string Build(string path)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();

        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 },
            new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 2, ApplyNumberFormat = true }) { Count = 2 },
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });

        var sharedStrings = wb.AddNewPart<SharedStringTablePart>();
        sharedStrings.SharedStringTable = new SharedStringTable(Strings.Select(s => new SharedStringItem(new Text(s)))) { Count = 4, UniqueCount = 3 };

        var data = wb.AddNewPart<WorksheetPart>();
        data.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:E5" },
            new SheetData(
                new Row(
                    Num("A1", "10", style: 1),
                    new Cell(new CellFormula("A1*2") { FormulaType = CellFormulaValues.Shared, Reference = "B1:B5", SharedIndex = 0 }, new CellValue("20")) { CellReference = "B1" },
                    Str("C1", 0),
                    Str("D1", 1)) { RowIndex = 1 },
                new Row(
                    Num("A2", "11"),
                    Follower("B2", "22"),
                    new Cell(new CellFormula("SUM(A1:A3)"), new CellValue("33")) { CellReference = "C2" }) { RowIndex = 2 },
                new Row(Num("A3", "12"), Follower("B3", "24"), Num("E3", "99")) { RowIndex = 3 },
                new Row(Num("A5", "14"), Follower("B5", "28"))
                {
                    RowIndex = 5,
                    StyleIndex = 1,
                    CustomFormat = true,
                    Spans = new ListValue<StringValue>([new StringValue("1:2")]),
                }),
            new AutoFilter { Reference = "A1:C5" },
            new MergeCells(new MergeCell { Reference = "D1:E1" }) { Count = 1 },
            new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("11")) { Type = ConditionalFormatValues.CellIs, Operator = ConditionalFormattingOperatorValues.GreaterThan, Priority = 1 },
                new ConditionalFormattingRule(new Formula("NOT(ISERROR(SEARCH(\"al\",A1)))")) { Type = ConditionalFormatValues.ContainsText, Operator = ConditionalFormattingOperatorValues.ContainsText, Text = "al", Priority = 2 })
            { SequenceOfReferences = new ListValue<StringValue>([new StringValue("A1:A5")]) },
            new DataValidations(
                new DataValidation(new Formula1("\"x,y\"")) { Type = DataValidationValues.List, AllowBlank = true, SequenceOfReferences = new ListValue<StringValue>([new StringValue("C3:C5")]) })
            { Count = 1 });

        var notes = wb.AddNewPart<WorksheetPart>();
        var table = notes.AddNewPart<TableDefinitionPart>("rId1");
        table.Table = new Table(
            new AutoFilter { Reference = "A1:A2" },
            new TableColumns(new TableColumn { Id = 1, Name = "beta" }) { Count = 1 })
        { Id = 1, Name = "Notes", DisplayName = "Notes", Reference = "A1:A2" };
        notes.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1:F2" },
            new SheetData(
                new Row(
                    Str("A1", 1),
                    new Cell(new CellFormula("LEN(A1:A2)") { FormulaType = CellFormulaValues.Array, Reference = "D1:D2" }, new CellValue("4")) { CellReference = "D1" },
                    new Cell(new CellFormula("SUM(LEN(A1:A2))") { FormulaType = CellFormulaValues.Array, Reference = "E1" }, new CellValue("9")) { CellReference = "E1" },
                    new Cell(new CellFormula { FormulaType = CellFormulaValues.DataTable, Reference = "F1:F2", DataTable2D = false, DataTableRow = false, R1 = "E1" }, new CellValue("9")) { CellReference = "F1" })
                { RowIndex = 1 },
                new Row(Str("A2", 2), Num("D2", "5"), Num("F2", "9")) { RowIndex = 2 }),
            new TableParts(new TablePart { Id = "rId1" }) { Count = 1 });

        var chain = wb.AddNewPart<CalculationChainPart>();
        chain.CalculationChain = new CalculationChain(
            new CalculationCell { CellReference = "B1", SheetId = 1 },
            new CalculationCell { CellReference = "B2" },
            new CalculationCell { CellReference = "B3" },
            new CalculationCell { CellReference = "B5" },
            new CalculationCell { CellReference = "C2" });

        wb.Workbook = new Workbook(
            new Sheets(
                new Sheet { Name = "Data", SheetId = 1, Id = wb.GetIdOfPart(data) },
                new Sheet { Name = "Notes", SheetId = 2, Id = wb.GetIdOfPart(notes) }),
            new DefinedNames(
                new DefinedName("Data!$A$1:$C$5") { Name = "_xlnm._FilterDatabase", LocalSheetId = 0, Hidden = true },
                new DefinedName("Data!$C$2") { Name = "Total" }),
            new CalculationProperties { CalculationId = 191029 });
        return path;
    }

    private static Cell Num(string at, string value, uint? style = null)
    {
        var c = new Cell(new CellValue(value)) { CellReference = at };
        if (style is { } s) c.StyleIndex = s;
        return c;
    }

    private static Cell Str(string at, int index) =>
        new(new CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture))) { CellReference = at, DataType = CellValues.SharedString };

    private static Cell Follower(string at, string cached) =>
        new(new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 }, new CellValue(cached)) { CellReference = at };

    /// <summary>Every zip entry's uncompressed bytes, by entry name.</summary>
    public static Dictionary<string, byte[]> Entries(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var s = e.Open();
            using var m = new MemoryStream();
            s.CopyTo(m);
            return m.ToArray();
        });
    }
}

/// <summary>A per-test scratch directory under the system temp path.</summary>
internal sealed class TempDir : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeproof-xlsx-" + Guid.NewGuid().ToString("N"));

    public TempDir() => Directory.CreateDirectory(Path);

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose() => Directory.Delete(Path, recursive: true);
}
