using System.Text;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeProof.Spreadsheet.Tests;

public sealed class CellEditorTests
{
    private const string CalcChainRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

    [Fact]
    public void NumberStringAndFormulaLeaveEveryOtherPartByteIdentical()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output,
        [
            new CellEdit("Data", "A1", "42"),          // existing styled number
            new CellEdit("data", "c3", "delta"),       // new string, inserted between B3 and E3
            new CellEdit("Data", "A4", "=A1+1"),       // formula in a new row between rows 3 and 5
        ]);

        Assert.True(result.Succeeded, string.Join("; ", result.Errors.Select(e => e.Message)));
        Assert.True(result.CalcChainRemoved);
        Assert.Equal(
            [
                new CellChange("Data", "A1", CellValueKind.Number, "42", "10"),
                new CellChange("Data", "C3", CellValueKind.Text, "delta", null),
                new CellChange("Data", "A4", CellValueKind.Formula, "=A1+1", null),
            ],
            result.Changes);

        using (var doc = SpreadsheetDocument.Open(output, isEditable: false))
        {
            var wb = doc.WorkbookPart!;
            var data = Sheet(wb, "Data");
            var a1 = CellAt(data, "A1");
            Assert.Equal("42", a1.CellValue?.Text);
            Assert.Null(a1.DataType);
            Assert.Equal(1u, a1.StyleIndex?.Value);

            var c3 = CellAt(data, "C3");
            Assert.Equal(CellValues.SharedString, c3.DataType?.Value);
            var items = wb.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ToList();
            Assert.Equal("delta", items[int.Parse(c3.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture)].Text?.Text);
            Assert.Equal(4u, wb.SharedStringTablePart.SharedStringTable!.UniqueCount?.Value);
            Assert.Equal(5u, wb.SharedStringTablePart.SharedStringTable!.Count?.Value);

            var a4 = CellAt(data, "A4");
            Assert.Equal("A1+1", a4.CellFormula?.Text);
            Assert.Null(a4.CellValue);

            var sheetData = data.GetFirstChild<SheetData>()!;
            Assert.Equal([1u, 2u, 3u, 4u, 5u], sheetData.Elements<Row>().Select(r => r.RowIndex!.Value));
            Assert.Equal(["A3", "B3", "C3", "E3"], sheetData.Elements<Row>().ElementAt(2).Elements<Cell>().Select(c => c.CellReference!.Value));

            Assert.Null(wb.CalculationChainPart);
            Assert.True(wb.Workbook!.CalculationProperties!.FullCalculationOnLoad!.Value);
            Assert.Equal(191029u, wb.Workbook.CalculationProperties.CalculationId!.Value);

            Assert.Empty(Docx.Validate(doc));
        }

        var before = XlsxFixture.Entries(input);
        var after = XlsxFixture.Entries(output);
        Assert.False(after.ContainsKey("xl/calcChain.xml"));
        Assert.DoesNotContain(CalcChainRel, Encoding.UTF8.GetString(after["xl/_rels/workbook.xml.rels"]), StringComparison.Ordinal);
        Assert.DoesNotContain("calcChain", Encoding.UTF8.GetString(after["[Content_Types].xml"]), StringComparison.Ordinal);
        Assert.Contains("fullCalcOnLoad=\"1\"", Encoding.UTF8.GetString(after["xl/workbook.xml"]), StringComparison.Ordinal);

        string[] touched = ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/workbook.xml", "xl/calcChain.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"];
        var untouched = before.Keys.Except(touched).ToList();
        Assert.Contains("xl/styles.xml", untouched);
        Assert.Contains("xl/worksheets/sheet2.xml", untouched);
        Assert.Contains("xl/tables/table1.xml", untouched);
        foreach (var name in untouched)
        {
            Assert.True(after.TryGetValue(name, out var bytes), $"{name} missing from output");
            Assert.True(before[name].AsSpan().SequenceEqual(bytes), $"{name} changed");
        }
        Assert.Empty(after.Keys.Except(before.Keys));
    }

    [Fact]
    public void ExistingStringReusesSharedStringEntry()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        // C1 held "alpha" (a reference released) and now holds "gamma"; C4 gains "alpha".
        var result = CellEditor.Apply(input, output, [new CellEdit("Data", "C4", "alpha"), new CellEdit("Data", "C1", "gamma")]);

        Assert.True(result.Succeeded);
        Assert.Equal("alpha", result.Changes[1].Previous);
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        var wb = doc.WorkbookPart!;
        var sharedStrings = wb.SharedStringTablePart!.SharedStringTable!;
        Assert.Equal(3, sharedStrings.Elements<SharedStringItem>().Count());
        Assert.Equal((5u, 3u), (sharedStrings.Count!.Value, sharedStrings.UniqueCount!.Value));
        Assert.Equal("0", CellAt(Sheet(wb, "Data"), "C4").CellValue?.Text);
        Assert.Equal("2", CellAt(Sheet(wb, "Data"), "C1").CellValue?.Text);
    }

    [Theory]
    [InlineData("5")]
    [InlineData("TRUE")]
    [InlineData("=1+1")]
    public void NonTextEditOverSharedStringReleasesItsReference(string value)
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        // C1 held "alpha": the 4 references in the fixture drop to 3, the entry itself stays.
        var result = CellEditor.Apply(input, output, [new CellEdit("Data", "C1", value)]);

        Assert.True(result.Succeeded);
        Assert.Equal("alpha", result.Changes[0].Previous);
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        var sharedStrings = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable!;
        Assert.Equal((3u, 3u, 3), (sharedStrings.Count!.Value, sharedStrings.UniqueCount!.Value, sharedStrings.Elements<SharedStringItem>().Count()));
        Assert.NotEqual(CellValues.SharedString, CellAt(Sheet(doc.WorkbookPart, "Data"), "C1").DataType?.Value);
        Assert.Empty(Docx.Validate(doc));
    }

    [Fact]
    public void NumberEditOverNonStringCellLeavesSharedStringsByteIdentical()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        Assert.True(CellEditor.Apply(input, output, [new CellEdit("Data", "A2", "7")]).Succeeded);

        Assert.Equal(XlsxFixture.Entries(input)["xl/sharedStrings.xml"], XlsxFixture.Entries(output)["xl/sharedStrings.xml"]);
    }

    [Theory]
    [InlineData("D1", "=LEN(A1:A2)*2", "array", "D1:D2")]
    [InlineData("D2", "7", "array", "D1:D2")]
    [InlineData("F1", "=E1", "dataTable", "F1:F2")]
    [InlineData("F2", "text", "dataTable", "F1:F2")]
    public void AnyEditInsideArrayOrDataTableRangeIsRefused(string cell, string value, string type, string range)
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Notes", cell, value)]);

        var error = Assert.Single(result.Errors);
        Assert.Equal(("xlsx-array-formula", Severity.Error, $"Notes!{cell}"), (error.Rule, error.Severity, error.Location));
        Assert.Contains($"{type} formula range {range}", error.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void SingleCellArrayFormulaIsReplacedWhole()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Notes", "E1", "3")]);

        Assert.True(result.Succeeded);
        Assert.Equal("=SUM(LEN(A1:A2))", result.Changes[0].Previous);
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        var e1 = CellAt(Sheet(doc.WorkbookPart!, "Notes"), "E1");
        Assert.Equal(("3", null), (e1.CellValue?.Text, e1.CellFormula));
        Assert.Empty(Docx.Validate(doc));
    }

    [Fact]
    public void NewCellsInheritRowStyleAndGrowDimensionAndSpans()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output,
            [new CellEdit("Data", "C5", "TRUE"), new CellEdit("Data", "F7", "-2.50e1"), new CellEdit("Notes", "B1", "'=x", CellValueKind.Text)]);

        Assert.True(result.Succeeded);
        Assert.Equal([CellValueKind.Boolean, CellValueKind.Number, CellValueKind.Text], result.Changes.Select(c => c.Kind));
        Assert.Equal("-25", result.Changes[1].Value);
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        var data = Sheet(doc.WorkbookPart!, "Data");
        var c5 = CellAt(data, "C5");
        Assert.Equal((CellValues.Boolean, "1", 1u), (c5.DataType!.Value, c5.CellValue!.Text, c5.StyleIndex!.Value));
        Assert.Equal("1:3", data.GetFirstChild<SheetData>()!.Elements<Row>().Single(r => r.RowIndex!.Value == 5).Spans!.InnerText);
        Assert.Null(CellAt(data, "F7").StyleIndex);
        Assert.Equal("A1:F7", data.SheetDimension!.Reference!.Value);
        Assert.Equal("A1:F2", Sheet(doc.WorkbookPart!, "Notes").SheetDimension!.Reference!.Value);
        Assert.Empty(Docx.Validate(doc));
    }

    [Fact]
    public void ValueOverSharedFormulaFollowerIsRefusedAndWritesNothing()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Data", "A1", "1"), new CellEdit("Data", "B3", "5")]);

        Assert.False(result.Succeeded);
        var error = Assert.Single(result.Errors);
        Assert.Equal(("xlsx-shared-formula", Severity.Error, "Data!B3"), (error.Rule, error.Severity, error.Location));
        Assert.Contains("Data!B3", error.Message, StringComparison.Ordinal);
        Assert.Empty(result.Changes);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void FormulaOverSharedFormulaFollowerReplacesOnlyThatCell()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Data", "B3", "=A3*3")]);

        Assert.True(result.Succeeded);
        Assert.Equal("=(shared formula si=0)", result.Changes[0].Previous);
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        var data = Sheet(doc.WorkbookPart!, "Data");
        var b3 = CellAt(data, "B3").CellFormula!;
        Assert.Equal(("A3*3", null), (b3.Text, b3.FormulaType));
        Assert.Equal(("A1*2", "B1:B5"), (CellAt(data, "B1").CellFormula!.Text, CellAt(data, "B1").CellFormula!.Reference!.Value));
        Assert.Equal(CellFormulaValues.Shared, CellAt(data, "B5").CellFormula!.FormulaType!.Value);
        Assert.Empty(Docx.Validate(doc));
        Assert.Equal(new FormulaCounts(SharedMasters: 1, SharedFollowers: 2, Plain: 2, Array: 0), WorkbookInspector.Inspect(output).Sheets[0].Formulas);
    }

    [Theory]
    [InlineData("7")]
    [InlineData("=A1*5")]
    public void AnythingOverSharedFormulaMasterWithFollowersIsRefused(string value)
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Data", "B1", value)]);

        var error = Assert.Single(result.Errors);
        Assert.Equal(("xlsx-shared-formula", "Data!B1"), (error.Rule, error.Location));
        Assert.Contains("3 other cell(s)", error.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void StringsAreStoredAsEscapedXstrings()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output, [new CellEdit("Notes", "B1", " a_x0041_b\u0001 ")]);

        Assert.True(result.Succeeded);
        using (var stream = new MemoryStream(XlsxFixture.Entries(output)["xl/sharedStrings.xml"]))
        {
            var t = System.Xml.Linq.XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace).Root!.Elements().Last().Elements().Single();
            Assert.Equal((" a_x005F_x0041_b_x0001_ ", "preserve"), (t.Value, (string?)t.Attribute(System.Xml.Linq.XNamespace.Xml + "space")));
        }
        using var doc = SpreadsheetDocument.Open(output, isEditable: false);
        Assert.Empty(Docx.Validate(doc));
    }

    [Fact]
    public void BadEditsReportEveryProblem()
    {
        using var dir = new TempDir();
        var input = XlsxFixture.Build(dir.File("in.xlsx"));
        var output = dir.File("out.xlsx");

        var result = CellEditor.Apply(input, output,
        [
            new CellEdit("Missing", "A1", "1"),
            new CellEdit("Data", "A0", "1"),
            new CellEdit("Data", "XFE1", "1"),
            new CellEdit("Data", "A2", "1,000", CellValueKind.Number),
            new CellEdit("Data", "A3", "yes", CellValueKind.Boolean),
            new CellEdit("Data", "A3", "1"),
            new CellEdit("Data", "A6", ""),
        ]);

        Assert.Equal(
            ["xlsx-edit-sheet", "xlsx-edit-cell", "xlsx-edit-cell", "xlsx-edit-value", "xlsx-edit-value", "xlsx-edit-cell", "xlsx-edit-value"],
            result.Errors.Select(e => e.Rule));
        Assert.False(File.Exists(output));
        Assert.Throws<ArgumentException>(() => CellEditor.Apply(input, input, [new CellEdit("Data", "A1", "1")]));
    }

    private static Worksheet Sheet(WorkbookPart wb, string name)
    {
        var id = wb.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == name).Id!.Value!;
        return ((WorksheetPart)wb.GetPartById(id)).Worksheet!;
    }

    private static Cell CellAt(Worksheet ws, string reference) =>
        ws.Descendants<Cell>().Single(c => c.CellReference?.Value == reference);
}
