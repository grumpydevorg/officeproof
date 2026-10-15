using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeProof.Spreadsheet.Tests;

public sealed class WorkbookInspectorTests
{
    [Fact]
    public void InspectCountsFormulasMergesFiltersAndConditionalFormats()
    {
        using var dir = new TempDir();
        var report = WorkbookInspector.Inspect(XlsxFixture.Build(dir.File("in.xlsx")));

        Assert.Equal(["Data", "Notes"], report.Sheets.Select(s => s.Name));
        var data = report.Sheets[0];
        Assert.Equal("worksheet", data.Kind);
        Assert.Equal("visible", data.State);
        Assert.Equal("A1:E5", data.Dimension);
        Assert.Equal(12, data.CellCount);
        Assert.Equal(new FormulaCounts(SharedMasters: 1, SharedFollowers: 3, Plain: 1, Array: 0), data.Formulas);
        Assert.Equal(5, data.Formulas.Total);
        Assert.Equal(["D1:E1"], data.MergedRanges);
        var cf = Assert.Single(data.ConditionalFormats);
        Assert.Equal("A1:A5", cf.Range);
        Assert.Equal(["cellIs", "containsText"], cf.RuleTypes);
        Assert.False(cf.Extended);
        Assert.Equal([new DataValidationReport("C3:C5", "list", false)], data.DataValidations);
        Assert.Equal("A1:C5", data.AutoFilter);
        Assert.Empty(data.Tables);

        var notes = report.Sheets[1];
        Assert.Equal(new FormulaCounts(SharedMasters: 0, SharedFollowers: 0, Plain: 0, Array: 3), notes.Formulas);
        var table = Assert.Single(notes.Tables);
        Assert.Equal(("Notes", "A1:A2"), (table.Name, table.Range));
        Assert.Null(notes.AutoFilter);

        Assert.Equal(
            [new DefinedNameReport("_xlnm._FilterDatabase", "Data", true, "Data!$A$1:$C$5"), new DefinedNameReport("Total", null, false, "Data!$C$2")],
            report.DefinedNames);
        Assert.True(report.Package.CalcChain);
        Assert.Equal(5, report.Package.CalcChainEntries);
        Assert.False(report.Package.FullCalcOnLoad);
        Assert.Empty(report.Package.CustomXmlItems);
        Assert.Null(report.Package.SensitivityLabelPart);
        Assert.Equal(3, report.SharedStringCount);
        Assert.Equal(2, report.CellXfsCount);
    }

    [Fact]
    public void FindingsFlagLabelExternalLinkCalcChainAndSharedFormulas()
    {
        using var dir = new TempDir();
        var path = XlsxFixture.Build(dir.File("in.xlsx"));
        using (var doc = SpreadsheetDocument.Open(path, isEditable: true))
        {
            var label = doc.AddLabelInfoPart();
            using (var s = label.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
            {
                w.Write("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"><clbl:label id="{2e6ba7ff-0000-0000-0000-000000000000}" enabled="1" method="Privileged" siteId="{8c3c81bc-0000-0000-0000-000000000000}" removed="0" /></clbl:labelList>""");
            }
            var link = doc.WorkbookPart!.AddNewPart<ExternalWorkbookPart>();
            link.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                new Uri("other.xlsx", UriKind.Relative), "rId1");
            link.ExternalLink = new ExternalLink(new ExternalBook { Id = "rId1" });
        }

        var report = WorkbookInspector.Inspect(path);
        Assert.Equal("/docMetadata/LabelInfo.xml", report.Package.SensitivityLabelPart);
        var label1 = Assert.Single(report.Package.SensitivityLabels);
        Assert.Equal(("Privileged", true), (label1.Method, label1.Enabled));
        Assert.Equal(["other.xlsx"], report.Package.ExternalLinks);

        var findings = report.ToFindings();
        var byRule = findings.ToDictionary(f => f.Rule);
        Assert.Equal(Severity.Info, byRule["xlsx-sensitivity-label"].Severity);
        Assert.Contains("non-Excel rewrites can drop it", byRule["xlsx-sensitivity-label"].Message, StringComparison.Ordinal);
        Assert.Equal(Severity.Warning, byRule["xlsx-external-link"].Severity);
        Assert.Equal(Severity.Info, byRule["xlsx-calcchain"].Severity);
        Assert.Equal("Data", byRule["xlsx-shared-formula"].Location);
        Assert.Equal(0, FindingReport.ExitCode(findings));

        var json = report.ToJson();
        Assert.Contains("\"sensitivityLabels\"", json, StringComparison.Ordinal);
        Assert.Contains("\"sharedFollowers\": 3", json, StringComparison.Ordinal);
    }
}
