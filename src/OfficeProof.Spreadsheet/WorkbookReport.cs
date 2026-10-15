using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeProof.Core;

namespace OfficeProof.Spreadsheet;

/// <summary>
/// What an .xlsx contains that an edit could disturb. Produced by <see cref="WorkbookInspector.Inspect"/>;
/// serialise with <see cref="ToJson"/>, turn risks into findings with <see cref="ToFindings"/>.
/// </summary>
/// <param name="Sheets">Sheets in workbook order.</param>
/// <param name="DefinedNames">Workbook and sheet-scoped names.</param>
/// <param name="Package">Package-level parts that non-Excel writers tend to drop or leave stale.</param>
/// <param name="SharedStringCount">Unique entries (<c>si</c>) in xl/sharedStrings.xml.</param>
/// <param name="CellXfsCount">Cell formats (<c>cellXfs/xf</c>) in xl/styles.xml.</param>
public sealed record WorkbookReport(
    IReadOnlyList<SheetReport> Sheets,
    IReadOnlyList<DefinedNameReport> DefinedNames,
    PackageReport Package,
    int SharedStringCount,
    int CellXfsCount)
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>Indented camelCase JSON.</summary>
    public string ToJson() => JsonSerializer.Serialize(this, Json);

    /// <summary>
    /// The risks in this workbook as findings: nothing here is an error in the file itself, so
    /// severities are Info (preserved by officeproof, easy for other tools to lose) or Warning
    /// (values a reader sees may not be current).
    /// </summary>
    public IReadOnlyList<Finding> ToFindings()
    {
        var findings = new List<Finding>();
        var p = Package;
        // 2026-10-15, ecological-risk-assessment bats xlsx: a SharePoint-labelled workbook carries the
        // label only in docMetadata/LabelInfo.xml. openpyxl/pandas/LibreOffice rewrite the package
        // from their own model and do not write that part back.
        foreach (var label in p.SensitivityLabels)
        {
            findings.Add(new Finding("xlsx-sensitivity-label", Severity.Info,
                $"sensitivity label {label.Id} (method {label.Method ?? "unknown"}{(label.Enabled ? "" : ", disabled")}); " +
                "officeproof keeps it, but non-Excel rewrites can drop it",
                p.SensitivityLabelPart));
        }
        if (p.CustomXmlItems.Count > 0)
        {
            findings.Add(new Finding("xlsx-customxml", Severity.Info,
                $"{p.CustomXmlItems.Count} customXml item(s) (SharePoint/document-library metadata); non-Excel rewrites can drop them",
                string.Join(" ", p.CustomXmlItems)));
        }
        foreach (var link in p.ExternalLinks)
        {
            findings.Add(new Finding("xlsx-external-link", Severity.Warning,
                $"formulas read another workbook ({link}); their cached values are only as current as Excel's last refresh",
                Location: link));
        }
        if (p.CalcChain)
        {
            findings.Add(new Finding("xlsx-calcchain", Severity.Info,
                $"calcChain lists {p.CalcChainEntries} formula cell(s); officeproof cell edits delete it and set fullCalcOnLoad so Excel rebuilds it",
                "/xl/calcChain.xml"));
        }
        var formulas = Sheets.Sum(s => s.Formulas.Total);
        if (p.FullCalcOnLoad && formulas > 0)
        {
            // Cached <v> values are what DuckDB read_xlsx, pandas and openpyxl(data_only) return; after an
            // XML-level edit they stay stale until Excel opens and recalculates the workbook.
            findings.Add(new Finding("xlsx-stale-cache", Severity.Warning,
                $"fullCalcOnLoad is set: the cached values of {formulas} formula cell(s) may be stale until Excel recalculates; " +
                "readers of cached values (DuckDB read_xlsx, pandas) see the old results",
                "/xl/workbook.xml"));
        }
        if (p.VbaProject)
        {
            findings.Add(new Finding("xlsx-vba-project", Severity.Warning,
                "workbook carries a VBA project; macros run only from .xlsm and are not inspected",
                "/xl/vbaProject.bin"));
        }
        if (p.PivotCaches > 0)
        {
            findings.Add(new Finding("xlsx-pivot-cache", Severity.Info,
                $"{p.PivotCaches} pivot cache(s) hold copies of their source data; cell edits do not refresh them"));
        }
        foreach (var sheet in Sheets.Where(s => s.Formulas.SharedFollowers > 0))
        {
            findings.Add(new Finding("xlsx-shared-formula", Severity.Info,
                string.Create(CultureInfo.InvariantCulture,
                    $"{sheet.Formulas.SharedMasters} shared formula(s) cover {sheet.Formulas.SharedFollowers} more cell(s); cell edits refuse to overwrite them with values"),
                sheet.Part, sheet.Name));
        }
        return findings;
    }
}

/// <param name="Name">Sheet tab name.</param>
/// <param name="Kind">"worksheet", "chartsheet", "dialogsheet" or "unknown"; only worksheets have cells.</param>
/// <param name="State">"visible", "hidden" or "veryHidden".</param>
/// <param name="Part">Package part URI, e.g. "/xl/worksheets/sheet2.xml".</param>
/// <param name="Dimension">The sheet's stored <c>dimension ref</c>, if any.</param>
/// <param name="CellCount"><c>c</c> elements, including styled empty cells.</param>
/// <param name="Formulas">Formula cells by storage form.</param>
/// <param name="MergedRanges">mergeCell refs.</param>
/// <param name="ConditionalFormats">Conditional-format ranges with their rule types.</param>
/// <param name="DataValidations">Data-validation ranges with their types.</param>
/// <param name="Tables">Table (ListObject) parts on this sheet.</param>
/// <param name="AutoFilter">The sheet-level autoFilter range, if any (table filters are under <paramref name="Tables"/>).</param>
public sealed record SheetReport(
    string Name,
    string Kind,
    string State,
    string Part,
    string? Dimension,
    int CellCount,
    FormulaCounts Formulas,
    IReadOnlyList<string> MergedRanges,
    IReadOnlyList<ConditionalFormatReport> ConditionalFormats,
    IReadOnlyList<DataValidationReport> DataValidations,
    IReadOnlyList<TableReport> Tables,
    string? AutoFilter);

/// <summary>
/// Formula cells by storage form. A shared formula is written once on its master
/// (<c>f t="shared" ref=".." si="n"</c> with text) and referenced by followers (<c>f t="shared" si="n"</c>, no text).
/// </summary>
/// <param name="SharedMasters">Cells holding a shared formula's text and range.</param>
/// <param name="SharedFollowers">Cells reusing a master's formula by si.</param>
/// <param name="Plain">Ordinary single-cell formulas.</param>
/// <param name="Array">Array and data-table formulas.</param>
public sealed record FormulaCounts(int SharedMasters, int SharedFollowers, int Plain, int Array)
{
    public int Total => SharedMasters + SharedFollowers + Plain + Array;
}

/// <param name="Range">The sqref, space-separated when it has several areas.</param>
/// <param name="RuleTypes">cfRule types in order, e.g. "containsText", "cellIs".</param>
/// <param name="Extended">True for x14 rules stored in the worksheet extLst (Excel 2010+ only).</param>
public sealed record ConditionalFormatReport(string Range, IReadOnlyList<string> RuleTypes, bool Extended);

/// <param name="Range">The sqref, space-separated when it has several areas.</param>
/// <param name="Type">"list", "whole", "decimal", … ("none" when absent).</param>
/// <param name="Extended">True for x14 validations stored in the worksheet extLst.</param>
public sealed record DataValidationReport(string Range, string Type, bool Extended);

public sealed record TableReport(string Name, string Range, string Part);

/// <param name="Name">The defined name, e.g. "_xlnm._FilterDatabase".</param>
/// <param name="Scope">Sheet name for a sheet-scoped name; null for workbook scope.</param>
/// <param name="Hidden">Hidden from Excel's Name Manager (Excel's own filter/print names usually are).</param>
/// <param name="Formula">The definition, e.g. "BatRisk!$A$1:$U$262".</param>
public sealed record DefinedNameReport(string Name, string? Scope, bool Hidden, string Formula);

/// <param name="CalcChain">xl/calcChain.xml exists.</param>
/// <param name="CalcChainEntries">Cells it lists (0 when absent).</param>
/// <param name="FullCalcOnLoad">workbook <c>calcPr fullCalcOnLoad</c> is set.</param>
/// <param name="CustomXmlItems">customXml item part URIs.</param>
/// <param name="SensitivityLabelPart">URI of the MIP label part (docMetadata/LabelInfo.xml), if present.</param>
/// <param name="SensitivityLabels">The labels it lists.</param>
/// <param name="VbaProject">xl/vbaProject.bin exists.</param>
/// <param name="ExternalLinks">Targets of external workbook links.</param>
/// <param name="PrinterSettings">Binary printer-settings parts.</param>
/// <param name="PivotCaches">Pivot cache definitions.</param>
public sealed record PackageReport(
    bool CalcChain,
    int CalcChainEntries,
    bool FullCalcOnLoad,
    IReadOnlyList<string> CustomXmlItems,
    string? SensitivityLabelPart,
    IReadOnlyList<SensitivityLabel> SensitivityLabels,
    bool VbaProject,
    IReadOnlyList<string> ExternalLinks,
    int PrinterSettings,
    int PivotCaches);

/// <summary>One <c>clbl:label</c> from docMetadata/LabelInfo.xml.</summary>
public sealed record SensitivityLabel(string Id, string? Method, bool Enabled, string? SiteId);
