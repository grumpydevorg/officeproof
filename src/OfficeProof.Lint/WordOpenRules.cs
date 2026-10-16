using OfficeProof.Core;

namespace OfficeProof.Lint;

// Microsoft documents that Word reads an omitted w:tblLook as 0x04A0 ([MS-OE376] 2.1.160), but
// Word for Mac 16.113 and Word for Windows 16.0.20430 read such a table as 0x0000 on some opens
// and 0x04A0 on others (2026-10-16: one in three to one in two opens of the same file). The
// table's header-row and first-column styling then changes from one open to the next.
/// <summary>table-look-missing: tables without w:tblLook, which Word styles at random per open.</summary>
public sealed class TableLookMissingRule : ILintRule
{
    public string Id => "table-look-missing";
    public Severity Severity => Severity.Warning;
    public string Description => "tables without w:tblLook, whose header-row and first-column styling Word picks at random on each open";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        foreach (var part in Docx.StoryParts(ctx.Main))
        {
            var count = ctx.Xml(part).Descendants(Ns.W + "tblPr")
                .Count(p => p.Parent?.Name == Ns.W + "tbl" && p.Element(Ns.W + "tblLook") is null);
            if (count > 0)
                yield return new Finding(Id, Severity,
                    $"{count} table(s) without w:tblLook: Word for Mac and Word for Windows style their header row and first column " +
                    "differently from one open to the next; setting Table Style Options in Word (or w:tblLook) fixes it",
                    LintContext.PartName(part));
        }
    }
}

/// <summary>update-fields-on-open: w:updateFields, so Word offers to update every field when the document opens.</summary>
public sealed class UpdateFieldsOnOpenRule : ILintRule
{
    public string Id => "update-fields-on-open";
    public Severity Severity => Severity.Info;
    public string Description => "w:updateFields set: Word asks to update every field (TOC, PAGE, cross-references) when the document opens";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        if (ctx.Main.DocumentSettingsPart is not { } settings) yield break;
        if (ctx.Xml(settings).Root?.Element(Ns.W + "updateFields") is null) yield break;
        yield return new Finding(Id, Severity,
            "Word asks on open whether to update every field; the reader then sees updated results, which verify --word does not compare " +
            "(generate and retemplate set this on purpose for our own documents)",
            LintContext.PartName(settings));
    }
}
