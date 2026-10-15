using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Lint;

/// <summary>unresolved-revision: tracked changes or comments still in the document, counted by kind per part.</summary>
public sealed class UnresolvedRevisionRule : ILintRule
{
    public string Id => "unresolved-revision";
    public Severity Severity => Severity.Warning;
    public string Description => "tracked changes or comments not yet accepted, rejected or deleted";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        foreach (var part in Docx.StoryParts(ctx.Main))
        {
            var root = ctx.Xml(part).Root;
            if (root is null) continue;
            var kinds = Wml.Revisions(root).GroupBy(e => e.Name.LocalName).OrderByDescending(g => g.Count()).ToList();
            if (kinds.Count > 0)
                yield return new Finding(Id, Severity,
                    $"{kinds.Sum(g => g.Count())} tracked change(s): {string.Join(", ", kinds.Select(g => $"{g.Count()} {g.Key}"))}",
                    LintContext.PartName(part));
            if (part is WordprocessingCommentsPart)
            {
                var comments = root.Elements(Ns.W + "comment").Count();
                if (comments > 0)
                    yield return new Finding(Id, Severity, $"{comments} comment(s)", LintContext.PartName(part));
            }
        }
    }
}

// Word does NOT undo tracked changes inside style or numbering definitions on Reject All.
// A 2026-10-15 redline linked numbering into the report's heading styles as a w:pPrChange in
// styles.xml; Reject All in Word for Mac left "2.0 1.0 Summary" and "1.0 Abbreviations", while
// a reject-all text diff (and Clippit's RejectRevisions) looked clean. Change formatting in a
// definition only by adding new styles and applying them as tracked paragraph changes.
/// <summary>revision-in-styles: tracked-change markup in styles.xml or numbering.xml.</summary>
public sealed class RevisionInStylesRule : ILintRule
{
    private static readonly XNamespace W = Ns.W;

    public string Id => "revision-in-styles";
    public Severity Severity => Severity.Error;
    public string Description => "tracked changes in style or numbering definitions, which Word's Reject All does not undo";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        foreach (var (part, xml) in new (OpenXmlPart?, XDocument?)[]
                 {
                     (ctx.Main.StyleDefinitionsPart, ctx.StylesXml),
                     (ctx.Main.NumberingDefinitionsPart, ctx.NumberingXml),
                 })
        {
            if (part is null || xml?.Root is not { } root) continue;
            foreach (var g in Wml.Revisions(root).GroupBy(Owner))
                yield return new Finding(Id, Severity,
                    $"{g.Count()} tracked change(s) ({string.Join(", ", g.Select(e => e.Name.LocalName).Distinct())}) in {g.Key}; "
                    + "Word's Reject All does not undo changes inside definitions",
                    LintContext.PartName(part), g.Key);
        }
    }

    private static string Owner(XElement revision)
    {
        foreach (var a in revision.Ancestors())
        {
            if (a.Name == W + "style") return $"style '{(string?)a.Attribute(W + "styleId")}'";
            if (a.Name == W + "abstractNum") return $"abstractNum {(string?)a.Attribute(W + "abstractNumId")}";
            if (a.Name == W + "num") return $"num {(string?)a.Attribute(W + "numId")}";
            if (a.Name == W + "docDefaults") return "docDefaults";
        }
        return revision.Parent?.Name.LocalName ?? "root";
    }
}
