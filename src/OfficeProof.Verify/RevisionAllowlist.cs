using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// Fails closed on tracked changes the Reject-All simulation has not been shown to handle: every
/// revision element in a story part must sit at a site on <see cref="Sites"/>, and every
/// WordprocessingML ancestor between it and the part root must be an allowed context. Anything
/// else is rule <c>unverifiable-revision</c> (Error): the verifier cannot say what Reject All
/// does with it, so it does not pass it.
/// </summary>
internal static class RevisionAllowlist
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>Every WordprocessingML revision element (ECMA-376 Part 1, 17.13.5 and 17.13.4.4).</summary>
    private static readonly HashSet<string> Kinds = new(StringComparer.Ordinal)
    {
        "ins", "del", "moveFrom", "moveTo",
        "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
        "pPrChange", "rPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
        "tblGridChange", "tblPrExChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
        "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
    };

    /// <summary>
    /// The allowlist: every site <see cref="RevisionEvidence"/> admits. "kind in container" admits
    /// a revision element under that parent (property containers are written
    /// grandparent/parent); "inside X" admits X as an ancestor.
    /// </summary>
    internal static readonly HashSet<string> Sites =
        new(RevisionEvidence.Sites.Keys.Where(s => RevisionEvidence.Refusal(s) is null), StringComparer.Ordinal);

    /// <summary>One finding per revision site not on the allowlist, per story part of <paramref name="package"/>.</summary>
    public static IEnumerable<Finding> Check(byte[] package, string label)
    {
        foreach (var (part, revisions) in Revisions(package))
        {
            var unverified = revisions
                .Select(e => (Element: e, Site: Site(e), Context: UnknownContext(e)))
                .Where(x => !Sites.Contains(x.Site) || x.Context is not null)
                .GroupBy(x => x.Context is null ? x.Site : $"{x.Site} inside {x.Context}");
            foreach (var group in unverified)
            {
                var (first, site, context) = group.First();
                var paragraph = first.AncestorsAndSelf(W + "p").FirstOrDefault();
                var where = paragraph is null ? null : $"first in paragraph \"{ParagraphText.Excerpt(Wml.VisibleText(paragraph))}\"";
                var reason = RevisionEvidence.Refusal(Sites.Contains(site) ? "inside " + context : site);
                yield return new Finding("unverifiable-revision", Severity.Error,
                    $"{label}: {group.Count()} tracked change(s) w:{group.Key} in {part}: {reason}, " +
                    "so the redline cannot be shown to reject back to the original",
                    part, where);
            }
        }
    }

    /// <summary>Every allowlist entry <paramref name="package"/> uses: its sites and their contexts.</summary>
    internal static IEnumerable<string> SitesUsed(byte[] package) =>
        Revisions(package).SelectMany(p => p.Revisions).SelectMany(e =>
            Owner(e)!.Ancestors().Where(a => a.Name.Namespace == W).Select(a => "inside " + Name(a)).Prepend(Site(e)));

    private static List<(string Part, List<XElement> Revisions)> Revisions(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        if (doc.MainDocumentPart is not { } main) return [];
        return [.. Docx.StoryParts(main)
            .Select(part => (part.Uri.OriginalString, Docx.Load(part).Root))
            .Where(x => x.Root is not null)
            .Select(x => (x.OriginalString, x.Root!.Descendants().Where(e => e.Name.Namespace == W && Kinds.Contains(e.Name.LocalName)).ToList()))];
    }

    /// <summary>"kind in container": the parent, or grandparent/parent for a property container.</summary>
    internal static string Site(XElement revision)
    {
        if (revision.Parent is not { } parent) return $"{revision.Name.LocalName} in (root)";
        var owner = Owner(revision);
        return $"{revision.Name.LocalName} in {(owner == parent ? Name(parent) : $"{Name(owner!)}/{Name(parent)}")}";
    }

    // The outermost element the site names: the parent, or the owner of a property container.
    private static XElement? Owner(XElement revision) =>
        revision.Parent is { } parent && (parent.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal) || parent.Name == W + "tblGrid") && parent.Parent is { } owner
            ? owner
            : revision.Parent;

    // The innermost WordprocessingML ancestor above the site that is not an allowed context, or null.
    private static string? UnknownContext(XElement revision) =>
        Owner(revision)?.Ancestors().Where(a => a.Name.Namespace == W).Select(Name).FirstOrDefault(n => !Sites.Contains("inside " + n));

    private static string Name(XElement e) => e.Name.Namespace == W ? e.Name.LocalName : $"{{{e.Name.NamespaceName}}}{e.Name.LocalName}";
}
