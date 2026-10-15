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
    /// The allowlist. "kind in container" admits a revision element under that parent (property
    /// containers are written grandparent/parent); "inside X" admits X as an ancestor. Evidence is
    /// a test that passes the tracked form through the pinned Clippit (zero errors) and fails the
    /// untracked form; PlantedDefectTests cases are named in brackets. Word's own behaviour for
    /// each container is recorded by ContainerProbeTests (Word, Clippit and ViewBytes agree).
    /// </summary>
    internal static readonly HashSet<string> Sites = new(StringComparer.Ordinal)
    {
        // Run-level insertion and deletion in a paragraph [body-text, paragraph-added,
        // table-cell-text, header-text, footer-text, footnote-text, endnote-text, image-*].
        "ins in p", "del in p",
        // Deleting an earlier insertion (Word writes w:del inside w:ins) [ins-then-del].
        "del in ins",
        // Inside a hyperlink, through ClippitWorkarounds' markers [hyperlink-add, hyperlink-remove,
        // hyperlink-retarget, hyperlink-anchor tracked forms; RevisionCheckTests hyperlink cases].
        "ins in hyperlink", "del in hyperlink",
        // Inserted/deleted paragraph marks [paragraph-added, paragraph-deleted, row-added, row-deleted].
        "ins in pPr/rPr", "del in pPr/rPr",
        // Inserted/deleted table rows [row-added, row-deleted].
        "ins in tr/trPr", "del in tr/trPr",
        // Property changes [paragraph-style, paragraph-properties, direct-numPr, run-formatting,
        // paragraph-mark-formatting, hyperlink-formatting, section-margins, section-break-margins,
        // table-properties, row-properties, cell-properties, table-grid, formatted-insertion,
        // formatted-deletion].
        "pPrChange in p/pPr", "rPrChange in r/rPr", "rPrChange in pPr/rPr",
        "sectPrChange in body/sectPr", "sectPrChange in pPr/sectPr",
        "tblPrChange in tbl/tblPr", "trPrChange in tr/trPr", "tcPrChange in tc/tcPr", "tblGridChange in tbl/tblGrid",
        // Moves; moved paragraph marks through ClippitWorkarounds (without it Clippit left the
        // moved-to paragraph behind as an empty paragraph) [move].
        "moveFrom in p", "moveTo in p", "moveFrom in pPr/rPr", "moveTo in pPr/rPr",
        "moveFromRangeStart in p", "moveFromRangeEnd in p", "moveToRangeStart in p", "moveToRangeEnd in p",
        // Content controls, smart tags [sdt-text, block-sdt-text, smartTag]; simple fields through
        // ClippitWorkarounds' markers (without them Clippit kept the inserted result) [fldSimple].
        "ins in sdtContent", "del in sdtContent", "ins in smartTag", "del in smartTag", "ins in fldSimple", "del in fldSimple",
        // Numbering and table-exception property changes: Word, the pinned Clippit and ViewBytes
        // reject and accept them alike, list strings included (RevisionKindProbeTests
        // numbering-change, tblprex-change; Word 16.113, 2026-10-15). Cell insertion, deletion
        // and merge stay off: Clippit leaves a short row on rejecting cellIns, throws on cellDel
        // after a cell without tcPr, and keeps a rejected cellMerge.
        "numberingChange in pPr/numPr", "tblPrExChange in tblPrEx",

        // Contexts. Story roots Clippit's RevisionProcessor rewrites (main, headers, footers,
        // footnotes, endnotes) [body-text, header-text, footer-text, footnote-text, endnote-text];
        // comments since the pinned Clippit 3.9.2-rmorgans (sergey-tihon/Clippit#553; 3.9.1
        // left them untouched), Word rejecting and accepting them alike (ContainerProbeTests "comment text")
        // [comment-changed, comment-formatting].
        "inside document", "inside body", "inside hdr", "inside ftr",
        "inside footnotes", "inside footnote", "inside endnotes", "inside endnote", "inside comments", "inside comment",
        // Block and run structure [body-text, table-cell-text, row-added, table-properties, cell-properties,
        // formatted-insertion, formatted-deletion, text-box-vml-text].
        "inside p", "inside r", "inside tbl", "inside tr", "inside tc", "inside ins", "inside del",
        // Hyperlinks and content controls [hyperlink-formatting, sdt-text, block-sdt-text].
        "inside hyperlink", "inside sdt", "inside sdtContent",
        // Text boxes, DrawingML and VML hosts [text-box-text, text-box-vml-text].
        "inside txbxContent", "inside drawing", "inside pict",
    };

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
                var first = group.First().Element;
                var paragraph = first.AncestorsAndSelf(W + "p").FirstOrDefault();
                var where = paragraph is null ? null : $"first in paragraph \"{ParagraphText.Excerpt(Wml.VisibleText(paragraph))}\"";
                yield return new Finding("unverifiable-revision", Severity.Error,
                    $"{label}: {group.Count()} tracked change(s) w:{group.Key} in {part}: officeproof has not verified what Reject All does with " +
                    "this revision, so the redline cannot be shown to reject back to the original",
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
