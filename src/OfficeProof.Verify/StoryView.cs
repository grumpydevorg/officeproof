using System.Xml;
using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// A story paragraph with its markup in canonical form: <paramref name="Properties"/> is the
/// paragraph properties, <paramref name="Content"/> the rest.
/// </summary>
internal sealed record StoryParagraph(OutlineEntry Entry, string Properties, string Content)
{
    /// <summary>What must match the original: everything but the position.</summary>
    public (string?, int?, string, string, string, string) Key =>
        (Entry.StyleId, Entry.Level, Entry.ListNumber, Entry.Text, Properties, Content);
}

/// <summary>
/// One story of a view: <paramref name="Name"/> identifies it across packages ("main story",
/// "header (section 1, default)", "footnote 3", "text box 1 in main story", "comment 2");
/// <paramref name="Skeleton"/> is its canonical markup with the paragraphs removed (tables,
/// rows, cells, section properties, story attributes). <paramref name="Defined"/> is false for
/// a note that is referenced but has no content.
/// </summary>
internal sealed record Story(string Name, string Part, IReadOnlyList<StoryParagraph> Paragraphs, string Skeleton, bool Defined = true);

/// <summary>
/// Every story of a package already in the wanted state (a Clippit view or a Word-saved copy),
/// reduced to what a reader can see so two packages compare story by story:
/// <list type="bullet">
/// <item>main story: the whole w:document, list numbers from Clippit's ListItemRetriever;</item>
/// <item>headers and footers: by section index and reference type, not by part name;</item>
/// <item>footnotes and endnotes: by id, those referenced from the main story, headers and footers,
/// plus the separator notes by type;</item>
/// <item>text boxes (w:txbxContent): a story of their own after their host; a box repeated in an
/// mc:Fallback stays inline in the host, so each mc:AlternateContent counts once;</item>
/// <item>comments: by id, those referenced from a story (kept apart in <see cref="Comments"/>).</item>
/// </list>
/// Canonical form: bookkeeping removed (rsids, w14 ids, proofing marks, rendered page breaks,
/// comment anchors, _GoBack and the redline's own _Ref bookmarks, bookmark ids, hyperlink
/// history), runs with identical formatting coalesced, and every relationship reference replaced
/// by what it resolves to (<see cref="Relationships"/>). Header and footer references become
/// "story": the referenced story is compared in its own right.
/// ListItemRetriever numbers every story but comments (Clippit numbers the main story, headers,
/// footers, footnotes, endnotes and each text box as separate sequences; it does not read the
/// comments part).
/// </summary>
internal sealed class StoryView
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace W14 = Ns.W14;
    private static readonly XNamespace Wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly HashSet<XName> Bookkeeping =
    [
        W + "proofErr", W + "lastRenderedPageBreak",
        W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference",
    ];

    private static readonly XName[] EmptyMeansAbsent = [W + "rPr", W + "pPr", W + "trPr", W + "tcPr", W + "tblPrEx"];

    private static readonly HashSet<string> SeparatorTypes = new(StringComparer.Ordinal) { "separator", "continuationSeparator", "continuationNotice" };
    private static readonly HashSet<string> ReferenceTypes = new(StringComparer.Ordinal) { "default", "first", "even" };

    private StoryView(IReadOnlyList<Story> stories, IReadOnlyList<Story> comments, IReadOnlySet<string> bookmarks, IReadOnlyList<Finding> problems)
    {
        Stories = stories;
        Comments = comments;
        Bookmarks = bookmarks;
        Problems = problems;
    }

    /// <summary>Main story, headers and footers, notes, and text boxes (each after its host).</summary>
    public IReadOnlyList<Story> Stories { get; }

    /// <summary>Referenced comments ("comment &lt;id&gt;") and their text boxes.</summary>
    public IReadOnlyList<Story> Comments { get; }

    /// <summary>Bookmark names in every story.</summary>
    public IReadOnlySet<string> Bookmarks { get; }

    /// <summary>
    /// Stories this view could not load or pair (an unknown header reference type, a dangling
    /// reference, a hyperlink Clippit separated from its markers): rule
    /// <c>unverifiable-revision</c>, Error. Whatever is listed here is not verified.
    /// </summary>
    public IReadOnlyList<Finding> Problems { get; }

    /// <summary>
    /// Reads every story of <paramref name="viewPackage"/>. <paramref name="original"/> (null when
    /// this is the original) supplies the bookmarks the original already had; any other _Ref
    /// bookmark is the redline's own caption or heading target and is ignored.
    /// </summary>
    public static StoryView Load(byte[] viewPackage, StoryView? original = null)
    {
        using var stream = new MemoryStream(viewPackage);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var builder = new Builder(doc, original?.Bookmarks);
        builder.LoadAll();
        return new StoryView(builder.Stories, builder.Comments, builder.BookmarkNames, builder.Problems);
    }

    private sealed class Builder(WordprocessingDocument doc, IReadOnlySet<string>? originalBookmarks)
    {
        private readonly MainDocumentPart _main = doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part");
        private readonly StyleChain _styles = StyleChain.Load(doc);
        private readonly Relationships _relationships = new();
        private readonly List<XElement> _roots = [];
        private readonly HashSet<string> _names = new(StringComparer.Ordinal);

        public List<Story> Stories { get; } = [];
        public List<Story> Comments { get; } = [];
        public HashSet<string> BookmarkNames { get; } = new(StringComparer.Ordinal);
        public List<Finding> Problems { get; } = [];

        public void LoadAll()
        {
            var document = Root(_main);
            if (document is null) return;
            Add("main story", _main, document, numbered: true, Stories);
            LoadHeadersAndFooters(document);
            LoadNotes(_main.FootnotesPart, "footnote");
            LoadNotes(_main.EndnotesPart, "endnote");
            LoadComments();
        }

        private void LoadHeadersAndFooters(XElement document)
        {
            var sections = document.Element(W + "body")?.Descendants(W + "sectPr")
                .Where(s => !s.Ancestors().Any(a => a.Name == W + "sectPrChange" || a.Name == W + "txbxContent"))
                .ToList() ?? [];
            for (var i = 0; i < sections.Count; i++)
            {
                foreach (var reference in sections[i].Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference"))
                {
                    var kind = reference.Name.LocalName[..^"Reference".Length];
                    var type = (string?)reference.Attribute(W + "type");
                    var section = $"section {i + 1}";
                    if (type is null || !ReferenceTypes.Contains(type))
                    {
                        Problem($"{kind} reference of type '{type ?? "(none)"}' in {section} is not modelled; that {kind} is not verified", _main, section);
                        continue;
                    }
                    var name = $"{kind} ({section}, {type})";
                    var id = (string?)reference.Attribute(Ns.R + "id");
                    var part = id is null ? null : Relationships.Part(_main, id);
                    if (part is not (HeaderPart or FooterPart) || (part is HeaderPart) != (kind == "header"))
                    {
                        Problem($"{name} refers to relationship '{id}', which is not a {kind} part; that {kind} is not verified", _main, name);
                        continue;
                    }
                    if (_names.Contains(name))
                    {
                        Problem($"{section} has more than one {type} {kind} reference; only the first is verified", _main, name);
                        continue;
                    }
                    if (Root(part) is { } root) Add(name, part, root, numbered: true, Stories);
                }
            }
        }

        private void LoadNotes(OpenXmlPart? part, string kind)
        {
            var referenced = _roots.SelectMany(r => r.Descendants(W + kind + "Reference"))
                .Select(r => (string?)r.Attribute(W + "id")).OfType<string>().Distinct(StringComparer.Ordinal).ToList();
            var notes = (part is null ? null : Root(part))?.Elements(W + kind).ToList() ?? [];
            if (part is not null)
            {
                foreach (var note in notes)
                {
                    var type = (string?)note.Attribute(W + "type");
                    if (type is not null && SeparatorTypes.Contains(type)) Add($"{kind} {type}", part, note, numbered: true, Stories);
                }
            }
            foreach (var id in referenced)
            {
                var name = $"{kind} {id}";
                var note = notes.FirstOrDefault(n => (string?)n.Attribute(W + "id") == id && ((string?)n.Attribute(W + "type") ?? "normal") == "normal");
                if (part is null || note is null)
                {
                    if (_names.Add(name)) Stories.Add(new Story(name, part?.Uri.OriginalString ?? $"/word/{kind}s.xml", [], "", Defined: false));
                    continue;
                }
                Add(name, part, note, numbered: true, Stories);
            }
        }

        private void LoadComments()
        {
            var referenced = _roots.SelectMany(r => r.Descendants(W + "commentReference"))
                .Select(r => (string?)r.Attribute(W + "id")).OfType<string>().ToHashSet(StringComparer.Ordinal);
            if (_main.WordprocessingCommentsPart is not { } part || referenced.Count == 0) return;
            var root = Root(part);
            if (root is null) return;
            foreach (var comment in root.Elements(W + "comment"))
            {
                var id = (string?)comment.Attribute(W + "id");
                if (id is not null && referenced.Contains(id)) Add($"comment {id}", part, comment, numbered: false, Comments);
            }
        }

        private XElement? Root(OpenXmlPart part)
        {
            try
            {
                // Clippit's cached XDocument: ListItemRetriever annotates these very elements.
                return part.GetXDocument().Root ?? throw new InvalidDataException("part has no root element");
            }
            catch (Exception e) when (e is XmlException or InvalidDataException)
            {
                Problem($"{part.Uri} cannot be read ({e.Message}); its stories are not verified", part, null);
                return null;
            }
        }

        private void Problem(string message, OpenXmlPart part, string? location) =>
            Problems.Add(new Finding("unverifiable-revision", Severity.Error, message, part.Uri.OriginalString, location));

        private void Add(string name, OpenXmlPart part, XElement root, bool numbered, List<Story> into)
        {
            var unique = name;
            for (var n = 2; !_names.Add(unique); n++) unique = $"{name} #{n}";
            name = unique;
            if (root.Parent is null || root.Parent == root.Document?.Root) _roots.Add(root);
            foreach (var b in root.Descendants(W + "bookmarkStart"))
                if ((string?)b.Attribute(W + "name") is { } bookmark) BookmarkNames.Add(bookmark);

            var paragraphs = Paragraphs(root).ToList();
            var entries = new List<OutlineEntry>(paragraphs.Count);
            foreach (var p in paragraphs)
            {
                var number = numbered ? ListItemRetriever.RetrieveListItem(doc, p) ?? "" : "";
                entries.Add(new OutlineEntry(entries.Count, Wml.StyleId(p), _styles.OutlineLevel(p), number, Wml.VisibleText(p)));
            }

            var boxes = TextBoxes(root).ToList();
            var copy = new XElement(root);
            var copyBoxes = TextBoxes(copy).ToList();
            for (var k = 0; k < copyBoxes.Count; k++)
                copyBoxes[k].ReplaceWith(new XElement(W + "txbxContent", new XAttribute(ClippitWorkarounds.Dw + "story", BoxName(k, name))));
            Resolve(copy, part);
            foreach (var wrapper in copy.Descendants().Where(h => h.Attribute(ClippitWorkarounds.Unrestored) is not null))
                Problem($"a w:{wrapper.Name.LocalName} in {name} could not be followed through accept/reject (Clippit separated its markers); it is not verified",
                    part, $"{name}: {Describe(wrapper)}");
            Normalize(copy, originalBookmarks);

            var canonical = Paragraphs(copy).ToList();
            if (canonical.Count != entries.Count)
                throw new InvalidOperationException($"paragraph count mismatch reading {name} ({canonical.Count} vs {entries.Count})");
            var result = new List<StoryParagraph>(entries.Count);
            for (var i = 0; i < entries.Count; i++)
            {
                var p = canonical[i];
                var pPr = p.Element(W + "pPr");
                var content = new XElement(p.Name, p.Attributes(), p.Nodes().Where(n => n != pPr));
                result.Add(new StoryParagraph(entries[i], pPr is null ? "" : Serialize(pPr), Serialize(content)));
            }
            var skeleton = new XElement(copy);
            skeleton.Descendants(W + "p").Where(p => !p.Ancestors(W + "p").Any()).ToList().Remove();
            into.Add(new Story(name, part.Uri.OriginalString, result, Serialize(skeleton)));

            for (var k = 0; k < boxes.Count; k++) Add(BoxName(k, name), part, boxes[k], numbered, into);
        }

        private static string BoxName(int k, string host) => $"text box {k + 1} in {host}";

        private static string Describe(XElement wrapper) =>
            string.Join(" ", wrapper.Attributes().Where(a => a.Name != ClippitWorkarounds.Unrestored).Select(a => $"{a.Name.LocalName}={a.Value}"));

        // Every relationship id becomes its target; header/footer references are paired by section instead.
        private void Resolve(XElement root, OpenXmlPart part)
        {
            foreach (var e in root.DescendantsAndSelf())
            {
                var isStoryReference = e.Name == W + "headerReference" || e.Name == W + "footerReference";
                foreach (var a in e.Attributes().Where(Relationships.IsReference))
                    a.Value = isStoryReference ? "story" : _relationships.Resolve(part, a.Value);
            }
        }
    }

    /// <summary>The story's own paragraphs: nested text boxes are stories of their own.</summary>
    private static IEnumerable<XElement> Paragraphs(XElement root) =>
        root.Descendants(W + "p").Where(p => !p.Ancestors().TakeWhile(a => a != root).Any(a => a.Name == W + "txbxContent"));

    /// <summary>The text boxes directly inside the story, in document order, skipping mc:Fallback copies.</summary>
    private static IEnumerable<XElement> TextBoxes(XElement root) =>
        root.Descendants(W + "txbxContent").Where(t => !t.Ancestors().TakeWhile(a => a != root).Any(a => a.Name == W + "txbxContent" || a.Name == Mc + "Fallback"));

    private static void Normalize(XElement root, IReadOnlySet<string>? originalBookmarks)
    {
        foreach (var e in root.DescendantsAndSelf())
            e.Attributes().Where(IsBookkeepingAttribute).ToList().Remove();

        var droppedBookmarks = new HashSet<string>(StringComparer.Ordinal);
        foreach (var start in root.Descendants(W + "bookmarkStart").ToList())
        {
            var name = (string?)start.Attribute(W + "name") ?? "";
            var ours = originalBookmarks is not null && name.StartsWith("_Ref", StringComparison.Ordinal) && !originalBookmarks.Contains(name);
            if (ours || name == "_GoBack")
            {
                droppedBookmarks.Add((string?)start.Attribute(W + "id") ?? "");
                start.Remove();
            }
        }
        foreach (var end in root.Descendants(W + "bookmarkEnd").ToList())
            if (droppedBookmarks.Contains((string?)end.Attribute(W + "id") ?? "")) end.Remove();
        // Bookmark ids are renumbered freely by Word and redline tools; the name is what matters.
        foreach (var b in root.Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd"))
            b.Attribute(W + "id")?.Remove();

        root.Descendants().Where(e => Bookkeeping.Contains(e.Name)).ToList().Remove();
        root.Descendants(W + "t").Where(t => t.Value.Length == 0).ToList().Remove();
        root.Descendants(W + "r").Where(r => r.Elements().All(c => c.Name == W + "rPr")).ToList().Remove();
        // An empty property container says nothing; Clippit leaves one behind after rejecting a
        // property change that had nothing to restore (PlantedDefectTests "row-properties").
        foreach (var name in EmptyMeansAbsent)
            root.Descendants(name).Where(e => !e.HasElements && !e.HasAttributes).ToList().Remove();

        foreach (var parent in root.DescendantsAndSelf().Where(e => e.Elements(W + "r").Skip(1).Any()).ToList())
            CoalesceRuns(parent);
        foreach (var run in root.Descendants(W + "r"))
            CoalesceText(run);
    }

    private static bool IsBookkeepingAttribute(XAttribute a) =>
        a.IsNamespaceDeclaration
        || a.Name == Ns.Xml + "space"
        || a.Name == Mc + "Ignorable"
        || (a.Name.Namespace == W && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
        || a.Name == W14 + "paraId" || a.Name == W14 + "textId"
        || a.Name == Wp14 + "anchorId" || a.Name == Wp14 + "editId"
        || (a.Name == W + "history" && a.Parent?.Name == W + "hyperlink");

    // Adjacent runs with the same formatting read the same as one run.
    private static void CoalesceRuns(XElement parent)
    {
        XElement? previous = null;
        foreach (var node in parent.Nodes().ToList())
        {
            if (node is XElement { } run && run.Name == W + "r")
            {
                if (previous is not null && XNode.DeepEquals(previous.Element(W + "rPr"), run.Element(W + "rPr")))
                {
                    previous.Add(run.Elements().Where(c => c.Name != W + "rPr"));
                    run.Remove();
                    continue;
                }
                previous = run;
            }
            else if (node is XElement)
            {
                previous = null;
            }
        }
    }

    private static void CoalesceText(XElement run)
    {
        XElement? previous = null;
        foreach (var child in run.Elements().ToList())
        {
            if (child.Name == W + "t" && previous is not null)
            {
                previous.Value += child.Value;
                child.Remove();
                continue;
            }
            previous = child.Name == W + "t" ? child : null;
        }
    }

    /// <summary>Attribute order and layout whitespace do not matter; text inside w:t does.</summary>
    private static string Serialize(XElement element) => Canonical(element).ToString(SaveOptions.DisableFormatting);

    private static XElement Canonical(XElement e) => new(e.Name,
        e.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)
            .Select(a => new XAttribute(a.Name, a.Value)),
        e.Nodes().Select(n => n switch
        {
            XElement c => (object?)Canonical(c),
            XText t when !(e.HasElements && string.IsNullOrWhiteSpace(t.Value)) => new XText(t.Value),
            _ => null,
        }).OfType<object>());
}
