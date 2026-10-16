using System.IO.Compression;
using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.TestKit.DocxBuilder;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

/// <summary>
/// For each construct of a document, an untracked change must give a specific Error and the
/// tracked equivalent must give none. The fixture (<see cref="Rich"/>) holds every construct
/// once; each case edits a copy of it.
/// </summary>
public sealed class PlantedDefectTests : IDisposable
{
    private static readonly XNamespace W = Ns.W;
    private readonly Redlines _r = new();
    private readonly string _original;

    public PlantedDefectTests() => _original = Rich(Path.Combine(_r.Dir, "rich.docx"));

    public void Dispose() => _r.Dispose();

    /// <summary>
    /// One construct. <paramref name="Rule"/> and <paramref name="Story"/> are what the untracked
    /// change must produce (the story appears in the finding's message or location).
    /// <paramref name="Tracked"/> null: Word has no tracked form (the reason is
    /// <paramref name="Note"/>). <paramref name="TrackedRule"/> set: the tracked form cannot be
    /// verified under Clippit and must fail closed with that rule instead of passing.
    /// <paramref name="DocumentUnchanged"/>: the untracked change leaves word/document.xml byte-identical.
    /// <paramref name="Extra"/>: a paragraph appended to the fixture for this case only.
    /// </summary>
    internal sealed record Case(
        string Rule, string Story, Action<PackageEdit>? Untracked, Action<PackageEdit>? Tracked,
        string? TrackedRule = null, bool DocumentUnchanged = false, string? Note = null, XElement? Extra = null);

    /// <summary>The fixture: one of every construct the cases edit.</summary>
    internal static string Rich(string path, XElement? extra = null) => new DocxBuilder()
        .Raw(P(new XElement(W + "pPr", new XElement(W + "sectPr", PageSize())), R("First section.")))
        .Para("Body text.")
        .Para("Formatted text.")
        .Para("Styled text.")
        .Para("List text.")
        .Raw(new XElement(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 2000), new XAttribute(W + "type", "dxa"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 2000))),
            Row("Cell text."), Row("Second row.")))
        .Raw(P(R("Link: "), Hyperlink("rIdLink1", R("link text")), R(" end.")))
        .Raw(P(new XElement(W + "hyperlink", new XAttribute(W + "anchor", "Target1"), R("anchor link"))))
        .Raw(P(Bookmark("bookmarkStart", "Target1"), R("Bookmarked text."), Bookmark("bookmarkEnd", null)))
        .Raw(P(FieldChar("begin"), Instr(" HYPERLINK \"https://example.com/field\" "), FieldChar("separate"), R("field link"), FieldChar("end")))
        .Raw(P(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 1)), R("Commented text."),
            new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 1)),
            new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 1)))))
        .Raw(P(R("Image: "), Picture("rIdImg1", 1)))
        .Raw(P(R("Other image: "), Picture("rIdImg2", 2)))
        .Raw(P(R("Box: "), VmlTextBox(P(R("Box text.")))))
        .Raw(P(R("Shape: "), ShapeTextBox(3, P(R("Shape text.")))))
        .Raw(P(Sdt(1, R("Control text."))))
        .Raw(Sdt(2, P(R("Block control text."))))
        .Raw(P(FldSimple(" DATE ", R("Date text."))))
        .Para("Move me.")
        .Para("Stay here.")
        .Raw(P(R("Noted."), NoteReference(1)))
        .Raw(P(R("End noted."), NoteReference(1, endnote: true)))
        .Para("Last paragraph.")
        .Raw(extra ?? P(R("No extra.")))
        .Header("rIdH1", P(R("Header text.")))
        .Footer("rIdF1", P(R("Footer text.")))
        .Footnote(1, P(R("Footnote text.")))
        .Endnote(1, P(R("Endnote text.")))
        .Link("rIdLink1", "https://example.com/a")
        .Link("rIdLink2", "https://example.com/b")
        .Comment(1, "Original comment.")
        .Image("rIdImg1", PngA)
        .Image("rIdImg2", PngB)
        .Save(path);

    private static XElement PageSize() => new(W + "pgSz", new XAttribute(W + "w", 11906), new XAttribute(W + "h", 16838));

    private static XElement Row(string text) => new(W + "tr", new XElement(W + "tc",
        new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 2000), new XAttribute(W + "type", "dxa"))),
        P(R(text))));

    internal static XElement Hyperlink(string relId, params object[] content) =>
        new(W + "hyperlink", new XAttribute(Ns.R + "id", relId), content);

    private static XElement Bookmark(string kind, string? name) => new(W + kind, new XAttribute(W + "id", 1),
        name is null ? null : new XAttribute(W + "name", name));

    private static XElement FieldChar(string kind) => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", kind)));

    private static XElement Instr(string code) => new(W + "r", new XElement(W + "instrText", new XAttribute(Ns.Xml + "space", "preserve"), code));

    private static XElement Margins() => new(W + "pgMar",
        new XAttribute(W + "top", 1440), new XAttribute(W + "right", 1440), new XAttribute(W + "bottom", 1440), new XAttribute(W + "left", 1440),
        new XAttribute(W + "header", 708), new XAttribute(W + "footer", 708), new XAttribute(W + "gutter", 0));

    // ---- editing helpers -------------------------------------------------------------------

    private static XElement Para(XElement root, string text) =>
        root.Descendants(W + "p").First(p => Wml.VisibleText(p) == text);

    private static IEnumerable<XElement> Paras(XElement root, string text) =>
        [.. root.Descendants(W + "p").Where(p => Wml.VisibleText(p) == text)];

    /// <summary>Replaces the runs of <paramref name="container"/> with one plain run, untracked.</summary>
    private static void Set(XElement container, string text)
    {
        container.Elements(W + "r").Remove();
        container.Add(Run(text));
    }

    /// <summary>Replaces the runs of <paramref name="container"/> with a tracked deletion and insertion.</summary>
    private static void Track(PackageEdit e, XElement container, string oldText, string newText)
    {
        container.Elements(W + "r").Remove();
        container.Add(e.R.Del(oldText), e.R.Ins(Run(newText)));
    }

    private static XElement Mark(PackageEdit e, string kind) => new(W + kind, e.R.Mark());

    private static XElement PropertyChange(PackageEdit e, string kind, params object[] old) => new(W + kind, e.R.Mark(), old);

    private static XElement Link(PackageEdit e) => e.Body.Descendants(W + "hyperlink").First(h => h.Attribute(Ns.R + "id") is not null);

    private static XElement AnchorLink(PackageEdit e) => e.Body.Descendants(W + "hyperlink").First(h => h.Attribute(W + "anchor") is not null);

    private static void ReplacePicture(PackageEdit e)
    {
        var image = e.Main.AddImagePart(ImagePartType.Png, "rIdImg3");
        using (var s = new MemoryStream(PngB)) image.FeedData(s);
        var run = Para(e.Body, "Image: ").Elements(W + "r").Single(r => r.Element(W + "drawing") is not null);
        run.ReplaceWith(new XElement(W + "del", e.R.Mark(), new XElement(run)), e.R.Ins(Picture("rIdImg3", 4)));
    }

    // ---- the cases ---------------------------------------------------------------------

    private static readonly Dictionary<string, Case> Cases = new()
    {
        ["body-text"] = new("reject-mismatch", "main story",
            e => Set(Para(e.Body, "Body text."), "Body text, edited."),
            e => Track(e, Para(e.Body, "Body text."), "Body text.", "Body text, edited.")),
        ["paragraph-added"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Body text.").AddAfterSelf(P(R("New paragraph."))),
            e => Para(e.Body, "Body text.").AddAfterSelf(P(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "ins"))), e.R.Ins(Run("New paragraph."))))),
        ["paragraph-deleted"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Formatted text.").Remove(),
            e =>
            {
                var p = Para(e.Body, "Formatted text.");
                p.Elements(W + "r").Remove();
                p.Add(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "del"))), e.R.Del("Formatted text."));
            }),
        ["run-formatting"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Formatted text.").Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"))),
            e => Para(e.Body, "Formatted text.").Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"),
                PropertyChange(e, "rPrChange", new XElement(W + "rPr"))))),
        ["paragraph-mark-formatting"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Body text.").AddFirst(new XElement(W + "pPr", new XElement(W + "rPr", new XElement(W + "b")))),
            e => Para(e.Body, "Body text.").AddFirst(new XElement(W + "pPr", new XElement(W + "rPr", new XElement(W + "b"),
                PropertyChange(e, "rPrChange", new XElement(W + "rPr")))))),
        ["paragraph-style"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Styled text.").AddFirst(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Head2")))),
            e => Para(e.Body, "Styled text.").AddFirst(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Head2")),
                PropertyChange(e, "pPrChange", new XElement(W + "pPr"))))),
        ["paragraph-properties"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Styled text.").AddFirst(new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")))),
            e => Para(e.Body, "Styled text.").AddFirst(new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")),
                PropertyChange(e, "pPrChange", new XElement(W + "pPr"))))),
        ["direct-numPr"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "List text.").AddFirst(new XElement(W + "pPr", NumPr())),
            e => Para(e.Body, "List text.").AddFirst(new XElement(W + "pPr", NumPr(), PropertyChange(e, "pPrChange", new XElement(W + "pPr"))))),
        // A list paragraph's tracked numbering, recorded as w:numberingChange in w:numPr as Word writes it.
        ["numbering-change"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "List text.").AddFirst(new XElement(W + "pPr", NumPr())),
            e => Para(e.Body, "List text.").AddFirst(new XElement(W + "pPr",
                new XElement(W + "numPr", NumPr().Elements(), new XElement(W + "numberingChange", e.R.Mark(), new XAttribute(W + "original", "%1."))),
                PropertyChange(e, "pPrChange", new XElement(W + "pPr"))))),
        ["table-cell-text"] = new("reject-mismatch", "main story",
            e => Set(Para(e.Body, "Cell text."), "Cell text, edited."),
            e => Track(e, Para(e.Body, "Cell text."), "Cell text.", "Cell text, edited.")),
        ["table-properties"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "tblW").Single().SetAttributeValue(W + "w", 4000),
            e =>
            {
                var tblPr = e.Body.Descendants(W + "tblPr").Single();
                var old = new XElement(tblPr);
                tblPr.Element(W + "tblW")!.SetAttributeValue(W + "w", 4000);
                tblPr.Add(PropertyChange(e, "tblPrChange", old));
            }),
        ["table-grid"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "gridCol").Single().SetAttributeValue(W + "w", 3000),
            e =>
            {
                var grid = e.Body.Descendants(W + "tblGrid").Single();
                var old = new XElement(grid);
                grid.Element(W + "gridCol")!.SetAttributeValue(W + "w", 3000);
                grid.Add(new XElement(W + "tblGridChange", new XAttribute(W + "id", 990), old));
            }),
        ["row-added"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "tbl").Single().Add(Row("New row.")),
            e => e.Body.Descendants(W + "tbl").Single().Add(new XElement(W + "tr",
                new XElement(W + "trPr", Mark(e, "ins")),
                new XElement(W + "tc",
                    new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 2000), new XAttribute(W + "type", "dxa"))),
                    P(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "ins"))), e.R.Ins(Run("New row."))))))),
        ["row-deleted"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Second row.").Ancestors(W + "tr").First().Remove(),
            e =>
            {
                var p = Para(e.Body, "Second row.");
                p.Ancestors(W + "tr").First().AddFirst(new XElement(W + "trPr", Mark(e, "del")));
                p.Elements(W + "r").Remove();
                p.Add(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "del"))), e.R.Del("Second row."));
            }),
        ["row-properties"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "tr").First().AddFirst(new XElement(W + "trPr", new XElement(W + "cantSplit"))),
            e => e.Body.Descendants(W + "tr").First().AddFirst(new XElement(W + "trPr", new XElement(W + "cantSplit"),
                PropertyChange(e, "trPrChange", new XElement(W + "trPr"))))),
        // Word for Windows and Word for Mac reject a tracked table-exception change differently
        // (RevisionKindProbeTests tblprex-change), so the tracked form fails closed.
        ["table-exception"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "tr").First().AddFirst(new XElement(W + "tblPrEx", new XElement(W + "jc", new XAttribute(W + "val", "center")))),
            e => e.Body.Descendants(W + "tr").First().AddFirst(new XElement(W + "tblPrEx", new XElement(W + "jc", new XAttribute(W + "val", "center")),
                PropertyChange(e, "tblPrExChange", new XElement(W + "tblPrEx")))),
            TrackedRule: "unverifiable-revision"),
        ["cell-properties"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "tcPr").First().Add(Shading()),
            e =>
            {
                var tcPr = e.Body.Descendants(W + "tcPr").First();
                var old = new XElement(tcPr);
                tcPr.Add(Shading(), PropertyChange(e, "tcPrChange", old));
            }),
        ["section-margins"] = new("reject-mismatch", "main story",
            e => e.Body.Element(W + "sectPr")!.Element(W + "pgSz")!.AddAfterSelf(Margins()),
            e =>
            {
                var sectPr = e.Body.Element(W + "sectPr")!;
                var old = new XElement(W + "sectPr", sectPr.Elements().Where(c => c.Name != W + "headerReference" && c.Name != W + "footerReference").Select(c => new XElement(c)));
                sectPr.Element(W + "pgSz")!.AddAfterSelf(Margins());
                sectPr.Add(PropertyChange(e, "sectPrChange", old));
            }),
        ["section-break-margins"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "First section.").Descendants(W + "pgSz").Single().AddAfterSelf(Margins()),
            e =>
            {
                var sectPr = Para(e.Body, "First section.").Descendants(W + "sectPr").Single();
                var old = new XElement(sectPr);
                sectPr.Element(W + "pgSz")!.AddAfterSelf(Margins());
                sectPr.Add(PropertyChange(e, "sectPrChange", old));
            }),
        ["header-text"] = new("reject-mismatch", "header (section 2, default)",
            e => Set(Para(e.Header, "Header text."), "Header text, edited."),
            e => Track(e, Para(e.Header, "Header text."), "Header text.", "Header text, edited.")),
        ["footer-text"] = new("reject-mismatch", "footer (section 2, default)",
            e => Set(Para(e.Footer, "Footer text."), "Footer text, edited."),
            e => Track(e, Para(e.Footer, "Footer text."), "Footer text.", "Footer text, edited.")),
        ["footnote-text"] = new("reject-mismatch", "footnote 1",
            e => Set(Para(e.Footnotes, "Footnote text."), "Footnote text, edited."),
            e => Track(e, Para(e.Footnotes, "Footnote text."), "Footnote text.", "Footnote text, edited.")),
        ["endnote-text"] = new("reject-mismatch", "endnote 1",
            e => Set(Para(e.Endnotes, "Endnote text."), "Endnote text, edited."),
            e => Track(e, Para(e.Endnotes, "Endnote text."), "Endnote text.", "Endnote text, edited.")),
        ["hyperlink-add"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Link: link text end.").Elements(W + "r").Last().ReplaceWith(Hyperlink("rIdLink2", R(" end."))),
            e => Para(e.Body, "Link: link text end.").Add(Hyperlink("rIdLink2", e.R.Ins(Run(" more"))))),
        ["hyperlink-remove"] = new("reject-mismatch", "main story",
            e => Link(e).ReplaceWith(Link(e).Nodes()),
            e =>
            {
                var link = Link(e);
                link.ReplaceNodes(e.R.Del("link text"));
                link.AddAfterSelf(e.R.Ins(Run("link text")));
            }),
        ["hyperlink-retarget"] = new("reject-mismatch", "main story",
            e =>
            {
                e.Main.DeleteReferenceRelationship("rIdLink1");
                e.Main.AddHyperlinkRelationship(new Uri("https://example.com/changed"), isExternal: true, "rIdLink1");
            },
            e =>
            {
                var link = Link(e);
                link.ReplaceNodes(e.R.Del("link text"));
                link.AddAfterSelf(Hyperlink("rIdLink2", e.R.Ins(Run("link text"))));
            },
            DocumentUnchanged: true),
        ["hyperlink-anchor"] = new("reject-mismatch", "main story",
            e => AnchorLink(e).SetAttributeValue(W + "anchor", "Target2"),
            e =>
            {
                var link = AnchorLink(e);
                link.ReplaceNodes(e.R.Del("anchor link"));
                link.AddAfterSelf(new XElement(W + "hyperlink", new XAttribute(W + "anchor", "Target2"), e.R.Ins(Run("anchor link"))));
            }),
        ["hyperlink-formatting"] = new("reject-mismatch", "main story",
            e => Link(e).Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"))),
            e => Link(e).Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"),
                PropertyChange(e, "rPrChange", new XElement(W + "rPr"))))),
        ["hyperlink-field-instruction"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "instrText").Single().Value = " HYPERLINK \"https://example.com/other\" ",
            e =>
            {
                var run = e.Body.Descendants(W + "instrText").Single().Parent!;
                run.ReplaceWith(
                    new XElement(W + "del", e.R.Mark(), new XElement(W + "r",
                        new XElement(W + "delInstrText", new XAttribute(Ns.Xml + "space", "preserve"), " HYPERLINK \"https://example.com/field\" "))),
                    e.R.Ins(Instr(" HYPERLINK \"https://example.com/other\" ")));
            }),
        ["bookmark-removed"] = new("reject-mismatch", "main story",
            e => Para(e.Body, "Bookmarked text.").Elements().Where(c => c.Name == W + "bookmarkStart" || c.Name == W + "bookmarkEnd").Remove(),
            e => Para(e.Body, "Bookmarked text.").Element(W + "r")!.ReplaceWith(e.R.Del("Bookmarked text."), e.R.Ins(Run("Rewritten text."))),
            Note: "a bookmark cannot be tracked; the tracked form deletes the bookmarked text and keeps the bookmark"),
        ["comment-changed"] = new("comment-changed", "comment 1",
            e => Set(Para(e.Comments, "Original comment."), "Changed comment."),
            e => Track(e, Para(e.Comments, "Original comment."), "Original comment.", "Changed comment.")),
        ["comment-formatting"] = new("comment-changed", "comment 1",
            e => Para(e.Comments, "Original comment.").Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"))),
            e => Para(e.Comments, "Original comment.").Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"),
                PropertyChange(e, "rPrChange", new XElement(W + "rPr"))))),
        ["comment-removed"] = new("comment-changed", "comment 1",
            e =>
            {
                e.Comments.Elements(W + "comment").Single().Remove();
                var p = Para(e.Body, "Commented text.");
                p.Elements().Where(c => c.Name != W + "r" || c.Element(W + "commentReference") is not null).Remove();
            },
            e =>
            {
                var p = Para(e.Body, "Commented text.");
                var reference = p.Elements(W + "r").Single(r => r.Element(W + "commentReference") is not null);
                reference.ReplaceWith(new XElement(W + "del", e.R.Mark(), new XElement(reference)));
                p.Elements(W + "r").Single(r => r.Element(W + "t") is not null).ReplaceWith(e.R.Del("Commented text."));
            },
            Note: "removing a comment cannot be tracked; the tracked form deletes the commented text and its anchor with tracking"),
        ["style-changed"] = new("definition-changed", "style Normal",
            e => Style(e.Root(e.Main.StyleDefinitionsPart!), "Normal").Descendants(W + "sz").Single().SetAttributeValue(W + "val", 28),
            null, DocumentUnchanged: true,
            Note: "Word cannot track a style definition (revisions in styles.xml survive Reject All: revision-in-styles); " +
                  "a change declared in VerifyOptions.IntendedStyles is a warning instead (RevisionCheckTests)"),
        ["style-removed"] = new("definition-changed", "style TOC3",
            e => Style(e.Root(e.Main.StyleDefinitionsPart!), "TOC3").Remove(),
            null, DocumentUnchanged: true, Note: "removing a style cannot be tracked"),
        ["document-defaults"] = new("definition-changed", "docDefaults",
            e => e.Root(e.Main.StyleDefinitionsPart!).Element(W + "docDefaults")!.Descendants(W + "rFonts").Single().SetAttributeValue(W + "ascii", "Arial"),
            null, DocumentUnchanged: true, Note: "document defaults cannot be tracked"),
        ["layout-setting"] = new("definition-changed", "evenAndOddHeaders",
            e => e.Root(e.Main.DocumentSettingsPart!).Add(new XElement(W + "evenAndOddHeaders")),
            null, DocumentUnchanged: true, Note: "document settings cannot be tracked"),
        ["image-bytes"] = new("reject-mismatch", "main story",
            e =>
            {
                var image = (ImagePart)e.Main.GetPartById("rIdImg1");
                using var s = new MemoryStream(PngB);
                image.FeedData(s);
            },
            ReplacePicture, DocumentUnchanged: true),
        ["image-retarget"] = new("reject-mismatch", "main story",
            e =>
            {
                var target = e.Main.GetPartById("rIdImg2").Uri.OriginalString.Replace("/word/", "", StringComparison.Ordinal);
                var current = e.Main.GetPartById("rIdImg1").Uri.OriginalString.Replace("/word/", "", StringComparison.Ordinal);
                e.Rewrite("word/_rels/document.xml.rels", rels => rels.Replace($"Target=\"/word/{current}\"", $"Target=\"/word/{target}\"", StringComparison.Ordinal)
                    .Replace($"Target=\"{current}\"", $"Target=\"{target}\"", StringComparison.Ordinal));
            },
            ReplacePicture, DocumentUnchanged: true),
        ["text-box-text"] = new("reject-mismatch", "text box 2 in main story",
            e => { foreach (var p in Paras(e.Body, "Shape text.")) Set(p, "Shape text, edited."); },
            e => { foreach (var p in Paras(e.Body, "Shape text.")) Track(e, p, "Shape text.", "Shape text, edited."); }),
        ["text-box-vml-text"] = new("reject-mismatch", "text box 1 in main story",
            e => Set(Para(e.Body, "Box text."), "Box text, edited."),
            e => Track(e, Para(e.Body, "Box text."), "Box text.", "Box text, edited.")),
        ["move"] = new("reject-mismatch", "main story",
            e =>
            {
                var p = Para(e.Body, "Move me.");
                p.Remove();
                Para(e.Body, "Stay here.").AddAfterSelf(p);
            },
            e =>
            {
                XElement Range(string kind)
                {
                    var start = new XElement(W + kind + "RangeStart", e.R.Mark(), new XAttribute(W + "name", "move1"));
                    return start;
                }
                XElement End(string kind, XElement start) => new(W + kind + "RangeEnd", new XAttribute(W + "id", (string)start.Attribute(W + "id")!));
                var from = Para(e.Body, "Move me.");
                from.Elements(W + "r").Remove();
                var fromStart = Range("moveFrom");
                from.Add(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "moveFrom"))),
                    fromStart, new XElement(W + "moveFrom", e.R.Mark(), Run("Move me.")), End("moveFrom", fromStart));
                var toStart = Range("moveTo");
                Para(e.Body, "Stay here.").AddAfterSelf(P(new XElement(W + "pPr", new XElement(W + "rPr", Mark(e, "moveTo"))),
                    toStart, new XElement(W + "moveTo", e.R.Mark(), Run("Move me.")), End("moveTo", toStart)));
            }),
        ["sdt-text"] = new("reject-mismatch", "main story",
            e => Set(e.Body.Descendants(W + "sdtContent").First(), "Control text, edited."),
            e => Track(e, e.Body.Descendants(W + "sdtContent").First(), "Control text.", "Control text, edited.")),
        ["block-sdt-text"] = new("reject-mismatch", "main story",
            e => Set(Para(e.Body, "Block control text."), "Block control text, edited."),
            e => Track(e, Para(e.Body, "Block control text."), "Block control text.", "Block control text, edited.")),
        ["smartTag"] = new("reject-mismatch", "main story",
            e => Set(e.Body.Descendants(W + "smartTag").Single(), "Smart text, edited."),
            e => Track(e, e.Body.Descendants(W + "smartTag").Single(), "Smart text.", "Smart text, edited."),
            Extra: P(SmartTag(R("Smart text."))),
            Note: "the Office 2019 schema no longer allows w:smartTag (Word 2010 dropped smart tags), so this fixture carries that schema error from the original"),
        ["fldSimple"] = new("reject-mismatch", "main story",
            e => Set(e.Body.Descendants(W + "fldSimple").Single(), "Date text, edited."),
            e => Track(e, e.Body.Descendants(W + "fldSimple").Single(), "Date text.", "Date text, edited.")),
        ["fldSimple-instruction"] = new("reject-mismatch", "main story",
            e => e.Body.Descendants(W + "fldSimple").Single().SetAttributeValue(W + "instr", " TIME "),
            null,
            Note: "w:fldSimple cannot sit inside w:ins or w:del (CT_RunTrackChange), so its instruction cannot be changed with tracking"),
        // Revision sites Word writes that the cases above do not: evidence for the allowlist.
        ["ins-then-del"] = new("", "", null,
            e => Para(e.Body, "Last paragraph.").Add(e.R.Ins(e.R.Del(" gone"))),
            Note: "tracked-only: Word's deletion of an earlier insertion"),
        ["formatted-insertion"] = new("", "", null,
            e => Para(e.Body, "Last paragraph.").Add(e.R.Ins(new XElement(W + "r",
                new XElement(W + "rPr", new XElement(W + "b"), PropertyChange(e, "rPrChange", new XElement(W + "rPr"))),
                new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), " bold")))),
            Note: "tracked-only: a formatting change on inserted text"),
        ["formatted-deletion"] = new("", "", null,
            e =>
            {
                var p = Para(e.Body, "Last paragraph.");
                p.Elements(W + "r").Remove();
                p.Add(new XElement(W + "del", e.R.Mark(), new XElement(W + "r",
                        new XElement(W + "rPr", new XElement(W + "b"), PropertyChange(e, "rPrChange", new XElement(W + "rPr"))),
                        new XElement(W + "delText", "Last paragraph."))),
                    e.R.Ins(Run("Final paragraph.")));
            },
            Note: "tracked-only: deleted text that had a tracked formatting change"),
    };

    private static XElement NumPr() => new(W + "numPr",
        new XElement(W + "ilvl", new XAttribute(W + "val", 0)), new XElement(W + "numId", new XAttribute(W + "val", 1)));

    private static XElement Shading() => new(W + "shd",
        new XAttribute(W + "val", "clear"), new XAttribute(W + "color", "auto"), new XAttribute(W + "fill", "FF0000"));

    public static TheoryData<string> Names => [.. Cases.Keys];

    private static List<Finding> Errors(IEnumerable<Finding> findings) => [.. findings.Where(f => f.Severity == Severity.Error)];

    private static string Show(IEnumerable<Finding> findings) =>
        string.Join("\n", findings.Select(f => $"{f.Severity} {f.Rule}: {f.Message} [{f.Part} {f.Location}]"));

    private string Original(Case c) => c.Extra is null ? _original : Rich(Path.Combine(_r.Dir, "rich-extra.docx"), c.Extra);

    /// <summary>Errors the redline adds: schema errors the original already has are not the redline's.</summary>
    private static List<Finding> NewErrors(string original, IEnumerable<Finding> findings)
    {
        var existing = Docx.Validate(original).Select(f => f.Message).ToHashSet(StringComparer.Ordinal);
        return [.. Errors(findings).Where(f => f.Rule != "schema" || !existing.Contains(f.Message))];
    }

    [Fact]
    public void Fixture_is_schema_valid_and_verifies_against_itself()
    {
        Assert.Empty(Docx.Validate(_original));
        var findings = RevisionCheck.Run(_original, _r.Edit("copy", _ => { }, from: _original));
        Assert.True(findings.All(f => f.Rule == "accept-outline"), Show(findings));
    }

    [Fact]
    public void Every_allowlist_entry_is_exercised_by_a_passing_tracked_case()
    {
        // The allowlist's evidence: each site and context occurs in a tracked case that
        // Tracked_change_passes shows to reject back to the original with zero errors.
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (name, c) in Cases.Where(c => c.Value.Tracked is not null && c.Value.TrackedRule is null))
            used.UnionWith(RevisionAllowlist.SitesUsed(File.ReadAllBytes(_r.Edit("sites-" + name, c.Tracked!, from: Original(c)))));

        Assert.Empty(RevisionAllowlist.Sites.Except(used).Order(StringComparer.Ordinal));
    }

    [Theory]
    [MemberData(nameof(Names))]
    public void Untracked_change_is_an_error(string name)
    {
        var c = Cases[name];
        if (c.Untracked is null) Assert.Skip($"no untracked form: {c.Note}");
        var original = Original(c);
        var redline = _r.Edit(name, c.Untracked, from: original);
        if (c.DocumentUnchanged) Assert.Equal(Entry(original, "word/document.xml"), Entry(redline, "word/document.xml"));

        var findings = RevisionCheck.Run(original, redline);

        var errors = NewErrors(original, findings);
        Assert.True(errors.Any(f => f.Rule == c.Rule
                                    && ((f.Location ?? "").Contains(c.Story, StringComparison.Ordinal) || f.Message.Contains(c.Story, StringComparison.Ordinal))),
            $"expected {c.Rule} naming {c.Story}; got\n{Show(findings)}");
        Assert.DoesNotContain(errors, f => f.Rule is "schema" or "unverifiable-revision");
    }

    [Theory]
    [MemberData(nameof(Names))]
    public void Tracked_change_passes(string name)
    {
        var c = Cases[name];
        if (c.Tracked is null) Assert.Skip($"no tracked form: {c.Note}");
        var original = Original(c);
        var redline = _r.Edit(name, c.Tracked, from: original);

        var findings = RevisionCheck.Run(original, redline);

        var errors = NewErrors(original, findings);
        if (c.TrackedRule is null)
        {
            Assert.True(errors.Count == 0, Show(findings));
        }
        else
        {
            // Fails closed: the unverifiable revision is reported, next to whatever the view shows.
            Assert.Contains(errors, f => f.Rule == c.TrackedRule);
            Assert.True(errors.All(f => f.Rule == c.TrackedRule || f.Rule == c.Rule), Show(findings));
        }
        Assert.DoesNotContain(findings, f => f.Rule == "untracked-change");
    }

    private static byte[] Entry(string path, string name)
    {
        using var zip = ZipFile.OpenRead(path);
        using var s = zip.GetEntry(name)!.Open();
        using var m = new MemoryStream();
        s.CopyTo(m);
        return m.ToArray();
    }
}
