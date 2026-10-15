// Container probe, recorded 2026-10-15: Microsoft Word for Mac 16.113.3 (16.113.26092714) on
// macOS 26.6.2. Each fixture holds "Alpha [-beta-][+delta+] gamma bold" with a
// tracked bold on "bold" (w:rPrChange) inside one container. Word's views come from
// WordOracle.SaveView (Word's own Reject/Accept All Changes, then save); "Clippit" is
// RevisionProcessor on the unmodified bytes; "ViewBytes" is Outline.ViewBytes as of this date
// (ClippitWorkarounds: w:hyperlink and w:fldSimple carried through Clippit as marker pairs; the
// reject view puts the redline's styles and numbering back). Clippit columns are for the pinned
// 3.9.2-rmorgans.2 (3.9.1 plus our fixes, see eng/pack-clippit.sh); published 3.9.1 gave K/A for
// w:hyperlink, w:fldSimple, w:dir and w:bdo, U/U for comment text, and threw accepting a
// w:cellDel after a cell without w:tcPr. Outcomes on the probed paragraph:
//   R  rejected: "Alpha beta gamma bold", plain
//   A  accepted: "Alpha delta gamma bold", "bold" bold
//   U  untouched: every revision left in place
//   K  deletion dropped, insertion kept ("Alpha delta gamma bold"); the bold is rejected
//   X  refused: Word shows a modal dialog on open that blocks automation (WordOracle refuses first)
//
// | container                  | Word rej | Word acc | Clippit rej | Clippit acc | W=C rej | W=C acc | ViewBytes rej | ViewBytes acc |
// |----------------------------|----------|----------|-------------|-------------|---------|---------|---------------|---------------|
// | header                     | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | footer                     | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | footnote                   | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | endnote                    | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | comment text               | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | text box (VML)             | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:hyperlink                | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:sdt, run level           | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:sdt, block level         | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:smartTag                 | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:fldSimple                | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:customXml, run level     | X        | X        | K           | A           | -       | -       | K             | A             |
// | table cell                 | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | moveFrom/moveTo (+ edits)  | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | deleted paragraph mark     | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:dir (rtl), run level     | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:bdo (rtl), run level     | R        | A        | R           | A           | yes     | yes     | R             | A             |
// | w:hyperlink in run w:sdt   | R        | A        | R           | A           | yes     | yes     | R             | A             |
//
// Word rejects and accepts tracked changes in headers and footers like any other story, and
// records them too: with Track Changes on, text set and inserted through AppleScript in the
// primary header and footer was saved as w:ins/w:del in header1.xml and footer1.xml (same
// date). Word's document-wide Reject/Accept All reaches every story above, although its
// "revisions" collection counts the main story only.
// Wrappers in Word's saved views: w:hyperlink and w:sdt kept, w:smartTag dropped (content
// kept), w:fldSimple rewritten as a complex field (fldChar), the VML text box kept as VML.
// Word keeps w:dir and w:bdo on save and ends their content with an added run holding U+202C
// POP DIRECTIONAL FORMATTING. Clippit's K for w:customXml is RevisionProcessor's
// ReverseRevisionsTransform: it reverses w:ins/w:del only when their nearest ancestor other than
// w:sdt, w:sdtContent, w:smartTag, w:hyperlink, w:fldSimple, w:dir and w:bdo is w:p.
//
// Revision kinds (RevisionKindProbeTests, same date and versions; KindDump of the main story:
// text, ListItemRetriever list numbers, table structure). Outcomes:
//   R  as Word's Reject All;  A  as Word's Accept All (in a reject column: rejected as accepted)
//   G  short row: the inserted cell removed without widening the cell before it (Word gives that
//      cell gridSpan 2)
//   T  throws NullReferenceException (3.9.1's AcceptDeletedCellsTransform read the w:tcPr of the
//      cell before the deleted one, which has none)
//
// | kind                                   | Word rej | Word acc | Clippit rej | Clippit acc | W=C rej | W=C acc | ViewBytes rej | ViewBytes acc | allowlist today |
// |----------------------------------------|----------|----------|-------------|-------------|---------|---------|---------------|---------------|-----------------|
// | w:numberingChange in w:numPr           | R        | A        | R           | A           | yes     | yes     | R             | A             | allowed         |
// | w:pPrChange of ilvl and numId          | R        | A        | R           | A           | yes     | yes     | R             | A             | allowed         |
// | w:cellIns                              | R        | A        | G           | A           | no      | yes     | G             | A             | refused         |
// | w:cellDel                              | R        | A        | R           | A           | yes     | yes     | R             | A             | refused         |
// | w:cellDel, cell before has no w:tcPr   | R        | A        | R           | A           | yes     | yes     | R             | A             | refused         |
// | w:cellMerge (vMerge rest/cont)         | R        | A        | A           | A           | no      | yes     | A             | A             | refused         |
// | w:tblPrExChange (row jc)               | R        | A        | R           | A           | yes     | yes     | R             | A             | allowed         |
// | w:del (author B) in w:ins (author A)   | R        | A        | R           | A           | yes     | yes     | R             | A             | allowed         |
//
// Word's list strings (WordOracle.Outline) equal ListItemRetriever's on the ViewBytes views:
// "1. 2. 3." for the numberingChange list in both views (Word and Clippit drop the markup and
// leave the numbering alone), "1. 2. 3. 4." rejected and "1. 1.1. a) 2." accepted for the
// numPr change. Word opened every fixture without a dialog. Its plain resave keeps
// w:numberingChange (rewriting w:original "%1)" as "%1:1:0:)"), w:cellIns, w:cellDel and
// w:cellMerge; for w:tblPrExChange it copies the row's w:jc into w:trPr and adds a w:tcPrChange
// to each cell of the row and a w:tblGridChange. A w:numberingChange in w:pPr/w:rPr is not
// schema-valid (CT_ParaRPr has none) and was not shown to Word.
using System.Text;
using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.TestKit.DocxBuilder;

namespace OfficeProof.Verify.Tests;

/// <summary>
/// Which containers of a tracked change Word and Clippit accept and reject alike (table above).
/// Each view is reduced to a <see cref="StoryDump"/> and compared with the recorded outcome. The
/// Clippit and ViewBytes tests run everywhere, so a Clippit upgrade or a ViewBytes change that
/// alters an agreement fails the suite; the Word test (OFFICEPROOF_WORD_TESTS=1) catches a Word
/// upgrade.
/// </summary>
public sealed class ContainerProbeTests : IDisposable
{
    internal enum Outcome { Rejected, Accepted, Untouched, InsertionKept, Refused }

    private sealed record Row(string Container, Outcome WordReject, Outcome WordAccept,
        Outcome ClippitReject, Outcome ClippitAccept, Outcome ViewReject, Outcome ViewAccept);

    private const Outcome R = Outcome.Rejected, A = Outcome.Accepted, U = Outcome.Untouched, K = Outcome.InsertionKept, X = Outcome.Refused;

    private static readonly Row[] Recorded =
    [
        new("header", R, A, R, A, R, A),
        new("footer", R, A, R, A, R, A),
        new("footnote", R, A, R, A, R, A),
        new("endnote", R, A, R, A, R, A),
        new("comment", R, A, R, A, R, A),
        new("textbox", R, A, R, A, R, A),
        new("hyperlink", R, A, R, A, R, A),
        new("sdt-run", R, A, R, A, R, A),
        new("sdt-block", R, A, R, A, R, A),
        new("smarttag", R, A, R, A, R, A),
        new("fldsimple", R, A, R, A, R, A),
        new("customxml", X, X, K, A, K, A),
        new("table-cell", R, A, R, A, R, A),
        new("move", R, A, R, A, R, A),
        new("deleted-paragraph-mark", R, A, R, A, R, A),
        new("dir", R, A, R, A, R, A),
        new("bdo", R, A, R, A, R, A),
        new("sdt-hyperlink", R, A, R, A, R, A),
    ];

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "officeproof-container-probe", Guid.NewGuid().ToString("N"));

    public ContainerProbeTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); }
        catch (IOException) { }
    }

    public static TheoryData<string, OutlineView> Cases()
    {
        var data = new TheoryData<string, OutlineView>();
        foreach (var row in Recorded)
        {
            data.Add(row.Container, OutlineView.RejectAll);
            data.Add(row.Container, OutlineView.AcceptAll);
        }
        return data;
    }

    [Fact]
    public void Every_probe_container_is_recorded() =>
        Assert.Equal(ProbeFixture.Containers, Recorded.Select(r => r.Container));

    [Theory]
    [MemberData(nameof(Cases))]
    public void Clippit_matches_the_recorded_outcome(string container, OutlineView view)
    {
        var row = Recorded.Single(r => r.Container == container);
        var bytes = File.ReadAllBytes(Fixture(container));

        var document = new WmlDocument("in.docx", bytes);
        var result = view == OutlineView.RejectAll ? RevisionProcessor.RejectRevisions(document) : RevisionProcessor.AcceptRevisions(document);

        Assert.Equal(ProbeFixture.Expected(container, view, Pick(view, row.ClippitReject, row.ClippitAccept)), StoryDump.Of(result.DocumentByteArray));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void View_bytes_match_the_recorded_outcome(string container, OutlineView view)
    {
        var row = Recorded.Single(r => r.Container == container);
        var bytes = File.ReadAllBytes(Fixture(container));

        Assert.Equal(ProbeFixture.Expected(container, view, Pick(view, row.ViewReject, row.ViewAccept)), StoryDump.Of(Outline.ViewBytes(bytes, view)));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Word_matches_the_recorded_outcome(string container, OutlineView view)
    {
        Assert.SkipUnless(Environment.GetEnvironmentVariable("OFFICEPROOF_WORD_TESTS") == "1" && WordOracle.IsAvailable,
            "set OFFICEPROOF_WORD_TESTS=1 on macOS with Microsoft Word installed");
        var row = Recorded.Single(r => r.Container == container);
        var outcome = Pick(view, row.WordReject, row.WordAccept);
        var path = Fixture(container);
        var saved = Path.Combine(_dir, $"{container}-{view}-word.docx");

        if (outcome == Outcome.Refused)
        {
            var e = Assert.Throws<WordOracleException>(() => WordOracle.SaveView(path, view, saved, TimeSpan.FromMinutes(2)));
            Assert.Contains("custom XML", e.Message, StringComparison.Ordinal);
            return;
        }
        WordOracle.SaveView(path, view, saved, TimeSpan.FromMinutes(2));

        Assert.Equal(ProbeFixture.Expected(container, view, outcome, wordSave: true), StoryDump.Of(File.ReadAllBytes(saved)));
    }

    private static Outcome Pick(OutlineView view, Outcome reject, Outcome accept) => view == OutlineView.RejectAll ? reject : accept;

    private string Fixture(string container) => ProbeFixture.Build(container, Path.Combine(_dir, container + ".docx"));
}

/// <summary>
/// Probe fixtures: "Body before.", the container, "Body after.". Inside the container the
/// paragraph reads "Alpha [-beta-][+delta+] gamma bold" with "bold" made bold as a tracked
/// formatting change, so Reject All should give "Alpha beta gamma bold" (plain) and Accept All
/// "Alpha delta gamma «b»bold«/»".
/// </summary>
internal sealed class ProbeFixture
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private int _id = 100;

    public static readonly string[] Containers =
    [
        "header", "footer", "footnote", "endnote", "comment", "textbox", "hyperlink", "sdt-run", "sdt-block",
        "smarttag", "fldsimple", "customxml", "table-cell", "move", "deleted-paragraph-mark", "dir", "bdo", "sdt-hyperlink",
    ];

    public static string Build(string container, string path) => new ProbeFixture().Make(container, path);

    /// <summary>
    /// The <see cref="StoryDump"/> of <paramref name="container"/>'s fixture in
    /// <paramref name="view"/> when the probed paragraph came out as <paramref name="outcome"/>.
    /// <paramref name="wordSave"/>: as Word saves it, which ends the content of w:dir and w:bdo
    /// with a run holding U+202C POP DIRECTIONAL FORMATTING.
    /// </summary>
    public static string Expected(string container, OutlineView view, ContainerProbeTests.Outcome outcome, bool wordSave = false)
    {
        var x = outcome switch
        {
            ContainerProbeTests.Outcome.Rejected => "Alpha beta gamma bold",
            ContainerProbeTests.Outcome.Accepted => "Alpha delta gamma «b»bold«/»",
            ContainerProbeTests.Outcome.Untouched => "Alpha [-beta-][+delta+] gamma «b,Δ»bold«/»",
            ContainerProbeTests.Outcome.InsertionKept => "Alpha delta gamma bold",
            _ => throw new ArgumentOutOfRangeException(nameof(outcome), outcome, "no dump for this outcome"),
        };
        if (wordSave && container is "dir" or "bdo") x += "\u202C";
        var reject = view == OutlineView.RejectAll;
        string[] lines = container switch
        {
            "header" or "footer" => ["main: Body before.", "main: Body after.", $"{container} default (section 1): {x}"],
            "footnote" or "endnote" => ["main: Body before.", $"main: Noted[{container}]", "main: Body after.", $"{container} 1: {x}"],
            "comment" => ["main: Body before.", "main: Commented", "main: Body after.", $"comment 1: {x}"],
            "textbox" => ["main: Body before.", "main: Box: ", "main: Body after.", $"main textbox 1: {x}"],
            "sdt-run" or "sdt-hyperlink" => ["main: Body before.", $"main: Before {x}", "main: Body after."],
            "move" => ["main: Body before.", $"main: Source: {(reject ? x : "")}", "main: Body middle.", $"main: Target: {(reject ? "" : x)}", "main: Body after."],
            "deleted-paragraph-mark" => reject
                ? ["main: Body before.", $"main: {x}", "main: Next paragraph.", "main: Body after."]
                : ["main: Body before.", $"main: {x}Next paragraph.", "main: Body after."],
            _ => ["main: Body before.", $"main: {x}", "main: Body after."],
        };
        return string.Join("\n", lines);
    }

    private object[] Mark() => [new XAttribute(W + "id", _id++), new XAttribute(W + "author", "probe"), new XAttribute(W + "date", "2026-10-15T00:00:00Z")];

    private XElement Del(string text) => new(W + "del", Mark(),
        new XElement(W + "r", new XElement(W + "delText", new XAttribute(Ns.Xml + "space", "preserve"), text)));

    private XElement Ins(string text) => new(W + "ins", Mark(), R(text));

    private XElement TrackedBold(string text) => new(W + "r",
        new XElement(W + "rPr", new XElement(W + "b"), new XElement(W + "rPrChange", Mark(), new XElement(W + "rPr"))),
        new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text));

    private object[] Edits() => [R("Alpha "), Del("beta"), Ins("delta"), R(" gamma "), TrackedBold("bold")];

    private string Make(string container, string path)
    {
        var b = new DocxBuilder().Para("Body before.");
        Action<WordprocessingDocument>? post = null;
        switch (container)
        {
            case "header":
                b.Header("rIdProbeHeader", P(Edits()));
                break;
            case "footer":
                b.Footer("rIdProbeFooter", P(Edits()));
                break;
            case "footnote":
                b.Raw(P(R("Noted"), NoteReference(1))).Footnote(1, P(Edits()));
                break;
            case "endnote":
                b.Raw(P(R("Noted"), NoteReference(1, endnote: true))).Endnote(1, P(Edits()));
                break;
            case "comment":
                b.Raw(P(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 0)), R("Commented"),
                        new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 0)),
                        new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)))))
                    .Comment(0, "placeholder");
                post = doc =>
                {
                    var part = doc.MainDocumentPart!.WordprocessingCommentsPart!;
                    var xml = Docx.Load(part);
                    xml.Root!.Element(W + "comment")!.Element(W + "p")!.ReplaceWith(P(Edits()));
                    Docx.Save(part, xml);
                };
                break;
            case "textbox":
                b.Raw(P(R("Box: "), new XElement(W + "r", new XElement(W + "pict",
                    new XElement(V + "rect", new XAttribute(XNamespace.Xmlns + "v", V.NamespaceName),
                        new XAttribute("style", "width:200pt;height:60pt"),
                        new XElement(V + "textbox", new XElement(W + "txbxContent", P(Edits()))))))));
                break;
            case "hyperlink":
                b.Link("rIdProbeLink", "https://example.com/")
                    .Raw(P(new XElement(W + "hyperlink", new XAttribute(Ns.R + "id", "rIdProbeLink"), Edits())));
                break;
            case "sdt-run":
                b.Raw(P(R("Before "), new XElement(W + "sdt",
                    new XElement(W + "sdtPr", new XElement(W + "id", new XAttribute(W + "val", 11))),
                    new XElement(W + "sdtContent", Edits()))));
                break;
            case "sdt-block":
                b.Raw(new XElement(W + "sdt",
                    new XElement(W + "sdtPr", new XElement(W + "id", new XAttribute(W + "val", 12))),
                    new XElement(W + "sdtContent", P(Edits()))));
                break;
            case "smarttag":
                b.Raw(P(new XElement(W + "smartTag", new XAttribute(W + "uri", "urn:schemas-microsoft-com:office:smarttags"),
                    new XAttribute(W + "element", "place"), Edits())));
                break;
            case "fldsimple":
                b.Raw(P(new XElement(W + "fldSimple", new XAttribute(W + "instr", " QUOTE probe "), Edits())));
                break;
            case "customxml":
                b.Raw(P(new XElement(W + "customXml", new XAttribute(W + "element", "tag"), Edits())));
                break;
            case "table-cell":
                b.Raw(new XElement(W + "tbl",
                    new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 0), new XAttribute(W + "type", "auto"))),
                    new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 4000))),
                    new XElement(W + "tr", new XElement(W + "tc", P(Edits())))));
                break;
            case "move":
                // "Alpha beta gamma bold" moves from the first paragraph to the second; the moved
                // text is edited at its destination.
                int from = _id++, to = _id++;
                b.Raw(P(R("Source: "), RangeMark("moveFromRangeStart", from, "probeMove"),
                        new XElement(W + "moveFrom", Mark(), R("Alpha beta gamma "), R("bold")),
                        RangeMark("moveFromRangeEnd", from)))
                    .Para("Body middle.")
                    .Raw(P(R("Target: "), RangeMark("moveToRangeStart", to, "probeMove"),
                        new XElement(W + "moveTo", Mark(), Edits()),
                        RangeMark("moveToRangeEnd", to)));
                break;
            case "deleted-paragraph-mark":
                b.Raw(P(new XElement(W + "pPr", new XElement(W + "rPr", new XElement(W + "del", Mark()))), Edits()))
                    .Para("Next paragraph.");
                break;
            case "dir":
                b.Raw(P(new XElement(W + "dir", new XAttribute(W + "val", "rtl"), Edits())));
                break;
            case "bdo":
                b.Raw(P(new XElement(W + "bdo", new XAttribute(W + "val", "rtl"), Edits())));
                break;
            case "sdt-hyperlink":
                b.Link("rIdProbeLink", "https://example.com/")
                    .Raw(P(R("Before "), Sdt(13, new XElement(W + "hyperlink", new XAttribute(Ns.R + "id", "rIdProbeLink"), Edits()))));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(container), container, "unknown probe container");
        }
        b.Para("Body after.").Save(path);
        if (post is not null)
        {
            using var doc = WordprocessingDocument.Open(path, isEditable: true);
            post(doc);
        }
        return path;
    }

    private static XElement RangeMark(string name, int id, string? moveName = null)
    {
        var e = new XElement(W + name, new XAttribute(W + "id", id));
        if (moveName is not null)
            e.Add(new XAttribute(W + "author", "probe"), new XAttribute(W + "date", "2026-10-15T00:00:00Z"), new XAttribute(W + "name", moveName));
        return e;
    }
}

/// <summary>
/// Every story of a package as text, one paragraph per line: "story: text". Formatting that
/// differs from plain shows as «b,i»text«/»; revision markup left in the view shows as
/// [+inserted+], [-deleted-], «...Δ» (formatting change) and a leading ¶+ / ¶- / ¶Δ on the
/// paragraph mark. Stories: main, text boxes (DrawingML choice; VML fallbacks skipped),
/// headers and footers per section reference, footnotes, endnotes and comments in reference
/// order. Field instructions are not text; field results are.
/// </summary>
internal static class StoryDump
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    public static string Of(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var main = doc.MainDocumentPart ?? throw new InvalidDataException("no main document part");
        var lines = new List<string>();
        var body = Docx.Load(main).Root!.Element(W + "body")!;
        Story("main", body);

        var section = 0;
        foreach (var sectPr in Own(body, W + "sectPr"))
        {
            section++;
            foreach (var reference in sectPr.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference"))
            {
                var kind = reference.Name.LocalName.Replace("Reference", "", StringComparison.Ordinal);
                var part = main.GetPartById((string)reference.Attribute(Ns.R + "id")!);
                Story($"{kind} {(string?)reference.Attribute(W + "type")} (section {section})", Docx.Load(part).Root!);
            }
        }
        Notes("footnote", main.FootnotesPart);
        Notes("endnote", main.EndnotesPart);
        Notes("comment", main.WordprocessingCommentsPart);
        return string.Join("\n", lines);

        void Notes(string kind, OpenXmlPart? part)
        {
            if (part is null) return;
            var root = Docx.Load(part).Root!;
            var referenceName = W + (kind == "comment" ? "commentReference" : kind + "Reference");
            var k = 0;
            foreach (var reference in Own(body, referenceName))
            {
                var id = (string?)reference.Attribute(W + "id");
                var note = root.Elements(W + kind).FirstOrDefault(n => (string?)n.Attribute(W + "id") == id);
                Story($"{kind} {++k}", note ?? new XElement(W + "missing"));
            }
        }

        void Story(string name, XElement root)
        {
            foreach (var p in Own(root, W + "p")) lines.Add($"{name}: {Paragraph(p)}");
            var box = 0;
            foreach (var content in Own(root, W + "txbxContent")) Story($"{name} textbox {++box}", content);
        }
    }

    // Descendants named name that belong to root's own story: not inside a nested text box
    // and not in a markup-compatibility fallback (Word writes each text box twice).
    private static IEnumerable<XElement> Own(XElement root, XName name) =>
        root.Descendants(name).Where(e => e.Ancestors().TakeWhile(a => a != root)
            .All(a => a.Name != W + "txbxContent" && a.Name != Mc + "Fallback"));

    internal static string Paragraph(XElement p)
    {
        var text = new StringBuilder();
        var pPr = p.Element(W + "pPr");
        if (pPr?.Element(W + "rPr")?.Element(W + "ins") is not null) text.Append("¶+ ");
        if (pPr?.Element(W + "rPr")?.Element(W + "del") is not null) text.Append("¶- ");
        if (pPr?.Element(W + "pPrChange") is not null) text.Append("¶Δ ");

        var segments = new List<(string Key, StringBuilder Text)>();
        foreach (var run in p.Descendants(W + "r").Where(r => r.Ancestors(W + "p").First() == p
                     && !r.Ancestors().TakeWhile(a => a != p).Any(a => a.Name == Mc + "Fallback")))
        {
            var runText = RunText(run);
            if (runText.Length == 0) continue;
            var key = RunKey(run, p);
            if (segments.Count > 0 && segments[^1].Key == key) segments[^1].Text.Append(runText);
            else segments.Add((key, new StringBuilder(runText)));
        }
        foreach (var (key, segment) in segments)
        {
            var parts = key.Split('|');
            var (revision, format) = (parts[0], parts[1]);
            var s = format.Length == 0 ? segment.ToString() : $"«{format}»{segment}«/»";
            text.Append(revision switch { "+" => $"[+{s}+]", "-" => $"[-{s}-]", _ => s });
        }
        return text.ToString();
    }

    private static string RunKey(XElement run, XElement p)
    {
        var wrappers = run.Ancestors().TakeWhile(a => a != p).Select(a => a.Name.LocalName).ToList();
        var revision = wrappers.Any(n => n is "del" or "moveFrom") ? "-" : wrappers.Any(n => n is "ins" or "moveTo") ? "+" : "";
        var rPr = run.Element(W + "rPr");
        var flags = new List<string>();
        if (rPr is not null)
        {
            if ((string?)rPr.Element(W + "rStyle")?.Attribute(W + "val") is { } style) flags.Add("style=" + style);
            foreach (var toggle in new[] { "b", "i", "strike", "dstrike", "caps", "smallCaps", "vanish" })
                if (On(rPr.Element(W + toggle))) flags.Add(toggle);
            if ((string?)rPr.Element(W + "u")?.Attribute(W + "val") is { } u && u != "none") flags.Add("u");
            if ((string?)rPr.Element(W + "vertAlign")?.Attribute(W + "val") is { } va && va != "baseline") flags.Add(va);
            if (rPr.Element(W + "rPrChange") is not null) flags.Add("Δ");
        }
        return revision + "|" + string.Join(",", flags);
    }

    private static bool On(XElement? toggle) =>
        toggle is not null && (string?)toggle.Attribute(W + "val") is null or "1" or "true" or "on";

    private static string RunText(XElement run) => string.Concat(run.Elements().Select(e => e.Name.LocalName switch
    {
        "t" or "delText" => e.Value,
        "tab" => "\t",
        "br" or "cr" => "\n",
        "noBreakHyphen" => "-",
        "footnoteReference" => "[footnote]",
        "endnoteReference" => "[endnote]",
        _ => "",
    }));
}

/// <summary>
/// Revision kinds outside the container probe's paragraph: list numbering, table cells, row
/// exceptions and a deletion by another author inside an insertion (second table above). Each view
/// is reduced to a <see cref="KindDump"/> (text, list numbers, table structure) and compared with
/// the recorded outcome; the Word tests (OFFICEPROOF_WORD_TESTS=1) also compare Word's list strings
/// with Clippit's ListItemRetriever. The allowlist test records which probed sites
/// <see cref="RevisionAllowlist"/> refuses today.
/// </summary>
public sealed class RevisionKindProbeTests : IDisposable
{
    internal enum Outcome { Rejected, Accepted, ShortRow, Throws }

    private sealed record Row(string Kind, Outcome WordReject, Outcome WordAccept,
        Outcome ClippitReject, Outcome ClippitAccept, Outcome ViewReject, Outcome ViewAccept);

    private const Outcome R = Outcome.Rejected, A = Outcome.Accepted, G = Outcome.ShortRow, T = Outcome.Throws;

    private static readonly Row[] Recorded =
    [
        new("numbering-change", R, A, R, A, R, A),
        new("numpr-change", R, A, R, A, R, A),
        new("cell-ins", R, A, G, A, G, A),
        new("cell-del", R, A, R, A, R, A),
        new("cell-del-no-tcpr", R, A, R, A, R, A),
        new("cell-merge", R, A, A, A, A, A),
        new("tblprex-change", R, A, R, A, R, A),
        new("del-in-ins-other-author", R, A, R, A, R, A),
    ];

    /// <summary>
    /// The allowlist keys (<c>w:site</c> or <c>w:site inside context</c> of an
    /// unverifiable-revision finding) each fixture produces with today's
    /// <see cref="RevisionAllowlist.Sites"/>; empty: the fixture passes the allowlist.
    /// </summary>
    private static readonly Dictionary<string, string[]> Refused = new(StringComparer.Ordinal)
    {
        ["numbering-change"] = [],
        ["numpr-change"] = [],
        ["cell-ins"] = ["cellIns in tc/tcPr"],
        ["cell-del"] = ["cellDel in tc/tcPr"],
        ["cell-del-no-tcpr"] = ["cellDel in tc/tcPr"],
        ["cell-merge"] = ["cellMerge in tc/tcPr"],
        ["tblprex-change"] = [],
        ["del-in-ins-other-author"] = [],
        ["dir"] = ["del in dir", "ins in dir", "rPrChange in r/rPr inside dir"],
        ["bdo"] = ["del in bdo", "ins in bdo", "rPrChange in r/rPr inside bdo"],
        ["sdt-hyperlink"] = [],
    };

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "officeproof-kind-probe", Guid.NewGuid().ToString("N"));

    public RevisionKindProbeTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); }
        catch (IOException) { }
    }

    public static TheoryData<string, OutlineView> Cases()
    {
        var data = new TheoryData<string, OutlineView>();
        foreach (var row in Recorded)
        {
            data.Add(row.Kind, OutlineView.RejectAll);
            data.Add(row.Kind, OutlineView.AcceptAll);
        }
        return data;
    }

    public static TheoryData<string, OutlineView> NumberedCases() =>
        [("numbering-change", OutlineView.RejectAll), ("numbering-change", OutlineView.AcceptAll),
         ("numpr-change", OutlineView.RejectAll), ("numpr-change", OutlineView.AcceptAll)];

    public static TheoryData<string> AllFixtures() => [.. KindFixture.Kinds.Concat(ProbeFixture.Containers)];

    public static TheoryData<string> AllowlistCases() => [.. Refused.Keys];

    [Fact]
    public void Every_probe_kind_is_recorded() =>
        Assert.Equal(KindFixture.Kinds, Recorded.Select(r => r.Kind));

    // The Open XML SDK 3.5 validator rejects w:smartTag as a child of w:p under both Office 2007
    // and Office 2019 ("invalid child element ... smartTag"), although ECMA-376 CT_P admits it
    // (EG_PContent) and Word 16.113 opens, accepts and rejects it without a dialog; that one
    // fixture is checked for every other finding.
    [Theory]
    [MemberData(nameof(AllFixtures))]
    public void Every_probe_fixture_is_schema_valid(string name) =>
        Assert.DoesNotContain(Docx.Validate(Build(name)), f => !(name == "smarttag" && f.Message.Contains(":smartTag'", StringComparison.Ordinal)));

    [Fact]
    public void Numbering_change_in_paragraph_mark_run_properties_is_not_schema_valid()
    {
        // CT_ParaRPr has no w:numberingChange; it belongs in w:numPr (CT_NumPr), as Word writes it.
        var path = Build("numbering-change");
        using (var doc = WordprocessingDocument.Open(path, isEditable: true))
        {
            var main = doc.MainDocumentPart!;
            var xml = Docx.Load(main);
            var change = xml.Descendants(Ns.W + "numberingChange").Single();
            var pPr = change.Parent!.Parent!;
            change.Remove();
            pPr.Add(new XElement(Ns.W + "rPr", change));
            Docx.Save(main, xml);
        }
        Assert.Contains(Docx.Validate(path), f => f.Message.Contains("numberingChange", StringComparison.Ordinal));
    }

    [Theory]
    [MemberData(nameof(AllowlistCases))]
    public void Allowlist_refuses_the_recorded_sites(string name)
    {
        var keys = RevisionAllowlist.Check(File.ReadAllBytes(Build(name)), "probe")
            .Select(f => System.Text.RegularExpressions.Regex.Match(f.Message, @"tracked change\(s\) w:(.+?) in /").Groups[1].Value)
            .Order(StringComparer.Ordinal);
        Assert.Equal(Refused[name], keys);
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Clippit_matches_the_recorded_outcome(string kind, OutlineView view)
    {
        var row = Recorded.Single(r => r.Kind == kind);
        var outcome = Pick(view, row.ClippitReject, row.ClippitAccept);
        var document = new WmlDocument("in.docx", File.ReadAllBytes(Build(kind)));
        byte[] Run() => (view == OutlineView.RejectAll ? RevisionProcessor.RejectRevisions(document) : RevisionProcessor.AcceptRevisions(document)).DocumentByteArray;

        if (outcome == Outcome.Throws)
        {
            Assert.Throws<NullReferenceException>(Run);
            return;
        }
        Assert.Equal(KindFixture.Expected(kind, view, outcome), KindDump.Of(Run()));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void View_bytes_match_the_recorded_outcome(string kind, OutlineView view)
    {
        var row = Recorded.Single(r => r.Kind == kind);
        var outcome = Pick(view, row.ViewReject, row.ViewAccept);
        var bytes = File.ReadAllBytes(Build(kind));

        if (outcome == Outcome.Throws)
        {
            Assert.Throws<NullReferenceException>(() => Outline.ViewBytes(bytes, view));
            return;
        }
        Assert.Equal(KindFixture.Expected(kind, view, outcome), KindDump.Of(Outline.ViewBytes(bytes, view)));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Word_matches_the_recorded_outcome(string kind, OutlineView view)
    {
        SkipUnlessWord();
        var row = Recorded.Single(r => r.Kind == kind);
        var saved = Path.Combine(_dir, $"{kind}-{view}-word.docx");
        WordOracle.SaveView(Build(kind), view, saved, TimeSpan.FromMinutes(2));

        Assert.Equal(KindFixture.Expected(kind, view, Pick(view, row.WordReject, row.WordAccept)), KindDump.Of(File.ReadAllBytes(saved)));
    }

    [Theory]
    [MemberData(nameof(NumberedCases))]
    public void Word_list_strings_match_list_item_retriever(string kind, OutlineView view)
    {
        SkipUnlessWord();
        var path = Build(kind);
        var expected = KindFixture.Expected(kind, view, view == OutlineView.RejectAll ? R : A).Split('\n').Where(l => l.StartsWith('['));

        var word = WordOracle.Outline(path, view, TimeSpan.FromMinutes(2));

        Assert.Equal(expected, word.Select(e => $"[{e.ListNumber}] {e.Text}"));
        Assert.Equal(Outline.Read(path, view).Select(e => (e.Index, e.ListNumber, e.Text)), word.Select(e => (e.Index, e.ListNumber, e.Text)));
    }

    private static void SkipUnlessWord() =>
        Assert.SkipUnless(Environment.GetEnvironmentVariable("OFFICEPROOF_WORD_TESTS") == "1" && WordOracle.IsAvailable,
            "set OFFICEPROOF_WORD_TESTS=1 on macOS with Microsoft Word installed");

    private static Outcome Pick(OutlineView view, Outcome reject, Outcome accept) => view == OutlineView.RejectAll ? reject : accept;

    private string Build(string name) => KindFixture.Kinds.Contains(name)
        ? KindFixture.Build(name, Path.Combine(_dir, name + ".docx"))
        : ProbeFixture.Build(name, Path.Combine(_dir, name + ".docx"));
}

/// <summary>
/// Fixtures of the kind probe: "Body before.", the construct, "Body after.". Lists use numId 1
/// (decimal, "%1." then "%1.%2.") and numId 2 (lower letter, "%1)"); tables have two columns
/// and, as Word writes them, a w:tcPr with w:tcW in every cell ("cell-del-no-tcpr" leaves it out).
/// </summary>
internal sealed class KindFixture
{
    private static readonly XNamespace W = Ns.W;
    private int _id = 200;

    public static readonly string[] Kinds =
    [
        "numbering-change", "numpr-change", "cell-ins", "cell-del", "cell-del-no-tcpr", "cell-merge", "tblprex-change",
        "del-in-ins-other-author",
    ];

    public static string Build(string kind, string path) => new KindFixture().Make(kind, path);

    /// <summary>The <see cref="KindDump"/> of <paramref name="kind"/>'s fixture in <paramref name="view"/> when it came out as <paramref name="outcome"/>.</summary>
    public static string Expected(string kind, OutlineView view, RevisionKindProbeTests.Outcome outcome)
    {
        var accepted = outcome switch
        {
            RevisionKindProbeTests.Outcome.Rejected => false,
            RevisionKindProbeTests.Outcome.Accepted => true,
            RevisionKindProbeTests.Outcome.ShortRow when kind == "cell-ins" && view == OutlineView.RejectAll => false,
            _ => throw new ArgumentOutOfRangeException(nameof(outcome), outcome, "no dump for this outcome"),
        };
        string[] lines = kind switch
        {
            "numbering-change" => ["[1.] One", "[2.] Two", "[3.] Three"],
            "numpr-change" => accepted
                ? ["[1.] One", "[1.1.] Two", "[a)] Three", "[2.] Four"]
                : ["[1.] One", "[2.] Two", "[3.] Three", "[4.] Four"],
            "cell-ins" => ["row 1: A1 | B1", accepted ? "row 2: A2 | B2"
                : outcome == RevisionKindProbeTests.Outcome.ShortRow ? "row 2: A2" : "row 2: (span=2) A2"],
            "cell-del" or "cell-del-no-tcpr" => ["row 1: A1 | B1", accepted ? "row 2: (span=2) A2" : "row 2: A2 | B2"],
            "cell-merge" => accepted
                ? ["row 1: (vMerge=restart) A1 | B1", "row 2: (vMerge=continue)  | B2"]
                : ["row 1: A1 | B1", "row 2:  | B2"],
            "tblprex-change" => ["row 1: A1 | B1", accepted ? "row 2 jc=center: A2 | B2" : "row 2: A2 | B2"],
            "del-in-ins-other-author" => [accepted ? "Alpha delta gamma" : "Alpha gamma"],
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "unknown probe kind"),
        };
        return string.Join("\n", ["Body before.", .. lines, "Body after."]);
    }

    private object[] Mark(string author = "probe") => [new XAttribute(W + "id", _id++), new XAttribute(W + "author", author), new XAttribute(W + "date", "2026-10-15T00:00:00Z")];

    private static XElement NumPr(int ilvl, int numId, params object[] extra) => new(W + "numPr",
        new XElement(W + "ilvl", new XAttribute(W + "val", ilvl)), new XElement(W + "numId", new XAttribute(W + "val", numId)), extra);

    private static XElement ListP(string text, XElement numPr, params object[] pPrExtra) => P(new XElement(W + "pPr", numPr, pPrExtra), R(text));

    private XElement PPrChange(XElement oldNumPr) => new(W + "pPrChange", Mark(), new XElement(W + "pPr", oldNumPr));

    private static XElement Table(params XElement[] rows) => new(W + "tbl",
        new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 0), new XAttribute(W + "type", "auto"))),
        new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 2000)), new XElement(W + "gridCol", new XAttribute(W + "w", 2000))),
        rows);

    private static XElement Row(params object[] content) => new(W + "tr", content);

    private static XElement Cell(string text, params XElement[] tcPr) => new(W + "tc",
        new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 0), new XAttribute(W + "type", "auto")), tcPr),
        text.Length > 0 ? P(R(text)) : P());

    private static XElement BareCell(string text) => new(W + "tc", P(R(text)));

    private static XElement VMerge(string? val) => val is null ? new XElement(W + "vMerge") : new XElement(W + "vMerge", new XAttribute(W + "val", val));

    private string Make(string kind, string path)
    {
        var b = new DocxBuilder().Para("Body before.");
        var numbered = false;
        switch (kind)
        {
            case "numbering-change":
                numbered = true;
                b.Raw(ListP("One", NumPr(0, 1)))
                    .Raw(ListP("Two", NumPr(0, 1, new XElement(W + "numberingChange", Mark(), new XAttribute(W + "original", "%1)")))))
                    .Raw(ListP("Three", NumPr(0, 1)));
                break;
            case "numpr-change":
                numbered = true;
                b.Raw(ListP("One", NumPr(0, 1)))
                    .Raw(ListP("Two", NumPr(1, 1), PPrChange(NumPr(0, 1))))
                    .Raw(ListP("Three", NumPr(0, 2), PPrChange(NumPr(0, 1))))
                    .Raw(ListP("Four", NumPr(0, 1)));
                break;
            case "cell-ins":
                b.Raw(Table(Row(Cell("A1"), Cell("B1")), Row(Cell("A2"), Cell("B2", new XElement(W + "cellIns", Mark())))));
                break;
            case "cell-del":
                b.Raw(Table(Row(Cell("A1"), Cell("B1")), Row(Cell("A2"), Cell("B2", new XElement(W + "cellDel", Mark())))));
                break;
            case "cell-del-no-tcpr":
                b.Raw(Table(Row(BareCell("A1"), BareCell("B1")), Row(BareCell("A2"), Cell("B2", new XElement(W + "cellDel", Mark())))));
                break;
            case "cell-merge":
                // A1 and A2 merged vertically as a tracked change; A2's content was empty.
                b.Raw(Table(
                    Row(Cell("A1", VMerge("restart"), new XElement(W + "cellMerge", Mark(), new XAttribute(W + "vMerge", "rest"))), Cell("B1")),
                    Row(Cell("", VMerge(null), new XElement(W + "cellMerge", Mark(), new XAttribute(W + "vMerge", "cont"))), Cell("B2"))));
                break;
            case "tblprex-change":
                b.Raw(Table(Row(Cell("A1"), Cell("B1")),
                    Row(new XElement(W + "tblPrEx", new XElement(W + "jc", new XAttribute(W + "val", "center")),
                            new XElement(W + "tblPrExChange", Mark(), new XElement(W + "tblPrEx"))),
                        Cell("A2"), Cell("B2"))));
                break;
            case "del-in-ins-other-author":
                b.Raw(P(R("Alpha "), new XElement(W + "ins", Mark("author A"), R("delta "),
                        new XElement(W + "del", Mark("author B"), new XElement(W + "r",
                            new XElement(W + "delText", new XAttribute(Ns.Xml + "space", "preserve"), "beta ")))),
                    R("gamma")));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "unknown probe kind");
        }
        b.Para("Body after.").Save(path);
        if (numbered)
        {
            using var doc = WordprocessingDocument.Open(path, isEditable: true);
            Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(Numbering()));
        }
        return path;
    }

    private static XElement Numbering()
    {
        static XElement Lvl(int ilvl, string format, string text) => new(W + "lvl", new XAttribute(W + "ilvl", ilvl),
            new XElement(W + "start", new XAttribute(W + "val", 1)),
            new XElement(W + "numFmt", new XAttribute(W + "val", format)),
            new XElement(W + "lvlText", new XAttribute(W + "val", text)),
            new XElement(W + "lvlJc", new XAttribute(W + "val", "left")));
        return new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 0),
                new XElement(W + "multiLevelType", new XAttribute(W + "val", "multilevel")),
                Lvl(0, "decimal", "%1."), Lvl(1, "decimal", "%1.%2.")),
            new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 1),
                new XElement(W + "multiLevelType", new XAttribute(W + "val", "singleLevel")),
                Lvl(0, "lowerLetter", "%1)")),
            new XElement(W + "num", new XAttribute(W + "numId", 1), new XElement(W + "abstractNumId", new XAttribute(W + "val", 0))),
            new XElement(W + "num", new XAttribute(W + "numId", 2), new XElement(W + "abstractNumId", new XAttribute(W + "val", 1))));
    }
}

/// <summary>
/// The main story as text with list numbers and table structure, one line per paragraph or
/// table row. Paragraphs read "[list number] text" (list number from Clippit's
/// ListItemRetriever; <see cref="StoryDump"/> notation for formatting and revisions left in; a
/// leading #Δ for a w:numberingChange left in). A row reads "row n{flags}: cell | cell"; row
/// flags: jc=… (row alignment from w:tblPrEx or w:trPr, which Word copies it to), Δ (w:tblPrExChange
/// left in), row+ / row-; cell flags in parentheses: span, vMerge, +cell / -cell / merge(new/orig)
/// for cell revisions left in, Δ for w:tcPrChange. Other table properties (Word adds tcW, tblInd,
/// tblCellMar) are not shown.
/// </summary>
internal static class KindDump
{
    private static readonly XNamespace W = Ns.W;

    public static string Of(byte[] package)
    {
        var numbers = Outline.ParagraphsOfView(package);
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var root = Docx.Load(doc.MainDocumentPart!).Root!;
        var order = Outline.MainStoryParagraphs(root).ToList();
        var lines = new List<string>();
        foreach (var block in root.Element(W + "body")!.Elements())
        {
            if (block.Name == W + "p") lines.Add(Para(block));
            else if (block.Name == W + "tbl")
            {
                var r = 0;
                foreach (var tr in block.Elements(W + "tr"))
                    lines.Add($"row {++r}{RowFlags(tr)}: " + string.Join(" | ", tr.Elements(W + "tc").Select(tc =>
                        CellFlags(tc) + string.Join(" / ", tc.Elements(W + "p").Select(Para)))));
            }
        }
        return string.Join("\n", lines);

        string Para(XElement p)
        {
            var number = numbers[order.IndexOf(p)].ListNumber;
            var numberingChange = p.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numberingChange") is not null;
            return (number.Length > 0 ? $"[{number}] " : "") + (numberingChange ? "#Δ " : "") + StoryDump.Paragraph(p);
        }
    }

    private static string RowFlags(XElement tr)
    {
        var flags = new List<string>();
        var ex = tr.Element(W + "tblPrEx");
        var trPr = tr.Element(W + "trPr");
        if ((string?)(ex?.Element(W + "jc") ?? trPr?.Element(W + "jc"))?.Attribute(W + "val") is { } jc) flags.Add("jc=" + jc);
        if (ex?.Element(W + "tblPrExChange") is not null) flags.Add("Δ");
        if (trPr?.Element(W + "ins") is not null) flags.Add("row+");
        if (trPr?.Element(W + "del") is not null) flags.Add("row-");
        return flags.Count == 0 ? "" : " " + string.Join(" ", flags);
    }

    private static string CellFlags(XElement tc)
    {
        var tcPr = tc.Element(W + "tcPr");
        if (tcPr is null) return "";
        var flags = new List<string>();
        if ((string?)tcPr.Element(W + "gridSpan")?.Attribute(W + "val") is { } span) flags.Add("span=" + span);
        if (tcPr.Element(W + "vMerge") is { } vm) flags.Add("vMerge=" + ((string?)vm.Attribute(W + "val") ?? "continue"));
        if (tcPr.Element(W + "cellIns") is not null) flags.Add("+cell");
        if (tcPr.Element(W + "cellDel") is not null) flags.Add("-cell");
        if (tcPr.Element(W + "cellMerge") is { } cm)
            flags.Add($"merge({(string?)cm.Attribute(W + "vMerge")}/{(string?)cm.Attribute(W + "vMergeOrig")})");
        if (tcPr.Element(W + "tcPrChange") is not null) flags.Add("Δ");
        return flags.Count == 0 ? "" : $"({string.Join(",", flags)}) ";
    }
}
