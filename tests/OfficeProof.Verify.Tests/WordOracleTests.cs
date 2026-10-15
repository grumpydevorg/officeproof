using System.IO.Compression;
using System.Security.Cryptography;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using static OfficeProof.TestKit.DocxBuilder;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

/// <summary>
/// Drives the real Microsoft Word. Opt-in: set OFFICEPROOF_WORD_TESTS=1 on a Mac with Word
/// installed and an unlocked GUI session.
/// </summary>
public sealed class WordOracleTests : IDisposable
{
    private readonly Redlines _r = new();

    public void Dispose() => _r.Dispose();

    private static void RequireWord() =>
        Assert.SkipUnless(Environment.GetEnvironmentVariable("OFFICEPROOF_WORD_TESTS") == "1" && WordOracle.IsAvailable,
            "set OFFICEPROOF_WORD_TESTS=1 on macOS with Microsoft Word installed");

    private static List<(string, string)> Pairs(IEnumerable<OutlineEntry> outline) =>
        [.. outline.Select(e => (e.ListNumber, e.Text))];

    [Theory]
    [InlineData(OutlineView.AcceptAll)]
    [InlineData(OutlineView.RejectAll)]
    [InlineData(OutlineView.AsIs)]
    public void Word_and_clippit_outlines_agree(OutlineView view)
    {
        RequireWord();
        var redline = _r.NumberedHeadings("numbered");

        var word = WordOracle.Outline(redline, view, TimeSpan.FromMinutes(2));

        Assert.Equal(Pairs(Outline.Read(redline, view)), Pairs(word));
    }

    [Fact]
    public void Word_paragraph_positions_count_table_row_ends()
    {
        RequireWord();
        XElement Cell(string text) => new(W + "tc", new XElement(W + "p", Run(text)));
        var path = new DocxBuilder()
            .Raw(new XElement(W + "tbl",
                new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 2000)), new XElement(W + "gridCol", new XAttribute(W + "w", 2000))),
                new XElement(W + "tr", Cell("a"), Cell("b")),
                new XElement(W + "tr", Cell("c"), Cell("d"))))
            .Heading("Head1", "After table")
            .Save(Path.Combine(_r.Dir, "table.docx"));

        var entry = Assert.Single(WordOracle.Outline(path, OutlineView.AsIs, TimeSpan.FromMinutes(2)));

        Assert.Equal("After table", entry.Text);
        Assert.Equal(0, entry.Level);
    }

    [Fact]
    public void Word_reports_broken_references_after_updating_fields()
    {
        RequireWord();
        var redline = _r.Make("broken-ref", body: b => Paragraph(b, "Aims text.").Add(_r.Ins(
            new XElement(W + "fldSimple", new XAttribute(W + "instr", " REF _Ref999 \\h "), Run("Figure 9")))));

        var report = WordOracle.UpdateFields(redline, OutlineView.AcceptAll, TimeSpan.FromMinutes(2));

        Assert.Contains(report.ErrorResults, l => l.StartsWith("Aims text.", StringComparison.Ordinal));
        Assert.Contains(report.TocLines, l => l.Contains("Summary", StringComparison.Ordinal));
    }

    [Fact]
    public void Word_views_refuse_to_write_over_their_input()
    {
        Assert.Throws<ArgumentException>(() => WordOracle.SaveView(_r.Original, OutlineView.RejectAll, _r.Original));
        Assert.Throws<ArgumentException>(() => WordOracle.ExportPdf(_r.Original, OutlineView.AsIs, _r.Original));
    }

    [Fact]
    public void Word_reject_all_of_a_redline_reads_as_word_resave_of_its_original()
    {
        RequireWord();
        var redline = _r.Make("replace", body: b =>
        {
            _r.TrackedReplacePrefix(Paragraph(b, "Aims text."), "Aims", "Goals");
            Paragraph(b, "Summary text.").Add(_r.Ins(Run(" Added.")));
        });
        var resaved = Path.Combine(_r.Dir, "original-word.docx");
        var rejected = Path.Combine(_r.Dir, "replace-word-reject.docx");
        var inputBytes = File.ReadAllBytes(redline);

        WordOracle.SaveView(_r.Original, OutlineView.AsIs, resaved, TimeSpan.FromMinutes(2));
        WordOracle.SaveView(redline, OutlineView.RejectAll, rejected, TimeSpan.FromMinutes(2));

        Assert.Equal(inputBytes, File.ReadAllBytes(redline));
        Assert.Equal(StoryDump.Of(File.ReadAllBytes(resaved)), StoryDump.Of(File.ReadAllBytes(rejected)));
    }

    /// <summary>
    /// Measured 2026-10-15 (Word 16.113): two Word saves of one revision-free document, and its
    /// saves after Reject All and Accept All, differ only in what <see cref="WordSave.Canonical"/>
    /// strips. Word saves are not idempotent across generations (a resave of a Word save dropped
    /// a w:tblPrEx the first save had added), so compare only saves of the same generation.
    /// </summary>
    [Fact]
    public void Word_saves_of_one_document_differ_only_in_bookkeeping()
    {
        RequireWord();
        var clean = WordSave.RichFixture(Path.Combine(_r.Dir, "rich.docx"));
        var saves = new Dictionary<string, OutlineView>
        {
            ["resave-1"] = OutlineView.AsIs, ["resave-2"] = OutlineView.AsIs,
            ["reject"] = OutlineView.RejectAll, ["accept"] = OutlineView.AcceptAll,
        };
        var canonical = saves.ToDictionary(s => s.Key, s =>
        {
            var path = Path.Combine(_r.Dir, s.Key + ".docx");
            WordOracle.SaveView(clean, s.Value, path, TimeSpan.FromMinutes(2));
            return WordSave.Canonical(File.ReadAllBytes(path));
        });

        var expected = canonical["resave-1"];
        foreach (var (name, parts) in canonical)
        {
            Assert.Equal(expected.Keys, parts.Keys);
            foreach (var (part, content) in expected)
                Assert.True(content == parts[part], $"{name}: {part} differs from resave-1 beyond bookkeeping");
        }
    }

    [Fact]
    public void Word_pdf_of_one_document_twice_renders_identically()
    {
        RequireWord();
        PageDiffTests.RequirePdftoppm();
        var a = Path.Combine(_r.Dir, "a.pdf");
        var b = Path.Combine(_r.Dir, "b.pdf");

        WordOracle.ExportPdf(_r.Original, OutlineView.AsIs, a, TimeSpan.FromMinutes(2));
        WordOracle.ExportPdf(_r.Original, OutlineView.AsIs, b, TimeSpan.FromMinutes(2));

        Assert.True(PageDiff.Compare(a, b).Identical);
    }

    [Fact]
    public void Word_pdf_after_reject_all_matches_the_original_and_after_accept_all_does_not()
    {
        RequireWord();
        PageDiffTests.RequirePdftoppm();
        // The redline also carries a comment: its anchor would shift justified text, which is
        // why ExportPdf deletes comments first.
        var redline = _r.Make("pdf", body: b =>
        {
            Paragraph(b, "Aims text.").Add(_r.Ins(Run(" Added words.")));
            var commented = Paragraph(b, "Summary text.");
            commented.AddFirst(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 0)));
            commented.Add(new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 0)),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0))));
        }, package: doc => Docx.Save(doc.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>(), new XDocument(
            new XElement(W + "comments", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                new XElement(W + "comment", new XAttribute(W + "id", 0), new XAttribute(W + "author", "review"),
                    P(R("A review note.")))))));
        var original = Path.Combine(_r.Dir, "original.pdf");
        var rejected = Path.Combine(_r.Dir, "rejected.pdf");
        var accepted = Path.Combine(_r.Dir, "accepted.pdf");

        WordOracle.ExportPdf(_r.Original, OutlineView.AsIs, original, TimeSpan.FromMinutes(2));
        WordOracle.ExportPdf(redline, OutlineView.RejectAll, rejected, TimeSpan.FromMinutes(2));
        WordOracle.ExportPdf(redline, OutlineView.AcceptAll, accepted, TimeSpan.FromMinutes(2));

        Assert.True(PageDiff.Compare(original, rejected).Identical);
        var changed = Assert.Single(PageDiff.Compare(original, accepted).Differing);
        Assert.NotNull(changed.Box);
    }
}

/// <summary>What Word rewrites on every save, and a fixture that exercises it.</summary>
internal static class WordSave
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace W14 = Ns.W14;
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace App = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    /// <summary>
    /// Every part of a Word-saved package (XML as text, other parts as SHA-256) without what Word
    /// regenerates on each save: w14:paraId, w14:textId and w14:anchorId, every w:rsid*
    /// attribute, settings' w:rsids and w14:docId, the comment paragraph ids and durable ids in
    /// commentsExtended (w15:paraId) and commentsIds (w16cid:paraId, w16cid:durableId), and the
    /// save stamps in docProps (cp:revision, cp:lastModifiedBy, dcterms:created,
    /// dcterms:modified, TotalTime).
    /// </summary>
    public static SortedDictionary<string, string> Canonical(byte[] package)
    {
        var parts = new SortedDictionary<string, string>(StringComparer.Ordinal);
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            if (!entry.FullName.EndsWith(".xml", StringComparison.Ordinal) && !entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            {
                parts[entry.FullName] = Convert.ToHexString(SHA256.HashData(buffer.ToArray()));
                continue;
            }
            buffer.Position = 0;
            var xml = XDocument.Load(buffer);
            foreach (var e in xml.Descendants().ToList())
            {
                if (e.Name == W + "rsids" || e.Name == W14 + "docId" || e.Name == Cp + "revision" || e.Name == Cp + "lastModifiedBy"
                    || e.Name == DcTerms + "created" || e.Name == DcTerms + "modified" || e.Name == App + "TotalTime")
                {
                    e.Remove();
                    continue;
                }
                e.Attributes().Where(IsBookkeeping).Remove();
            }
            parts[entry.FullName] = xml.ToString(SaveOptions.DisableFormatting);
        }
        return parts;
    }

    private static bool IsBookkeeping(XAttribute a) =>
        (a.Name.Namespace == W && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
        || (a.Name.Namespace == W14 && a.Name.LocalName is "paraId" or "textId" or "anchorId")
        || (a.Name.Namespace == W15 && a.Name.LocalName == "paraId")
        || (a.Name.Namespace == W16Cid && a.Name.LocalName is "paraId" or "durableId");

    /// <summary>A revision-free document with a header, footer, footnote, comment, hyperlink, VML text box and table.</summary>
    public static string RichFixture(string path)
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        return new DocxBuilder()
            .Heading("Head1", "1.0 Summary").Para("Body text.")
            .Header("rIdHeader", P(R("Header text"))).Footer("rIdFooter", P(R("Footer text")))
            .Raw(P(R("Noted"), NoteReference(1))).Footnote(1, P(R("Footnote text")))
            .Raw(P(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 0)), R("Commented"),
                new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 0)),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)))))
            .Comment(0, "Comment text")
            .Link("rIdLink", "https://example.com/")
            .Raw(P(new XElement(W + "hyperlink", new XAttribute(Ns.R + "id", "rIdLink"), R("link text"))))
            .Raw(P(R("Box: "), new XElement(W + "r", new XElement(W + "pict", new XElement(v + "rect",
                new XAttribute(XNamespace.Xmlns + "v", v.NamespaceName), new XAttribute("style", "width:200pt;height:60pt"),
                new XElement(v + "textbox", new XElement(W + "txbxContent", P(R("Box text")))))))))
            .Raw(new XElement(W + "tbl",
                new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 0), new XAttribute(W + "type", "auto"))),
                new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 4000))),
                new XElement(W + "tr", new XElement(W + "tc", P(R("Cell text"))))))
            .Para("Body after.")
            .Save(path);
    }
}

/// <summary>
/// Manual smoke run on real documents that must never enter the repository. Opt-in: set
/// OFFICEPROOF_SMOKE_ORIGINAL and OFFICEPROOF_SMOKE_REDLINE (and OFFICEPROOF_WORD_TESTS=1 to add the
/// Word comparison); the report is written to OFFICEPROOF_SMOKE_OUT or the temp directory.
/// </summary>
public sealed class RealDocumentSmoke
{
    [Fact]
    public void Revision_check_on_real_documents()
    {
        var original = Environment.GetEnvironmentVariable("OFFICEPROOF_SMOKE_ORIGINAL");
        var redline = Environment.GetEnvironmentVariable("OFFICEPROOF_SMOKE_REDLINE");
        Assert.SkipWhen(string.IsNullOrEmpty(original) || string.IsNullOrEmpty(redline), "set OFFICEPROOF_SMOKE_ORIGINAL and OFFICEPROOF_SMOKE_REDLINE");
        var outPath = Environment.GetEnvironmentVariable("OFFICEPROOF_SMOKE_OUT") ?? Path.Combine(Path.GetTempPath(), "officeproof-verify-smoke.txt");

        using var output = new StringWriter();
        var started = DateTime.UtcNow;
        var findings = RevisionCheck.Run(original!, redline!);
        output.WriteLine($"RevisionCheck.Run ({(DateTime.UtcNow - started).TotalSeconds:0.0} s)");
        FindingReport.WriteText(output, findings);

        if (Environment.GetEnvironmentVariable("OFFICEPROOF_WORD_TESTS") == "1" && WordOracle.IsAvailable)
        {
            foreach (var view in new[] { OutlineView.RejectAll, OutlineView.AcceptAll })
            {
                var clippit = Outline.Read(redline!, view);
                output.WriteLine($"--- {view}: Clippit vs Word");
                try
                {
                    started = DateTime.UtcNow;
                    var word = WordOracle.Outline(redline!, view, TimeSpan.FromMinutes(5));
                    output.WriteLine($"Word outline ({(DateTime.UtcNow - started).TotalSeconds:0.0} s, {word.Count} entries)");
                    var disagreements = 0;
                    foreach (var (c, w) in clippit.Zip(word))
                    {
                        var same = c.ListNumber == w.ListNumber && c.Text == w.Text;
                        if (!same) disagreements++;
                        output.WriteLine($"{(same ? "  " : "!!")} #{c.Index} clippit [{c.ListNumber}] {c.Text} | word [{w.ListNumber}] L{w.Level} {w.Text}");
                    }
                    output.WriteLine($"{disagreements} disagreement(s)");
                }
                catch (WordOracleException e)
                {
                    output.WriteLine("WordOracle failed: " + e.Message);
                    break;
                }
            }
            try
            {
                started = DateTime.UtcNow;
                var fields = WordOracle.UpdateFields(redline!, OutlineView.AcceptAll, TimeSpan.FromMinutes(5));
                output.WriteLine($"--- Word update fields, Accept All ({(DateTime.UtcNow - started).TotalSeconds:0.0} s)");
                foreach (var l in fields.TocLines) output.WriteLine("TOC  " + l);
                foreach (var l in fields.ErrorResults) output.WriteLine("ERR  " + l);
            }
            catch (WordOracleException e)
            {
                output.WriteLine("WordOracle.UpdateFields failed: " + e.Message);
            }
            WordViewsAndPages(output, original!, redline!);
        }
        File.WriteAllText(outPath, output.ToString());
    }

    // Word's resave of the original against Word's Reject All of the redline, as stories and as
    // pages. The derived copies live in a temporary folder that is deleted afterwards.
    private static void WordViewsAndPages(StringWriter output, string original, string redline)
    {
        var dir = Path.Combine(Path.GetTempPath(), "officeproof-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var timeout = TimeSpan.FromMinutes(10);
            string Timed(string what, Action action)
            {
                var clock = System.Diagnostics.Stopwatch.StartNew();
                action();
                return $"{what}: {clock.Elapsed.TotalSeconds:0.0} s";
            }
            var resaved = Path.Combine(dir, "original-resave.docx");
            var rejected = Path.Combine(dir, "redline-reject.docx");
            var originalPdf = Path.Combine(dir, "original.pdf");
            var rejectedPdf = Path.Combine(dir, "redline-reject.pdf");
            output.WriteLine("--- Word views and pages");
            output.WriteLine(Timed("SaveView original AsIs", () => WordOracle.SaveView(original, OutlineView.AsIs, resaved, timeout)));
            output.WriteLine(Timed("SaveView redline RejectAll", () => WordOracle.SaveView(redline, OutlineView.RejectAll, rejected, timeout)));
            var resavedStories = StoryDump.Of(File.ReadAllBytes(resaved)).Split('\n');
            var rejectedStories = StoryDump.Of(File.ReadAllBytes(rejected)).Split('\n');
            var onlyResaved = resavedStories.Except(rejectedStories).ToList();
            var onlyRejected = rejectedStories.Except(resavedStories).ToList();
            output.WriteLine($"story lines: {resavedStories.Length} in Word resave, {rejectedStories.Length} in Word Reject All; {onlyResaved.Count} only in the resave, {onlyRejected.Count} only in Reject All");
            foreach (var line in onlyResaved.Take(20)) output.WriteLine("-  " + line);
            foreach (var line in onlyRejected.Take(20)) output.WriteLine("+  " + line);
            output.WriteLine(Timed("ExportPdf original AsIs", () => WordOracle.ExportPdf(original, OutlineView.AsIs, originalPdf, timeout)));
            output.WriteLine(Timed("ExportPdf redline RejectAll", () => WordOracle.ExportPdf(redline, OutlineView.RejectAll, rejectedPdf, timeout)));
            var clock = System.Diagnostics.Stopwatch.StartNew();
            var pages = PageDiff.Compare(originalPdf, rejectedPdf);
            output.WriteLine($"PageDiff: {clock.Elapsed.TotalSeconds:0.0} s, pages {pages.PagesA} / {pages.PagesB}, identical: {pages.Identical}");
            foreach (var p in pages.Differing)
                output.WriteLine($"!! page {p.Page}: {p.ChangedPixels} px ({p.ChangedFraction:P3}), box {p.Box}");
        }
        catch (Exception e) when (e is WordOracleException or InvalidOperationException or InvalidDataException)
        {
            output.WriteLine("Word views and pages failed: " + e.Message);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
}
