using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.TestKit.DocxBuilder;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

/// <summary>Stories, their pairing, relationship resolution and the Clippit round trip.</summary>
public sealed class StoryViewTests : IDisposable
{
    private static readonly XNamespace W = Ns.W;
    private readonly Redlines _r = new();
    private readonly string _original;

    public StoryViewTests() => _original = PlantedDefectTests.Rich(Path.Combine(_r.Dir, "rich.docx"));

    public void Dispose() => _r.Dispose();

    private static List<Finding> Errors(IEnumerable<Finding> findings) => [.. findings.Where(f => f.Severity == Severity.Error)];

    private static string Show(IEnumerable<Finding> findings) =>
        string.Join("\n", findings.Select(f => $"{f.Severity} {f.Rule}: {f.Message} [{f.Part} {f.Location}]"));

    private static XElement Link(PackageEdit e) => e.Body.Descendants(W + "hyperlink").First(h => h.Attribute(Ns.R + "id") is not null);

    private static XElement MainXml(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        return Docx.Load(doc.MainDocumentPart!).Root!;
    }

    [Fact]
    public void Every_story_is_loaded_with_text_boxes_counted_once()
    {
        var view = StoryView.Load(File.ReadAllBytes(_original));

        Assert.Equal(
            ["main story", "text box 1 in main story", "text box 2 in main story", "header (section 2, default)", "footer (section 2, default)",
             "footnote separator", "footnote continuationSeparator", "footnote 1", "endnote separator", "endnote continuationSeparator", "endnote 1"],
            view.Stories.Select(s => s.Name));
        Assert.Equal("Shape text.", Assert.Single(view.Stories.Single(s => s.Name == "text box 2 in main story").Paragraphs).Entry.Text);
        Assert.Equal(["comment 1"], view.Comments.Select(s => s.Name));
        Assert.Empty(view.Problems);
    }

    [Fact]
    public void Story_comparison_works_on_any_package_bytes()
    {
        // The Word side feeds Word-saved copies straight in, without Clippit.
        var bytes = File.ReadAllBytes(_original);
        var copy = File.ReadAllBytes(_r.Edit("copy", e => _ = e.Body, from: _original));

        Assert.Empty(StoryComparison.Compare(StoryView.Load(bytes), StoryView.Load(copy)));
    }

    [Fact]
    public void List_numbers_come_from_clippit_in_headers_too()
    {
        var path = new DocxBuilder()
            .Header("rIdH1", P(new XElement(W + "pPr", new XElement(W + "numPr",
                new XElement(W + "ilvl", new XAttribute(W + "val", 0)), new XElement(W + "numId", new XAttribute(W + "val", 1)))), R("Numbered header line")))
            .Para("Body.")
            .Save(Path.Combine(_r.Dir, "numbered-header.docx"));
        var numbered = _r.Make("numbered-header-with-numbering", from: path,
            package: doc => Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(Numbering())));

        var header = StoryView.Load(File.ReadAllBytes(numbered)).Stories.Single(s => s.Name == "header (section 1, default)");

        Assert.Equal("1.0", Assert.Single(header.Paragraphs).Entry.ListNumber);
    }

    [Fact]
    public void Headers_pair_by_section_and_type_not_by_part_name()
    {
        var redline = _r.Edit("header-moved", e =>
        {
            var old = e.Main.HeaderParts.Single();
            var moved = e.Main.AddNewPart<HeaderPart>("rIdH9");
            Docx.Save(moved, Docx.Load(old));
            e.Body.Descendants(W + "headerReference").Single().SetAttributeValue(Ns.R + "id", "rIdH9");
            e.Main.DeletePart(old);
        }, from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.True(Errors(findings).Count == 0, Show(findings));
    }

    [Fact]
    public void Header_moved_to_another_reference_type_is_a_mismatch()
    {
        var redline = _r.Edit("header-first", e =>
        {
            var sectPr = e.Body.Element(W + "sectPr")!;
            sectPr.Element(W + "headerReference")!.SetAttributeValue(W + "type", "first");
            sectPr.Add(new XElement(W + "titlePg"));
        }, from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.Contains(findings, f => f.Rule == "reject-mismatch" && f.Message.Contains("loses header (section 2, default)", StringComparison.Ordinal));
        Assert.Contains(findings, f => f.Rule == "reject-mismatch" && f.Message.Contains("has header (section 2, first)", StringComparison.Ordinal));
    }

    [Fact]
    public void Unknown_header_reference_type_fails_closed()
    {
        var redline = _r.Edit("header-odd", e => e.Body.Descendants(W + "headerReference").Single().SetAttributeValue(W + "type", "odd"), from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.Contains(findings, f => f is { Rule: "unverifiable-revision", Severity: Severity.Error } && f.Message.Contains("type 'odd'", StringComparison.Ordinal));
    }

    [Fact]
    public void Undefined_footnote_is_a_mismatch()
    {
        var redline = _r.Edit("footnote-gone", e => e.Footnotes.Elements(W + "footnote").Single(n => (string?)n.Attribute(W + "id") == "1").Remove(), from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.Contains(findings, f => f.Rule == "reject-mismatch" && f.Message.Contains("footnote 1", StringComparison.Ordinal)
                                       && f.Message.Contains("no longer defined", StringComparison.Ordinal));
    }

    [Fact]
    public void Added_comment_is_info()
    {
        var redline = _r.Edit("comment-added", e =>
        {
            e.Comments.Add(new XElement(W + "comment", new XAttribute(W + "id", 2), new XAttribute(W + "author", "officeproof"),
                new XAttribute(W + "date", "2026-10-15T00:00:00Z"), P(R("Why this changed."))));
            var p = Para(e.Body, "Body text.");
            p.AddFirst(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 2)));
            p.Add(new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 2)),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 2))));
        }, from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.True(Errors(findings).Count == 0, Show(findings));
        var added = Assert.Single(findings, f => f.Rule == "comment-added");
        Assert.Equal(Severity.Info, added.Severity);
        Assert.Contains("comment 2", added.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OutlineView.RejectAll, "Original comment.")]
    [InlineData(OutlineView.AcceptAll, "Changed comment.")]
    public void Tracked_edit_inside_a_comment_is_rejected_and_accepted(OutlineView view, string text)
    {
        // ContainerProbeTests "comment text": Word and the pinned Clippit both process revisions
        // inside comments, so the views carry the comment's text in that state.
        var redline = _r.Edit("comment-tracked", e =>
        {
            var p = Para(e.Comments, "Original comment.");
            p.Elements(W + "r").Remove();
            p.Add(e.R.Del("Original comment."), e.R.Ins(Run("Changed comment.")));
        }, from: _original);

        var comment = StoryView.Load(Outline.ViewBytes(File.ReadAllBytes(redline), view)).Comments.Single(c => c.Name == "comment 1");

        Assert.Equal(text, Assert.Single(comment.Paragraphs).Entry.Text);
        var findings = RevisionCheck.Run(_original, redline);
        Assert.True(Errors(findings).Count == 0, Show(findings));
    }

    [Fact]
    public void Pure_relationship_renumbering_passes()
    {
        var redline = _r.Edit("renumbered", e =>
        {
            Link(e).SetAttributeValue(Ns.R + "id", "rIdLinkX");
            e.Body.Descendants().Attributes(Ns.R + "embed").Single(a => a.Value == "rIdImg1").Value = "rIdImgX";
            e.Rewrite("word/_rels/document.xml.rels", rels => rels
                .Replace("Id=\"rIdLink1\"", "Id=\"rIdLinkX\"", StringComparison.Ordinal)
                .Replace("Id=\"rIdImg1\"", "Id=\"rIdImgX\"", StringComparison.Ordinal));
        }, from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.True(Errors(findings).Count == 0, Show(findings));
        Assert.DoesNotContain(findings, f => f.Rule == "untracked-change");
    }

    [Theory]
    [InlineData(OutlineView.RejectAll, "https://example.com/a")]
    [InlineData(OutlineView.AcceptAll, "https://example.com/b")]
    public void Wrappers_survive_both_views_with_their_identity(OutlineView view, string url)
    {
        // A tracked retarget: the old link's text deleted, a new link's text inserted.
        var redline = _r.Edit("retarget", e =>
        {
            var link = Link(e);
            link.ReplaceNodes(e.R.Del("link text"));
            link.AddAfterSelf(PlantedDefectTests.Hyperlink("rIdLink2", e.R.Ins(Run("link text"))));
            var field = e.Body.Descendants(W + "fldSimple").Single();
            field.Elements(W + "r").Remove();
            field.Add(e.R.Del("Date text."), e.R.Ins(Run("Date text, edited.")));
        }, from: _original);

        var bytes = Outline.ViewBytes(File.ReadAllBytes(redline), view);
        var main = MainXml(bytes);
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);

        Assert.DoesNotContain(main.Descendants(W + "bookmarkStart"), b => ((string?)b.Attribute(W + "name"))!.StartsWith("_officeproof", StringComparison.Ordinal));
        var links = main.Descendants(W + "hyperlink").ToList();
        var external = Assert.Single(links, h => h.Attribute(Ns.R + "id") is not null);
        Assert.Equal("link text", string.Concat(external.Descendants(W + "t").Select(t => t.Value)));
        Assert.Equal(url, doc.MainDocumentPart!.HyperlinkRelationships.Single(r => r.Id == (string)external.Attribute(Ns.R + "id")!).Uri.OriginalString.TrimEnd('/'));
        Assert.Single(links, h => (string?)h.Attribute(W + "anchor") == "Target1");
        var field2 = Assert.Single(main.Descendants(W + "fldSimple"));
        Assert.Equal(" DATE ", (string?)field2.Attribute(W + "instr"));
        Assert.Equal(view == OutlineView.RejectAll ? "Date text." : "Date text, edited.", string.Concat(field2.Descendants(W + "t").Select(t => t.Value)));
    }

    [Fact]
    public void Hyperlink_inside_an_insertion_is_a_schema_error_and_rejects_cleanly()
    {
        // CT_RunTrackChange (w:ins) admits runs, sdt, smartTag and customXml, not w:hyperlink.
        var redline = _r.Edit("link-in-ins", e =>
            Para(e.Body, "Last paragraph.").Add(e.R.Ins(PlantedDefectTests.Hyperlink("rIdLink2", Run(" new link")))), from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.Contains(findings, f => f.Rule == "schema");
        Assert.True(Errors(findings).All(f => f.Rule == "schema"), Show(findings));
        Assert.Contains(MainXml(Outline.ViewBytes(File.ReadAllBytes(redline), OutlineView.AcceptAll)).Descendants(W + "hyperlink"),
            h => (string?)h.Attribute(Ns.R + "id") == "rIdLink2");
    }

    [Fact]
    public void Wrapper_separated_from_its_markers_fails_closed()
    {
        var marked = _r.Edit("unrestored", e => Link(e).SetAttributeValue(ClippitWorkarounds.Unrestored, "true"), from: _original);

        var problem = Assert.Single(StoryView.Load(File.ReadAllBytes(marked)).Problems);

        Assert.Equal("unverifiable-revision", problem.Rule);
        Assert.Equal(Severity.Error, problem.Severity);
    }

    // "ins in customXml": WordOracle's probe (2026-10-15) found Clippit's Reject keeps the
    // insertion and drops the deletion inside run-level w:customXml, and Word cannot be asked.
    public static TheoryData<string> UnlistedSites => ["cellDel", "cellIns", "ins in body", "customXmlInsRangeStart", "ins in customXml"];

    [Theory]
    [MemberData(nameof(UnlistedSites))]
    public void Revision_site_off_the_allowlist_is_unverifiable(string site)
    {
        var redline = _r.Edit(site.Replace(' ', '-'), e =>
        {
            switch (site)
            {
                case "cellDel":
                    e.Body.Descendants(W + "tcPr").First().Add(new XElement(W + "cellDel", e.R.Mark()));
                    break;
                case "cellIns":
                    e.Body.Descendants(W + "tcPr").First().Add(new XElement(W + "cellIns", e.R.Mark()));
                    break;
                case "ins in body":
                    e.Body.AddFirst(e.R.Ins(Run("Floating insertion.")));
                    break;
                case "ins in customXml":
                    var tagged = Para(e.Body, "Body text.");
                    var run = tagged.Element(W + "r")!;
                    run.ReplaceWith(new XElement(W + "customXml", new XAttribute(W + "element", "tag"), e.R.Del("Body text."), e.R.Ins(Run("Body text, edited."))));
                    break;
                case "customXmlInsRangeStart":
                    var p = Para(e.Body, "Body text.");
                    var start = new XElement(W + "customXmlInsRangeStart", e.R.Mark());
                    p.AddFirst(start);
                    p.Add(new XElement(W + "customXmlInsRangeEnd", new XAttribute(W + "id", (string)start.Attribute(W + "id")!)));
                    break;
            }
        }, from: _original);

        var findings = RevisionCheck.Run(_original, redline);

        Assert.Contains(findings, f => f is { Rule: "unverifiable-revision", Severity: Severity.Error }
                                       && f.Message.Contains("w:" + site.Split(' ')[0], StringComparison.Ordinal));
    }

    [Fact]
    public void Allowlisted_sites_name_their_container()
    {
        var p = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "rPr", new XElement(W + "ins"))), new XElement(W + "ins"));

        Assert.Equal("ins in pPr/rPr", RevisionAllowlist.Site(p.Element(W + "pPr")!.Element(W + "rPr")!.Element(W + "ins")!));
        Assert.Equal("ins in p", RevisionAllowlist.Site(p.Element(W + "ins")!));
        Assert.Contains("ins in pPr/rPr", RevisionAllowlist.Sites);
    }

    private static XElement Para(XElement root, string text) =>
        root.Descendants(W + "p").First(p => Wml.VisibleText(p) == text);
}
