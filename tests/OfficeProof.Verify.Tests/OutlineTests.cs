using OfficeProof.Core;
using OfficeProof.TestKit;
using System.Xml.Linq;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

public sealed class OutlineTests : IDisposable
{
    private readonly Redlines _r = new();

    public void Dispose() => _r.Dispose();

    private static List<string> Headings(IEnumerable<OutlineEntry> outline) =>
        [.. outline.Select(e => (e.ListNumber.Length > 0 ? e.ListNumber + " " : "") + e.Text)];

    [Fact]
    public void Headings_come_from_outline_levels_through_the_style_chain()
    {
        var outline = Outline.Read(_r.Original);

        Assert.Equal(
            ["Abbreviations", "1.0 Summary", "2.0 Background", "2.0 Aims", "3.0 Data", "3.1 Frog Pond", "3.2 Birds", "3.2.1 Detail", "Appendix 1 Species"],
            Headings(outline));
        Assert.Equal([0, 0, 0, 0, 0, 1, 1, 2, 0], outline.Select(e => e.Level));
        Assert.All(outline, e => Assert.Equal("", e.ListNumber));
    }

    [Fact]
    public void Accept_all_gives_list_numbers_and_reject_all_gives_the_typed_original()
    {
        var redline = _r.NumberedHeadings("numbered");

        Assert.Equal(
            ["Abbreviations", "1.0 Summary", "2.0 Background", "2.1 Aims", "3.0 Data", "3.1 Frog Pond", "3.2 Birds", "3.2.1 Detail", "Appendix 1 Species"],
            Headings(Outline.Read(redline, OutlineView.AcceptAll)));
        Assert.Equal(Headings(Outline.Read(_r.Original)), Headings(Outline.Read(redline, OutlineView.RejectAll)));
        Assert.All(Outline.Read(redline, OutlineView.RejectAll), e => Assert.Equal("", e.ListNumber));
    }

    [Fact]
    public void As_is_view_shows_insertions_and_hides_deletions()
    {
        var redline = _r.Make("typed", body: b => _r.TrackedReplacePrefix(Paragraph(b, "2.0 Aims"), "2.0", "2.1"));

        Assert.Contains(Outline.Read(redline), e => e.Text == "2.1 Aims");
        Assert.Contains(Outline.Read(redline, OutlineView.RejectAll), e => e.Text == "2.0 Aims");
    }

    [Fact]
    public void Reject_all_keeps_revisions_in_style_definitions_like_word()
    {
        // Client redline v2 (2026-10-15): numPr added to Head1 as a tracked pPrChange in styles.xml. Word's
        // Reject All leaves the style numbered, so the restored typed number follows the list
        // number ("2.0 1.0 Summary"); the reject view must show the same.
        var redline = _r.NumberedHeadings("style-numbered", body: _ => { });
        using (var doc = Docx.OpenCopy(redline, Path.Combine(_r.Dir, "v2.docx")))
        {
            var part = doc.MainDocumentPart!.StyleDefinitionsPart!;
            var styles = Docx.Load(part);
            var pPr = Style(styles.Root!, "Head1").Element(W + "pPr")!;
            pPr.Add(new XElement(W + "numPr",
                new XElement(W + "ilvl", new XAttribute(W + "val", 0)),
                new XElement(W + "numId", new XAttribute(W + "val", 1))));
            pPr.Add(new XElement(W + "pPrChange", _r.Mark(), new XElement(W + "pPr", new XElement(W + "keepNext"), new XElement(W + "outlineLvl", new XAttribute(W + "val", 0)))));
            Docx.Save(part, styles);
        }

        var rejected = Outline.Read(Path.Combine(_r.Dir, "v2.docx"), OutlineView.RejectAll);

        Assert.Contains(rejected, e => e.Text == "1.0 Summary" && e.ListNumber.Length > 0);
    }

    [Fact]
    public void Read_from_an_open_document_leaves_it_unchanged()
    {
        var redline = _r.NumberedHeadings("numbered");
        var before = File.ReadAllBytes(redline);
        using (var doc = Docx.OpenRead(redline))
        {
            Assert.Equal(
                Headings(Outline.Read(redline, OutlineView.AcceptAll)),
                Headings(Outline.Read(doc, OutlineView.AcceptAll)));
        }
        Assert.Equal(before, File.ReadAllBytes(redline));
    }

    [Fact]
    public void Body_text_paragraphs_with_outline_level_nine_are_not_headings()
    {
        var path = new DocxBuilder()
            .Raw(new XElement(W + "p",
                new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Head1")), new XElement(W + "outlineLvl", new XAttribute(W + "val", 9))),
                Run("Not a heading")))
            .Heading("Head2", "A heading")
            .Save(Path.Combine(_r.Dir, "lvl9.docx"));

        var entry = Assert.Single(Outline.Read(path));
        Assert.Equal("A heading", entry.Text);
        Assert.Equal(1, entry.Index);
    }
}
