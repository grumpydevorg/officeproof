using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using static OfficeProof.Redline.Tests.Support;

namespace OfficeProof.Redline.Tests;

public sealed class EditTests
{
    private static string DelText(XElement root) =>
        string.Concat(root.Descendants(W + "del").SelectMany(d => d.Descendants())
            .Select(e => e.Name == W + "delText" ? e.Value : e.Name == W + "tab" ? "\t" : ""));

    private static string InsText(XElement root) =>
        string.Concat(root.Descendants(W + "ins").SelectMany(d => d.Descendants())
            .Select(e => e.Name == W + "t" ? e.Value : e.Name == W + "tab" ? "\t" : ""));

    private static string Doc(params string[] paragraphs)
    {
        var b = new DocxBuilder();
        foreach (var p in paragraphs) b.Para(p);
        return b.Save(Path.Combine(TempDir(), "doc.docx"));
    }

    private static string RunScript(string input, string json, out EditScriptResult result)
    {
        EditScriptResult? r = null;
        var output = Edit(input, ed => r = EditScript.Parse(json).Apply(ed));
        result = r!;
        return output;
    }

    [Fact]
    public void Replace_tracks_only_the_changed_span()
    {
        var input = Doc("Formal tests of evening richness by year and distance found nothing.");
        var output = RunScript(input, """[{ "find": "evening richness", "replace": "year and distance", "with": "year, distance and elevation" }]""", out var result);
        var p = MainXml(output).Descendants(W + "p").Single();
        Assert.Equal(" and distance", DelText(p));
        Assert.Equal(", distance and elevation", InsText(p));
        Assert.Equal(1, result.Replacements);
        Assert.Equal("Formal tests of evening richness by year and distance found nothing.", Outline(output, accept: false).Single().Text);
        Assert.Equal("Formal tests of evening richness by year, distance and elevation found nothing.", Outline(output, accept: true).Single().Text);
    }

    [Fact]
    public void Replace_snaps_to_whole_words()
    {
        var output = RunScript(TypedNumberReport(), """[{ "find": "Aims text", "replace": "Aims text", "with": "Aim text" }]""", out _);
        var p = MainXml(output).Descendants(W + "p").Single(p => Wml.VisibleText(p).EndsWith(" text.", StringComparison.Ordinal) && Wml.HasRevisions(p));
        Assert.Equal("Aims", DelText(p));
        Assert.Equal("Aim", InsText(p));
    }

    [Fact]
    public void Whitespace_only_deletion_survives_save_and_reject()
    {
        var input = TypedNumberReport();
        var output = RunScript(input, """[{ "find": "metres above", "replace": "asl  metres", "with": "asl metres" }]""", out _);
        var del = MainXml(output).Descendants(W + "delText").Single();
        Assert.Equal(" ", del.Value);
        Assert.Equal("preserve", (string?)del.Attribute(Ns.Xml + "space"));
        Assert.Contains("m asl  metres above sea level", Outline(output, accept: false).Select(p => p.Text));
        Assert.Contains("m asl metres above sea level", Outline(output, accept: true).Select(p => p.Text));
    }

    [Theory]
    [InlineData("text.", "more than one")]
    [InlineData("no such words", "no paragraph")]
    public void A_snippet_must_match_exactly_one_paragraph(string snippet, string expected)
    {
        var e = Assert.Throws<InvalidOperationException>(() =>
            RunScript(TypedNumberReport(), $$"""[{ "find": "{{snippet}}", "comment": "x" }]""", out _));
        Assert.Contains(snippet, e.Message, StringComparison.Ordinal);
        Assert.Contains(expected, e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Repeated_old_text_needs_nth()
    {
        var input = Doc("one two one");
        var e = Assert.Throws<InvalidOperationException>(() => RunScript(input, """[{ "find": "two", "replace": "one", "with": "three" }]""", out _));
        Assert.Contains("nth", e.Message, StringComparison.Ordinal);
        var output = RunScript(input, """[{ "find": "two", "replace": "one", "with": "three", "nth": 1 }]""", out _);
        Assert.Equal("one two three", Outline(output, accept: true).Single().Text);
    }

    [Theory]
    [InlineData("""[{ "find": "a", "with": "b" }]""")]
    [InlineData("""[{ "find": "a", "replace": "b" }]""")]
    [InlineData("""[{ "find": "a" }]""")]
    [InlineData("""[{ "find": "a", "comment": "c", "colour": "red" }]""")]
    [InlineData("""{ "find": "a" }""")]
    public void Malformed_scripts_are_rejected(string json) =>
        Assert.Throws<FormatException>(() => EditScript.Parse(json));

    [Fact]
    public void Comments_anchor_on_a_substring_or_a_replacement()
    {
        var input = Doc("Calculated a set of 52 soundscape indices (Table 3).");
        var output = RunScript(input, """
            [
              { "find": "52 soundscape", "comment": "52 or 63?", "anchor": "52" },
              { "find": "52 soundscape", "replace": "Table 3", "with": "Appendix 2", "comment": "Index list is Appendix 2." }
            ]
            """, out var result);
        Assert.Equal(2, result.Comments.Count);
        var main = MainXml(output);
        var comments = PartXml(output, m => m.WordprocessingCommentsPart).Root!.Elements(W + "comment").ToList();
        Assert.Equal(["52 or 63?", "Index list is Appendix 2."], comments.Select(c => string.Concat(c.Descendants(W + "t").Select(t => t.Value))));
        foreach (var id in comments.Select(c => (string)c.Attribute(W + "id")!))
        {
            Assert.Single(main.Descendants(W + "commentRangeStart"), e => (string?)e.Attribute(W + "id") == id);
            Assert.Single(main.Descendants(W + "commentRangeEnd"), e => (string?)e.Attribute(W + "id") == id);
            Assert.Single(main.Descendants(W + "commentReference"), e => (string?)e.Attribute(W + "id") == id);
        }
        var start = main.Descendants(W + "commentRangeStart").First();
        Assert.Equal("52", (string?)start.ElementsAfterSelf(W + "r").First().Element(W + "t"));
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Splitting_keeps_formatting_multiple_texts_and_tabs()
    {
        var bold = new XElement(W + "rPr", new XElement(W + "b"));
        var input = new DocxBuilder().Raw(new XElement(W + "p", new XElement(W + "r", bold,
                new XElement(W + "t", "ab"), new XElement(W + "tab"), new XElement(W + "t", "cd"))))
            .Save(Path.Combine(TempDir(), "runs.docx"));
        var output = Edit(input, ed =>
        {
            var p = ed.Paragraphs.Single();
            ed.Replace(p, 1, 4, [Inline.Text("X")]);
        });
        var para = MainXml(output).Descendants(W + "p").Single();
        Assert.Equal("b\tc", DelText(para));
        Assert.Equal("X", InsText(para));
        Assert.All(para.Descendants(W + "r"), r => Assert.NotNull(r.Element(W + "rPr")?.Element(W + "b")));
        Assert.Equal("ab\tcd", Outline(output, accept: false).Single().Text);
        Assert.Equal("aXd", Outline(output, accept: true).Single().Text);
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Edits_inside_another_authors_insertion_are_refused()
    {
        XElement Run(string text) => new(W + "r", new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text));
        var input = new DocxBuilder().Raw(new XElement(W + "p", Run("Hello "),
                new XElement(W + "ins", new XAttribute(W + "id", 1), new XAttribute(W + "author", "Reviewer"),
                    new XAttribute(W + "date", "2026-10-14T00:00:00Z"), Run("big ")),
                Run("world")))
            .Save(Path.Combine(TempDir(), "foreign-ins.docx"));
        var e = Assert.Throws<InvalidOperationException>(() =>
            RunScript(input, """[{ "find": "Hello", "replace": "big world", "with": "small world" }]""", out _));
        Assert.Contains("Reviewer", e.Message, StringComparison.Ordinal);
        // Text around the insertion is still editable.
        var output = RunScript(input, """[{ "find": "Hello", "replace": "world", "with": "planet" }]""", out _);
        Assert.Equal("Hello big planet", Outline(output, accept: true).Single().Text);
    }

    [Fact]
    public void Paragraph_property_changes_are_tracked_with_the_prior_properties()
    {
        var input = new DocxBuilder().Heading("Head1", "Title").Save(Path.Combine(TempDir(), "ppr.docx"));
        var output = Edit(input, ed => ed.ChangeParagraph(ed.Paragraphs.Single(), pPr => TrackedEditor.SetStyle(pPr, "Head2")));
        var pPr = MainXml(output).Descendants(W + "pPr").Single(e => e.Parent!.Name == W + "p");
        Assert.Equal("Head2", (string?)pPr.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.Equal("Head1", (string?)pPr.Element(W + "pPrChange")?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.Equal("Head1", Outline(output, accept: false).Single().Style);
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Splitting_a_run_with_a_tracked_format_change_gives_each_half_its_own_id()
    {
        var input = new DocxBuilder().Raw(new XElement(W + "p", new XElement(W + "r",
                new XElement(W + "rPr", new XElement(W + "b"),
                    new XElement(W + "rPrChange", new XAttribute(W + "id", 1), new XAttribute(W + "author", "Reviewer"),
                        new XAttribute(W + "date", "2026-10-14T00:00:00Z"), new XElement(W + "rPr"))),
                new XElement(W + "t", "abcdef"))))
            .Save(Path.Combine(TempDir(), "rprchange.docx"));
        var output = Edit(input, ed => ed.Replace(ed.Paragraphs.Single(), 2, 4, [Inline.Text("X")]));
        var ids = MainXml(output).Descendants().Attributes(W + "id").Select(a => a.Value).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Equal(3, MainXml(output).Descendants(W + "rPrChange").Count()); // ab, cd (deleted), ef; the insertion has none
        Assert.Empty(Docx.Validate(output));
    }

    [Theory]
    [InlineData("Entry two", "Entry two extra")]
    [InlineData("Entry two", "New Entry two")]
    [InlineData("Entry one", "Entry one extra")]
    public void Insertions_next_to_table_of_contents_text_are_refused_naming_the_paragraph(string old, string replacement)
    {
        var input = new DocxBuilder().Toc((1, "Entry one"), (1, "Entry two")).Para("Body text.")
            .Save(Path.Combine(TempDir(), "toc.docx"));
        var e = Assert.Throws<InvalidOperationException>(() =>
            RunScript(input, $$"""[{ "find": "{{old}}", "replace": "{{old}}", "with": "{{replacement}}" }]""", out _));
        Assert.Contains($"in paragraph \"{old}\"", e.Message, StringComparison.Ordinal);
        Assert.Contains("field", e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_comment_anchored_on_text_this_script_inserted_sits_outside_the_insertion()
    {
        var input = Doc("Calculated a set of 52 soundscape indices.");
        var output = RunScript(input, """
            [
              { "find": "soundscape indices", "replace": "52", "with": "63" },
              { "find": "soundscape indices", "comment": "Was 52.", "anchor": "63" }
            ]
            """, out _);
        var p = MainXml(output).Descendants(W + "p").Single();
        var markers = p.Descendants().Where(e => e.Name == W + "commentRangeStart" || e.Name == W + "commentRangeEnd" || e.Name == W + "commentReference").ToList();
        Assert.Equal(3, markers.Count);
        Assert.All(markers, m => Assert.DoesNotContain(m.Ancestors(), a => a.Name == W + "ins"));
        Assert.Equal("Calculated a set of 52 soundscape indices.", Outline(output, accept: false).Single().Text);
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Text_box_fallback_copies_are_not_paragraphs_to_edit()
    {
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        XElement Box(string text) => new(W + "txbxContent", new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "t", text))));
        var input = new DocxBuilder()
            .Raw(new XElement(W + "p", new XElement(W + "r", new XElement(mc + "AlternateContent",
                new XElement(mc + "Choice", new XAttribute("Requires", "wps"), Box("Figure 1. Boxed caption.")),
                new XElement(mc + "Fallback", Box("Figure 1. Boxed caption."))))))
            .Para("As Figure 1 shows.")
            .Save(Path.Combine(TempDir(), "textbox.docx"));
        CaptionFieldsResult? result = null;
        var output = Edit(input, ed =>
        {
            Assert.Equal("Figure 1. Boxed caption.", TrackedEditor.Text(ed.FindParagraph("Boxed caption")));
            result = CaptionFields.Apply(ed, options: new CaptionOptions { ExplainInComment = false });
        });
        Assert.Equal(1, result!.Captions);
        Assert.Empty(result.Notes);
        Assert.Empty(result.Comments);
        var fallback = MainXml(output).Descendants(mc + "Fallback").Single();
        Assert.False(Wml.HasRevisions(fallback));
    }
}
