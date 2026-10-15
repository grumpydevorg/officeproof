using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.Redline.Tests.Support;

namespace OfficeProof.Redline.Tests;

public sealed class HeadingAndCaptionTests
{
    private static readonly HeadingNumberingOptions AimsOverride = new()
    {
        LevelOverrides = new Dictionary<string, int> { ["Aims"] = 2 },
    };

    private static (string Input, string Output, HeadingNumberingResult Headings, CaptionFieldsResult Captions) Redline(HeadingNumberingOptions? options)
    {
        var input = TypedNumberReport();
        HeadingNumberingResult? headings = null;
        CaptionFieldsResult? captions = null;
        var output = Edit(input, ed =>
        {
            headings = HeadingNumbering.Apply(ed, ReportLevels, options);
            captions = CaptionFields.Apply(ed);
        });
        return (input, output, headings!, captions!);
    }

    private static readonly string[] Headings = ["Head1", "Head2", "Head3", "Head1numbered", "Head2numbered", "Head3numbered"];

    private static List<string> HeadingOutline(string path) =>
        Outline(path, accept: true).Where(p => Headings.Contains(p.Style))
            .Select(p => p.Number.Length == 0 ? p.Text : $"{p.Number} {p.Text}").ToList();

    [Fact]
    public void Reject_all_restores_every_paragraph_and_leaves_no_list_numbers()
    {
        var (input, output, _, _) = Redline(AimsOverride);
        var rejected = Outline(output, accept: false);
        Assert.Equal(BodyTexts(input), rejected.Select(p => p.Text).ToList());
        Assert.All(rejected, p => Assert.Equal("", p.Number));
    }

    [Fact]
    public void Accept_all_numbers_headings_with_the_level_override()
    {
        var (_, output, result, _) = Redline(AimsOverride);
        Assert.Equal(
            ["Abbreviations", "1.0 Summary", "2.0 Background", "2.1 Aims", "3.0 Data", "3.1 Frog Pond", "3.2 Birds", "3.2.1 Detail", "Appendix 1 Species"],
            HeadingOutline(output));
        Assert.Equal(7, result.NumberedHeadings);
        Assert.Equal(["Abbreviations", "Appendix 1 Species"], result.UnnumberedHeadings);
    }

    [Fact]
    public void Without_the_override_the_duplicate_two_becomes_the_next_top_level_section()
    {
        var (_, output, _, _) = Redline(null);
        Assert.Equal(
            ["Abbreviations", "1.0 Summary", "2.0 Background", "3.0 Aims", "4.0 Data", "4.1 Frog Pond", "4.2 Birds", "4.2.1 Detail", "Appendix 1 Species"],
            HeadingOutline(output));
    }

    [Fact]
    public void Fields_point_at_existing_bookmarks_and_the_package_is_valid()
    {
        var (_, output, headings, captions) = Redline(AimsOverride);
        var xml = MainXml(output);
        var instructions = xml.Descendants(W + "p").SelectMany(Wml.FieldInstructions).ToList();
        var seq = instructions.Where(i => i.StartsWith("SEQ ", StringComparison.Ordinal)).ToList();
        var refs = instructions.Where(i => i.StartsWith("REF ", StringComparison.Ordinal)).ToList();
        var bookmarks = xml.Descendants(W + "bookmarkStart").Select(b => (string)b.Attribute(W + "name")!).ToHashSet();

        Assert.Equal(3, captions.Captions);
        Assert.Equal(captions.Captions, seq.Count);
        Assert.Equal(6, captions.References);
        Assert.Equal(1, headings.SectionReferences);
        Assert.Equal(7, refs.Count);
        Assert.All(refs, r => Assert.Contains(r.Split(' ')[1], bookmarks));
        Assert.Contains(refs, r => r.EndsWith(@"\r \h", StringComparison.Ordinal));
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Style_and_numbering_definitions_are_untracked_and_reported()
    {
        var (_, output, headings, _) = Redline(AimsOverride);
        var styles = PartXml(output, m => m.StyleDefinitionsPart);
        var numbering = PartXml(output, m => m.NumberingDefinitionsPart);
        Assert.False(Wml.HasRevisions(styles.Root!));
        Assert.False(Wml.HasRevisions(numbering.Root!));
        Assert.Contains(headings.UntrackedChanges, c => c.Contains("Head1numbered", StringComparison.Ordinal));
        Assert.Contains(headings.UntrackedChanges, c => c.StartsWith("numbering.xml", StringComparison.Ordinal));
        // the override and the explanation each get a comment
        Assert.Equal(2, headings.Comments.Count);
    }

    [Fact]
    public void New_list_ids_go_above_existing_ones_in_schema_order()
    {
        var input = TypedNumberReport();
        using (var doc = WordprocessingDocument.Open(input, true))
        {
            Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(new XElement(W + "numbering",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 5),
                    new XElement(W + "lvl", new XAttribute(W + "ilvl", 0), new XElement(W + "numFmt", new XAttribute(W + "val", "bullet")),
                        new XElement(W + "lvlText", new XAttribute(W + "val", "•")))),
                new XElement(W + "num", new XAttribute(W + "numId", 7), new XElement(W + "abstractNumId", new XAttribute(W + "val", 5))))));
        }
        var output = Edit(input, ed => HeadingNumbering.Apply(ed, ReportLevels));
        var root = PartXml(output, m => m.NumberingDefinitionsPart).Root!;
        Assert.Equal(["abstractNum:5", "abstractNum:6", "num:7", "num:8"],
            root.Elements().Select(e => $"{e.Name.LocalName}:{(string?)e.Attribute(W + "abstractNumId") ?? (string?)e.Attribute(W + "numId")}"));
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void An_override_matching_no_heading_is_an_error()
    {
        var input = TypedNumberReport();
        var e = Assert.Throws<InvalidOperationException>(() => Edit(input, ed => HeadingNumbering.Apply(ed, ReportLevels,
            new HeadingNumberingOptions { LevelOverrides = new Dictionary<string, int> { ["Nowhere"] = 2 } })));
        Assert.Contains("Nowhere", e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_reference_fix_retargets_one_reference_and_explains_it()
    {
        var input = TypedNumberReport();
        CaptionFieldsResult? result = null;
        var output = Edit(input, ed => result = CaptionFields.Apply(ed, [new ReferenceFix("Detail text", "Figure", 1, 2)]));
        var detail = MainXml(output).Descendants(W + "p").Single(p => Wml.VisibleText(p).StartsWith("Detail text", StringComparison.Ordinal));
        var figure2 = Bookmark(output, "Figure 2. Ordination");
        Assert.Equal([$@"REF {figure2} \h", $@"REF {figure2} \h"], Wml.FieldInstructions(detail));
        Assert.Equal("Detail text, compare Figure 2 and Figure 2.", Wml.VisibleText(detail));
        Assert.Equal(1, result!.Retargeted);
        Assert.Single(result.Comments, c => c.Text.StartsWith("Was Figure 1", StringComparison.Ordinal));
        Assert.Equal("Detail text, compare Figure 1 and Figure 2.", Outline(output, accept: false).Single(p => p.Text.StartsWith("Detail", StringComparison.Ordinal)).Text);
    }

    [Fact]
    public void A_reference_fix_that_matches_nothing_is_an_error_naming_it()
    {
        var input = TypedNumberReport();
        var e = Assert.Throws<InvalidOperationException>(() => Edit(input, ed => CaptionFields.Apply(ed, [new ReferenceFix("Detail text", "Figure", 8, 9)])));
        Assert.Contains("Detail text", e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_reference_without_a_caption_gets_a_comment_not_a_field()
    {
        var input = new DocxBuilder()
            .Para("As Figure 1 and Figure 5 show.")
            .Para("Figure 1. Only caption.")
            .Save(Path.Combine(TempDir(), "orphan.docx"));
        CaptionFieldsResult? result = null;
        var output = Edit(input, ed => result = CaptionFields.Apply(ed, options: new CaptionOptions { ExplainInComment = false }));
        Assert.Equal(1, result!.References);
        Assert.Equal(["Figure 5 in \"As Figure 1 and Figure 5 show.\""], result.UnresolvedReferences);
        var comment = Assert.Single(result.Comments);
        Assert.Contains("No caption Figure 5", comment.Text, StringComparison.Ordinal);
        var p = MainXml(output).Descendants(W + "p").First();
        Assert.Single(Wml.FieldInstructions(p));
        Assert.Single(p.Descendants(W + "commentRangeStart"));
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Toc_font_goes_into_the_toc_styles_and_is_reported_untracked()
    {
        var input = TypedNumberReport();
        var output = Path.Combine(TempDir(), "toc.docx");
        TocFontResult result;
        using (var doc = Docx.OpenCopy(input, output)) result = TocStyles.SetFont(doc, "Arial");
        var styles = PartXml(output, m => m.StyleDefinitionsPart).Root!.Elements(W + "style")
            .Where(s => ((string)s.Attribute(W + "styleId")!).StartsWith("TOC", StringComparison.Ordinal)).ToList();
        Assert.Equal(3, styles.Count);
        Assert.All(styles, s => Assert.Equal("Arial", (string?)s.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii")));
        Assert.Equal(3, result.UntrackedChanges.Count);
        Assert.Contains("toc 4", Assert.Single(result.Notes), StringComparison.Ordinal);
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void Chapter_numbered_captions_and_references_are_left_alone()
    {
        var input = new DocxBuilder()
            .Para("Figure 2.1: Map of the stations.")
            .Para("See Figure 2.3 and Figure 2.1, and Figure 1.")
            .Para("Figure 1. Plain caption.")
            .Save(Path.Combine(TempDir(), "chapters.docx"));
        CaptionFieldsResult? result = null;
        var output = Edit(input, ed => result = CaptionFields.Apply(ed, options: new CaptionOptions { ExplainInComment = false }));
        Assert.Equal(1, result!.Captions);
        Assert.Equal(1, result.References);
        var paras = MainXml(output).Descendants(W + "p").ToList();
        Assert.False(Wml.HasRevisions(paras[0]));
        var refs = Wml.FieldInstructions(paras[1]).ToList();
        Assert.Single(refs);
        Assert.Equal("See Figure 2.3 and Figure 2.1, and Figure 1.", Wml.VisibleText(paras[1]));
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void A_plural_reference_becomes_one_number_field_per_figure_with_the_words_kept()
    {
        var input = new DocxBuilder()
            .Para("Figure 1. First.").Para("Figure 2. Second.").Para("Figure 3. Third.")
            .Para("Compare Figures 1 and 3, then Figures 1, 2 and 3.")
            .Save(Path.Combine(TempDir(), "plural.docx"));
        CaptionFieldsResult? result = null;
        var output = Edit(input, ed => result = CaptionFields.Apply(ed, [new ReferenceFix("Compare Figures", "Figure", 3, 2, Nth: 0)],
            new CaptionOptions { ExplainInComment = false }));
        Assert.Equal(5, result!.References);
        var xml = MainXml(output);
        var p = xml.Descendants(W + "p").Last();
        var refs = Wml.FieldInstructions(p).ToList();
        Assert.Equal(5, refs.Count);
        foreach (var name in refs.Select(r => r.Split(' ')[1]))
        {
            // The target bookmark spans the caption's SEQ field only, so REF shows "3", not "Figure 3".
            var start = xml.Descendants(W + "bookmarkStart").Single(b => (string?)b.Attribute(W + "name") == name);
            var end = start.ElementsAfterSelf(W + "bookmarkEnd").First(e => (string?)e.Attribute(W + "id") == (string?)start.Attribute(W + "id"));
            var inside = start.ElementsAfterSelf().TakeWhile(e => e != end).ToList();
            Assert.Matches(@"^\d+$", string.Concat(inside.SelectMany(e => e.DescendantsAndSelf(W + "t")).Select(t => t.Value)));
            Assert.Single(inside.SelectMany(e => e.Descendants(W + "instrText")), i => i.Value.Contains("SEQ", StringComparison.Ordinal));
        }
        Assert.Equal("Compare Figures 1 and 2, then Figures 1, 2 and 3.", Wml.VisibleText(p));
        Assert.Equal("Compare Figures 1 and 3, then Figures 1, 2 and 3.", Outline(output, accept: false)[^1].Text);
        Assert.Empty(Docx.Validate(output));
    }

    [Fact]
    public void The_caption_explanation_comment_is_not_inside_a_caption_bookmark()
    {
        var (_, output, _, captions) = Redline(AimsOverride);
        Assert.Contains(captions.Comments, c => c.Text.StartsWith("Figure and table numbers", StringComparison.Ordinal));
        var xml = MainXml(output);
        var captionParas = xml.Descendants(W + "p").Where(p => Wml.HasField(p, "SEQ")).ToList();
        Assert.NotEmpty(captionParas);
        foreach (var start in captionParas.SelectMany(p => p.Elements(W + "bookmarkStart")))
        {
            var end = start.ElementsAfterSelf(W + "bookmarkEnd").First(e => (string?)e.Attribute(W + "id") == (string?)start.Attribute(W + "id"));
            var inside = start.ElementsAfterSelf().TakeWhile(e => e != end).SelectMany(e => e.DescendantsAndSelf());
            Assert.DoesNotContain(inside, e => e.Name.LocalName is "commentRangeStart" or "commentRangeEnd" or "commentReference");
        }
    }

    [Fact]
    public void A_year_is_not_a_heading_number_and_section_references_need_a_dot()
    {
        var input = new DocxBuilder()
            .Heading("Head1", "1.0 Summary")
            .Heading("Head1", "2025 Field season")
            .Para("As in section 1, see section 1.0.")
            .Save(Path.Combine(TempDir(), "year.docx"));
        HeadingNumberingResult? result = null;
        var output = Edit(input, ed => result = HeadingNumbering.Apply(ed, ReportLevels, new HeadingNumberingOptions { ExplainInComment = false }));
        Assert.Equal(1, result!.NumberedHeadings);
        Assert.Equal(["2025 Field season"], result.UnnumberedHeadings);
        Assert.Equal(1, result.SectionReferences);
        var accepted = Outline(output, accept: true);
        Assert.Equal(("2025 Field season", ""), (accepted[1].Text, accepted[1].Number));
        Assert.Equal("As in section 1, see section 1.0.", accepted[2].Text);
        Assert.Single(Wml.FieldInstructions(MainXml(output).Descendants(W + "p").Last()));
    }

    private static string Bookmark(string path, string captionStart)
    {
        var p = MainXml(path).Descendants(W + "p").Single(p => Wml.VisibleText(p).StartsWith(captionStart, StringComparison.Ordinal));
        return (string)p.Descendants(W + "bookmarkStart").Single().Attribute(W + "name")!;
    }
}
