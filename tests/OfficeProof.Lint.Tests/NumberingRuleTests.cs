using OfficeProof.Core;
using OfficeProof.TestKit;
using static OfficeProof.Lint.Tests.Fixtures;

namespace OfficeProof.Lint.Tests;

public sealed class TypedHeadingNumberTests
{
    private const string Rule = "typed-heading-number";

    [Fact]
    public void LegacyReport_flags_every_typed_heading_number_and_nothing_else()
    {
        var findings = RunRule(LegacyReport(), Rule);

        Assert.Equal(["1.0", "2.0", "2.0", "3.0", "3.1", "3.2", "3.2.1"],
            findings.Select(f => f.Message.Split('\'')[1]));
        Assert.All(findings, f => Assert.Equal(Severity.Warning, f.Severity));
        Assert.All(findings, f => Assert.Equal("/word/document.xml", f.Part));
        Assert.DoesNotContain(findings, f => f.Location!.Contains("Abbreviations", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => f.Location!.Contains("Appendix", StringComparison.Ordinal));
    }

    [Fact]
    public void Location_is_paragraph_index_and_sixty_character_excerpt()
    {
        var longTitle = "4.0 " + new string('x', 80);
        var findings = RunRule(new DocxBuilder().Para("intro").Heading("Head1", longTitle), Rule);

        var location = Assert.Single(findings).Location!;
        Assert.StartsWith("¶2 \"4.0 xxx", location, StringComparison.Ordinal);
        Assert.EndsWith("…\"", location, StringComparison.Ordinal);
        Assert.Equal("¶2 \"\"".Length + 60, location.Length);
    }

    [Fact]
    public void Heading_numbered_through_its_style_numPr_is_clean()
    {
        var builder = new DocxBuilder()
            .Style(ParaStyle("Head1numbered", "Head1", NumPr(1)))
            .Heading("Head1numbered", "1.0 Summary");

        Assert.Empty(RunRule(builder, Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void Heading_with_direct_numPr_is_clean()
    {
        var builder = new DocxBuilder().Raw(PNum("Head1", 1, R("1.0 Summary")));

        Assert.Empty(RunRule(builder, Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void NumId_zero_on_the_paragraph_switches_style_numbering_off()
    {
        var builder = new DocxBuilder()
            .Style(ParaStyle("Head1numbered", "Head1", NumPr(1)))
            .Raw(PNum("Head1numbered", 0, R("1.0 Summary")));

        Assert.Single(RunRule(builder, Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void NumId_that_numbering_xml_does_not_define_is_no_numbering()
    {
        // numId 7 is not in numbering.xml: Word renders no number, so the typed one is all there is.
        var builder = new DocxBuilder().Raw(PNum("Head1", 7, R("1.0 Summary")));

        Assert.Single(RunRule(builder, Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void Outline_level_is_inherited_through_basedOn()
    {
        var builder = new DocxBuilder()
            .Style(ParaStyle("MySub", "Head2"))
            .Heading("MySub", "3.1 Frog Pond");

        Assert.Single(RunRule(builder, Rule));
    }

    [Fact]
    public void Body_text_and_year_headings_are_not_typed_heading_numbers()
    {
        var builder = new DocxBuilder()
            .Para("1.0 is a number in body text")
            .Heading("Head1", "2025 Field season")
            .Style(ParaStyle("BodyLevel", "Head1", OutlineLvl(9)))
            .Heading("BodyLevel", "1.0 Explicit body-text level");

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Deleted_typed_number_is_not_visible()
    {
        var builder = new DocxBuilder().Raw(P("Head1", Del(1, "1.0 "), R("Summary")));

        Assert.Empty(RunRule(builder, Rule));
    }
}

public sealed class DoubleHeadingNumberTests
{
    private const string Rule = "double-heading-number";

    private static DocxBuilder Numbered(params System.Xml.Linq.XElement[] heading) =>
        new DocxBuilder().Style(ParaStyle("Head1numbered", "Head1", NumPr(1))).Raw(P("Head1numbered", heading));

    [Fact]
    public void Typed_number_on_a_list_numbered_heading_is_an_error()
    {
        var findings = RunRule(Numbered(R("1.0 Summary")), Rule, path => AddNumbering(path));

        var finding = Assert.Single(findings);
        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Contains("'1.0'", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Numbered_heading_without_typed_number_is_clean()
    {
        Assert.Empty(RunRule(Numbered(R("Summary")), Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void Tracked_deletion_of_the_typed_number_is_clean()
    {
        // The redline's shape: numbered style applied, typed "1.0 " deleted as w:del.
        Assert.Empty(RunRule(Numbered(Del(1, "1.0 "), R("Summary")), Rule, path => AddNumbering(path)));
    }

    [Fact]
    public void Typed_number_without_list_numbering_is_not_doubled()
    {
        Assert.Empty(RunRule(new DocxBuilder().Heading("Head1", "1.0 Summary"), Rule));
    }
}

public sealed class DuplicateHeadingNumberTests
{
    private const string Rule = "duplicate-heading-number";

    [Fact]
    public void LegacyReport_reports_the_second_2_0()
    {
        var finding = Assert.Single(RunRule(LegacyReport(), Rule));

        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Contains("'2.0'", finding.Message, StringComparison.Ordinal);
        Assert.Contains("2.0 Background", finding.Message, StringComparison.Ordinal);
        Assert.Contains("2.0 Aims", finding.Location, StringComparison.Ordinal);
    }

    [Fact]
    public void Distinct_numbers_and_repeats_in_body_text_are_clean()
    {
        var builder = new DocxBuilder()
            .Heading("Head1", "1.0 Summary")
            .Para("1.0 again but in body text")
            .Heading("Head1", "2.0 Background");

        Assert.Empty(RunRule(builder, Rule));
    }
}

public sealed class TypedCaptionNumberTests
{
    private const string Rule = "typed-caption-number";

    [Fact]
    public void LegacyReport_flags_each_typed_caption()
    {
        var findings = RunRule(LegacyReport(), Rule);

        Assert.Equal(["Table 1", "Figure 1", "Figure 2"], findings.Select(f => f.Message.Split('\'')[1]));
        Assert.All(findings, f => Assert.Equal(Severity.Warning, f.Severity));
    }

    [Fact]
    public void Seq_field_caption_is_clean()
    {
        var builder = new DocxBuilder().Raw(P(null, R("Figure "), Field("SEQ Figure \\* ARABIC", "1"), R(". Photographs.")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Sentence_starting_with_figure_is_not_a_caption()
    {
        Assert.Empty(RunRule(new DocxBuilder().Para("Figure 1 shows the stations."), Rule));
    }

    [Theory]
    [InlineData("Table 2: Counts.")]
    [InlineData("Figure 3 – Map of sites.")]
    public void Colon_and_dash_captions_are_typed_captions(string text)
    {
        Assert.Single(RunRule(new DocxBuilder().Para(text), Rule));
    }

    [Fact]
    public void Table_of_figures_lines_are_generated_not_typed()
    {
        var builder = new DocxBuilder().Raw(P("TOC1",
            FldChar("begin"), new System.Xml.Linq.XElement(Ns.W + "r", new System.Xml.Linq.XElement(Ns.W + "instrText", " TOC \\h \\z \\c \"Figure\" ")),
            FldChar("separate"), R("Figure 1. Photographs.\t3"), FldChar("end")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Theory]
    [InlineData("Figure 2.1: Map of sites.")]
    [InlineData("Table 3.2. Counts.")]
    public void Chapter_numbered_captions_are_not_typed_single_counter_captions(string text)
    {
        Assert.Empty(RunRule(new DocxBuilder().Para(text), Rule));
    }
}

public sealed class TypedCrossReferenceTests
{
    private const string Rule = "typed-cross-reference";

    [Fact]
    public void LegacyReport_flags_typed_references_but_not_captions()
    {
        var findings = RunRule(LegacyReport(), Rule);

        Assert.Equal(4, findings.Count);
        Assert.Contains(findings, f => f.Message.Contains("'section 3.1'", StringComparison.Ordinal));
        Assert.Contains(findings, f => f.Message.Contains("'Figure 1', 'Table 1'", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => f.Location!.Contains("Figure 1. Photographs", StringComparison.Ordinal));
        Assert.DoesNotContain(findings, f => f.Location!.Contains("Table 1. Recordings", StringComparison.Ordinal));
    }

    [Fact]
    public void Ref_field_result_is_not_a_typed_reference()
    {
        var builder = new DocxBuilder()
            .Raw(P(null, BookmarkStart(1, "_Ref1"), R("Figure "), Field("SEQ Figure \\* ARABIC", "1"), BookmarkEnd(1), R(". Photographs.")))
            .Raw(P(null, R("Stations ("), Field("REF _Ref1 \\h", "Figure 1"), R("); see "), Field("REF _Ref2 \\r \\h", "3.1"), R(".")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Typed_label_before_a_field_number_is_not_a_typed_reference()
    {
        var builder = new DocxBuilder().Raw(P(null, R("as in Figure "), Field("REF _Ref1 \\# 0 \\h", "1"), R(".")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Field_instruction_text_is_never_a_reference()
    {
        // The instruction names "Table 9"; only the result ("x") is visible.
        var builder = new DocxBuilder().Raw(P(null, Field("QUOTE \"Table 9\"", "x")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Deleted_reference_is_clean_and_inserted_reference_is_flagged()
    {
        var deleted = new DocxBuilder().Raw(P(null, R("See "), Del(1, "Figure 1"), R(".")));
        var inserted = new DocxBuilder().Raw(P(null, R("See "), Ins(1, R("Figure 1")), R(".")));

        Assert.Empty(RunRule(deleted, Rule));
        Assert.Single(RunRule(inserted, Rule));
    }

    [Fact]
    public void Plural_reference_lists_every_number()
    {
        var finding = Assert.Single(RunRule(new DocxBuilder().Para("Compare Figures 7 and 9, and Tables 1, 2 and 4."), Rule));

        Assert.Contains("'Figure 7', 'Figure 9', 'Table 1', 'Table 2', 'Table 4'", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Chapter_numbered_references_are_not_single_number_references()
    {
        Assert.Empty(RunRule(new DocxBuilder().Para("As Figure 2.3 and Table 4.1 show."), Rule));
    }
}

public sealed class MissingReferenceTargetTests
{
    private const string Rule = "missing-reference-target";

    [Fact]
    public void LegacyReport_has_every_target()
    {
        Assert.Empty(RunRule(LegacyReport(), Rule));
    }

    [Fact]
    public void Typed_reference_to_a_caption_that_does_not_exist()
    {
        var builder = new DocxBuilder()
            .Para("Figure 1. Photographs.")
            .Para("Compare Figure 1 with Figure 5 and Table 2.");

        var findings = RunRule(builder, Rule);

        Assert.Equal(["Figure 5", "Table 2"], findings.Select(f => f.Message.Split('\'')[1]));
        Assert.All(findings, f => Assert.Equal(Severity.Error, f.Severity));
    }

    [Fact]
    public void Ref_to_a_missing_bookmark()
    {
        var builder = new DocxBuilder()
            .Raw(P(null, BookmarkStart(1, "_Ref1"), R("Figure 1. Photographs."), BookmarkEnd(1)))
            .Raw(P(null, R("See "), Field("REF _Ref1 \\h", "Figure 1"), R(" and "), Field("REF _Ref404 \\h", "Figure 1"), R(".")));

        var finding = Assert.Single(RunRule(builder, Rule));

        Assert.Contains("_Ref404", finding.Message, StringComparison.Ordinal);
        Assert.Contains("Error! Reference source not found", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Ref_result_naming_a_caption_that_does_not_exist()
    {
        var builder = new DocxBuilder()
            .Raw(P(null, BookmarkStart(1, "_Ref1"), R("Figure 1. Photographs."), BookmarkEnd(1)))
            .Raw(P(null, R("See "), Field("REF _Ref1 \\h", "Figure 3"), R(".")));

        var finding = Assert.Single(RunRule(builder, Rule));

        Assert.Contains("'Figure 3'", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Ref_already_showing_the_broken_reference_text()
    {
        var builder = new DocxBuilder()
            .Raw(P(null, BookmarkStart(1, "_Ref1"), R("Heading"), BookmarkEnd(1)))
            .Raw(P(null, R("See "), Field("REF _Ref1 \\h", "Error! Reference source not found."), R(".")));

        Assert.Single(RunRule(builder, Rule));
    }

    [Fact]
    public void Toc_pageref_fields_are_not_checked()
    {
        // A stale TOC's PAGEREF _Toc… bookmarks are rebuilt by Update Field.
        var builder = new DocxBuilder().Raw(P("TOC1",
            FldChar("begin"), new System.Xml.Linq.XElement(Ns.W + "r", new System.Xml.Linq.XElement(Ns.W + "instrText", " TOC \\o \"1-3\" ")),
            FldChar("separate"), R("Summary\t"), Field("PAGEREF _Toc999 \\h", "3"), FldChar("end")));

        Assert.Empty(RunRule(builder, Rule));
    }

    [Fact]
    public void Every_number_of_a_plural_reference_needs_a_caption()
    {
        var builder = new DocxBuilder()
            .Para("Figure 1. Photographs.")
            .Para("Compare Figures 1 and 7.");

        var finding = Assert.Single(RunRule(builder, Rule));

        Assert.Contains("'Figure 7'", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Chapter_numbered_references_do_not_point_at_single_number_captions()
    {
        // "Figure 2.3" is chapter 2's third figure, not a reference to a missing "Figure 2".
        var builder = new DocxBuilder()
            .Para("Figure 2.1: Map of sites.")
            .Para("As Figure 2.3 shows.");

        Assert.Empty(RunRule(builder, Rule));
    }
}
