using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using static OfficeProof.Lint.Tests.Fixtures;

namespace OfficeProof.Lint.Tests;

public sealed class DanglingStyleTests
{
    private const string Rule = "dangling-style";

    [Fact]
    public void Undefined_paragraph_style_is_counted_per_style()
    {
        var builder = new DocxBuilder()
            .Para("one", "ListParagraph")
            .Para("two", "ListParagraph")
            .Para("three", "TOCHeading");

        var findings = RunRule(builder, Rule);

        Assert.Equal(2, findings.Count);
        Assert.Equal(Severity.Error, findings[0].Severity);
        Assert.StartsWith("2 pStyle reference(s) to undefined style 'ListParagraph'", findings[0].Message, StringComparison.Ordinal);
        Assert.Equal("¶1 \"one\"", findings[0].Location);
        Assert.Contains("'TOCHeading'", findings[1].Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Undefined_character_style_is_dangling_too()
    {
        var rPr = new XElement(Ns.W + "rPr", new XElement(Ns.W + "rStyle", new XAttribute(Ns.W + "val", "Emphasis")));
        var finding = Assert.Single(RunRule(new DocxBuilder().Raw(P(null, R("text", rPr))), Rule));

        Assert.Contains("rStyle", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LegacyReport_references_only_defined_styles()
    {
        Assert.Empty(RunRule(LegacyReport(), Rule));
    }
}

public sealed class BrandColourTests
{
    private const string Rule = "brand-colour";

    private static readonly BrandPalette Palette = new(
        new Dictionary<string, string> { ["1f4e79"] = "dark blue", ["7f7f7f"] = "grey", ["000000"] = "black" },
        new Dictionary<string, string>
        {
            ["7f7f7e"] = "7f7f7f (grey) - one digit out",
            ["2f5496"] = "not in the palette at all - Word's stock accent5 blue",
        });

    private static IReadOnlyList<Finding> Run(DocxBuilder builder) =>
        Linter.Run(builder.Save(NewPath(Rule)), Linter.Rules([Rule], Palette));

    private static XElement Coloured(string hex) =>
        P(null, R("text", new XElement(Ns.W + "rPr", new XElement(Ns.W + "color", new XAttribute(Ns.W + "val", hex)))));

    private static XElement Shaded(string hex) =>
        P(null, new XElement(Ns.W + "pPr", new XElement(Ns.W + "shd",
            new XAttribute(Ns.W + "val", "clear"), new XAttribute(Ns.W + "color", "auto"), new XAttribute(Ns.W + "fill", hex))), R("band"));

    [Fact]
    public void Near_miss_colours_are_named()
    {
        var findings = Run(new DocxBuilder().Raw(Shaded("7F7F7E")).Raw(Shaded("7F7F7E")).Raw(Coloured("2F5496")));

        Assert.Equal(2, findings.Count);
        Assert.Equal("#7f7f7e x2 outside the brand palette -> 7f7f7f (grey) - one digit out", findings[0].Message);
        Assert.Contains("accent5", findings[1].Message, StringComparison.Ordinal);
        Assert.All(findings, f => Assert.Equal(Severity.Warning, f.Severity));
    }

    [Fact]
    public void Palette_theme_and_auto_colours_are_clean()
    {
        var themed = P(null, R("t", new XElement(Ns.W + "rPr", new XElement(Ns.W + "color",
            new XAttribute(Ns.W + "val", "auto"), new XAttribute(Ns.W + "themeColor", "accent1")))));

        Assert.Empty(Run(new DocxBuilder().Raw(Shaded("7F7F7F")).Raw(Coloured("1F4E79")).Raw(themed)));
    }

    [Fact]
    public void Style_colours_are_located_by_style()
    {
        var style = ParaStyle("Banner", "Normal");
        style.Add(new XElement(Ns.W + "rPr", new XElement(Ns.W + "color", new XAttribute(Ns.W + "val", "123456"))));

        var finding = Assert.Single(Run(new DocxBuilder().Style(style)));

        Assert.Equal("/word/styles.xml", finding.Part);
        Assert.Equal("style 'Banner'", finding.Location);
    }

    [Fact]
    public void Palette_is_injectable()
    {
        var path = new DocxBuilder().Raw(Coloured("123456")).Raw(Coloured("ABCDEF")).Save(NewPath("palette"));
        var palette = new BrandPalette(
            new Dictionary<string, string> { ["123456"] = "house blue" },
            new Dictionary<string, string> { ["abcdef"] = "use 123456" });

        var finding = Assert.Single(Linter.Run(path, [new BrandColourRule(palette)]));

        Assert.Equal("#abcdef x1 outside the brand palette -> use 123456", finding.Message);
    }

    [Fact]
    public void Without_a_palette_the_rule_is_skipped_with_one_info_finding()
    {
        var finding = Assert.Single(RunRule(new DocxBuilder().Raw(Coloured("123456")), Rule));

        Assert.Equal(Severity.Info, finding.Severity);
        Assert.StartsWith("skipped: no brand palette", finding.Message, StringComparison.Ordinal);
        Assert.Null(finding.Part);
    }
}

public sealed class TocDirectFormattingTests
{
    private const string Rule = "toc-direct-formatting";

    [Fact]
    public void LegacyReport_direct_arial_on_toc_lines_is_reported_per_toc_style()
    {
        var findings = RunRule(LegacyReport(), Rule);

        Assert.Equal(2, findings.Count);
        Assert.StartsWith("4 TOC1 line(s)", findings[0].Message, StringComparison.Ordinal);
        Assert.StartsWith("2 TOC2 line(s)", findings[1].Message, StringComparison.Ordinal);
        Assert.All(findings, f => Assert.Contains("ascii=Arial, style gives Times New Roman", f.Message, StringComparison.Ordinal));
        Assert.All(findings, f => Assert.Contains("Update Field will discard it", f.Message, StringComparison.Ordinal));
        Assert.Equal("¶1 \"1.0 Summary\"", findings[0].Location);
    }

    [Fact]
    public void Toc_styles_that_give_the_font_clear_the_finding()
    {
        var path = LegacyReport();
        EditStyles(path, styles =>
        {
            StyleIn(styles, "TOC1").Add(new XElement(Ns.W + "rPr", RFonts("Arial")));
            StyleIn(styles, "TOC2").Add(new XElement(Ns.W + "rPr", RFonts("Arial")));
        });

        Assert.Empty(RunRule(path, Rule));
    }

    [Fact]
    public void Matching_size_is_harmless_but_a_different_size_is_lost()
    {
        static XElement TocLine(int sz, bool first) => P("TOC1",
            first ? new object[] { FldChar("begin"), new XElement(Ns.W + "r", new XElement(Ns.W + "instrText", " TOC \\o ")), FldChar("separate") } : null,
            R("Line", new XElement(Ns.W + "rPr", new XElement(Ns.W + "sz", new XAttribute(Ns.W + "val", sz)))),
            first ? null : FldChar("end"));

        var same = RunRule(new DocxBuilder().Raw(TocLine(24, true)).Raw(TocLine(24, false)), Rule);
        var different = RunRule(new DocxBuilder().Raw(TocLine(24, true)).Raw(TocLine(18, false)), Rule);

        Assert.Empty(same);
        Assert.Contains("sz=18, style gives 24", Assert.Single(different).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Toc_styled_lines_outside_a_toc_field_are_not_regenerated()
    {
        var builder = new DocxBuilder().Raw(P("TOC1", R("1.0 Summary", new XElement(Ns.W + "rPr", RFonts("Arial")))));

        Assert.Empty(RunRule(builder, Rule));
    }
}

public sealed class UnresolvedRevisionTests
{
    private const string Rule = "unresolved-revision";

    [Fact]
    public void Tracked_changes_and_comments_are_counted_by_kind()
    {
        var builder = new DocxBuilder()
            .Raw(P(null, R("Kept "), Ins(1, R("added")), Del(2, "removed"), Del(3, "also removed")));

        var findings = RunRule(builder, Rule, path => AddComments(path, 2));

        Assert.Equal(2, findings.Count);
        Assert.Equal("3 tracked change(s): 2 del, 1 ins", findings[0].Message);
        Assert.Equal("/word/document.xml", findings[0].Part);
        Assert.Equal("2 comment(s)", findings[1].Message);
        Assert.Equal("/word/comments.xml", findings[1].Part);
    }

    [Fact]
    public void Clean_document_has_no_revisions()
    {
        Assert.Empty(RunRule(LegacyReport(), Rule));
    }
}

public sealed class RevisionInStylesTests
{
    private const string Rule = "revision-in-styles";

    private static XElement PPrChange() =>
        new(Ns.W + "pPrChange", new XAttribute(Ns.W + "id", 9), new XAttribute(Ns.W + "author", "Test"),
            new XAttribute(Ns.W + "date", "2026-10-15T00:00:00Z"), new XElement(Ns.W + "pPr"));

    [Fact]
    public void Tracked_change_in_a_style_definition_is_an_error()
    {
        var path = LegacyReport();
        EditStyles(path, styles => StyleIn(styles, "Head1").Element(Ns.W + "pPr")!.Add(NumPr(1), PPrChange()));

        var finding = Assert.Single(RunRule(path, Rule));

        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Equal("/word/styles.xml", finding.Part);
        Assert.Equal("style 'Head1'", finding.Location);
        Assert.Contains("Reject All does not undo", finding.Message, StringComparison.Ordinal);
        // Story-part revision counting does not double-report it.
        Assert.Empty(RunRule(path, "unresolved-revision"));
    }

    [Fact]
    public void Tracked_change_in_numbering_is_an_error()
    {
        var path = LegacyReport();
        AddNumbering(path, new XElement(Ns.W + "pPr", PPrChange()));

        var finding = Assert.Single(RunRule(path, Rule));

        Assert.Equal("/word/numbering.xml", finding.Part);
        Assert.Equal("abstractNum 0", finding.Location);
    }

    [Fact]
    public void Tracked_changes_in_the_body_are_not_style_revisions()
    {
        Assert.Empty(RunRule(new DocxBuilder().Raw(P(null, Ins(1, R("added")))), Rule));
    }
}

public sealed class LinterTests
{
    [Fact]
    public void Default_rules_have_unique_ids_and_documented_severities()
    {
        var rules = Linter.DefaultRules();

        Assert.Equal(
            [
                ("dangling-style", Severity.Error), ("brand-colour", Severity.Warning),
                ("typed-heading-number", Severity.Warning), ("double-heading-number", Severity.Error),
                ("duplicate-heading-number", Severity.Error),
                ("typed-caption-number", Severity.Warning), ("typed-cross-reference", Severity.Warning),
                ("missing-reference-target", Severity.Error), ("toc-direct-formatting", Severity.Warning),
                ("unresolved-revision", Severity.Warning), ("revision-in-styles", Severity.Error),
            ],
            rules.Select(r => (r.Id, r.Severity)));
        Assert.All(rules, r => Assert.False(string.IsNullOrWhiteSpace(r.Description)));
    }

    [Fact]
    public void Rules_are_found_by_id_and_unknown_ids_are_rejected()
    {
        Assert.IsType<TocDirectFormattingRule>(Linter.FindRule("TOC-Direct-Formatting"));
        Assert.Null(Linter.FindRule("nope"));
        var ex = Assert.Throws<ArgumentException>(() => Linter.Rules(["typed-heading-number", "nope"]));
        Assert.Contains("'nope'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("revision-in-styles", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LegacyReport_produces_the_typed_number_problems_and_no_missing_targets()
    {
        var findings = Linter.Run(LegacyReport());

        // Without a template palette brand-colour only reports that it was skipped.
        Assert.Equal(
            ["brand-colour", "typed-heading-number", "duplicate-heading-number", "typed-caption-number", "typed-cross-reference", "toc-direct-formatting"],
            findings.Select(f => f.Rule).Distinct());
        Assert.Equal(Severity.Info, Assert.Single(findings, f => f.Rule == "brand-colour").Severity);
        Assert.Equal(1, FindingReport.ExitCode(findings));
    }

    [Fact]
    public void A_palette_reaches_brand_colour_through_rule_selection()
    {
        var palette = new BrandPalette(new Dictionary<string, string>(), new Dictionary<string, string>());

        var rule = Assert.IsType<BrandColourRule>(Linter.FindRule("brand-colour", palette));
        var path = new DocxBuilder().Raw(P(null, R("x", new XElement(Ns.W + "rPr", new XElement(Ns.W + "color", new XAttribute(Ns.W + "val", "123456"))))))
            .Save(NewPath("palette-selection"));

        Assert.Equal(Severity.Warning, Assert.Single(Linter.Run(path, [rule])).Severity);
        Assert.Equal(Severity.Warning, Assert.Single(Linter.Run(path, Linter.Rules(["brand-colour"], palette))).Severity);
        Assert.Contains(Linter.Run(path, Linter.DefaultRules(palette)), f => f.Rule == "brand-colour" && f.Severity == Severity.Warning);
    }

    [Fact]
    public void Run_leaves_the_file_untouched()
    {
        var path = LegacyReport();
        var before = SHA256.HashData(File.ReadAllBytes(path));

        _ = Linter.Run(path);

        Assert.Equal(before, SHA256.HashData(File.ReadAllBytes(path)));
    }
}
