using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

public sealed class RevisionCheckTests : IDisposable
{
    private readonly Redlines _r = new();

    public void Dispose() => _r.Dispose();

    private static List<Finding> Errors(IEnumerable<Finding> findings) => [.. findings.Where(f => f.Severity == Severity.Error)];

    [Fact]
    public void Tracked_typed_heading_number_change_passes()
    {
        var redline = _r.Make("typed", body: b => _r.TrackedReplacePrefix(Paragraph(b, "2.0 Aims"), "2.0", "2.1"));

        var findings = RevisionCheck.Run(_r.Original, redline);

        Assert.Empty(Errors(findings));
        Assert.DoesNotContain(findings, f => f.Severity == Severity.Warning);
        var outline = Assert.Single(findings, f => f.Rule == "accept-outline");
        Assert.Contains("2.0 Background | 2.1 Aims | 3.0 Data", outline.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Tracked_numbered_heading_styles_pass_and_accept_reads_the_list_numbers()
    {
        var redline = _r.NumberedHeadings("numbered");

        var findings = RevisionCheck.Run(_r.Original, redline);

        Assert.Empty(Errors(findings));
        var outline = Assert.Single(findings, f => f.Rule == "accept-outline");
        Assert.Contains("1.0 Summary | 2.0 Background | 2.1 Aims | 3.0 Data | 3.1 Frog Pond | 3.2 Birds", outline.Message, StringComparison.Ordinal);
        // The new styles and the numbering part cannot be tracked: the author must be told.
        Assert.Contains(findings, f => f is { Rule: "untracked-change", Part: "/word/styles.xml" }
                                       && f.Message.Contains("added style Head1num, style Head2num", StringComparison.Ordinal));
        Assert.Contains(findings, f => f is { Rule: "untracked-change", Part: "/word/numbering.xml" });
    }

    [Fact]
    public void Untracked_paragraph_deletion_is_a_reject_mismatch()
    {
        var redline = _r.Make("deleted", body: b =>
        {
            _r.TrackedReplacePrefix(Paragraph(b, "2.0 Aims"), "2.0", "2.1");
            Paragraph(b, "Aims text.").Remove();
        });

        var finding = Assert.Single(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "reject-mismatch");

        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Contains("loses original paragraph \"Aims text.\"", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Untracked_text_edit_reports_the_changed_paragraph()
    {
        var redline = _r.Make("edited", body: b =>
        {
            var p = Paragraph(b, "Summary text.");
            p.Elements(W + "r").Remove();
            p.Add(Run("Summary text, edited."));
        });

        var finding = Assert.Single(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "reject-mismatch");

        Assert.Contains("gives \"Summary text, edited.\" where the original has \"Summary text.\"", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Revision_inside_a_style_definition_is_an_error()
    {
        // Client redline v2 (2026-10-15): numbering added to Head1 itself as a tracked pPrChange.
        var redline = _r.Make("style-revision", styles: s =>
        {
            var pPr = Style(s, "Head1").Element(W + "pPr")!;
            pPr.Add(new XElement(W + "pPrChange", _r.Mark(), new XElement(W + "pPr", new XElement(W + "keepNext"))));
            pPr.AddFirst(new XElement(W + "spacing", new XAttribute(W + "before", 240)));
        });

        var finding = Assert.Single(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "revision-in-styles");

        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Equal("style Head1", finding.Location);
        Assert.Contains("w:pPrChange", finding.Message, StringComparison.Ordinal);
    }

    /// <summary>A tracked insertion " See {REF bookmark \h}" with a cached result.</summary>
    private XElement InsertedReference(string bookmark, string result) => _r.Ins(
        Run(" See "),
        new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
        new XElement(W + "r", new XElement(W + "instrText", new XAttribute(Ns.Xml + "space", "preserve"), $" REF {bookmark} \\h ")),
        new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
        Run(result),
        new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));

    [Fact]
    public void Reference_to_a_missing_bookmark_breaks_after_accept()
    {
        var redline = _r.Make("broken-ref", body: b => Paragraph(b, "Aims text.").Add(InsertedReference("_Ref999", "Figure 9")));

        var findings = RevisionCheck.Run(_r.Original, redline);

        var finding = Assert.Single(findings, f => f.Rule == "broken-reference");
        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Contains("'_Ref999'", finding.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(findings, f => f.Rule == "reject-mismatch");
    }

    [Fact]
    public void Reference_to_an_existing_bookmark_is_fine()
    {
        var redline = _r.Make("good-ref", body: b =>
        {
            Paragraph(b, "Figure 1. Photographs of the stations.").AddFirst(
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", 1), new XAttribute(W + "name", "_Ref1")));
            Paragraph(b, "Figure 1. Photographs of the stations.").Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 1)));
            Paragraph(b, "Aims text.").Add(InsertedReference("_Ref1", "Figure 1"));
        });

        // The redline's own _Ref bookmark is not a reject mismatch either.
        Assert.Empty(Errors(RevisionCheck.Run(_r.Original, redline)));
    }

    private static XElement TocArial() => new(W + "rPr",
        new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial")));

    [Fact]
    public void Undeclared_change_to_an_existing_style_is_an_error_naming_the_style()
    {
        // Review 2026-10-15: an untracked style change was only a warning, so a silent change to
        // an existing style passed verify while Reject All kept it.
        var redline = _r.Make("toc-fonts", styles: s => Style(s, "TOC1").Add(TocArial()));

        var findings = RevisionCheck.Run(_r.Original, redline);

        var finding = Assert.Single(Errors(findings));
        Assert.Equal("definition-changed", finding.Rule);
        Assert.Equal("/word/styles.xml", finding.Part);
        Assert.Equal("style TOC1", finding.Location);
        Assert.Contains("existing style TOC1 (\"toc 1\") changed", finding.Message, StringComparison.Ordinal);
        Assert.Contains("--allow-style TOC1", finding.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(findings, f => f.Rule == "untracked-change");
    }

    [Fact]
    public void Declared_style_change_is_a_warning_naming_the_style()
    {
        var redline = _r.Make("toc-fonts-declared", styles: s => Style(s, "TOC1").Add(TocArial()));

        var findings = RevisionCheck.Run(_r.Original, redline, new VerifyOptions { IntendedStyles = new HashSet<string>(["TOC1"]) });

        Assert.Empty(Errors(findings));
        var finding = Assert.Single(findings, f => f.Rule == "untracked-change");
        Assert.Equal(Severity.Warning, finding.Severity);
        Assert.Equal("/word/styles.xml", finding.Part);
        Assert.Contains("changed style TOC1 (declared intended)", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Declaring_one_style_does_not_cover_another()
    {
        var redline = _r.Make("toc-fonts-two", styles: s =>
        {
            Style(s, "TOC1").Add(TocArial());
            Style(s, "TOC2").Add(TocArial());
        });

        var findings = RevisionCheck.Run(_r.Original, redline, new VerifyOptions { IntendedStyles = new HashSet<string>(["TOC1"]) });

        Assert.Equal("style TOC2", Assert.Single(Errors(findings)).Location);
    }

    [Fact]
    public void Removed_style_and_changed_document_defaults_are_errors()
    {
        var redline = _r.Make("defaults", styles: s =>
        {
            Style(s, "Head3").Remove();
            s.Element(W + "docDefaults")!.Descendants(W + "rFonts").Single().SetAttributeValue(W + "ascii", "Arial");
        });

        var errors = Errors(RevisionCheck.Run(_r.Original, redline, new VerifyOptions { IntendedStyles = new HashSet<string>(["TOC1"]) }));

        Assert.Equal(["docDefaults", "style Head3"], errors.Where(f => f.Rule == "definition-changed").Select(f => f.Location).Order(StringComparer.Ordinal));
        Assert.Contains(errors, f => f.Location == "style Head3" && f.Message.Contains("removed", StringComparison.Ordinal));
    }

    [Fact]
    public void Style_interface_flags_are_a_warning()
    {
        // Word sets these as styles are used (gallery and styles pane); a reader sees nothing.
        // CT_Style order: name, basedOn, ..., uiPriority, unhideWhenUsed, ..., rsid, pPr.
        var redline = _r.Make("ui-flags", styles: s =>
        {
            var style = Style(s, "TOC2");
            style.Element(W + "basedOn")!.AddAfterSelf(new XElement(W + "uiPriority", new XAttribute(W + "val", 39)), new XElement(W + "unhideWhenUsed"));
            style.Element(W + "pPr")!.AddBeforeSelf(new XElement(W + "rsid", new XAttribute(W + "val", "00AB12CD")));
        });

        var findings = RevisionCheck.Run(_r.Original, redline);

        Assert.Empty(Errors(findings));
        Assert.Contains(findings, f => f.Rule == "untracked-change" && f.Message.Contains("changed style TOC2 (user-interface flags only)", StringComparison.Ordinal));
    }

    [Fact]
    public void Changing_an_existing_list_is_an_error_and_adding_one_is_not()
    {
        var original = _r.Make("listed", package: doc => Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(Numbering())));
        static XElement LoadNumbering(WordprocessingDocument doc, out NumberingDefinitionsPart part)
        {
            part = doc.MainDocumentPart!.NumberingDefinitionsPart!;
            return Docx.Load(part).Root!;
        }
        var changed = _r.Make("list-changed", from: original, package: doc =>
        {
            var root = LoadNumbering(doc, out var part);
            root.Descendants(W + "lvlText").First().SetAttributeValue(W + "val", "%1.");
            Docx.Save(part, root.Document!);
        });
        var added = _r.Make("list-added", from: original, package: doc =>
        {
            var root = LoadNumbering(doc, out var part);
            var abstractNum = new XElement(root.Element(W + "abstractNum")!);
            abstractNum.SetAttributeValue(W + "abstractNumId", 1);
            root.Element(W + "abstractNum")!.AddAfterSelf(abstractNum);
            root.Add(new XElement(W + "num", new XAttribute(W + "numId", 2), new XElement(W + "abstractNumId", new XAttribute(W + "val", 1))));
            Docx.Save(part, root.Document!);
        });

        var error = Assert.Single(Errors(RevisionCheck.Run(original, changed)));
        var addedFindings = RevisionCheck.Run(original, added);

        Assert.Equal(("definition-changed", "abstractNum 0"), (error.Rule, error.Location));
        Assert.Empty(Errors(addedFindings));
        Assert.Contains(addedFindings, f => f is { Rule: "untracked-change", Part: "/word/numbering.xml" }
                                            && f.Message.Contains("added abstractNum 1, num 2", StringComparison.Ordinal));
    }

    [Fact]
    public void Review_state_settings_are_a_warning_and_layout_settings_an_error()
    {
        static void Settings(WordprocessingDocument doc, params XElement[] add)
        {
            var part = doc.MainDocumentPart!.DocumentSettingsPart!;
            var xml = Docx.Load(part);
            xml.Root!.Add(add);
            Docx.Save(part, xml);
        }
        var review = _r.Make("settings-review", package: doc => Settings(doc,
            new XElement(W + "trackRevisions"),
            new XElement(W + "rsids", new XElement(W + "rsidRoot", new XAttribute(W + "val", "00AB12CD")))));
        var layout = _r.Make("settings-layout", package: doc => Settings(doc, new XElement(W + "defaultTabStop", new XAttribute(W + "val", 360))));

        var reviewFindings = RevisionCheck.Run(_r.Original, review);
        var error = Assert.Single(Errors(RevisionCheck.Run(_r.Original, layout)));

        Assert.Empty(Errors(reviewFindings));
        Assert.Contains(reviewFindings, f => f.Rule == "untracked-change" && f.Message.Contains("added trackRevisions, rsids (review state)", StringComparison.Ordinal));
        Assert.Equal(("definition-changed", "defaultTabStop"), (error.Rule, error.Location));
    }

    [Fact]
    public void Theme_change_is_an_error()
    {
        static void Theme(WordprocessingDocument doc, string name)
        {
            var main = doc.MainDocumentPart!;
            var part = main.ThemePart ?? main.AddNewPart<ThemePart>();
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            Docx.Save(part, new XDocument(new XElement(a + "theme", new XAttribute(XNamespace.Xmlns + "a", a.NamespaceName), new XAttribute("name", name))));
        }
        var themed = _r.Make("themed", package: doc => Theme(doc, "Office"));
        var rethemed = _r.Make("rethemed", from: themed, package: doc => Theme(doc, "Other"));

        var added = Errors(RevisionCheck.Run(_r.Original, themed)).Where(f => f.Rule == "definition-changed").ToList();
        var changed = Errors(RevisionCheck.Run(themed, rethemed)).Where(f => f.Rule == "definition-changed").ToList();

        Assert.Contains("added", Assert.Single(added).Message, StringComparison.Ordinal);
        Assert.Equal("theme", Assert.Single(changed).Location);
        Assert.Contains("changed", changed[0].Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unchanged_copy_has_no_findings_but_the_outline()
    {
        var findings = RevisionCheck.Run(_r.Original, _r.Make("copy"));

        var only = Assert.Single(findings);
        Assert.Equal("accept-outline", only.Rule);
    }

    [Fact]
    public void Untracked_paragraph_style_change_is_a_reject_mismatch()
    {
        // Review 2026-10-15 finding 1: "4.3 Methods" Head2 -> Normal without a pPrChange passed verify.
        var redline = _r.Make("restyled", body: b =>
            Paragraph(b, "3.1 Frog Pond").Element(W + "pPr")!.Element(W + "pStyle")!.SetAttributeValue(W + "val", "Normal"));

        var finding = Assert.Single(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "reject-mismatch");

        Assert.Equal(Severity.Error, finding.Severity);
        Assert.Contains("style Head2 -> Normal", finding.Message, StringComparison.Ordinal);
        Assert.Contains("3.1 Frog Pond", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Untracked_run_formatting_is_a_reject_mismatch()
    {
        var redline = _r.Make("bolded", body: b =>
            Paragraph(b, "Summary text.").Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "b"))));

        var finding = Assert.Single(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "reject-mismatch");

        Assert.Contains("run formatting or non-text content", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Untracked_table_property_change_is_a_reject_mismatch()
    {
        XElement Table(string width) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", width), new XAttribute(W + "type", "dxa"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 2000))),
            new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "p", Run("cell")))));
        var original = _r.Make("table", body: b => Paragraph(b, "Aims text.").AddAfterSelf(Table("2000")));
        var redline = _r.Make("table-wider", from: original, body: b => b.Descendants(W + "tblW").Single().SetAttributeValue(W + "w", "4000"));

        var finding = Assert.Single(RevisionCheck.Run(original, redline), f => f.Rule == "reject-mismatch");

        Assert.Contains("outside paragraph content", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Redline_bookkeeping_is_not_a_reject_mismatch()
    {
        // rsids, proofing marks, rendered page breaks, comment anchors and run splits are what
        // Word and redline tools add without changing what the reader sees.
        var redline = _r.Make("bookkeeping", body: b =>
        {
            var p = Paragraph(b, "Background text (Figure 1).");
            p.SetAttributeValue(W + "rsidR", "00AB12CD");
            p.Elements(W + "r").Remove();
            p.Add(
                new XElement(W + "commentRangeStart", new XAttribute(W + "id", 7)),
                new XElement(W + "r", new XAttribute(W + "rsidR", "00AB12CD"), new XElement(W + "lastRenderedPageBreak"),
                    new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), "Background ")),
                new XElement(W + "proofErr", new XAttribute(W + "type", "spellStart")),
                Run("text (Figure 1)."),
                new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 7)),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 7))));
        });

        Assert.DoesNotContain(RevisionCheck.Run(_r.Original, redline), f => f.Rule == "reject-mismatch");
    }

    [Fact]
    public void Tracked_replacement_inside_a_hyperlink_passes()
    {
        // Review 2026-10-15 finding 3: Clippit 3.9 RejectRevisions drops w:del text inside a
        // w:hyperlink and keeps the w:ins text, so a correct redline failed reject-mismatch.
        static XElement Link(params object[] content) =>
            new(W + "hyperlink", new XAttribute(W + "anchor", "_Toc1"), content);
        var original = _r.Make("linked", body: b => Paragraph(b, "Aims text.").AddAfterSelf(
            new XElement(W + "p", Run("see "), Link(Run("link")), Run(" after link"))));
        var redline = _r.Make("linked-edit", from: original, body: b =>
        {
            var link = b.Descendants(W + "hyperlink").Single();
            link.ReplaceNodes(_r.Del("link"), _r.Ins(Run("the new link")));
        });

        var findings = RevisionCheck.Run(original, redline);

        Assert.Empty(Errors(findings));
        Assert.Contains(Outline.Paragraphs(File.ReadAllBytes(redline), OutlineView.RejectAll), e => e.Text == "see link after link");
        Assert.Contains(Outline.Paragraphs(File.ReadAllBytes(redline), OutlineView.AcceptAll), e => e.Text == "see the new link after link");
    }

    [Theory]
    [InlineData(" REF _Ref123 \\h ", "_Ref123")]
    [InlineData("PAGEREF \"_Toc 1\" \\h", "_Toc 1")]
    [InlineData("ref _Ref5 \\r \\h", "_Ref5")]
    [InlineData("SEQ Figure \\* ARABIC", null)]
    [InlineData("REF \\h", null)]
    public void Referenced_bookmark_is_parsed_from_the_instruction(string instruction, string? expected) =>
        Assert.Equal(expected, RevisionCheck.ReferencedBookmark(instruction));
}
