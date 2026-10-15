using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.Templating.Tests.TestFiles;

namespace OfficeProof.Templating.Tests;

public sealed class RetemplaterTests
{
    private static readonly ReportMetadata Seed = ReportMetadata.Parse("""
        report:
          title: "Seed title"
          preparedFor: "Seed client"
          projectCode: "EX-999"
          version: "2026-10-15"
          author: "Seed Author"
        """);

    private static List<(string? Style, string Text)> Headings(string path)
    {
        using var doc = Docx.OpenRead(path);
        return Docx.Load(doc.MainDocumentPart!).Descendants(Ns.W + "p")
            .Select(p => (Style: Wml.StyleId(p), Text: Wml.VisibleText(p)))
            .Where(p => p.Style is not null && p.Style.StartsWith("Heading", StringComparison.Ordinal))
            .ToList();
    }

    [Fact]
    public void Legacy_headings_map_to_the_manifest_headings_without_typed_numbers_and_the_result_validates()
    {
        using var tmp = new TempDir();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        var output = tmp.File("legacy-templated.docx");

        var result = Retemplater.Retemplate(LegacyHeadingsTemplate(tmp), input, output, Seed, NoFonts(tmp));

        Assert.Equal(
        [
            ("Heading1", "Abbreviations"), ("Heading1", "Summary"), ("Heading1", "Background"),
            ("Heading1", "Aims"), ("Heading1", "Data"), ("Heading2", "Frog Pond"), ("Heading2", "Birds"),
            ("Heading3", "Detail"), ("Heading1", "Appendix 1 Species"),
        ], Headings(output));
        Assert.Equal(9, result.HeadingsRemapped);
        Assert.Equal(7, result.NumbersDropped);
        Assert.Equal(1, result.Sections);
        Assert.False(result.CoverBlockInserted);
        Assert.Null(result.Fonts);
        AssertValid(output);
    }

    [Fact]
    public void A_template_without_a_heading_map_remaps_nothing()
    {
        using var tmp = new TempDir();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        var output = tmp.File("out.docx");

        var result = Retemplater.Retemplate(ExampleTemplate(), input, output, Seed, NoFonts(tmp));

        Assert.Equal(0, result.HeadingsRemapped);
        Assert.Equal(0, result.NumbersDropped);
        Assert.Empty(Headings(output));
    }

    [Fact]
    public void The_cover_anchor_replaces_the_typed_cover_block_above_it()
    {
        using var tmp = new TempDir();
        var source = LegacyHeadingsTemplate(tmp, "coverAnchor: Registered office");
        var input = new DocxBuilder()
            .Para("Typed title")
            .Para("Prepared for someone")
            .Para("Registered office: 1 Example Street")
            .Heading("Head1", "1 Introduction")
            .Save(tmp.File("cover.docx"));
        var output = tmp.File("out.docx");

        var result = Retemplater.Retemplate(source, input, output, Seed, NoFonts(tmp));

        Assert.True(result.CoverBlockInserted);
        Assert.Equal(2, result.CoverParagraphsReplaced);
        using var doc = Docx.OpenRead(output);
        var body = Docx.Load(doc.MainDocumentPart!).Root!.Element(Ns.W + "body")!;
        Assert.Equal("Seed title", Wml.VisibleText(body.Element(Ns.W + "sdt")!.Descendants(Ns.W + "p").Single()));
        Assert.DoesNotContain(body.Descendants(Ns.W + "p"), p => Wml.VisibleText(p) == "Typed title");
        Assert.Contains(body.Elements(Ns.W + "p"), p => Wml.VisibleText(p).StartsWith("Registered office", StringComparison.Ordinal));
        AssertValid(output);
    }

    [Fact]
    public void A_document_without_identity_is_seeded_from_the_metadata()
    {
        using var tmp = new TempDir();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        var output = tmp.File("out.docx");

        var result = Retemplater.Retemplate(ExampleTemplate(), input, output, Seed, NoFonts(tmp));

        Assert.Equal(IdentitySources.None, result.IdentitySources);
        Assert.Empty(result.IdentityKeys);
        using var doc = Docx.OpenRead(output);
        Assert.Equal(("Seed title", "Seed Author"), CoreTitleCreator(doc));
        Assert.Equal("EX-999", CustomProperty(doc, "ProjectCode"));
    }

    [Fact]
    public void The_documents_own_identity_wins_over_the_seed()
    {
        using var tmp = new TempDir();
        var source = ExampleTemplate();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        using (var doc = WordprocessingDocument.Open(input, true))
            DocPropsWriter.Write(doc, ReportMetadata.From([new("title", "Pilot study"), new("author", "An Author")]), source.Manifest);
        var output = tmp.File("out.docx");

        var result = Retemplater.Retemplate(source, input, output, Seed, NoFonts(tmp));

        Assert.Equal(IdentitySources.CoreProperties, result.IdentitySources);
        Assert.Equal(["title", "author"], result.IdentityKeys);
        using var outDoc = Docx.OpenRead(output);
        Assert.Equal(("Pilot study", "An Author"), CoreTitleCreator(outDoc));
        Assert.Equal("EX-999", CustomProperty(outDoc, "ProjectCode"));   // the document had none; the seed fills it
        var store = Assert.Single(ReportStores(outDoc.MainDocumentPart!, source.Manifest));
        Assert.Equal("Pilot study", store.Element(store.Name.Namespace + "title")!.Value);
        AssertValid(output);
    }

    [Fact]
    public void The_input_is_never_modified()
    {
        using var tmp = new TempDir();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        var before = SHA256.HashData(File.ReadAllBytes(input));

        Retemplater.Retemplate(ExampleTemplate(), input, tmp.File("out.docx"), Seed, NoFonts(tmp));

        Assert.Equal(before, SHA256.HashData(File.ReadAllBytes(input)));
        Assert.Throws<ArgumentException>(() => Retemplater.Retemplate(ExampleTemplate(), input, input, Seed, NoFonts(tmp)));
    }

    [Fact]
    public void Retemplating_a_generated_document_reuses_its_parts_instead_of_duplicating_them()
    {
        using var tmp = new TempDir();
        var source = DerivedTemplate(tmp, WithFontsYaml);
        var fontDir = tmp.Dir("fonts");
        File.WriteAllBytes(Path.Combine(fontDir, "TestSans-Regular.otf"), [.. Enumerable.Range(0, 64).Select(i => (byte)i)]);
        var generated = Generator.Generate(source, WriteMetadata(tmp, "report.yaml", ReportYaml(tmp)), tmp.Dir("gen"), fontDir);
        var output = tmp.File("again.docx");

        var result = Retemplater.Retemplate(source, generated.SamplePath, output, Seed, fontDir);

        // The document's store wins over the seed, and is rewritten in place rather than added to.
        Assert.True(result.IdentitySources.HasFlag(IdentitySources.DataStore));
        Assert.Equal("EX-999", result.Metadata.Get("projectCode"));
        Assert.Equal("Test report & friends", result.Metadata.Get("title"));
        Assert.Equal(3, result.StaleHeaderFooterParts);
        Assert.Equal(["TestSans-Regular.otf"], result.Fonts!.Embedded);

        using var doc = Docx.OpenRead(output);
        var parts = AllParts(doc);
        var core = Assert.Single(parts, p => p is CoreFilePropertiesPart);
        Assert.Equal("/docProps/core.xml", core.Uri.ToString());
        Assert.Single(parts, p => p is CustomFilePropertiesPart);
        Assert.Single(ReportStores(doc.MainDocumentPart!, source.Manifest));
        Assert.Single(parts, p => p is FontPart);
        Assert.Equal(2, doc.MainDocumentPart!.HeaderParts.Count());
        Assert.Single(doc.MainDocumentPart!.FooterParts);
        AssertValid(output);
    }

    [Fact]
    public void Typed_numbers_follow_the_shared_Core_definition_not_the_legacy_dotted_regex()
    {
        using var tmp = new TempDir();
        var input = new DocxBuilder()
            .Heading("Head1", "4. Results")
            .Heading("Head1", "1 Introduction")
            .Heading("Head2", "3.1\tFrog Pond")
            .Heading("Head1", "2025 Field season")
            .Save(tmp.File("numbers.docx"));
        var output = tmp.File("out.docx");

        var result = Retemplater.Retemplate(LegacyHeadingsTemplate(tmp), input, output, Seed, NoFonts(tmp));

        // The legacy regex required a dot inside the number, so it kept "4. " and "1 ".
        Assert.Equal(
            [("Heading1", "Results"), ("Heading1", "Introduction"), ("Heading2", "Frog Pond"), ("Heading1", "2025 Field season")],
            Headings(output));
        Assert.Equal(3, result.NumbersDropped);
    }

    [Fact]
    public void Retemplate_keeps_the_documents_edit_history_in_core_properties()
    {
        using var tmp = new TempDir();
        var input = DocxBuilder.TypedNumberReport(tmp.File("legacy.docx"));
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        using (var doc = WordprocessingDocument.Open(input, true))
            Docx.Save(doc.AddCoreFilePropertiesPart(), new XDocument(new XElement(cp + "coreProperties",
                new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName), new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName),
                new XElement(dc + "creator", "account-name"),
                new XElement(cp + "lastModifiedBy", "Last Editor"),
                new XElement(cp + "revision", "406"))));
        var output = tmp.File("out.docx");

        Retemplater.Retemplate(ExampleTemplate(), input, output, Seed, NoFonts(tmp));

        using var outDoc = Docx.OpenRead(output);
        var core = Docx.Load(outDoc.CoreFilePropertiesPart!).Root!;
        Assert.Equal("Last Editor", core.Element(cp + "lastModifiedBy")?.Value);
        Assert.Equal("406", core.Element(cp + "revision")?.Value);
        Assert.Equal("Seed title", core.Element(dc + "title")?.Value);
        AssertValid(output);
    }

    [Fact]
    public void Generate_still_stamps_revision_and_last_modified_by_on_the_documents_it_creates()
    {
        using var tmp = new TempDir();
        var result = Generator.Generate(ExampleTemplate(), WriteMetadata(tmp, "report.yaml", ReportYaml(tmp)), tmp.Dir("out"), NoFonts(tmp));

        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        using var doc = Docx.OpenRead(result.SamplePath);
        var core = Docx.Load(doc.CoreFilePropertiesPart!).Root!;
        Assert.Equal("Dr Test Author", core.Element(cp + "lastModifiedBy")?.Value);
        Assert.Equal("1", core.Element(cp + "revision")?.Value);
    }
}
