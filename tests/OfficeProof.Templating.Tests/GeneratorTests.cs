using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.Templating.Tests.TestFiles;

namespace OfficeProof.Templating.Tests;

public sealed class GeneratorTests
{
    [Fact]
    public void Generate_from_the_example_template_writes_two_schema_valid_outputs_named_by_the_manifest()
    {
        using var tmp = new TempDir();
        var meta = WriteMetadata(tmp, "report.yaml", ReportYaml(tmp));

        var result = Generator.Generate(ExampleTemplate(), meta, tmp.Dir("out"), NoFonts(tmp));

        Assert.EndsWith("Example-sample.docx", result.SamplePath);
        Assert.EndsWith("Example.dotx", result.TemplatePath);
        AssertValid(result.SamplePath);
        AssertValid(result.TemplatePath);
        using var dotx = WordprocessingDocument.Open(result.TemplatePath, false);
        Assert.Equal(DocumentFormat.OpenXml.WordprocessingDocumentType.Template, dotx.DocumentType);
    }

    [Fact]
    public void Generate_writes_metadata_from_any_path_into_core_custom_and_the_bound_store()
    {
        using var tmp = new TempDir();
        var meta = WriteMetadata(tmp, "elsewhere.yaml", ReportYaml(tmp));
        var source = ExampleTemplate();

        var result = Generator.Generate(source, meta, tmp.Dir("out"), NoFonts(tmp));

        using var doc = Docx.OpenRead(result.SamplePath);
        Assert.Equal(("Test report & friends", "Dr Test Author"), CoreTitleCreator(doc));
        Assert.Equal("EX-999", CustomProperty(doc, "ProjectCode"));
        Assert.Equal("2026-10-15", CustomProperty(doc, "Version"));

        var main = doc.MainDocumentPart!;
        var store = Assert.Single(main.CustomXmlParts, p => Docx.Load(p).Root!.Name == source.Manifest.DataStore.Root);
        Assert.Equal("EX-999", Docx.Load(store).Root!.Element(source.Manifest.DataStore.Root.Namespace + "projectCode")!.Value);

        // The controls bind to the store that exists, in the body and the running header.
        System.Xml.Linq.XNamespace ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
        var itemId = (string?)Docx.Load(store.CustomXmlPropertiesPart!).Root!.Attribute(ds + "itemID");
        Assert.NotNull(itemId);
        var bound = Docx.StoryParts(main).SelectMany(p => Docx.Load(p).Descendants(Ns.W + "dataBinding"))
            .Select(b => (string?)b.Attribute(Ns.W + "storeItemID")).Distinct().ToList();
        Assert.Equal([itemId], bound);

        Assert.Equal(CoverImageStatus.Embedded, result.Cover.Status);
        Assert.Null(result.Fonts);
        Assert.Null(main.FontTablePart);
    }

    [Fact]
    public void A_template_without_a_logo_writes_a_footer_without_an_image()
    {
        using var tmp = new TempDir();

        var result = Generator.Generate(ExampleTemplate(), WriteMetadata(tmp, "report.yaml", ReportYaml(tmp)), tmp.Dir("out"), NoFonts(tmp));

        using var doc = Docx.OpenRead(result.SamplePath);
        var footer = Assert.Single(doc.MainDocumentPart!.FooterParts);
        Assert.Empty(footer.ImageParts);
        Assert.Contains("NUMPAGES", Docx.Load(footer).ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void The_manifest_logo_fills_the_footer_and_its_fonts_are_embedded()
    {
        using var tmp = new TempDir();
        var dir = DerivedDirectory(tmp, WithFontsYaml + "\nlogo: logo.png");
        File.WriteAllBytes(Path.Combine(dir, "logo.png"), TestKit.DocxBuilder.PngB);
        File.WriteAllText(Path.Combine(dir, TemplateSource.Footer), """
            <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1080000" cy="198000"/><wp:docPr id="500" name="Logo"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="500" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="{{LOGOID}}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1080000" cy="198000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            </w:ftr>
            """);
        var source = new TemplateSource(dir);
        var fontDir = tmp.Dir("fonts");
        File.WriteAllBytes(Path.Combine(fontDir, "TestSans-Bold.otf"), [.. Enumerable.Range(0, 64).Select(i => (byte)i)]);

        var result = Generator.Generate(source, WriteMetadata(tmp, "report.yaml", ReportYaml(tmp)), tmp.Dir("out"), fontDir);

        Assert.Equal("Test Sans", result.Fonts!.Family);
        Assert.Equal(["TestSans-Bold.otf"], result.Fonts.Embedded);
        Assert.Equal(3, result.Fonts.Missing.Count);
        using var doc = Docx.OpenRead(result.SamplePath);
        var footer = Assert.Single(doc.MainDocumentPart!.FooterParts);
        var logo = Assert.Single(footer.ImageParts);
        var blip = (string?)Docx.Load(footer).Descendants(System.Xml.Linq.XName.Get("blip", "http://schemas.openxmlformats.org/drawingml/2006/main"))
            .Single().Attribute(Ns.R + "embed");
        Assert.Equal(footer.GetIdOfPart(logo), blip);
        AssertValid(result.SamplePath);
    }

    [Fact]
    public void Generating_twice_leaves_exactly_one_core_properties_part_at_its_canonical_name()
    {
        using var tmp = new TempDir();
        var meta = WriteMetadata(tmp, "report.yaml", ReportYaml(tmp));
        var outDir = tmp.Dir("out");

        Generator.Generate(ExampleTemplate(), meta, outDir, NoFonts(tmp));
        var result = Generator.Generate(ExampleTemplate(), meta, outDir, NoFonts(tmp));

        using var doc = Docx.OpenRead(result.SamplePath);
        var core = Assert.Single(AllParts(doc), p => p is CoreFilePropertiesPart);
        Assert.Equal("/docProps/core.xml", core.Uri.ToString());
    }

    [Fact]
    public void A_missing_cover_image_leaves_a_blank_first_page_and_still_validates()
    {
        using var tmp = new TempDir();
        var yaml = ReportYaml(tmp).Replace(tmp.File("cover.png"), "no-such-cover.jpg", StringComparison.Ordinal);
        var meta = WriteMetadata(tmp, "report.yaml", yaml);

        var result = Generator.Generate(ExampleTemplate(), meta, tmp.Dir("out"), NoFonts(tmp));

        Assert.Equal(CoverImageStatus.NotFound, result.Cover.Status);
        Assert.Equal("no-such-cover.jpg", result.Cover.Image);
        AssertValid(result.SamplePath);
        using var doc = Docx.OpenRead(result.SamplePath);
        var first = (HeaderPart)doc.MainDocumentPart!.GetPartById(result.Cover.RelationshipId);
        Assert.Empty(first.ImageParts);
        Assert.Empty(Docx.Load(first).Descendants(Ns.W + "drawing"));
    }
}
