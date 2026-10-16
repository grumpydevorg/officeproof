using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.Templating.Tests.TestFiles;

namespace OfficeProof.Templating.Tests;

public sealed class FontEmbedderTests
{
    private static FontEmbedder TestSans(params string[] fontDirectories) => new(fontDirectories, "Test Sans",
        [.. TestFaces.Select((f, i) => new FontFace(f, (FontSlot)i))]);

    [Fact]
    public void An_empty_font_directory_reports_every_face_missing_without_throwing()
    {
        using var tmp = new TempDir();
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using var doc = WordprocessingDocument.Open(path, true);

        var result = TestSans(NoFonts(tmp)).Embed(doc.MainDocumentPart!);

        Assert.Empty(result.Embedded);
        Assert.Equal(TestFaces, result.Missing);
        Assert.Empty(doc.MainDocumentPart!.FontTablePart!.FontParts);
    }

    [Fact]
    public void Obfuscate_xors_the_first_32_bytes_with_the_reversed_key_and_is_its_own_inverse()
    {
        var key = Guid.Parse("00112233-4455-6677-8899-AABBCCDDEEFF");
        var font = new byte[40];

        var obfuscated = FontEmbedder.Obfuscate(font, key);

        Assert.Equal(0xFF, obfuscated[0]);
        Assert.Equal(0xEE, obfuscated[1]);
        Assert.Equal(0x00, obfuscated[15]);
        Assert.Equal(0xFF, obfuscated[16]);
        Assert.All(obfuscated[32..], b => Assert.Equal(0, b));
        Assert.Equal(font, FontEmbedder.Obfuscate(obfuscated, key));
    }

    [Fact]
    public void An_embedded_face_round_trips_through_its_font_key()
    {
        using var tmp = new TempDir();
        var fontDir = tmp.Dir("fonts");
        byte[] face = [.. Enumerable.Range(0, 100).Select(i => (byte)(i * 7))];
        File.WriteAllBytes(Path.Combine(fontDir, "TestSans-Bold.otf"), face);
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using var doc = WordprocessingDocument.Open(path, true);

        var result = TestSans(fontDir).Embed(doc.MainDocumentPart!);

        Assert.Equal(["TestSans-Bold.otf"], result.Embedded);
        var table = doc.MainDocumentPart!.FontTablePart!;
        var embed = Assert.Single(Docx.Load(table).Descendants(Ns.W + "embedBold"));
        var part = (FontPart)table.GetPartById((string)embed.Attribute(Ns.R + "id")!);
        Assert.Equal("application/vnd.openxmlformats-officedocument.obfuscatedFont", part.ContentType);
        using var s = part.GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        var key = Guid.Parse(((string)embed.Attribute(Ns.W + "fontKey")!).Trim('{', '}'));
        Assert.Equal(face, FontEmbedder.Obfuscate(ms.ToArray(), key));
    }

    [Fact]
    public void A_face_comes_from_the_first_directory_holding_it_subdirectories_included()
    {
        using var tmp = new TempDir();
        var user = tmp.Dir("user");
        var system = tmp.Dir("system");
        var nested = tmp.Dir(Path.Combine("system", "truetype", "test-sans"));
        byte[] userFace = [1], systemFace = [2], nestedFace = [3];
        File.WriteAllBytes(Path.Combine(user, "TestSans-Regular.otf"), userFace);
        File.WriteAllBytes(Path.Combine(system, "TestSans-Regular.otf"), systemFace);
        File.WriteAllBytes(Path.Combine(nested, "TestSans-Bold.otf"), nestedFace);
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using var doc = WordprocessingDocument.Open(path, true);

        var result = TestSans(tmp.File("absent"), user, system).Embed(doc.MainDocumentPart!);

        Assert.Equal(["TestSans-Regular.otf", "TestSans-Bold.otf"], result.Embedded);
        Assert.Equal(["TestSans-Italic.otf", "TestSans-BoldItalic.otf"], result.Missing);
        var table = doc.MainDocumentPart!.FontTablePart!;
        var regular = Assert.Single(Docx.Load(table).Descendants(Ns.W + "embedRegular"));
        using var s = table.GetPartById((string)regular.Attribute(Ns.R + "id")!).GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        var key = Guid.Parse(((string)regular.Attribute(Ns.W + "fontKey")!).Trim('{', '}'));
        Assert.Equal(userFace, FontEmbedder.Obfuscate(ms.ToArray(), key));
    }

    [Fact]
    public void The_default_font_directories_are_the_users_then_the_machines()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var dirs = FontEmbedder.DefaultFontDirectories;

        if (OperatingSystem.IsWindows())
        {
            // Windows paths compare case-insensitively (%WINDIR% is C:\WINDOWS on some machines).
            Assert.Equal([
                Path.Combine(Environment.GetEnvironmentVariable("LOCALAPPDATA")!, "Microsoft", "Windows", "Fonts"),
                Path.Combine(Environment.GetEnvironmentVariable("WINDIR")!, "Fonts")], dirs, StringComparer.OrdinalIgnoreCase);
        }
        else if (OperatingSystem.IsMacOS())
        {
            Assert.Equal([Path.Combine(home, "Library", "Fonts"), "/Library/Fonts"], dirs);
        }
        else
        {
            var dataHome = Environment.GetEnvironmentVariable("XDG_DATA_HOME") is { } x && Path.IsPathFullyQualified(x)
                ? x
                : Path.Combine(home, ".local", "share");
            Assert.Equal([Path.Combine(dataHome, "fonts"), "/usr/local/share/fonts", "/usr/share/fonts"], dirs);
        }
    }
}

public sealed class SettingsInjectorTests
{
    private static readonly XNamespace W = Ns.W;

    [Fact]
    public void Settings_are_inserted_at_their_schema_positions()
    {
        using var tmp = new TempDir();
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var part = doc.MainDocumentPart!.DocumentSettingsPart!;
            Docx.Save(part, new XDocument(new XElement(W + "settings", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                new XElement(W + "zoom", new XAttribute(W + "percent", 100)),
                new XElement(W + "defaultTabStop", new XAttribute(W + "val", 720)),
                new XElement(W + "compat"))));
            var change = SettingsInjector.Apply(doc.MainDocumentPart!);
            Assert.Equal(new SettingsChange(true, true), change);
        }

        using (var doc = Docx.OpenRead(path))
        {
            var names = Docx.Load(doc.MainDocumentPart!.DocumentSettingsPart!).Root!.Elements().Select(e => e.Name.LocalName);
            Assert.Equal(["zoom", "embedTrueTypeFonts", "defaultTabStop", "updateFields", "compat"], names);
        }
        AssertValid(path);
    }

    [Fact]
    public void Existing_settings_are_left_alone()
    {
        using var tmp = new TempDir();
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using var doc = WordprocessingDocument.Open(path, true);
        SettingsInjector.Apply(doc.MainDocumentPart!);

        Assert.Equal(new SettingsChange(false, false), SettingsInjector.Apply(doc.MainDocumentPart!));
    }
}

public sealed class IdentityReaderTests
{
    private static readonly TemplateManifest Manifest = ExampleTemplate().Manifest;

    private static DocumentIdentity ReadAfter(TempDir tmp, ReportMetadata props)
    {
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using (var doc = WordprocessingDocument.Open(path, true))
            DocPropsWriter.Write(doc, props, Manifest);
        using var read = Docx.OpenRead(path);
        return IdentityReader.Read(read, Manifest);
    }

    [Fact]
    public void A_creator_alone_is_not_an_identity_title_and_project_code_are_missing()
    {
        using var tmp = new TempDir();

        // A real report's only identity was dc:creator, an account name.
        var identity = ReadAfter(tmp, ReportMetadata.From([new("author", "account-name")]));

        Assert.Equal(IdentitySources.CoreProperties, identity.Sources);
        Assert.Equal(["title", "projectCode"], identity.Missing(Manifest.IdentityKeys));
    }

    [Fact]
    public void Missing_lists_only_the_asked_keys_the_document_lacks_in_the_order_asked()
    {
        using var tmp = new TempDir();

        var identity = ReadAfter(tmp, ReportMetadata.From([new("title", "Pilot"), new("projectCode", "EX-039")]));

        Assert.Empty(identity.Missing(Manifest.IdentityKeys));
        Assert.Equal(["version", "preparedFor"], identity.Missing("version", "title", "preparedFor"));
        Assert.Equal(IdentitySources.CustomProperties | IdentitySources.CoreProperties, identity.Sources);
    }

    [Fact]
    public void A_blank_value_counts_as_missing()
    {
        using var tmp = new TempDir();

        var identity = ReadAfter(tmp, ReportMetadata.From([new("title", "  "), new("projectCode", "EX-039")]));

        Assert.Equal(["title"], identity.Missing(Manifest.IdentityKeys));
    }

    [Fact]
    public void Only_properties_the_manifest_binds_are_read_back()
    {
        using var tmp = new TempDir();
        var path = DocxBuilder.TypedNumberReport(tmp.File("doc.docx"));
        using (var doc = WordprocessingDocument.Open(path, true))
            DocPropsWriter.Write(doc, ReportMetadata.From([new("title", "Bound title"), new("projectCode", "EX-1")]), Manifest);
        var other = TemplateManifest.Parse("""
            name: Other
            dataStore: {namespace: "urn:other", prefix: o}
            coreProperties: {title: Fixed text}
            """);

        using var read = Docx.OpenRead(path);
        var identity = IdentityReader.Read(read, other);

        Assert.Equal(IdentitySources.None, identity.Sources);
        Assert.Empty(identity.Values.Keys);
    }
}

public sealed class ReportMetadataTests
{
    [Fact]
    public void Load_reads_the_report_map_from_any_path_in_order()
    {
        using var tmp = new TempDir();
        var meta = WriteMetadata(tmp, "nested-name.yml", """
            # comment
            other: ignored
            report:
              title: "A & B"
              projectCode: EX-001
              version:
            """);

        Assert.Equal(["title", "projectCode", "version"], meta.Keys);
        Assert.Equal("A & B", meta.Get("title"));
        Assert.Null(meta.Get("version"));
        Assert.Equal(tmp.File("nested-name.yml"), meta.Source);
    }

    [Theory]
    [InlineData("title: x")]
    [InlineData("report: just-a-string")]
    [InlineData("report:\n  list: [1, 2]")]
    [InlineData("report:\n  \"bad key\": x")]
    [InlineData("report: [unclosed")]
    public void Malformed_metadata_is_rejected_with_the_file_named(string yaml)
    {
        var e = Assert.Throws<InvalidDataException>(() => ReportMetadata.Parse(yaml, "case.yaml"));
        Assert.StartsWith("case.yaml:", e.Message);
    }

    [Fact]
    public void With_overrides_in_place_and_appends_new_keys()
    {
        var seed = ReportMetadata.From([new("title", "seed"), new("version", "1")]);

        var merged = seed.With([new("title", "doc"), new("author", "a")]);

        Assert.Equal([new("title", "doc"), new("version", "1"), new("author", "a")], merged.Entries);
        Assert.Equal("seed", seed.Get("title"));
    }
}

public sealed class TemplateManifestTests
{
    private const string Full = """
        name: House
        dataStore:
          namespace: https://example.org/house
          prefix: h
        logo: logo.png
        fonts:
          family: House Sans
          regular: HouseSans-Regular.otf
          boldItalic: HouseSans-BoldItalic.otf
        coreProperties:
          title: "{{title}}"
          category: Report
        customProperties:
          Reference: "{{projectCode}}"
          Department: Research
        identityKeys: [title, projectCode]
        headingStyles:
          OldHeading1: Heading1
        coverAnchor: Registered office
        palette:
          colours:
            "#1F4E79": dark blue
            "000000": black
          nearMisses:
            1f4e78: one digit out
        """;

    [Fact]
    public void Every_section_is_read()
    {
        var m = TemplateManifest.Parse(Full);

        Assert.Equal(("House.dotx", "House-sample.docx"), (m.TemplateFileName, m.SampleFileName));
        Assert.Equal(XName.Get("report", "https://example.org/house"), m.DataStore.Root);
        Assert.Equal("h", m.DataStore.Prefix);
        Assert.Equal("logo.png", m.Logo);
        Assert.Equal("House Sans", m.Fonts!.Family);
        Assert.Equal([new FontFace("HouseSans-Regular.otf", FontSlot.Regular), new FontFace("HouseSans-BoldItalic.otf", FontSlot.BoldItalic)], m.Fonts.Faces);
        Assert.Equal([new PropertyBinding("title", "title", null), new PropertyBinding("category", null, "Report")], m.CoreProperties);
        Assert.Equal([new PropertyBinding("Reference", "projectCode", null), new PropertyBinding("Department", null, "Research")], m.CustomProperties);
        Assert.Equal(["title", "projectCode"], m.IdentityKeys);
        Assert.Equal("Heading1", m.HeadingStyles["OldHeading1"]);
        Assert.Equal("Registered office", m.CoverAnchor);
        Assert.Equal(new Dictionary<string, string> { ["1f4e79"] = "dark blue", ["000000"] = "black" }, m.Palette!.Colours);
        Assert.Equal("one digit out", m.Palette.NearMisses["1f4e78"]);
    }

    [Fact]
    public void Only_name_and_data_store_are_required()
    {
        var m = TemplateManifest.Parse("name: Bare\ndataStore: {namespace: \"urn:bare\", prefix: b}");

        Assert.Null(m.Logo);
        Assert.Null(m.Fonts);
        Assert.Null(m.Palette);
        Assert.Null(m.CoverAnchor);
        Assert.Empty(m.CoreProperties);
        Assert.Empty(m.IdentityKeys);
        Assert.Empty(m.HeadingStyles);
    }

    [Fact]
    public void A_property_binding_resolves_from_metadata_or_fixed_text()
    {
        var m = TemplateManifest.Parse(Full);
        var meta = ReportMetadata.From([new("projectCode", "EX-7")]);

        Assert.Equal(["EX-7", "Research"], m.CustomProperties.Select(b => b.Resolve(meta)));
        Assert.Null(m.CoreProperties[0].Resolve(meta));
    }

    [Theory]
    [InlineData("dataStore: {namespace: n, prefix: p}", "name is required")]
    [InlineData("name: X", "dataStore is required")]
    [InlineData("name: X\ndataStore: {namespace: n, prefix: \"bad prefix\"}", "not a valid XML prefix")]
    [InlineData("name: X\ndataStore: {namespace: n, prefix: p}\nheadingStyle: {}", "unknown key 'headingStyle'")]
    [InlineData("name: X\ndataStore: {namespace: n, prefix: p}\npalette: {colours: {abc: short}}", "'abc' is not a 6-digit hex colour")]
    [InlineData("name: X\ndataStore: {namespace: n, prefix: p}\ncoreProperties: {revision: \"1\"}", "revision is not a core property")]
    [InlineData("name: X\ndataStore: {namespace: n, prefix: p}\nfonts: {regular: a.otf}", "fonts.family is required")]
    [InlineData("name: a/b\ndataStore: {namespace: n, prefix: p}", "not usable as a file name")]
    public void Malformed_manifests_are_rejected_with_the_file_and_key_named(string yaml, string expected)
    {
        var e = Assert.Throws<InvalidDataException>(() => TemplateManifest.Parse(yaml, "case.yaml"));

        Assert.StartsWith("case.yaml:", e.Message);
        Assert.Contains(expected, e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void The_example_template_has_a_palette_covering_every_explicit_colour_in_its_styles()
    {
        var source = ExampleTemplate();
        var colours = XDocument.Parse(source.ReadText(TemplateSource.Styles)).Descendants()
            .SelectMany(e => e.Attributes().Where(a => a.Name.LocalName is "color" or "fill" || (e.Name == Ns.W + "color" && a.Name == Ns.W + "val")))
            .Select(a => a.Value.ToLowerInvariant())
            .Where(v => v != "auto")
            .ToHashSet();

        Assert.NotEmpty(colours);
        Assert.Subset(source.Manifest.Palette!.Colours.Keys.ToHashSet(), colours);
    }
}

public sealed class TemplateSourceTests
{
    [Fact]
    public void A_directory_missing_parts_names_them_all()
    {
        using var tmp = new TempDir();
        File.WriteAllText(tmp.File(TemplateSource.Styles), "<x/>");

        var e = Assert.Throws<FileNotFoundException>(() => new TemplateSource(tmp.Path));

        Assert.Contains(TemplateManifest.FileName, e.Message);
        Assert.Contains(TemplateSource.Theme, e.Message);
        Assert.DoesNotContain(TemplateSource.Styles + ",", e.Message);
    }

    [Fact]
    public void Files_the_manifest_names_must_exist()
    {
        using var tmp = new TempDir();

        var e = Assert.Throws<FileNotFoundException>(() => DerivedTemplate(tmp, "logo: missing-logo.png\ncoverAnchor: Registered office"));

        Assert.Contains("missing-logo.png", e.Message);
        Assert.DoesNotContain(TemplateSource.CoverBlock, e.Message);   // the example ships one
    }

    [Fact]
    public void A_footer_that_places_a_logo_needs_one_in_the_manifest()
    {
        using var tmp = new TempDir();
        var dir = DerivedDirectory(tmp, "");
        File.WriteAllText(Path.Combine(dir, TemplateSource.Footer), "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><!-- {{LOGOID}} --></w:ftr>");

        var e = Assert.Throws<InvalidDataException>(() => new TemplateSource(dir));

        Assert.Contains("names no 'logo'", e.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_template_without_fonts_has_no_embedder()
    {
        Assert.Null(ExampleTemplate().FontEmbedder());
    }
}
