using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Templating.Tests;

/// <summary>Per-test scratch directory under the system temp path, removed on dispose.</summary>
internal sealed class TempDir : IDisposable
{
    public TempDir() => Directory.CreateDirectory(Path);

    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeproof-templating-" + Guid.NewGuid().ToString("N"));

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public string Dir(string name) => Directory.CreateDirectory(File(name)).FullName;

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path, recursive: true);
        }
        catch (IOException)
        {
            // best effort: a locked file must not fail the test that produced it
        }
    }
}

internal static class TestFiles
{
    /// <summary>The repository's neutral example template, found by walking up from the test assembly to officeproof.slnx.</summary>
    public static string ExampleDirectory
    {
        get
        {
            for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
                if (System.IO.File.Exists(Path.Combine(dir.FullName, "officeproof.slnx")))
                    return Path.Combine(dir.FullName, "examples", "minimal-template");
            throw new InvalidOperationException($"no officeproof.slnx above {AppContext.BaseDirectory}");
        }
    }

    public static TemplateSource ExampleTemplate() => new(ExampleDirectory);

    /// <summary>A copy of the example template in <paramref name="tmp"/> with <paramref name="manifestYaml"/> appended to its template.yaml.</summary>
    public static TemplateSource DerivedTemplate(TempDir tmp, string manifestYaml) => new(DerivedDirectory(tmp, manifestYaml));

    /// <inheritdoc cref="DerivedTemplate"/>
    public static string DerivedDirectory(TempDir tmp, string manifestYaml)
    {
        var dir = tmp.Dir("template-" + Guid.NewGuid().ToString("N"));
        foreach (var file in Directory.GetFiles(ExampleDirectory))
            System.IO.File.Copy(file, Path.Combine(dir, Path.GetFileName(file)));
        System.IO.File.AppendAllText(Path.Combine(dir, TemplateManifest.FileName), "\n" + manifestYaml + "\n");
        return dir;
    }

    /// <summary>The example template plus a map for <see cref="DocxBuilder"/>'s legacy heading styles.</summary>
    public static TemplateSource LegacyHeadingsTemplate(TempDir tmp, string extraYaml = "") => DerivedTemplate(tmp, """
        headingStyles:
          Head1: Heading1
          Head2: Heading2
          Head3: Heading3
        """ + "\n" + extraYaml);

    /// <summary>Font faces a <see cref="WithFontsYaml"/> template embeds, in slot order.</summary>
    public static readonly string[] TestFaces = ["TestSans-Regular.otf", "TestSans-Bold.otf", "TestSans-Italic.otf", "TestSans-BoldItalic.otf"];

    public const string WithFontsYaml = """
        fonts:
          family: Test Sans
          regular: TestSans-Regular.otf
          bold: TestSans-Bold.otf
          italic: TestSans-Italic.otf
          boldItalic: TestSans-BoldItalic.otf
        """;

    public static ReportMetadata WriteMetadata(TempDir tmp, string name, string yaml)
    {
        var path = tmp.File(name);
        System.IO.File.WriteAllText(path, yaml);
        return ReportMetadata.Load(path);
    }

    /// <summary>Report metadata whose cover image is a 1x1 PNG written into <paramref name="tmp"/>.</summary>
    public static string ReportYaml(TempDir tmp)
    {
        var cover = tmp.File("cover.png");
        System.IO.File.WriteAllBytes(cover, DocxBuilder.PngA);
        return $"""
            report:
              title: "Test report & friends"
              preparedFor: "Client Ltd"
              preparedBy: "Example Consulting Ltd"
              projectCode: "EX-999"
              version: "2026-10-15"
              author: "Dr Test Author"
              coverImage: "{cover}"
            """;
    }

    /// <summary>A font directory with no faces in it.</summary>
    public static string NoFonts(TempDir tmp) => tmp.Dir("no-fonts");

    public static void AssertValid(string path)
    {
        var errors = Docx.Validate(path);
        Assert.True(errors.Count == 0, $"{System.IO.Path.GetFileName(path)}: {errors.Count} schema error(s):\n"
            + string.Join("\n", errors.Take(10).Select(e => $"{e.Message} [{e.Part} {e.Location}]")));
    }

    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";

    public static (string? Title, string? Creator) CoreTitleCreator(WordprocessingDocument doc)
    {
        var xml = Docx.Load(doc.CoreFilePropertiesPart!);
        return (xml.Descendants(Dc + "title").SingleOrDefault()?.Value, xml.Descendants(Dc + "creator").SingleOrDefault()?.Value);
    }

    public static string? CustomProperty(WordprocessingDocument doc, string name) =>
        doc.CustomFilePropertiesPart?.Properties?
            .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .SingleOrDefault(p => p.Name?.Value == name)?.VTLPWSTR?.Text;

    /// <summary>Every part in the package, by URI, reached through relationships.</summary>
    public static List<OpenXmlPart> AllParts(OpenXmlPackage package)
    {
        var seen = new HashSet<Uri>();
        var result = new List<OpenXmlPart>();
        void Walk(OpenXmlPartContainer c)
        {
            foreach (var pair in c.Parts)
                if (seen.Add(pair.OpenXmlPart.Uri))
                {
                    result.Add(pair.OpenXmlPart);
                    Walk(pair.OpenXmlPart);
                }
        }
        Walk(package);
        return result;
    }

    public static List<XElement> ReportStores(MainDocumentPart main, TemplateManifest manifest) =>
        main.CustomXmlParts.Select(p => Docx.Load(p).Root!)
            .Where(r => r.Name == manifest.DataStore.Root).ToList();
}
