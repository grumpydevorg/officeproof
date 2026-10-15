using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>Outcome of <see cref="Generator.Generate"/>.</summary>
/// <param name="SamplePath">The reference sample document (.docx).</param>
/// <param name="TemplatePath">The Word template (.dotx).</param>
/// <param name="Fonts">Font embedding outcome (identical for both outputs); null when the template embeds no fonts.</param>
/// <param name="Cover">Cover image outcome (identical for both outputs).</param>
public sealed record GenerateResult(string SamplePath, string TemplatePath, FontEmbedResult? Fonts, CoverHeaderResult Cover);

/// <summary>Emits a Word template and a reference sample from a <see cref="TemplateSource"/> and report metadata.</summary>
public static class Generator
{
    /// <summary>
    /// Writes the manifest's <see cref="TemplateManifest.SampleFileName"/> and
    /// <see cref="TemplateManifest.TemplateFileName"/> into <paramref name="outputDirectory"/>
    /// (created if needed, existing files replaced).
    /// </summary>
    /// <param name="source">The canonical parts.</param>
    /// <param name="metadata">Cover, header, data store and property values.</param>
    /// <param name="outputDirectory">Where to write.</param>
    /// <param name="fontDirectory">Where the manifest's font faces are; default <see cref="FontEmbedder.UserFontDirectory"/>.</param>
    public static GenerateResult Generate(TemplateSource source, ReportMetadata metadata, string outputDirectory, string? fontDirectory = null)
    {
        var fonts = source.FontEmbedder(fontDirectory);
        Directory.CreateDirectory(outputDirectory);
        var sample = Path.GetFullPath(Path.Combine(outputDirectory, source.Manifest.SampleFileName));
        var template = Path.GetFullPath(Path.Combine(outputDirectory, source.Manifest.TemplateFileName));
        var (embedded, cover) = Build(source, metadata, fonts, sample, WordprocessingDocumentType.Document);
        Build(source, metadata, fonts, template, WordprocessingDocumentType.Template);
        return new GenerateResult(sample, template, embedded, cover);
    }

    private static (FontEmbedResult?, CoverHeaderResult) Build(
        TemplateSource source, ReportMetadata metadata, FontEmbedder? fonts, string path, WordprocessingDocumentType type)
    {
        if (File.Exists(path)) File.Delete(path);
        using var doc = WordprocessingDocument.Create(path, type);
        var main = doc.AddMainDocumentPart();

        // Assemble document.xml from the cover template + metadata; the data store, the running
        // header and the footer bind or point to it.
        var storeId = DataStoreWriter.Write(main, metadata, source.Manifest.DataStore);
        var headerId = HeaderFooterBuilder.AddRunningHeader(main, source, metadata, storeId);
        var footerId = HeaderFooterBuilder.AddFooter(main, source);
        var cover = HeaderFooterBuilder.AddCoverHeader(main, source, metadata);

        PartTemplate.WriteText(main, PartTemplate.Fill(source.ReadText(TemplateSource.CoverDocument), metadata,
            ("STOREID", storeId), ("HEADERID", headerId), ("FIRSTHEADERID", cover.RelationshipId), ("FOOTERID", footerId)));

        Docx.Write(main.AddNewPart<StyleDefinitionsPart>(), source.ReadBytes(TemplateSource.Styles));
        Docx.Write(main.AddNewPart<ThemePart>(), source.ReadBytes(TemplateSource.Theme));
        Docx.Write(main.AddNewPart<NumberingDefinitionsPart>(), source.ReadBytes(TemplateSource.Numbering));
        // The template's settings.xml is copied as is: it should carry embedTrueTypeFonts and updateFields.
        Docx.Write(main.AddNewPart<DocumentSettingsPart>(), source.ReadBytes(TemplateSource.Settings));
        DocPropsWriter.Write(doc, metadata, source.Manifest);
        var embedded = fonts?.Embed(main);
        return (embedded, cover);
    }
}
