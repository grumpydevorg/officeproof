using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>State of the full-bleed cover image in the first-page header.</summary>
public enum CoverImageStatus
{
    /// <summary>Metadata names no <c>coverImage</c>; the first page has no image.</summary>
    NotConfigured,
    /// <summary>The image was embedded.</summary>
    Embedded,
    /// <summary>Metadata names an image that does not exist; the first page is left blank rather than failing the build.</summary>
    NotFound,
}

/// <summary>Outcome of <see cref="HeaderFooterBuilder.AddCoverHeader"/>.</summary>
/// <param name="RelationshipId">Id of the first-page header, for <c>w:headerReference w:type="first"</c>.</param>
/// <param name="Image">The <c>coverImage</c> value from metadata, when set.</param>
/// <param name="Status">Whether it was embedded.</param>
public sealed record CoverHeaderResult(string RelationshipId, string? Image, CoverImageStatus Status);

/// <summary>Builds the template's header and footer parts. Each returns the relationship id a <c>w:sectPr</c> reference needs.</summary>
public static class HeaderFooterBuilder
{
    /// <summary>The running header (<see cref="TemplateSource.Header"/>), its controls data-bound to the store <paramref name="storeId"/>.</summary>
    public static string AddRunningHeader(MainDocumentPart main, TemplateSource source, ReportMetadata metadata, string storeId)
    {
        var header = main.AddNewPart<HeaderPart>();
        PartTemplate.WriteText(header, PartTemplate.Fill(source.ReadText(TemplateSource.Header), metadata, ("STOREID", storeId)));
        return main.GetIdOfPart(header);
    }

    /// <summary>
    /// The footer (<see cref="TemplateSource.Footer"/>); when the manifest names a logo, the image
    /// is added and <see cref="TemplateSource.LogoToken"/> becomes its relationship id.
    /// </summary>
    /// <exception cref="InvalidOperationException">The logo is not .png, .jpg or .jpeg.</exception>
    public static string AddFooter(MainDocumentPart main, TemplateSource source)
    {
        var footer = main.AddNewPart<FooterPart>();
        var xml = source.ReadText(TemplateSource.Footer);
        if (source.Manifest.Logo is { } logoFile)
        {
            var logo = footer.AddImagePart(ImageType(logoFile, "logo"));
            Docx.Write(logo, source.ReadBytes(logoFile));
            xml = xml.Replace(TemplateSource.LogoToken, footer.GetIdOfPart(logo), StringComparison.Ordinal);
        }
        PartTemplate.WriteText(footer, xml);
        return main.GetIdOfPart(footer);
    }

    /// <summary>
    /// First-page header: carries the full-bleed cover image and no running header.
    /// </summary>
    /// <remarks>
    /// The image lives here rather than in the body because the header layer draws beneath body
    /// text, so the title flows over it with no wrapping to fight.
    /// <para>
    /// Metadata <c>coverImage</c> names the asset, relative to the template directory (or
    /// absolute). If it is absent or the file is missing, the header degrades to an empty first
    /// page rather than failing the build, and the result says so.
    /// </para>
    /// </remarks>
    /// <exception cref="InvalidOperationException">The image is not .png, .jpg or .jpeg.</exception>
    public static CoverHeaderResult AddCoverHeader(MainDocumentPart main, TemplateSource source, ReportMetadata metadata)
    {
        var xml = source.ReadText(TemplateSource.FirstPageHeader);
        var asset = metadata.Get("coverImage");
        var assetPath = asset is null ? null : source.PathOf(asset);

        CoverImageStatus status;
        HeaderPart header;
        if (assetPath is not null && File.Exists(assetPath))
        {
            var type = ImageType(assetPath, "coverImage");
            header = main.AddNewPart<HeaderPart>();
            var image = header.AddImagePart(type);
            Docx.Write(image, File.ReadAllBytes(assetPath));
            PartTemplate.WriteText(header, xml.Replace("{{COVERIMAGE}}", header.GetIdOfPart(image), StringComparison.Ordinal));
            status = CoverImageStatus.Embedded;
        }
        else
        {
            // Drop the whole drawing rather than leave a dangling r:embed: the header keeps one
            // empty paragraph, which a w:hdr requires.
            header = main.AddNewPart<HeaderPart>();
            var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var root = doc.Root!;
            root.Elements().Remove();
            root.Add(new XElement(Ns.W + "p"));
            Docx.Save(header, doc);
            status = assetPath is null ? CoverImageStatus.NotConfigured : CoverImageStatus.NotFound;
        }
        return new CoverHeaderResult(main.GetIdOfPart(header), asset, status);
    }

    private static PartTypeInfo ImageType(string path, string what) => Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".png" => ImagePartType.Png,
        ".jpg" or ".jpeg" => ImagePartType.Jpeg,
        var e => throw new InvalidOperationException($"{what}: unsupported type '{e}'"),
    };
}
