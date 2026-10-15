namespace OfficeProof.Templating;

/// <summary>
/// A template directory: the canonical source parts that <see cref="Generator"/> and
/// <see cref="Retemplater"/> copy into documents (styles, theme, numbering, settings, the
/// cover/header/footer part templates) and the <see cref="TemplateManifest"/> that says
/// everything else about the house (names, data store, logo, fonts, properties, palette).
/// </summary>
/// <remarks>
/// There is no built-in template: callers name the directory. It is checked once, up front,
/// so a missing part fails with its name rather than half-way through a build.
/// </remarks>
public sealed class TemplateSource
{
    public const string Styles = "styles.xml";
    public const string Theme = "theme1.xml";
    public const string Numbering = "numbering.xml";
    public const string Settings = "settings.xml";
    public const string CoverDocument = "cover-document.xml";
    public const string CoverBlock = "cover-block.xml";
    public const string Header = "header.xml";
    public const string FirstPageHeader = "header-first.xml";
    public const string Footer = "footer.xml";
    public const string Metadata = "report.yaml";

    /// <summary>Token in <see cref="Footer"/> that becomes the logo's relationship id.</summary>
    public const string LogoToken = "{{LOGOID}}";

    /// <summary>
    /// The parts every template has. <see cref="Metadata"/> is optional because metadata may come
    /// from any path; the manifest's logo and, with a cover anchor, <see cref="CoverBlock"/> are
    /// required when the manifest asks for them.
    /// </summary>
    public static IReadOnlyList<string> RequiredParts { get; } =
        [TemplateManifest.FileName, Styles, Theme, Numbering, Settings, CoverDocument, Header, FirstPageHeader, Footer];

    /// <summary>Opens <paramref name="directory"/> as a template source.</summary>
    /// <exception cref="DirectoryNotFoundException">The directory does not exist.</exception>
    /// <exception cref="FileNotFoundException">One or more required parts is missing; the message names them all.</exception>
    /// <exception cref="InvalidDataException">The manifest is malformed, or the footer expects a logo the manifest does not name.</exception>
    public TemplateSource(string directory)
    {
        Directory = Path.GetFullPath(directory);
        if (!System.IO.Directory.Exists(Directory))
            throw new DirectoryNotFoundException($"template directory not found: {Directory}");
        RequireParts(RequiredParts);

        Manifest = TemplateManifest.Load(Directory);
        List<string> named = [];
        if (Manifest.Logo is { } logo) named.Add(logo);
        if (Manifest.CoverAnchor is not null) named.Add(CoverBlock);
        RequireParts(named);
        if (Manifest.Logo is null && ReadText(Footer).Contains(LogoToken, StringComparison.Ordinal))
            throw new InvalidDataException($"{PathOf(Footer)} places a logo ({LogoToken}) but {TemplateManifest.FileName} names no 'logo'");
    }

    /// <summary>Absolute path of the template directory.</summary>
    public string Directory { get; }

    /// <summary>The directory's <c>template.yaml</c>.</summary>
    public TemplateManifest Manifest { get; }

    /// <summary>The template's own <c>report.yaml</c>, the default metadata seed.</summary>
    public string MetadataPath => PathOf(Metadata);

    /// <summary>Absolute path of <paramref name="name"/>; absolute names are returned unchanged.</summary>
    public string PathOf(string name) => Path.GetFullPath(Path.Combine(Directory, name));

    public bool Contains(string name) => File.Exists(PathOf(name));

    public string ReadText(string name) => File.ReadAllText(PathOf(name));

    public byte[] ReadBytes(string name) => File.ReadAllBytes(PathOf(name));

    /// <summary>
    /// The embedder for the manifest's font family, looking for its faces in
    /// <paramref name="fontDirectory"/> (default <see cref="FontEmbedder.UserFontDirectory"/>);
    /// null when the template embeds no fonts.
    /// </summary>
    public FontEmbedder? FontEmbedder(string? fontDirectory = null) =>
        Manifest.Fonts is { } fonts
            ? new FontEmbedder(fontDirectory ?? Templating.FontEmbedder.UserFontDirectory, fonts.Family, fonts.Faces)
            : null;

    private void RequireParts(IEnumerable<string> parts)
    {
        var missing = parts.Where(p => !Contains(p)).ToList();
        if (missing.Count > 0)
            throw new FileNotFoundException($"template {Directory} is missing: {string.Join(", ", missing)}");
    }
}
