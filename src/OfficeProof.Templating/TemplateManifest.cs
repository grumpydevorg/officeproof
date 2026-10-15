using System.Collections.Frozen;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;

namespace OfficeProof.Templating;

/// <summary>The custom-XML store that cover and header content controls bind to.</summary>
/// <param name="Namespace">Namespace of the store's root element; the template's <c>w:dataBinding</c> prefix mappings name the same URI.</param>
/// <param name="Prefix">Prefix declared on the root element (<c>&lt;prefix:report&gt;</c>).</param>
public sealed record DataStoreSpec(string Namespace, string Prefix)
{
    /// <summary>The store's root element name, <c>report</c> in <see cref="Namespace"/>.</summary>
    public XName Root => XName.Get("report", Namespace);
}

/// <summary>A font family to embed and the face files that fill its slots.</summary>
public sealed record FontSpec(string Family, IReadOnlyList<FontFace> Faces);

/// <summary>
/// One document property: <paramref name="Name"/> takes the value of metadata key
/// <paramref name="Key"/> (manifest value <c>"{{key}}"</c>) or the fixed text
/// <paramref name="Literal"/> (any other manifest value). Exactly one of the two is set.
/// </summary>
public sealed record PropertyBinding(string Name, string? Key, string? Literal)
{
    /// <summary>The value to write: the literal, or the metadata value (null when the key is absent or blank).</summary>
    public string? Resolve(ReportMetadata metadata) => Key is null ? Literal : metadata.Get(Key);
}

/// <summary>
/// Explicit colours a template allows (lower-case 6-digit hex without '#' → name) and known
/// near-misses worth naming (hex → hint). The CLI hands it to the lint <c>brand-colour</c> rule.
/// </summary>
public sealed record PaletteSpec(IReadOnlyDictionary<string, string> Colours, IReadOnlyDictionary<string, string> NearMisses);

/// <summary>
/// A template directory's <c>template.yaml</c>: everything house-specific that is not a part
/// file -- output names, the data-store namespace, logo, fonts, which metadata keys fill which
/// document properties, the identity a document must carry, the legacy heading styles
/// <c>retemplate</c> remaps, the cover anchor and the colour palette. The schema is documented
/// in the README of each template directory (see <c>examples/minimal-template</c>).
/// </summary>
/// <remarks>
/// Unknown keys are errors rather than ignored, so a misspelt section fails loudly instead of
/// silently turning a feature off.
/// </remarks>
public sealed partial class TemplateManifest
{
    /// <summary>The manifest's file name inside a template directory.</summary>
    public const string FileName = "template.yaml";

    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    // Core properties a template may set, by manifest name. cp:lastModifiedBy and cp:revision
    // are the document's edit history, never template data (see DocPropsWriter).
    private static readonly FrozenDictionary<string, XName> CorePropertyNames = new Dictionary<string, XName>
    {
        ["title"] = Dc + "title",
        ["subject"] = Dc + "subject",
        ["creator"] = Dc + "creator",
        ["description"] = Dc + "description",
        ["language"] = Dc + "language",
        ["keywords"] = Cp + "keywords",
        ["category"] = Cp + "category",
        ["contentStatus"] = Cp + "contentStatus",
    }.ToFrozenDictionary();

    private static readonly FrozenSet<string> TopLevelKeys = new[]
    {
        "name", "dataStore", "logo", "fonts", "coreProperties", "customProperties",
        "identityKeys", "headingStyles", "coverAnchor", "palette",
    }.ToFrozenSet();

    private TemplateManifest(string source) => Source = source;

    /// <summary>Where the manifest was read from (a path, or the label given to <see cref="Parse"/>).</summary>
    public string Source { get; }

    /// <summary>Base name of <c>generate</c>'s outputs.</summary>
    public string Name { get; private init; } = "";

    /// <summary><c>&lt;name&gt;-sample.docx</c>, the reference sample <c>generate</c> writes.</summary>
    public string SampleFileName => Name + "-sample.docx";

    /// <summary><c>&lt;name&gt;.dotx</c>, the Word template <c>generate</c> writes.</summary>
    public string TemplateFileName => Name + ".dotx";

    /// <summary>The bound data store.</summary>
    public DataStoreSpec DataStore { get; private init; } = new("", "");

    /// <summary>Footer logo image, relative to the template directory; null for no logo.</summary>
    public string? Logo { get; private init; }

    /// <summary>The font family embedded in every output; null embeds none.</summary>
    public FontSpec? Fonts { get; private init; }

    /// <summary>docProps/core.xml properties, in write order.</summary>
    public IReadOnlyList<PropertyBinding> CoreProperties { get; private init; } = [];

    /// <summary>docProps/custom.xml properties, in write order.</summary>
    public IReadOnlyList<PropertyBinding> CustomProperties { get; private init; } = [];

    /// <summary>
    /// Metadata keys a document must carry itself before <c>retemplate</c> may run without an
    /// explicit metadata file. Empty means a document never identifies itself: metadata is always required.
    /// </summary>
    /// <remarks>
    /// Any one identity value is not enough: a report whose only identity was an account-name
    /// dc:creator let the template's report.yaml stamp another report's title and project code
    /// onto it on 2026-10-15 (review finding 6). List the keys the running header and cover show first.
    /// </remarks>
    public IReadOnlyList<string> IdentityKeys { get; private init; } = [];

    /// <summary>Legacy paragraph style id → the canonical heading style <c>retemplate</c> applies.</summary>
    public FrozenDictionary<string, string> HeadingStyles { get; private init; } = FrozenDictionary<string, string>.Empty;

    /// <summary>
    /// Text that marks the paragraph ending a hand-typed cover block: <c>retemplate</c> replaces
    /// every body paragraph before the first paragraph containing it with the data-bound
    /// cover block. Null leaves every cover alone.
    /// </summary>
    public string? CoverAnchor { get; private init; }

    /// <summary>The colour palette for lint; null when the template declares none.</summary>
    public PaletteSpec? Palette { get; private init; }

    /// <summary>The core property element for a manifest property name.</summary>
    public static XName CorePropertyElement(string name) => CorePropertyNames[name];

    /// <summary>Reads <see cref="FileName"/> from <paramref name="directory"/>.</summary>
    /// <exception cref="FileNotFoundException">The directory has no manifest.</exception>
    /// <exception cref="InvalidDataException">The manifest is malformed; the message names the file and the key.</exception>
    public static TemplateManifest Load(string directory)
    {
        var path = Path.GetFullPath(Path.Combine(directory, FileName));
        if (!File.Exists(path)) throw new FileNotFoundException($"template manifest not found: {path}", path);
        return Parse(File.ReadAllText(path), path);
    }

    /// <summary>Parses manifest YAML; <paramref name="sourceName"/> labels errors.</summary>
    /// <exception cref="InvalidDataException">See <see cref="Load"/>.</exception>
    public static TemplateManifest Parse(string yaml, string sourceName = FileName)
    {
        var stream = new YamlStream();
        try
        {
            stream.Load(new StringReader(yaml));
        }
        catch (YamlException e)
        {
            throw new InvalidDataException($"{sourceName}: not valid YAML: {e.Message}", e);
        }
        if (stream.Documents.Count == 0 || stream.Documents[0].RootNode is not YamlMappingNode root)
            throw new InvalidDataException($"{sourceName}: expected a map at the top level");

        var r = new Reader(sourceName);
        var map = r.Map(root, "", TopLevelKeys);
        var store = r.RequiredMap(map, "dataStore", ["namespace", "prefix"]);
        var prefix = r.RequiredScalar(store, "dataStore.prefix");
        try
        {
            XmlConvert.VerifyNCName(prefix);
        }
        catch (XmlException)
        {
            throw r.Error($"dataStore.prefix '{prefix}' is not a valid XML prefix");
        }

        return new TemplateManifest(sourceName)
        {
            Name = FileNameSafe(r, r.RequiredScalar(map, "name"), "name"),
            DataStore = new DataStoreSpec(r.RequiredScalar(store, "dataStore.namespace"), prefix),
            Logo = r.Scalar(map, "logo"),
            Fonts = ReadFonts(r, map),
            CoreProperties = ReadProperties(r, map, "coreProperties", core: true),
            CustomProperties = ReadProperties(r, map, "customProperties", core: false),
            IdentityKeys = r.List(map, "identityKeys"),
            HeadingStyles = r.StringMap(map, "headingStyles").ToFrozenDictionary(StringComparer.Ordinal),
            CoverAnchor = r.Scalar(map, "coverAnchor"),
            Palette = ReadPalette(r, map),
        };
    }

    private static string FileNameSafe(Reader r, string value, string key) =>
        value.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || value.Contains('/', StringComparison.Ordinal)
            ? throw r.Error($"{key} '{value}' is not usable as a file name")
            : value;

    private static FontSpec? ReadFonts(Reader r, Dictionary<string, YamlNode> map)
    {
        var slots = Enum.GetValues<FontSlot>().ToDictionary(s => JsonName(s.ToString()));
        var fonts = r.OptionalMap(map, "fonts", ["family", .. slots.Keys]);
        if (fonts is null) return null;
        var faces = slots
            .Select(s => (Slot: s.Value, File: r.Scalar(fonts, s.Key, "fonts." + s.Key)))
            .Where(f => f.File is not null)
            .Select(f => new FontFace(f.File!, f.Slot))
            .ToList();
        return new FontSpec(r.RequiredScalar(fonts, "fonts.family"), faces);
    }

    private static List<PropertyBinding> ReadProperties(Reader r, Dictionary<string, YamlNode> map, string key, bool core)
    {
        var result = new List<PropertyBinding>();
        foreach (var (name, value) in r.StringMap(map, key))
        {
            if (core && !CorePropertyNames.ContainsKey(name))
                throw r.Error($"{key}.{name} is not a core property; known: {string.Join(", ", CorePropertyNames.Keys.Order(StringComparer.Ordinal))}");
            var bound = Binding().Match(value);
            result.Add(bound.Success ? new PropertyBinding(name, bound.Groups[1].Value, null) : new PropertyBinding(name, null, value));
        }
        return result;
    }

    private static PaletteSpec? ReadPalette(Reader r, Dictionary<string, YamlNode> map)
    {
        var palette = r.OptionalMap(map, "palette", ["colours", "nearMisses"]);
        if (palette is null) return null;
        return new PaletteSpec(Hexes(r, palette, "colours"), Hexes(r, palette, "nearMisses"));
    }

    private static Dictionary<string, string> Hexes(Reader r, Dictionary<string, YamlNode> palette, string key)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (hex, label) in r.StringMap(palette, key, "palette." + key))
        {
            var normal = hex.TrimStart('#').ToLowerInvariant();
            if (!Hex().IsMatch(normal)) throw r.Error($"palette.{key}: '{hex}' is not a 6-digit hex colour");
            result[normal] = label;
        }
        return result;
    }

    private static string JsonName(string pascal) => char.ToLowerInvariant(pascal[0]) + pascal[1..];

    [GeneratedRegex(@"^\{\{([A-Za-z_][\w.-]*)\}\}$")]
    private static partial Regex Binding();

    [GeneratedRegex("^[0-9a-f]{6}$")]
    private static partial Regex Hex();

    /// <summary>Typed access to the YAML tree with errors that name the file and the key.</summary>
    private sealed class Reader(string source)
    {
        public InvalidDataException Error(string message) => new($"{source}: {message}");

        public Dictionary<string, YamlNode> Map(YamlMappingNode node, string path, IReadOnlyCollection<string> allowed)
        {
            var result = new Dictionary<string, YamlNode>(StringComparer.Ordinal);
            foreach (var (k, v) in node.Children)
            {
                var key = (k as YamlScalarNode)?.Value ?? k.ToString();
                if (!allowed.Contains(key))
                    throw Error($"unknown key '{path}{key}'; expected one of: {string.Join(", ", allowed)}");
                result[key] = v;
            }
            return result;
        }

        public Dictionary<string, YamlNode>? OptionalMap(Dictionary<string, YamlNode> map, string key, IReadOnlyCollection<string> allowed) =>
            !map.TryGetValue(key, out var node) || IsNull(node) ? null
            : node is YamlMappingNode m ? Map(m, key + ".", allowed)
            : throw Error($"{key} must be a map");

        public Dictionary<string, YamlNode> RequiredMap(Dictionary<string, YamlNode> map, string key, IReadOnlyCollection<string> allowed) =>
            OptionalMap(map, key, allowed) ?? throw Error($"{key} is required");

        public string? Scalar(Dictionary<string, YamlNode> map, string key, string? path = null)
        {
            if (!map.TryGetValue(key, out var node) || IsNull(node)) return null;
            return node is YamlScalarNode { Value: { } v } ? (string.IsNullOrWhiteSpace(v) ? null : v)
                : throw Error($"{path ?? key} must be a single value");
        }

        public string RequiredScalar(Dictionary<string, YamlNode> map, string path) =>
            Scalar(map, path[(path.LastIndexOf('.') + 1)..], path) ?? throw Error($"{path} is required");

        public List<string> List(Dictionary<string, YamlNode> map, string key)
        {
            if (!map.TryGetValue(key, out var node) || IsNull(node)) return [];
            if (node is not YamlSequenceNode seq) throw Error($"{key} must be a list");
            return [.. seq.Children.Select(c => c is YamlScalarNode { Value: { } v } && !string.IsNullOrWhiteSpace(v)
                ? v : throw Error($"{key}: every entry must be a single value"))];
        }

        public List<KeyValuePair<string, string>> StringMap(Dictionary<string, YamlNode> map, string key, string? path = null)
        {
            path ??= key;
            if (!map.TryGetValue(key, out var node) || IsNull(node)) return [];
            if (node is not YamlMappingNode m) throw Error($"{path} must be a map");
            return [.. m.Children.Select(c => new KeyValuePair<string, string>(
                (c.Key as YamlScalarNode)?.Value ?? throw Error($"{path}: keys must be single values"),
                c.Value is YamlScalarNode { Value: { } v } ? v : throw Error($"{path}.{c.Key}: must be a single value")))];
        }

        private static bool IsNull(YamlNode node) =>
            node is YamlScalarNode { Style: ScalarStyle.Plain, Value: null or "" or "~" or "null" };
    }
}
