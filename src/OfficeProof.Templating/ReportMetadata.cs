using System.Xml;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;

namespace OfficeProof.Templating;

/// <summary>
/// Report metadata: an ordered map of key to text that fills the cover, the running header,
/// the template's report data store and the package properties. Immutable.
/// </summary>
/// <remarks>
/// Keys become element names in the data store (<c>&lt;prefix:title&gt;</c>), so each must be a valid XML
/// NCName. Order is preserved because the data store is written in it.
/// <para>
/// The legacy CLI read only <c>template/report.yaml</c>, hand-edited global state with no
/// --metadata flag: whatever sat there was stamped onto the next document processed, which is
/// how one report's identity ended up on another retemplated report on 2026-08-26. <see cref="Load"/>
/// accepts any path, so each report can carry its own file.
/// </para>
/// </remarks>
public sealed class ReportMetadata
{
    private readonly List<KeyValuePair<string, string>> _entries;

    private ReportMetadata(List<KeyValuePair<string, string>> entries, string? source)
    {
        _entries = entries;
        Source = source;
    }

    /// <summary>No metadata.</summary>
    public static ReportMetadata Empty { get; } = new([], null);

    /// <summary>Where the metadata was loaded from, when it came from a file.</summary>
    public string? Source { get; }

    /// <summary>The entries in source order.</summary>
    public IReadOnlyList<KeyValuePair<string, string>> Entries => _entries;

    public IEnumerable<string> Keys => _entries.Select(e => e.Key);

    /// <summary>The value for <paramref name="key"/>, or null when it is absent or blank.</summary>
    public string? Get(string key)
    {
        foreach (var (k, v) in _entries)
            if (k == key) return string.IsNullOrWhiteSpace(v) ? null : v;
        return null;
    }

    /// <summary>Builds metadata from pairs; a later duplicate key replaces the earlier value in place.</summary>
    /// <exception cref="ArgumentException">A key is not a valid XML NCName.</exception>
    public static ReportMetadata From(IEnumerable<KeyValuePair<string, string>> entries) => Empty.With(entries);

    /// <summary>
    /// A copy with <paramref name="overrides"/> applied: an existing key keeps its position and
    /// takes the new value, a new key is appended.
    /// </summary>
    /// <exception cref="ArgumentException">A key is not a valid XML NCName.</exception>
    public ReportMetadata With(IEnumerable<KeyValuePair<string, string>> overrides)
    {
        var entries = new List<KeyValuePair<string, string>>(_entries);
        foreach (var (key, value) in overrides)
        {
            CheckKey(key);
            var i = entries.FindIndex(e => e.Key == key);
            if (i >= 0) entries[i] = new(key, value);
            else entries.Add(new(key, value));
        }
        return new ReportMetadata(entries, Source);
    }

    /// <summary>A copy with <paramref name="overrides"/> applied (see <see cref="With(IEnumerable{KeyValuePair{string, string}})"/>).</summary>
    public ReportMetadata With(ReportMetadata overrides) => With(overrides.Entries);

    /// <summary>Loads the top-level <c>report:</c> map of a YAML file at any path.</summary>
    /// <exception cref="InvalidDataException">The file is not YAML, has no <c>report:</c> map, or a value under it is not a scalar.</exception>
    public static ReportMetadata Load(string path)
    {
        var full = Path.GetFullPath(path);
        var parsed = Parse(File.ReadAllText(full), full);
        return new ReportMetadata(parsed._entries, full);
    }

    /// <summary>Parses YAML text holding a top-level <c>report:</c> map. <paramref name="sourceName"/> labels errors.</summary>
    /// <exception cref="InvalidDataException">See <see cref="Load"/>.</exception>
    public static ReportMetadata Parse(string yaml, string sourceName = "metadata")
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

        if (stream.Documents.Count == 0
            || stream.Documents[0].RootNode is not YamlMappingNode root
            || !root.Children.TryGetValue(new YamlScalarNode("report"), out var reportNode)
            || reportNode is not YamlMappingNode report)
            throw new InvalidDataException($"{sourceName}: no top-level 'report:' map");

        var entries = new List<KeyValuePair<string, string>>();
        foreach (var (k, v) in report.Children)
        {
            var key = (k as YamlScalarNode)?.Value ?? k.ToString();
            if (v is not YamlScalarNode scalar)
                throw new InvalidDataException($"{sourceName}: report.{key} must be a single value");
            entries.Add(new(key, scalar.Value ?? ""));
        }
        try
        {
            return From(entries);
        }
        catch (ArgumentException e)
        {
            throw new InvalidDataException($"{sourceName}: {e.Message}", e);
        }
    }

    private static void CheckKey(string key)
    {
        try
        {
            XmlConvert.VerifyNCName(key);
        }
        catch (XmlException)
        {
            throw new ArgumentException($"metadata key '{key}' is not a valid XML element name", nameof(key));
        }
    }
}
