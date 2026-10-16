using System.Text.Json;

namespace OfficeProof.Verify;

/// <summary>
/// The measured evidence behind <see cref="RevisionAllowlist"/>, read from the embedded
/// <c>revision-evidence.json</c>: each engine (Word for Mac, Word for Windows, Clippit,
/// officeproof's view) and the build it was measured on, each probe's outcome per engine, and
/// each revision site's evidence. A site is admitted only when a planted test case exercises it
/// and every probe it cites comes out the same in every engine; otherwise
/// <see cref="Refusal"/> says why not.
/// </summary>
internal static class RevisionEvidence
{
    /// <summary>One engine; <paramref name="Measured"/> and <paramref name="Build"/> are set for Word.</summary>
    internal sealed record Engine(string Key, string Name, string? Version, string? Measured, string? Build)
    {
        public string Label => Version is null ? Name : $"{Name} {Version}";
    }

    /// <summary>A probe's outcome letters per engine key: Reject All, then Accept All.</summary>
    internal sealed record Probe(string Name, IReadOnlyDictionary<string, string> Outcomes)
    {
        public bool Agrees => Outcomes.Values.Distinct(StringComparer.Ordinal).Count() == 1;
    }

    /// <summary>A revision site; <paramref name="Word"/>: how Word's Reject All of its planted case failed, if it did.</summary>
    internal sealed record Site(string Key, bool Planted, IReadOnlyList<string> Probes, string? Word);

    private static readonly Lazy<(IReadOnlyDictionary<string, string> Outcomes, IReadOnlyList<Engine> Engines,
        IReadOnlyDictionary<string, Probe> Probes, IReadOnlyDictionary<string, Site> Sites)> Data = new(Load);

    /// <summary>Outcome letter to its meaning.</summary>
    public static IReadOnlyDictionary<string, string> Outcomes => Data.Value.Outcomes;

    public static IReadOnlyList<Engine> Engines => Data.Value.Engines;

    public static IReadOnlyDictionary<string, Probe> Probes => Data.Value.Probes;

    public static IReadOnlyDictionary<string, Site> Sites => Data.Value.Sites;

    /// <summary>
    /// The Word build measured for this platform's backend (<c>word-mac</c>, <c>word-windows</c>):
    /// its name, version and the date (the loader requires both for every Word engine).
    /// </summary>
    public static (string Name, string Version, string Measured) Word(string key)
    {
        var e = Engines.Single(e => e.Key == key);
        return (e.Name, e.Version!, e.Measured!);
    }

    /// <summary>
    /// Why <paramref name="site"/> is not admitted, or null when it is: no evidence at all, a
    /// cited probe the engines disagree on (with each engine's outcome), Word not restoring the
    /// planted case, or no planted case.
    /// </summary>
    public static string? Refusal(string site)
    {
        if (!Sites.TryGetValue(site, out var s))
            return "officeproof has not verified what Reject All does with this revision";
        var disagreeing = s.Probes.Select(p => Probes[p]).Where(p => !p.Agrees).ToList();
        if (disagreeing.Count > 0)
            return string.Join("; ", disagreeing.Select(Disagreement));
        if (s.Word is not null)
            return s.Word + "; only verify --word can check this revision";
        return s.Planted ? null : "officeproof has no planted test case for this revision yet, although Word and Clippit agree on it";
    }

    private static string Disagreement(Probe probe)
    {
        var parts = new List<string>();
        foreach (var (index, view) in new[] { (0, "Reject All"), (1, "Accept All") })
        {
            var groups = Engines.GroupBy(e => probe.Outcomes[e.Key][index]).ToList();
            if (groups.Count < 2) continue;
            parts.Add(view + " " + string.Join("; ", groups.Select(g =>
                $"in {Join(g.Select(e => e.Label))} {Outcomes[g.Key.ToString()]}")));
        }
        return $"the engines disagree on it ({string.Join("; ", parts)}; probe {probe.Name})";
    }

    private static string Join(IEnumerable<string> items)
    {
        var list = items.ToList();
        return list.Count == 1 ? list[0] : string.Join(", ", list.Take(list.Count - 1)) + " and " + list[^1];
    }

    private static (IReadOnlyDictionary<string, string>, IReadOnlyList<Engine>, IReadOnlyDictionary<string, Probe>, IReadOnlyDictionary<string, Site>) Load()
    {
        using var stream = typeof(RevisionEvidence).Assembly.GetManifestResourceStream("OfficeProof.Verify.revision-evidence.json")
            ?? throw new InvalidOperationException("embedded revision-evidence.json is missing");
        using var json = JsonDocument.Parse(stream);
        var root = json.RootElement;

        var outcomes = root.GetProperty("outcomes").EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!, StringComparer.Ordinal);
        var engines = root.GetProperty("engines").EnumerateObject().Select(p => new Engine(p.Name,
            p.Value.GetProperty("name").GetString()!, Optional(p.Value, "version"),
            Optional(p.Value, "measured"), Optional(p.Value, "build"))).ToList();
        foreach (var word in engines.Where(e => e.Key.StartsWith("word-", StringComparison.Ordinal)))
            if (word.Version is null || word.Measured is null)
                throw new InvalidDataException($"revision-evidence.json: engine {word.Key} needs the Word version and the date it was measured");
        var probes = root.GetProperty("probes").EnumerateObject().ToDictionary(p => p.Name, p => new Probe(p.Name,
            p.Value.EnumerateObject().ToDictionary(o => o.Name, o => o.Value.GetString()!, StringComparer.Ordinal)), StringComparer.Ordinal);
        var sites = root.GetProperty("sites").EnumerateObject().ToDictionary(p => p.Name, p => new Site(p.Name,
            p.Value.GetProperty("planted").GetBoolean(),
            p.Value.TryGetProperty("probes", out var cited) ? [.. cited.EnumerateArray().Select(c => c.GetString()!)] : [],
            Optional(p.Value, "word")), StringComparer.Ordinal);

        foreach (var probe in probes.Values)
        {
            if (!probe.Outcomes.Keys.Order(StringComparer.Ordinal).SequenceEqual(engines.Select(e => e.Key).Order(StringComparer.Ordinal)))
                throw new InvalidDataException($"revision-evidence.json: probe {probe.Name} must give an outcome for every engine");
            foreach (var o in probe.Outcomes.Values)
                if (o.Length != 2 || o.Any(c => !outcomes.ContainsKey(c.ToString())))
                    throw new InvalidDataException($"revision-evidence.json: probe {probe.Name} outcome \"{o}\" is not two known letters");
        }
        foreach (var site in sites.Values)
            foreach (var p in site.Probes)
                if (!probes.ContainsKey(p))
                    throw new InvalidDataException($"revision-evidence.json: site \"{site.Key}\" cites unknown probe {p}");
        return (outcomes, engines, probes, sites);
    }

    private static string? Optional(JsonElement e, string name) => e.TryGetProperty(name, out var v) ? v.GetString() : null;
}
