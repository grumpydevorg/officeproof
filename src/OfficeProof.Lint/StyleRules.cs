using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Lint;

// Every pStyle a document references must exist in the styles part. Word resolves
// a dangling reference to Normal SILENTLY, so this is invisible in the rendered
// document and in OpenXmlValidator - it shows up only as formatting that quietly
// went missing. Retemplating a report on 2026-08-26 left 75 such paragraphs
// (ListParagraph x54, TOC1/TOC2/TOCHeading x20, NormalLatinArial x1) because the
// canonical bundle came from a report that never used them; the document grew
// from 24 to 29 pages and nothing reported an error.
// Character (rStyle) and table (tblStyle) references fall back the same silent way.
/// <summary>dangling-style: a pStyle/rStyle/tblStyle that styles.xml does not define.</summary>
public sealed class DanglingStyleRule : ILintRule
{
    private static readonly XNamespace W = Ns.W;

    public string Id => "dangling-style";
    public Severity Severity => Severity.Error;
    public string Description => "style references that styles.xml does not define (Word silently falls back to the default style)";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        // Old properties kept inside a pPrChange/rPrChange count too: Reject All makes them live.
        var dangling = Docx.StoryParts(ctx.Main)
            .SelectMany(part => ctx.Xml(part).Descendants()
                .Where(e => e.Name == W + "pStyle" || e.Name == W + "rStyle" || e.Name == W + "tblStyle")
                .Select(e => (Part: part, Element: e, Id: (string?)e.Attribute(W + "val"))))
            .Where(r => r.Id is not null && !ctx.IsStyleDefined(r.Id))
            .GroupBy(r => (r.Element.Name.LocalName, r.Id!));
        foreach (var g in dangling.OrderByDescending(g => g.Count()).ThenBy(g => g.Key.Item2, StringComparer.Ordinal))
        {
            var (kind, id) = g.Key;
            var (part, element, _) = g.First();
            yield return new Finding(Id, Severity,
                $"{g.Count()} {kind} reference(s) to undefined style '{id}' - Word silently falls back to the default style",
                LintContext.PartName(part), ctx.Locate(part, element));
        }
    }
}

/// <summary>
/// The colours a brand allows (<see cref="Colours"/>: lower-case 6-digit hex without '#' → name)
/// and known near-misses worth naming rather than just flagging (<see cref="NearMisses"/>: hex → hint).
/// There is no built-in palette: it comes from the template (the CLI reads <c>palette</c> in
/// the template directory's <c>template.yaml</c>).
/// </summary>
public sealed record BrandPalette(IReadOnlyDictionary<string, string> Colours, IReadOnlyDictionary<string, string> NearMisses);

// Every explicit colour must come from the brand palette.
//
// The case this exists for: a report-template lineage shaded its header band one digit off
// the brand's olive grey, and every document built from those templates inherited it. Nobody
// was going to catch that by eye. The same documents also carried Word's stock accent5 blue
// for defined terms, which is not in the palette at all. A palette names such near-misses.
//
// Explicit colours are w:fill/w:color attributes (shading, borders) and w:color/@w:val
// (run colour); theme colours and "auto" are not explicit and pass.
/// <summary>
/// brand-colour: an explicit hex colour outside the <see cref="BrandPalette"/>. Without a
/// palette the rule is skipped and says so in one <see cref="Severity.Info"/> finding.
/// </summary>
public sealed partial class BrandColourRule(BrandPalette? palette = null) : ILintRule
{
    private static readonly XNamespace W = Ns.W;

    public string Id => "brand-colour";
    public Severity Severity => Severity.Warning;
    public string Description => "explicit colours outside the brand palette (skipped without a template palette)";

    [GeneratedRegex("^[0-9a-f]{6}$")]
    private static partial Regex Hex();

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        if (palette is null)
        {
            yield return new Finding(Id, Severity.Info, "skipped: no brand palette (a template with a 'palette' in its template.yaml supplies one)");
            yield break;
        }
        var allowed = new HashSet<string>(palette.Colours.Keys, StringComparer.OrdinalIgnoreCase);
        var hints = new Dictionary<string, string>(palette.NearMisses, StringComparer.OrdinalIgnoreCase);
        var parts = Docx.StoryParts(ctx.Main).ToList<OpenXmlPart>();
        if (ctx.Main.StyleDefinitionsPart is { } sp) parts.Add(sp);
        if (ctx.Main.NumberingDefinitionsPart is { } np) parts.Add(np);

        // One finding per colour across the whole package, as the original lint printed it: a
        // template colour repeats in every header, and fourteen copies of one hint are noise.
        var offBrand = parts
            .SelectMany(part => ctx.Xml(part).Descendants()
                .SelectMany(e => e.Attributes().Where(a => IsColour(e, a)).Select(a => (Part: part, Element: e, Hex: a.Value.ToLowerInvariant()))))
            .Where(c => Hex().IsMatch(c.Hex) && !allowed.Contains(c.Hex))
            .GroupBy(c => c.Hex);
        foreach (var g in offBrand.OrderByDescending(g => g.Count()).ThenBy(g => g.Key, StringComparer.Ordinal))
        {
            var hint = hints.TryGetValue(g.Key, out var h) ? $" -> {h}" : "";
            var partCount = g.Select(c => c.Part.Uri).Distinct().Count();
            var spread = partCount > 1 ? string.Create(CultureInfo.InvariantCulture, $" in {partCount} parts") : "";
            var (part, element, _) = g.First();
            yield return new Finding(Id, Severity,
                $"#{g.Key} x{g.Count()}{spread} outside the brand palette{hint}",
                LintContext.PartName(part), Owner(ctx, part, element));
        }
    }

    private static bool IsColour(XElement e, XAttribute a) =>
        a.Name.LocalName is "fill" or "color" or "themeColor"
        || (e.Name == W + "color" && a.Name == W + "val");

    private static string? Owner(LintContext ctx, OpenXmlPart part, XElement e) =>
        e.AncestorsAndSelf(W + "style").FirstOrDefault()?.Attribute(W + "styleId") is { } id
            ? string.Create(CultureInfo.InvariantCulture, $"style '{id.Value}'")
            : ctx.Locate(part, e);
}
