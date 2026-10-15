using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeProof.Core;

namespace OfficeProof.Lint;

// Update Field rebuilds every TOC line from the TOC1..TOC9 paragraph styles and drops the
// cached lines, direct run formatting included. A 2026-10-15 report had Arial applied by hand
// to its TOC lines while TOC1-3 were based on Normal (Times New Roman): the first Update Table
// turned the whole TOC into Times New Roman. The fix is to put the formatting in the TOCn
// styles. Direct formatting that matches what the style already gives is harmless.
/// <summary>
/// toc-direct-formatting: TOC field-result lines (TOC1..9 styles) with direct rFonts/sz/color
/// that their TOC style does not give; Update Field will discard it.
/// </summary>
public sealed partial class TocDirectFormattingRule : ILintRule
{
    private static readonly XNamespace W = Ns.W;

    // Each font slot and the theme attribute that overrides it at the same level.
    private static readonly (string Font, string Theme)[] FontSlots =
        [("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"), ("eastAsia", "eastAsiaTheme"), ("cs", "cstheme")];

    public string Id => "toc-direct-formatting";
    public Severity Severity => Severity.Warning;
    public string Description => "direct formatting on TOC lines that Update Field will discard";

    [GeneratedRegex("^toc ?[1-9]$", RegexOptions.IgnoreCase)]
    private static partial Regex TocStyle();

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var byStyle = new Dictionary<string, (ParagraphText First, int Lines, SortedSet<string> Lost)>(StringComparer.Ordinal);
        foreach (var p in ctx.Paragraphs(ctx.Main))
        {
            var styleId = Wml.StyleId(p.Paragraph);
            if (styleId is null || !IsTocStyle(ctx, styleId) || !p.InField("TOC")) continue;
            var lost = new SortedSet<string>(StringComparer.Ordinal);
            foreach (var rPr in p.Segments.Where(s => s.InField("TOC")).Select(s => s.Node.Parent?.Element(W + "rPr")).OfType<XElement>().Distinct())
                lost.UnionWith(Discarded(ctx, styleId, rPr));
            if (lost.Count == 0) continue;
            if (byStyle.TryGetValue(styleId, out var entry))
            {
                entry.Lost.UnionWith(lost);
                byStyle[styleId] = (entry.First, entry.Lines + 1, entry.Lost);
            }
            else byStyle[styleId] = (p, 1, lost);
        }

        var part = LintContext.PartName(ctx.Main);
        foreach (var (styleId, (first, lines, lost)) in byStyle.OrderBy(kv => kv.Key, StringComparer.Ordinal))
            yield return new Finding(Id, Severity,
                $"{lines} {styleId} line(s) carry direct formatting the {styleId} style does not give ({string.Join("; ", lost)}); "
                + $"Update Field will discard it - put it in the {styleId} style",
                part, first.Location);
    }

    private static bool IsTocStyle(LintContext ctx, string styleId) =>
        TocStyle().IsMatch(styleId)
        || TocStyle().IsMatch((string?)ctx.Styles.Style(styleId)?.Element(W + "name")?.Attribute(W + "val") ?? "");

    /// <summary>Direct rFonts/sz/color values on <paramref name="rPr"/> that differ from the style's effective value.</summary>
    private static IEnumerable<string> Discarded(LintContext ctx, string styleId, XElement rPr)
    {
        if (rPr.Element(W + "rFonts") is { } fonts)
        {
            foreach (var (font, theme) in FontSlots)
            {
                var direct = FontValue(fonts, font, theme);
                if (direct is null) continue;
                var styled = ctx.StyleRunLevels(styleId).Select(l => l?.Element(W + "rFonts") is { } f ? FontValue(f, font, theme) : null)
                    .FirstOrDefault(v => v is not null);
                if (!string.Equals(direct, styled, StringComparison.OrdinalIgnoreCase))
                    yield return $"rFonts {font}={direct}, style gives {styled ?? "nothing"}";
            }
        }
        if ((string?)rPr.Element(W + "sz")?.Attribute(W + "val") is { } sz)
        {
            // Word's default when nothing sets a size is 10 pt (20 half-points).
            var styled = ctx.StyleRunProperty(styleId, "sz", "val") ?? "20";
            if (sz != styled) yield return $"sz={sz}, style gives {styled}";
        }
        if ((string?)rPr.Element(W + "color")?.Attribute(W + "val") is { } color)
        {
            var styled = ctx.StyleRunProperty(styleId, "color", "val") ?? "auto";
            if (!string.Equals(Colour(color), Colour(styled), StringComparison.OrdinalIgnoreCase))
                yield return $"color={color}, style gives {styled}";
        }
    }

    // A theme font attribute overrides the explicit one on the same element.
    private static string? FontValue(XElement rFonts, string font, string theme) =>
        (string?)rFonts.Attribute(W + theme) is { } t ? $"theme:{t}" : (string?)rFonts.Attribute(W + font);

    // "auto" renders black on a white page; treat the two as the same colour.
    private static string Colour(string value) => value.Equals("auto", StringComparison.OrdinalIgnoreCase) ? "000000" : value;
}
