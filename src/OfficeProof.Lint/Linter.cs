using OfficeProof.Core;

namespace OfficeProof.Lint;

/// <summary>Runs lint rules over a .docx without modifying it.</summary>
public static class Linter
{
    /// <summary>
    /// Every built-in rule in report order. <paramref name="palette"/> configures
    /// <c>brand-colour</c>; without one that rule is skipped with an Info finding.
    /// </summary>
    public static IReadOnlyList<ILintRule> DefaultRules(BrandPalette? palette = null) =>
    [
        new DanglingStyleRule(),
        new BrandColourRule(palette),
        new TypedHeadingNumberRule(),
        new DoubleHeadingNumberRule(),
        new DuplicateHeadingNumberRule(),
        new TypedCaptionNumberRule(),
        new TypedCrossReferenceRule(),
        new MissingReferenceTargetRule(),
        new TocDirectFormattingRule(),
        new UnresolvedRevisionRule(),
        new RevisionInStylesRule(),
    ];

    /// <summary>The built-in rule with <paramref name="id"/> (configured with <paramref name="palette"/>), or null.</summary>
    public static ILintRule? FindRule(string id, BrandPalette? palette = null) =>
        DefaultRules(palette).FirstOrDefault(r => string.Equals(r.Id, id, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// The built-in rules named by <paramref name="ids"/>, in the order given, configured with <paramref name="palette"/>.
    /// Throws <see cref="ArgumentException"/> naming the known ids when one is unknown.
    /// </summary>
    public static IReadOnlyList<ILintRule> Rules(IEnumerable<string> ids, BrandPalette? palette = null)
    {
        ArgumentNullException.ThrowIfNull(ids);
        return [.. ids.Select(id => FindRule(id, palette) ?? throw new ArgumentException(
            $"unknown lint rule '{id}'; known: {string.Join(", ", DefaultRules().Select(r => r.Id))}", nameof(ids)))];
    }

    /// <summary>
    /// Opens <paramref name="path"/> read-only and runs <paramref name="rules"/> (default:
    /// <see cref="DefaultRules"/> without a palette). Findings come back grouped by rule, in document order within a rule.
    /// </summary>
    public static IReadOnlyList<Finding> Run(string path, IEnumerable<ILintRule>? rules = null)
    {
        using var doc = Docx.OpenRead(path);
        var ctx = new LintContext(doc);
        return [.. (rules ?? DefaultRules()).SelectMany(r => r.Check(ctx))];
    }
}
