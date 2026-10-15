using OfficeProof.Core;

namespace OfficeProof.Edit;

/// <summary>The story parts a replacement searches.</summary>
[Flags]
public enum Stories
{
    None = 0,
    /// <summary>The main document part (body, tables, text boxes).</summary>
    Main = 1,
    Headers = 2,
    Footers = 4,
    /// <summary>Footnotes and endnotes.</summary>
    Notes = 8,
    /// <summary>Comment text; never part of <see cref="All"/>.</summary>
    Comments = 16,
    /// <summary>Every printed story: main, headers, footers, footnotes and endnotes.</summary>
    All = Main | Headers | Footers | Notes,
}

/// <summary>
/// One find/replace pair. <paramref name="Find"/> is literal text unless
/// <see cref="ReplaceOptions.Regex"/> is set, in which case it is a .NET regular expression and
/// <paramref name="With"/> may use <c>$1</c>, <c>${name}</c> and <c>$$</c>.
/// </summary>
public sealed record ReplacePair(string Find, string With);

/// <summary>How <see cref="Replacer.Apply"/> searches.</summary>
public sealed record ReplaceOptions
{
    /// <summary>Treat every <see cref="ReplacePair.Find"/> as a regular expression (case-sensitive).</summary>
    public bool Regex { get; init; }

    public Stories Stories { get; init; } = Stories.All;

    /// <summary>When set, the total number of matches over all pairs must equal this, or nothing is written.</summary>
    public int? Expect { get; init; }
}

/// <summary>Matches of one pair in one story part, e.g. ("/word/header1.xml", 2).</summary>
public sealed record PartMatches(string Part, int Matches);

/// <summary>What one pair matched, per story part (parts without a match are omitted).</summary>
public sealed record PairResult(ReplacePair Pair, IReadOnlyList<PartMatches> Parts)
{
    public int Total => Parts.Sum(p => p.Matches);
}

/// <summary>
/// Outcome of <see cref="Replacer.Apply"/>. When <see cref="Errors"/> is non-empty nothing was
/// written. Otherwise <see cref="EditedParts"/> are the only parts whose bytes changed and
/// <see cref="NewSchemaErrors"/> are schema errors the output has beyond the input's own.
/// </summary>
/// <param name="Pairs">Per pair, in order; each pair searched the result of the pairs before it.</param>
/// <param name="Errors">Why the replace was refused (rules replace-revisions, replace-no-match,
/// replace-expect, replace-non-text, replace-unaligned).</param>
/// <param name="EditedParts">Story parts written, as package URIs.</param>
/// <param name="NewSchemaErrors">Open XML SDK schema errors absent from the input.</param>
public sealed record ReplaceResult(
    IReadOnlyList<PairResult> Pairs,
    IReadOnlyList<Finding> Errors,
    IReadOnlyList<string> EditedParts,
    IReadOnlyList<Finding> NewSchemaErrors)
{
    public bool Succeeded => Errors.Count == 0;

    public int Total => Pairs.Sum(p => p.Total);
}
