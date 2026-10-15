using System.Globalization;
using System.Text.RegularExpressions;
using OfficeProof.Core;

namespace OfficeProof.Verify;

/// <summary>
/// Compares two <see cref="StoryView"/>s story by story: the original (or Reject All of the
/// original) against Reject All of a redline, or any two packages already in the wanted state.
/// Rules: <c>reject-mismatch</c> (Error) for a story whose paragraphs (style, outline level, list
/// number, text, canonical markup) or skeleton differ, or that only one side has;
/// <c>comment-changed</c> (Error) for a pre-existing comment altered or removed (its text after
/// Reject All, so a tracked edit inside a comment passes);
/// <c>comment-added</c> (Info) for a comment only the candidate has; and the views' own
/// <c>unverifiable-revision</c> problems (Error).
/// </summary>
internal static class StoryComparison
{
    /// <summary>How many paragraph differences per story are reported before summarising.</summary>
    public const int MaxReportedDifferences = 5;

    public static IEnumerable<Finding> Compare(StoryView original, StoryView candidate)
    {
        foreach (var p in original.Problems) yield return p with { Message = "original: " + p.Message };
        foreach (var p in candidate.Problems) yield return p with { Message = "reject-all: " + p.Message };

        var candidates = candidate.Stories.ToDictionary(s => s.Name, StringComparer.Ordinal);
        var originals = original.Stories.Select(s => s.Name).ToHashSet(StringComparer.Ordinal);
        foreach (var o in original.Stories)
        {
            if (!candidates.TryGetValue(o.Name, out var c))
            {
                yield return Mismatch($"Reject All loses {o.Name}{Excerpt(o)} (removed without tracking?)", o.Part, o.Name);
                continue;
            }
            foreach (var f in CompareStory(o, c)) yield return f;
        }
        foreach (var c in candidate.Stories.Where(s => !originals.Contains(s.Name)))
            yield return Mismatch($"Reject All has {c.Name}{Excerpt(c)} that the original does not have (added without tracking?)", c.Part, c.Name);

        foreach (var f in CompareComments(original, candidate)) yield return f;
    }

    private static Finding Mismatch(string message, string part, string? location) =>
        new("reject-mismatch", Severity.Error, message, part, location);

    private static string Excerpt(Story s) =>
        s.Paragraphs.Select(p => p.Entry.Text).FirstOrDefault(t => t.Length > 0) is { } text ? $" (\"{ParagraphText.Excerpt(text)}\")" : "";

    private static IEnumerable<Finding> CompareStory(Story original, Story rejected)
    {
        var story = original.Name;
        if (original.Defined != rejected.Defined)
        {
            yield return Mismatch(original.Defined
                ? $"Reject All loses the content of {story}: the note is referenced but no longer defined"
                : $"Reject All gives {story} content the original does not define", rejected.Part, story);
            yield break;
        }
        var diffs = SequenceDiff.Compare(original.Paragraphs, rejected.Paragraphs, p => p.Key).ToList();
        foreach (var d in diffs.Take(MaxReportedDifferences))
        {
            var (message, location) = d switch
            {
                { Original: { } o, Rejected: { } r } =>
                    (Changed(o, r), $"{story} paragraph {o.Entry.Index} (reject-all paragraph {r.Entry.Index})"),
                { Original: { } o } =>
                    ($"Reject All loses original paragraph \"{RevisionCheck.Describe(o.Entry)}\" (deleted without tracking?)",
                     $"{story} paragraph {o.Entry.Index}"),
                { Rejected: { } r } =>
                    ($"Reject All keeps paragraph \"{RevisionCheck.Describe(r.Entry)}\" that the original does not have (inserted without tracking?)",
                     $"{story} reject-all paragraph {r.Entry.Index}"),
                _ => throw new InvalidOperationException("empty difference"),
            };
            yield return Mismatch($"{story}: {message}", rejected.Part, location);
        }
        if (diffs.Count > MaxReportedDifferences)
            yield return Mismatch($"{story}: {diffs.Count - MaxReportedDifferences} further paragraph difference(s) after Reject All not shown", rejected.Part, story);
        if (original.Skeleton != rejected.Skeleton)
            yield return Mismatch(
                $"{story}: Reject All differs from the original outside paragraph content: table, row, cell, section or story properties changed without tracking",
                rejected.Part, story);
    }

    // Review 2026-10-15 finding 1: an untracked pStyle change (Head2 -> Normal on "4.3 Methods")
    // passed when only list number and text were compared; Word's Reject All keeps such changes.
    private static string Changed(StoryParagraph o, StoryParagraph r)
    {
        var (a, b) = (o.Entry, r.Entry);
        var what = (a.ListNumber, a.Text) != (b.ListNumber, b.Text) ? null
            : a.StyleId != b.StyleId ? $"style {a.StyleId ?? "(default)"} -> {b.StyleId ?? "(default)"}"
            : a.Level != b.Level ? $"outline level {Level(a.Level)} -> {Level(b.Level)}"
            : o.Properties != r.Properties ? "paragraph properties differ"
            : Strip(o.Content, ResolvedTarget) == Strip(r.Content, ResolvedTarget)
                ? "a relationship target differs (hyperlink URL, image or embedded object; the story markup may be byte-identical)"
            : Strip(o.Content, HyperlinkTag) == Strip(r.Content, HyperlinkTag) ? "a hyperlink was added, removed or re-anchored"
            : "run formatting or non-text content differs (formatting, field, bookmark, drawing or other markup)";
        return what is null
            ? $"Reject All gives \"{RevisionCheck.Describe(b)}\" where the original has \"{RevisionCheck.Describe(a)}\""
            : $"Reject All does not restore \"{RevisionCheck.Describe(a)}\": {what} (changed without tracking)";

        static string Level(int? level) => level?.ToString(CultureInfo.InvariantCulture) ?? "body text";
    }

    private static string Strip(string content, Regex pattern) => pattern.Replace(content, "");

    // Values Relationships.Resolve writes in place of relationship ids.
    private static readonly Regex ResolvedTarget =
        new("\"(?:part [^\"]*|external [^\"]*|internal-link [^\"]*|data [^\"]*|unresolved [^\"]*)\"", RegexOptions.CultureInvariant);

    private static readonly Regex HyperlinkTag =
        new("</?(?:\\w+:)?hyperlink\\b[^>]*>", RegexOptions.CultureInvariant);

    // A comment is not a tracked change: adding or removing one cannot be tracked, a new one is
    // the redline's explanation. Its text can: Word and Clippit reject and accept revisions inside
    // a comment like any story (ContainerProbeTests), so after Reject All a pre-existing comment
    // must come back exactly.
    private static IEnumerable<Finding> CompareComments(StoryView original, StoryView candidate)
    {
        var candidates = candidate.Comments.ToDictionary(s => s.Name, StringComparer.Ordinal);
        var originals = original.Comments.Select(s => s.Name).ToHashSet(StringComparer.Ordinal);
        foreach (var o in original.Comments)
        {
            if (!candidates.TryGetValue(o.Name, out var c))
            {
                yield return new Finding("comment-changed", Severity.Error,
                    $"pre-existing {o.Name}{Excerpt(o)} is gone after Reject All (deleted or unanchored; removing a comment cannot be tracked)", o.Part, o.Name);
                continue;
            }
            if (Same(o, c)) continue;
            var before = string.Join(" / ", o.Paragraphs.Select(p => p.Entry.Text));
            var after = string.Join(" / ", c.Paragraphs.Select(p => p.Entry.Text));
            var what = before == after ? "author, date, formatting or other markup differs" : $"\"{ParagraphText.Excerpt(before)}\" -> \"{ParagraphText.Excerpt(after)}\"";
            yield return new Finding("comment-changed", Severity.Error,
                $"pre-existing {o.Name} changed: {what} (changed without tracking; Reject All does not restore it)", c.Part, o.Name);
        }
        foreach (var c in candidate.Comments.Where(s => !originals.Contains(s.Name)))
            yield return new Finding("comment-added", Severity.Info, $"{c.Name} added{Excerpt(c)}", c.Part, c.Name);

        static bool Same(Story a, Story b) =>
            a.Skeleton == b.Skeleton && a.Paragraphs.Select(p => p.Key).SequenceEqual(b.Paragraphs.Select(p => p.Key));
    }
}
