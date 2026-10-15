using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeProof.Core;

namespace OfficeProof.Redline;

/// <summary>
/// One edit-script step. <c>find</c> is a snippet identifying exactly one paragraph.
/// Replace: <c>replace</c> (old text in that paragraph) + <c>with</c> (new text), <c>nth</c>
/// (0-based) when the old text occurs more than once, optional <c>comment</c> on the change.
/// Comment: <c>comment</c> alone, anchored on <c>anchor</c> (a substring, <c>nth</c> if repeated)
/// or on the whole paragraph.
/// </summary>
public sealed record EditOperation(
    string Find,
    string? Replace = null,
    string? With = null,
    int? Nth = null,
    string? Comment = null,
    string? Anchor = null);

/// <summary>
/// A JSON list of <see cref="EditOperation"/>s applied in order as tracked changes, e.g.
/// <code>[{ "find": "year and distance", "replace": "year and distance", "with": "year, distance and elevation" },
///  { "find": "Calculated a set of 52", "comment": "52 or 63?", "anchor": "52" }]</code>
/// </summary>
public sealed class EditScript
{
    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    public EditScript(IReadOnlyList<EditOperation> operations)
    {
        ArgumentNullException.ThrowIfNull(operations);
        for (var i = 0; i < operations.Count; i++) Check(i, operations[i]);
        Operations = operations;
    }

    /// <summary>The steps, in application order.</summary>
    public IReadOnlyList<EditOperation> Operations { get; }

    /// <summary>Parses a script; malformed JSON or an invalid step throws naming the step.</summary>
    public static EditScript Parse(string json)
    {
        List<EditOperation>? ops;
        try
        {
            ops = JsonSerializer.Deserialize<List<EditOperation>>(json, Json);
        }
        catch (JsonException e)
        {
            throw new FormatException($"edit script: {e.Message}", e);
        }
        return new EditScript(ops ?? throw new FormatException("edit script: expected a JSON array of operations"));
    }

    /// <summary>Reads and parses a script file.</summary>
    public static EditScript Load(string path) => Parse(File.ReadAllText(path));

    /// <summary>
    /// Applies every step through <paramref name="editor"/>. Replacements track only the span
    /// that differs. Throws (naming the snippet) when a snippet matches 0 or several paragraphs
    /// or the old text is missing or ambiguous.
    /// </summary>
    public EditScriptResult Apply(TrackedEditor editor)
    {
        ArgumentNullException.ThrowIfNull(editor);
        var replacements = 0;
        List<AddedComment> comments = [];
        for (var i = 0; i < Operations.Count; i++)
        {
            var op = Operations[i];
            var p = Step(i, () => editor.FindParagraph(op.Find));
            var text = TrackedEditor.Text(p);
            if (op.Replace is { } old)
            {
                var at = Occurrence(i, op, text, old);
                var change = Step(i, () => editor.ReplaceText(p, at, at + old.Length, op.With!), text);
                if (change.IsEmpty) throw new InvalidOperationException($"edit {i + 1} (\"{op.Find}\"): replace and with are identical");
                replacements++;
                if (op.Comment is { } c) comments.Add(editor.Comment(change.First, change.Last, c));
            }
            else if (op.Anchor is { } anchor)
            {
                var at = Occurrence(i, op, text, anchor);
                var runs = Step(i, () => editor.IsolateRuns(p, at, at + anchor.Length), text);
                comments.Add(editor.Comment(runs[0], runs[^1], op.Comment!));
            }
            else
            {
                comments.Add(editor.Comment(p, p, op.Comment!));
            }
        }
        return new EditScriptResult(replacements, comments);
    }

    private static void Check(int i, EditOperation op)
    {
        string Where() => $"edit {i + 1}";
        if (string.IsNullOrEmpty(op.Find)) throw new FormatException($"{Where()}: \"find\" is required");
        if (op.Replace is null && op.With is not null) throw new FormatException($"{Where()} (\"{op.Find}\"): \"with\" needs \"replace\"");
        if (op.Replace is not null && op.With is null) throw new FormatException($"{Where()} (\"{op.Find}\"): \"replace\" needs \"with\"");
        if (op.Replace is "") throw new FormatException($"{Where()} (\"{op.Find}\"): \"replace\" is empty");
        if (op.Replace is null && string.IsNullOrEmpty(op.Comment)) throw new FormatException($"{Where()} (\"{op.Find}\"): needs \"replace\"/\"with\" or \"comment\"");
        if (op.Replace is not null && op.Anchor is not null) throw new FormatException($"{Where()} (\"{op.Find}\"): \"anchor\" applies to comment-only steps");
        if (op.Anchor is "") throw new FormatException($"{Where()} (\"{op.Find}\"): \"anchor\" is empty");
        if (op.Nth is < 0) throw new FormatException($"{Where()} (\"{op.Find}\"): \"nth\" must be >= 0");
    }

    private static int Occurrence(int i, EditOperation op, string text, string needle)
    {
        List<int> hits = [];
        for (var at = text.IndexOf(needle, StringComparison.Ordinal); at >= 0; at = text.IndexOf(needle, at + needle.Length, StringComparison.Ordinal))
            hits.Add(at);
        if (hits.Count == 0) throw new InvalidOperationException($"edit {i + 1} (\"{op.Find}\"): \"{needle}\" not found in the paragraph");
        if (op.Nth is { } n)
            return n < hits.Count ? hits[n] : throw new InvalidOperationException($"edit {i + 1} (\"{op.Find}\"): nth {n} but \"{needle}\" occurs {hits.Count} time(s)");
        return hits.Count == 1 ? hits[0] : throw new InvalidOperationException($"edit {i + 1} (\"{op.Find}\"): \"{needle}\" occurs {hits.Count} times; give \"nth\"");
    }

    private static T Step<T>(int i, Func<T> action, string? paragraph = null)
    {
        try
        {
            return action();
        }
        catch (InvalidOperationException e)
        {
            var where = paragraph is null ? "" : $" in paragraph \"{ParagraphText.Excerpt(paragraph)}\"";
            throw new InvalidOperationException($"edit {i + 1}{where}: {e.Message}", e);
        }
    }
}
