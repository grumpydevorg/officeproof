using System.Globalization;
using System.Xml.Linq;

namespace OfficeProof.Redline;

/// <summary>
/// Who and when every tracked change and comment in one redline pass is attributed to.
/// Word shows <paramref name="Name"/> in the balloon and <paramref name="Initials"/> on comments.
/// </summary>
public sealed record RevisionAuthor(string Name, string Initials, DateTimeOffset Date)
{
    /// <summary>The w:date value: UTC, ISO 8601, whole seconds (Word's own format).</summary>
    public string WordDate => Date.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
}

/// <summary>Content inserted by a tracked replacement: plain text or a complex field.</summary>
public abstract record Inline
{
    /// <summary>Plain text; '\t' becomes a w:tab.</summary>
    public static Inline Text(string text) => new TextInline(text);

    /// <summary>A complex field <c>{ code }</c> showing <paramref name="cachedResult"/> until Word updates it.</summary>
    public static Inline Field(string code, string cachedResult) => new FieldInline(code, cachedResult);
}

/// <summary>Plain inserted text.</summary>
public sealed record TextInline(string Value) : Inline;

/// <summary>An inserted complex field (begin, instrText, separate, cached result, end).</summary>
public sealed record FieldInline(string Code, string CachedResult) : Inline;

/// <summary>
/// The markup one tracked replacement produced. A deletion spanning runs with different
/// parents (e.g. part inside a hyperlink) becomes several adjacent w:del elements.
/// </summary>
public sealed record TrackedReplacement(IReadOnlyList<XElement> Deletions, XElement? Insertion)
{
    /// <summary>First element of the change, for anchoring a comment.</summary>
    public XElement First => Deletions.Count > 0 ? Deletions[0] : Insertion!;

    /// <summary>Last element of the change, for anchoring a comment.</summary>
    public XElement Last => Insertion ?? Deletions[^1];

    /// <summary>True when nothing was changed (identical old and new text).</summary>
    public bool IsEmpty => Deletions.Count == 0 && Insertion is null;
}

/// <summary>A comment the redline added: its id, the text it is anchored to and its body.</summary>
public sealed record AddedComment(int Id, string Anchor, string Text);

/// <summary>What every redline operation reports so the CLI can tell the author.</summary>
public interface IRedlineResult
{
    /// <summary>
    /// Changes Word cannot track (style and numbering definitions). Reject All does not undo
    /// them, so the author must be told about each.
    /// </summary>
    IReadOnlyList<string> UntrackedChanges { get; }

    /// <summary>Comments added to the document.</summary>
    IReadOnlyList<AddedComment> Comments { get; }

    /// <summary>Things the author should check that did not stop the operation.</summary>
    IReadOnlyList<string> Notes { get; }

    /// <summary>
    /// Ids of existing styles whose definitions the operation changed on purpose. verify fails
    /// any other change to an existing style (rule definition-changed), so the redline declares these.
    /// </summary>
    IReadOnlyList<string> ChangedStyles { get; }
}

/// <summary>Result of <see cref="EditScript.Apply"/>.</summary>
public sealed record EditScriptResult(
    int Replacements,
    IReadOnlyList<AddedComment> Comments) : IRedlineResult
{
    public IReadOnlyList<string> UntrackedChanges => [];
    public IReadOnlyList<string> Notes => [];
    public IReadOnlyList<string> ChangedStyles => [];
}

/// <summary>Result of <see cref="CaptionFields.Apply"/>.</summary>
public sealed record CaptionFieldsResult(
    int Captions,
    int References,
    int Retargeted,
    IReadOnlyList<string> UnresolvedReferences,
    IReadOnlyList<AddedComment> Comments,
    IReadOnlyList<string> Notes) : IRedlineResult
{
    public IReadOnlyList<string> UntrackedChanges => [];
    public IReadOnlyList<string> ChangedStyles => [];
}

/// <summary>Result of <see cref="HeadingNumbering.Apply"/>.</summary>
public sealed record HeadingNumberingResult(
    int NumberedHeadings,
    IReadOnlyList<string> UnnumberedHeadings,
    int SectionReferences,
    IReadOnlyList<string> UntrackedChanges,
    IReadOnlyList<AddedComment> Comments,
    IReadOnlyList<string> Notes) : IRedlineResult
{
    /// <summary>Heading numbering only adds styles and lists; it changes no existing definition.</summary>
    public IReadOnlyList<string> ChangedStyles => [];
}

/// <summary>Result of <see cref="TocStyles.SetFont"/>.</summary>
public sealed record TocFontResult(
    IReadOnlyList<string> ChangedStyles,
    IReadOnlyList<string> UntrackedChanges,
    IReadOnlyList<string> Notes) : IRedlineResult
{
    public IReadOnlyList<AddedComment> Comments => [];
}
