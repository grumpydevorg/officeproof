using System.Globalization;
using OfficeProof.Core;

namespace OfficeProof.Lint;

// The 2026-10-15 pilot report numbered everything by hand: headings, captions and
// cross-references were typed characters and only the TOC was a live field. Inserting or
// reordering a section therefore renumbers nothing, and it had already gone wrong: two
// headings read "2.0". These rules find the typed numbers a redline has to replace with
// numbered heading styles, SEQ captions and REF cross-references.

/// <summary>
/// typed-heading-number: a heading whose number is typed text, with no list numbering in effect.
/// A heading with both is <see cref="DoubleHeadingNumberRule"/>'s (an error, not debt).
/// </summary>
public sealed class TypedHeadingNumberRule : ILintRule
{
    public string Id => "typed-heading-number";
    public Severity Severity => Severity.Warning;
    public string Description => "headings numbered by typed text instead of a numbered heading style";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        foreach (var (p, number) in TextPatterns.Headings(ctx))
        {
            if (number is null || ctx.HasNumbering(p.Paragraph)) continue;
            yield return new Finding(Id, Severity,
                $"heading number '{number}' is typed text; adding or moving a section will not renumber it",
                part, p.Location);
        }
    }
}

// A heading that already carries list numbering AND keeps its typed number shows both:
// "1.0 1.0 Summary". The pilot report's v2 redline (2026-10-15) produced exactly that in Word after
// Reject All left the numbering linked into the style but restored the typed text. It is a
// separate rule rather than a severity switch inside typed-heading-number so each rule keeps
// one severity: a typed number alone is debt (Warning), a doubled number is visibly wrong (Error).
/// <summary>double-heading-number: a heading with list numbering in effect that also starts with a typed number.</summary>
public sealed class DoubleHeadingNumberRule : ILintRule
{
    public string Id => "double-heading-number";
    public Severity Severity => Severity.Error;
    public string Description => "headings showing a list number and a typed number (\"1.0 1.0 Summary\")";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        foreach (var (p, number) in TextPatterns.Headings(ctx))
        {
            if (number is null || !ctx.HasNumbering(p.Paragraph)) continue;
            yield return new Finding(Id, Severity,
                $"heading has list numbering and also starts with the typed number '{number}'; Word shows both - delete the typed number",
                part, p.Location);
        }
    }
}

/// <summary>duplicate-heading-number: two headings carry the same typed number (the 2026-10-15 pilot report had two "2.0").</summary>
public sealed class DuplicateHeadingNumberRule : ILintRule
{
    public string Id => "duplicate-heading-number";
    public Severity Severity => Severity.Error;
    public string Description => "the same typed number on two headings";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        var first = new Dictionary<string, ParagraphText>();
        foreach (var (p, number) in TextPatterns.Headings(ctx))
        {
            if (number is null) continue;
            if (first.TryGetValue(number, out var earlier))
                yield return new Finding(Id, Severity,
                    $"heading number '{number}' already used at {earlier.Location}", part, p.Location);
            else first[number] = p;
        }
    }
}

/// <summary>typed-caption-number: a "Figure N." / "Table N:" caption without a SEQ field.</summary>
public sealed class TypedCaptionNumberRule : ILintRule
{
    public string Id => "typed-caption-number";
    public Severity Severity => Severity.Warning;
    public string Description => "figure/table captions numbered by typed text instead of a SEQ field";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        foreach (var p in TextPatterns.BodyParagraphs(ctx))
        {
            if (TypedNumbers.Caption().Match(p.Text) is not { Success: true } m || Wml.HasField(p.Paragraph, "SEQ")) continue;
            yield return new Finding(Id, Severity,
                $"caption number '{m.Groups["label"].Value} {m.Groups["n"].Value}' is typed text, not a SEQ field; it will not renumber",
                part, p.Location);
        }
    }
}

/// <summary>
/// typed-cross-reference: running text naming "Figure N", "Table N" or "section N.N" as typed
/// characters rather than a REF field. Field results (REF, SEQ, TOC) and a caption's own label
/// are not references.
/// </summary>
public sealed class TypedCrossReferenceRule : ILintRule
{
    public string Id => "typed-cross-reference";
    public Severity Severity => Severity.Warning;
    public string Description => "cross-references typed as text instead of REF fields";

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        foreach (var (p, refs) in TypedReferences(ctx))
        {
            var list = string.Join(", ", refs.Select(r => $"'{r}'").Distinct());
            yield return new Finding(Id, Severity,
                $"typed cross-reference(s) {list}; they will not follow renumbering - use REF fields",
                part, p.Location);
        }
    }

    /// <summary>
    /// Per paragraph, the typed references in document order ("Figure 1", "section 3.1"). Every
    /// number of a plural list counts ("Figures 7 and 9" → Figure 7, Figure 9); chapter-style
    /// "Figure 2.1" is not a reference to Figure 2.
    /// </summary>
    internal static IEnumerable<(ParagraphText Paragraph, IReadOnlyList<string> References)> TypedReferences(LintContext ctx)
    {
        var captionLabelEnd = TextPatterns.Captions(ctx).ToDictionary(c => c.Paragraph, c => c.LabelEnd);
        foreach (var p in TextPatterns.BodyParagraphs(ctx))
        {
            var typed = p.TypedText();
            var refs = TextPatterns.FigureAndTableReferences(typed, captionLabelEnd.GetValueOrDefault(p))
                .Concat(TypedNumbers.SectionReference().Matches(typed).Select(m => (m.Index, Text: $"section {m.Groups["n"].Value}")))
                .OrderBy(r => r.Index)
                .Select(r => r.Text)
                .ToList();
            if (refs.Count > 0) yield return (p, refs);
        }
    }
}

/// <summary>
/// missing-reference-target: a reference Word cannot resolve. Either a typed or REF reference to
/// "Figure N"/"Table N" with no such caption, or a REF/PAGEREF/NOTEREF whose bookmark is missing
/// (Word prints "Error! Reference source not found." on the next field update).
/// </summary>
public sealed class MissingReferenceTargetRule : ILintRule
{
    public string Id => "missing-reference-target";
    public Severity Severity => Severity.Error;
    public string Description => "references to captions or bookmarks that do not exist";

    private static readonly string[] BookmarkFields = ["REF", "PAGEREF", "NOTEREF"];

    public IEnumerable<Finding> Check(LintContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        var part = LintContext.PartName(ctx.Main);
        var captions = TextPatterns.Captions(ctx).Select(c => $"{c.Label} {c.Number}").ToHashSet();
        var bookmarks = Docx.StoryParts(ctx.Main)
            .SelectMany(sp => ctx.Xml(sp).Descendants(Ns.W + "bookmarkStart"))
            .Select(b => (string?)b.Attribute(Ns.W + "name")).OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var (p, refs) in TypedCrossReferenceRule.TypedReferences(ctx))
            foreach (var r in refs.Where(r => !r.StartsWith("section", StringComparison.Ordinal) && !captions.Contains(r)).Distinct())
                yield return new Finding(Id, Severity, $"typed reference to '{r}' but no caption '{r}' exists", part, p.Location);

        foreach (var p in TextPatterns.BodyParagraphs(ctx))
        {
            foreach (var instr in Wml.FieldInstructions(p.Paragraph))
            {
                var type = FieldText.Type(instr);
                if (!BookmarkFields.Contains(type, StringComparer.OrdinalIgnoreCase)) continue;
                var bookmark = FieldText.Argument(instr);
                if (bookmark is not null && !bookmarks.Contains(bookmark))
                    yield return new Finding(Id, Severity,
                        $"{type.ToUpperInvariant()} field points at missing bookmark '{bookmark}'; Word will print 'Error! Reference source not found.'",
                        part, p.Location);
            }
            foreach (var (instr, result) in RefResults(p))
            {
                if (result.Contains(TypedNumbers.BrokenReference, StringComparison.OrdinalIgnoreCase))
                    yield return new Finding(Id, Severity, $"field '{instr}' already shows '{TypedNumbers.BrokenReference}.'", part, p.Location);
                else if (TypedNumbers.CaptionStart().Match(result) is { Success: true } m)
                {
                    var target = string.Create(CultureInfo.InvariantCulture, $"{TypedNumbers.Label(m.Groups["label"].Value)} {m.Groups["n"].Value}");
                    if (!captions.Contains(target))
                        yield return new Finding(Id, Severity, $"REF result '{target}' but no caption '{target}' exists", part, p.Location);
                }
            }
        }
    }

    /// <summary>Each REF field's instruction and cached result text within the paragraph.</summary>
    private static IEnumerable<(string Instruction, string Result)> RefResults(ParagraphText p)
    {
        string? current = null;
        var text = new System.Text.StringBuilder();
        foreach (var s in p.Segments)
        {
            var instr = s.Fields.LastOrDefault(f => FieldText.IsType(f, "REF"));
            if (!string.Equals(instr, current, StringComparison.Ordinal) && current is not null)
            {
                yield return (current, text.ToString());
                text.Clear();
            }
            current = instr;
            if (instr is not null) text.Append(s.Text);
        }
        if (current is not null) yield return (current, text.ToString());
    }
}
