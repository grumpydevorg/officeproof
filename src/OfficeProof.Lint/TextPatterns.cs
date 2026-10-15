using OfficeProof.Core;

namespace OfficeProof.Lint;

/// <summary>The paragraph sets the numbering rules share, built on Core's <see cref="TypedNumbers"/>.</summary>
internal static class TextPatterns
{
    /// <summary>Main-body paragraphs that are not generated TOC / table-of-figures lines.</summary>
    public static IEnumerable<ParagraphText> BodyParagraphs(LintContext ctx) =>
        ctx.Paragraphs(ctx.Main).Where(p => !p.InField("TOC"));

    /// <summary>Headings (effective outline level 0-8) with the typed number at their start, if any.</summary>
    public static IEnumerable<(ParagraphText Paragraph, string? Number)> Headings(LintContext ctx) =>
        BodyParagraphs(ctx)
            .Where(p => ctx.Styles.OutlineLevel(p.Paragraph) is not null)
            .Select(p => (p, TypedNumbers.HeadingNumber().Match(p.TypedText()) is { Success: true } m ? m.Groups["n"].Value : null));

    /// <summary>
    /// Caption paragraphs: a caption-shaped opening ("Figure 1."), a SEQ field, or a style in the
    /// "caption" chain, provided the paragraph opens with "Figure N"/"Table N" (chapter-style
    /// "Figure 2.1" is not a single counter and is never a caption here).
    /// Returns the label, the displayed number, and the length of the caption label.
    /// </summary>
    public static IEnumerable<(ParagraphText Paragraph, string Label, string Number, int LabelEnd)> Captions(LintContext ctx)
    {
        foreach (var p in BodyParagraphs(ctx))
        {
            var start = TypedNumbers.CaptionStart().Match(p.Text);
            if (!start.Success) continue;
            var isCaption = TypedNumbers.Caption().IsMatch(p.Text)
                            || Wml.HasField(p.Paragraph, "SEQ")
                            || ctx.Styles.Chain(p.Paragraph).Any(s => string.Equals(
                                (string?)s.Element(Ns.W + "name")?.Attribute(Ns.W + "val"), "caption", StringComparison.OrdinalIgnoreCase));
            if (isCaption) yield return (p, TypedNumbers.Label(start.Groups["label"].Value), start.Groups["n"].Value, start.Length);
        }
    }

    /// <summary>
    /// Typed figure/table references in <paramref name="typed"/> (field results already masked),
    /// from <paramref name="from"/> on: each number of a plural list ("Figures 7 and 9") counts.
    /// Returns (offset, "Figure 7") pairs.
    /// </summary>
    public static IEnumerable<(int Index, string Text)> FigureAndTableReferences(string typed, int from)
    {
        foreach (var m in TypedNumbers.Reference().Matches(typed).Where(m => m.Index >= from))
            yield return (m.Index, $"{TypedNumbers.Label(m.Groups["label"].Value)} {m.Groups["n"].Value}");
        foreach (var m in TypedNumbers.PluralReference().Matches(typed).Where(m => m.Index >= from))
        {
            var label = TypedNumbers.Label(m.Groups["label"].Value);
            foreach (var c in m.Groups["n"].Captures.Cast<System.Text.RegularExpressions.Capture>())
                yield return (c.Index, $"{label} {c.Value}");
        }
    }
}
