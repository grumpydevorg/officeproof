using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeProof.Core;

namespace OfficeProof.Redline;

/// <summary>
/// Retargets one typed in-text reference before it becomes a field: in the paragraph identified
/// by <paramref name="ParagraphSnippet"/>, "<paramref name="Label"/> <paramref name="Number"/>"
/// (or that number inside "Figures 7 and 8") points at <paramref name="NewNumber"/> instead.
/// <paramref name="Nth"/> (0-based, among references to that label and number in the paragraph)
/// picks one occurrence; null retargets all of them.
/// A client report (2026-10-15): two "Figure 8" references meant Figures 7 and 9.
/// </summary>
public sealed record ReferenceFix(string ParagraphSnippet, string Label, int Number, int NewNumber, int? Nth = null);

/// <summary>Options for <see cref="CaptionFields.Apply"/>.</summary>
public sealed record CaptionOptions
{
    /// <summary>Add a comment on the first caption explaining the change and the F9 update.</summary>
    public bool ExplainInComment { get; init; } = true;
}

/// <summary>
/// Typed caption numbers → SEQ fields inside "_Ref" bookmarks; typed references → REF fields to
/// those bookmarks; all tracked, so Reject All restores the typed text. What counts as a typed
/// caption or reference is <see cref="TypedNumbers"/>: chapter-style numbers ("Figure 2.1") are
/// left alone because they are not one SEQ counter.
/// </summary>
public static class CaptionFields
{
    private sealed class Caption(XElement paragraph, XElement insertion, string bookmark)
    {
        public XElement Paragraph { get; } = paragraph;

        /// <summary>The w:ins holding the SEQ field.</summary>
        public XElement Insertion { get; } = insertion;

        /// <summary>Bookmark over label and number ("Figure 7"), for single references.</summary>
        public string Bookmark { get; } = bookmark;

        /// <summary>Bookmark over the number alone ("7"), made on demand for plural references.</summary>
        public string? NumberBookmark { get; set; }
    }

    private sealed record Reference(int Start, int End, string Label, int Number, bool Plural);

    /// <summary>
    /// Converts captions then references. "Figures 7 and 9" becomes one REF per number (to a
    /// bookmark over the caption's number alone) with the text between kept. A reference with no
    /// matching caption gets a comment instead of a field. Each <see cref="ReferenceFix"/> must
    /// match one paragraph and a reference in it, else this throws naming it.
    /// </summary>
    public static CaptionFieldsResult Apply(TrackedEditor editor, IReadOnlyList<ReferenceFix>? referenceFixes = null, CaptionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(editor);
        options ??= new CaptionOptions();

        var fixes = (referenceFixes ?? []).Select(f => (Fix: f with { Label = TypedNumbers.Label(f.Label) }, Paragraph: editor.FindParagraph(f.ParagraphSnippet))).ToList();
        var usedFixes = new HashSet<ReferenceFix>();

        List<AddedComment> comments = [];
        List<string> notes = [];
        List<string> unresolved = [];
        var captions = new Dictionary<(string, int), Caption>();
        var captionLabelAt = new Dictionary<XElement, int>();
        XElement? firstCaption = null;

        // Captions: "Figure 7." → "Figure " + { SEQ Figure \* ARABIC } inside bookmark _Ref…
        foreach (var p in editor.Paragraphs.ToList())
        {
            if (editor.StartsInsideField(p)) continue;
            var text = TrackedEditor.Text(p);
            var m = TypedNumbers.Caption().Match(text);
            if (!m.Success) continue;
            var labelAt = m.Groups["label"].Index;
            var n = m.Groups["n"];
            if (!editor.IsPlainText(p, labelAt, n.Index + n.Length)) continue;
            var label = TypedNumbers.Label(m.Groups["label"].Value);
            var number = int.Parse(n.Value, CultureInfo.InvariantCulture);
            captionLabelAt[p] = labelAt;
            if (captions.ContainsKey((label, number)))
            {
                notes.Add($"duplicate caption {label} {number}: left typed; references go to the first");
                comments.Add(editor.Comment(p, p, $"A second caption numbered {label} {number}; it was left as typed text. Renumber it and its references."));
                continue;
            }
            var change = editor.Replace(p, n.Index, n.Index + n.Length, [Inline.Field($@"SEQ {label} \* ARABIC", n.Value)]);
            var labelRuns = editor.IsolateRuns(p, labelAt, n.Index);
            var name = editor.NewRefBookmarkName();
            editor.Bookmark(p, labelRuns[0], change.Insertion!, name);
            captions[(label, number)] = new Caption(p, change.Insertion!, name);
            firstCaption ??= p;
        }
        foreach (var label in captions.Keys.Select(k => k.Item1).Distinct().Order(StringComparer.Ordinal))
        {
            var nums = captions.Keys.Where(k => k.Item1 == label).Select(k => k.Item2).Order().ToList();
            if (!nums.SequenceEqual(Enumerable.Range(1, nums.Count)))
                notes.Add($"{label} captions are numbered {string.Join(", ", nums)}; SEQ fields renumber them 1..{nums.Count} when updated");
        }

        // References: "Figure 7" → { REF _Ref… \h } showing "Figure 7"; "Figures 7 and 9" →
        // "Figures " { REF …7 } " and " { REF …9 }, each showing the number only.
        var references = 0;
        var retargeted = 0;
        foreach (var p in editor.Paragraphs.ToList())
        {
            if (editor.StartsInsideField(p)) continue;
            var text = TrackedEditor.Text(p);
            List<Reference> found = [];
            foreach (Match m in TypedNumbers.Reference().Matches(text))
            {
                if (captionLabelAt.TryGetValue(p, out var at) && m.Index == at) continue;
                found.Add(new Reference(m.Index, m.Index + m.Length, TypedNumbers.Label(m.Groups["label"].Value),
                    int.Parse(m.Groups["n"].Value, CultureInfo.InvariantCulture), Plural: false));
            }
            foreach (Match m in TypedNumbers.PluralReference().Matches(text))
            {
                var label = TypedNumbers.Label(m.Groups["label"].Value);
                foreach (Capture c in m.Groups["n"].Captures)
                    found.Add(new Reference(c.Index, c.Index + c.Length, label, int.Parse(c.Value, CultureInfo.InvariantCulture), Plural: true));
            }
            found = [.. found.Where(r => editor.IsPlainText(p, r.Start, r.End)).OrderBy(r => r.Start)];

            var seen = new Dictionary<(string, int), int>();
            var planned = found.Select(r =>
            {
                var occurrence = seen[(r.Label, r.Number)] = seen.GetValueOrDefault((r.Label, r.Number), -1) + 1;
                var fix = fixes.FirstOrDefault(f => f.Paragraph == p && f.Fix.Label == r.Label && f.Fix.Number == r.Number
                                                    && (f.Fix.Nth is null || f.Fix.Nth == occurrence)).Fix;
                return (Ref: r, Fix: fix);
            }).ToList();

            // Last to first, so earlier offsets stay valid.
            foreach (var (r, fix) in Enumerable.Reverse(planned))
            {
                var target = fix?.NewNumber ?? r.Number;
                if (fix is not null) usedFixes.Add(fix);
                if (!captions.TryGetValue((r.Label, target), out var caption))
                {
                    unresolved.Add($"{r.Label} {target} in \"{ParagraphText.Excerpt(text)}\"");
                    var runs = editor.IsolateRuns(p, r.Start, r.End);
                    comments.Add(editor.Comment(runs[0], runs[^1],
                        $"No caption {r.Label} {target} found, so this reference is left as typed text. Check which {r.Label.ToLowerInvariant()} is meant."));
                    continue;
                }
                TrackedReplacement change;
                if (r.Plural)
                {
                    if (caption.NumberBookmark is null)
                    {
                        caption.NumberBookmark = editor.NewRefBookmarkName();
                        editor.Bookmark(caption.Paragraph, caption.Insertion, caption.Insertion, caption.NumberBookmark);
                    }
                    var shown = target.ToString(CultureInfo.InvariantCulture);
                    change = editor.Replace(p, r.Start, r.End, [Inline.Field($@"REF {caption.NumberBookmark} \h", shown)]);
                }
                else
                {
                    change = editor.Replace(p, r.Start, r.End, [Inline.Field($@"REF {caption.Bookmark} \h", $"{r.Label} {target}")]);
                }
                references++;
                if (fix is not null)
                {
                    retargeted++;
                    comments.Add(editor.Comment(change.First, change.Last,
                        $"Was {r.Label} {r.Number}; now refers to {r.Label} {target}. Revert if {r.Label} {r.Number} was intended."));
                }
            }
        }
        foreach (var (fix, _) in fixes.Where(f => !usedFixes.Contains(f.Fix)))
            throw new InvalidOperationException($"reference fix \"{fix.ParagraphSnippet}\": no convertible {fix.Label} {fix.Number} reference (nth {fix.Nth?.ToString(CultureInfo.InvariantCulture) ?? "any"}) in that paragraph");

        // Anchored on the whole caption paragraph, not the SEQ insertion: markers inside the
        // _Ref bookmark would travel into every REF result (2026-10-15 review).
        if (options.ExplainInComment && firstCaption is not null)
            comments.Insert(0, editor.Comment(firstCaption, firstCaption,
                "Figure and table numbers were typed by hand, so they do not follow when figures move. Every caption "
                + "number is now a SEQ field and every in-text reference a cross-reference (REF) field. After accepting: "
                + "Ctrl+A, F9 to update fields before exporting the PDF. New references: Insert > Cross-reference > "
                + "'Only label and number'."));

        return new CaptionFieldsResult(captions.Count, references, retargeted, unresolved, comments, notes);
    }
}
