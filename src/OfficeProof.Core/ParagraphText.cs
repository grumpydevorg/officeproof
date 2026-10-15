using System.Text;
using System.Xml.Linq;

namespace OfficeProof.Core;

/// <summary>
/// A run of visible text (one w:t or w:tab) and the fields whose cached result it belongs to.
/// <paramref name="Fields"/> lists the enclosing field instructions, outermost first; it is empty
/// for ordinary typed text.
/// </summary>
public sealed record TextSegment(XElement Node, string Text, IReadOnlyList<string> Fields)
{
    /// <summary>True when this text is (part of) the cached result of any field.</summary>
    public bool IsFieldResult => Fields.Count > 0;

    /// <summary>True when an enclosing field's instruction starts with <paramref name="fieldType"/> ("TOC", "REF", "SEQ").</summary>
    public bool InField(string fieldType) => Fields.Any(f => FieldText.IsType(f, fieldType));
}

/// <summary>
/// A paragraph's visible text as Word displays it with all markup shown (see
/// <see cref="Wml.VisibleText"/>), split into segments that remember which field results they sit in.
/// </summary>
public sealed class ParagraphText
{
    internal ParagraphText(XElement paragraph, int index, IReadOnlyList<TextSegment> segments)
    {
        Paragraph = paragraph;
        Index = index;
        Segments = segments;
        Text = string.Concat(segments.Select(s => s.Text));
    }

    /// <summary>The w:p element.</summary>
    public XElement Paragraph { get; }

    /// <summary>1-based position among the part's paragraphs in document order (table cells included).</summary>
    public int Index { get; }

    /// <summary>Visible text segments in order.</summary>
    public IReadOnlyList<TextSegment> Segments { get; }

    /// <summary>The full visible text, field results included.</summary>
    public string Text { get; }

    /// <summary>True when any visible text of the paragraph sits in the result of a <paramref name="fieldType"/> field.</summary>
    public bool InField(string fieldType) => Segments.Any(s => s.InField(fieldType));

    /// <summary>
    /// <see cref="Text"/> with every field-result character replaced by U+FFFC, so offsets still line
    /// up but generated text (REF/SEQ/TOC results) can never match a pattern for typed text.
    /// </summary>
    public string TypedText()
    {
        var sb = new StringBuilder(Text.Length);
        foreach (var s in Segments)
            sb.Append(s.IsFieldResult ? new string(FieldText.Masked, s.Text.Length) : s.Text);
        return sb.ToString();
    }

    /// <summary>Human locator: "¶12 \"first 60 characters…\"".</summary>
    public string Location => $"¶{Index} \"{Excerpt(Text)}\"";

    /// <summary>Whitespace-collapsed text cut to <paramref name="max"/> characters with an ellipsis, for messages.</summary>
    public static string Excerpt(string text, int max = 60)
    {
        var flat = string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return flat.Length <= max ? flat : string.Concat(flat.AsSpan(0, max - 1), "…");
    }
}

/// <summary>Field-instruction parsing: the field type ("REF") and first argument (a bookmark name).</summary>
public static class FieldText
{
    public const char Masked = '\uFFFC';

    /// <summary>True when <paramref name="instruction"/> is a field of <paramref name="fieldType"/>.</summary>
    public static bool IsType(string instruction, string fieldType) =>
        string.Equals(Type(instruction), fieldType, StringComparison.OrdinalIgnoreCase);

    /// <summary>The field type: the instruction's first word ("REF", "SEQ", "TOC").</summary>
    public static string Type(string instruction)
    {
        var t = instruction.TrimStart();
        var end = t.IndexOfAny([' ', '\t', '\\']);
        return end < 0 ? t : t[..end];
    }

    /// <summary>The instruction's first argument (bookmark for REF, identifier for SEQ), unquoted.</summary>
    public static string? Argument(string instruction)
    {
        var parts = instruction.Trim().Split((char[]?)null, 3, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length < 2 || parts[1].StartsWith('\\') ? null : parts[1].Trim('"');
    }
}

/// <summary>
/// Walks a story part in document order and attributes every visible text node to the
/// complex (fldChar begin/separate/end) and simple (w:fldSimple) fields whose cached result
/// contains it. Fields span paragraphs (a TOC result is many paragraphs), so this is done
/// once per part rather than per paragraph.
/// </summary>
public static class FieldWalker
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private sealed class Frame
    {
        public StringBuilder Instruction { get; } = new();
        public bool InResult { get; set; }
    }

    public static IReadOnlyList<ParagraphText> Walk(XElement root)
    {
        var paragraphs = new List<XElement>();
        var segments = new Dictionary<XElement, List<TextSegment>>();
        var stack = new List<Frame>();

        foreach (var e in root.Descendants())
        {
            if (e.Name == W + "p")
            {
                if (!InFallback(e))
                {
                    paragraphs.Add(e);
                    segments[e] = [];
                }
                continue;
            }
            if (e.Name == W + "fldChar")
            {
                switch ((string?)e.Attribute(W + "fldCharType"))
                {
                    case "begin":
                        stack.Add(new Frame());
                        break;
                    case "separate" when stack.Count > 0:
                        stack[^1].InResult = true;
                        break;
                    case "end" when stack.Count > 0:
                        stack.RemoveAt(stack.Count - 1);
                        break;
                }
                continue;
            }
            if (e.Name == W + "instrText" || e.Name == W + "delInstrText")
            {
                if (stack.Count > 0 && !stack[^1].InResult) stack[^1].Instruction.Append(e.Value);
                continue;
            }
            var isText = e.Name == W + "t" || (e.Name == W + "tab" && e.Parent?.Name == W + "r");
            if (!isText) continue;
            // Text between a field's begin and separate is part of its instruction: Word never shows it.
            if (stack.Any(f => !f.InResult)) continue;
            if (e.Ancestors().Any(a => a.Name == W + "del" || a.Name == W + "moveFrom" || a.Name == Mc + "Fallback")) continue;
            var p = e.Ancestors(W + "p").FirstOrDefault();
            if (p is null || !segments.TryGetValue(p, out var list)) continue;

            var fields = stack.Select(f => f.Instruction.ToString().Trim()).ToList();
            fields.AddRange(e.Ancestors(W + "fldSimple").Reverse()
                .Select(s => ((string?)s.Attribute(W + "instr") ?? "").Trim()));
            list.Add(new TextSegment(e, e.Name == W + "tab" ? "\t" : e.Value, fields));
        }

        return [.. paragraphs.Select((p, i) => new ParagraphText(p, i + 1, segments[p]))];
    }

    private static bool InFallback(XElement e) => e.Ancestors(Mc + "Fallback").Any();
}
