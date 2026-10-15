using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Core;

/// <summary>
/// Resolves paragraph properties through the style basedOn chain: outline level and list
/// numbering, the two things that decide what Word shows as a numbered heading.
/// </summary>
public sealed class StyleChain
{
    private static readonly XNamespace W = Ns.W;
    private readonly Dictionary<string, XElement> _styles;

    private StyleChain(Dictionary<string, XElement> styles, string? defaultParagraphStyle)
    {
        _styles = styles;
        DefaultParagraphStyle = defaultParagraphStyle;
    }

    /// <summary>Reads the styles part of <paramref name="document"/> (an empty chain when there is none).</summary>
    public static StyleChain Load(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var part = document.MainDocumentPart?.StyleDefinitionsPart;
        return FromStyles(part is null ? null : Docx.Load(part).Root);
    }

    /// <summary>Builds a chain from a w:styles element (null for none).</summary>
    public static StyleChain FromStyles(XElement? styles)
    {
        var map = new Dictionary<string, XElement>(StringComparer.Ordinal);
        string? defaultStyle = null;
        foreach (var s in styles?.Elements(W + "style") ?? [])
        {
            if ((string?)s.Attribute(W + "styleId") is not { } id) continue;
            map.TryAdd(id, s);
            if ((string?)s.Attribute(W + "type") == "paragraph" && (string?)s.Attribute(W + "default") is "1" or "true")
                defaultStyle ??= id;
        }
        return new StyleChain(map, defaultStyle);
    }

    /// <summary>The default paragraph style id (usually "Normal"), or null.</summary>
    public string? DefaultParagraphStyle { get; }

    /// <summary>The w:style with this id, or null.</summary>
    public XElement? Style(string id) => _styles.GetValueOrDefault(id);

    /// <summary>The style and its basedOn ancestors, nearest first; cycles are cut.</summary>
    public IEnumerable<XElement> Chain(string? styleId)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var id = styleId;
        while (id is not null && seen.Add(id) && _styles.TryGetValue(id, out var style))
        {
            yield return style;
            id = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
        }
    }

    /// <summary>
    /// The paragraph's style chain: its pStyle, else the default paragraph style. A pStyle naming
    /// a style that does not exist also falls back to the default, as Word does (silently; see
    /// lint rule dangling-style).
    /// </summary>
    public IEnumerable<XElement> Chain(XElement paragraph)
    {
        var id = Wml.StyleId(paragraph);
        return Chain(id is not null && _styles.ContainsKey(id) ? id : DefaultParagraphStyle);
    }

    /// <summary>Effective outline level 0-8 (direct w:outlineLvl first, then the chain), or null for body text (absent or 9).</summary>
    public int? OutlineLevel(XElement paragraph)
    {
        var direct = paragraph.Element(W + "pPr")?.Element(W + "outlineLvl")
                     ?? Chain(paragraph).Select(s => s.Element(W + "pPr")?.Element(W + "outlineLvl")).FirstOrDefault(e => e is not null);
        return direct is not null && int.TryParse((string?)direct.Attribute(W + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)
               && v is >= 0 and <= 8 ? v : null;
    }

    /// <summary>
    /// Effective list instance: the paragraph's own w:numPr/w:numId, else the nearest style's.
    /// 0 (explicitly unnumbered) and absence both return null.
    /// </summary>
    public int? NumId(XElement paragraph)
    {
        var numId = paragraph.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")
                    ?? Chain(paragraph).Select(s => s.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")).FirstOrDefault(e => e is not null);
        return numId is not null && int.TryParse((string?)numId.Attribute(W + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)
               && v > 0 ? v : null;
    }
}
