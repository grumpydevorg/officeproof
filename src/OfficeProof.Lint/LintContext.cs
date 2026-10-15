using System.Globalization;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Lint;

/// <summary>
/// What a rule sees: the open (read-only) document, each part's XML loaded once, field-aware
/// paragraph text, paragraph-style resolution (Core <see cref="StyleChain"/>), and the run
/// properties a paragraph style gives. One context serves every rule in a run.
/// </summary>
public sealed class LintContext
{
    private static readonly XNamespace W = Ns.W;
    private readonly Dictionary<Uri, XDocument> _xml = [];
    private readonly Dictionary<Uri, IReadOnlyList<ParagraphText>> _paragraphs = [];
    private readonly Dictionary<Uri, Dictionary<XElement, ParagraphText>> _byElement = [];
    private readonly HashSet<int> _numIds;

    /// <summary>Wraps an open document; the caller owns (and disposes) it.</summary>
    public LintContext(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        Document = document;
        Main = document.MainDocumentPart ?? throw new InvalidOperationException("the package has no main document part");
        StylesXml = Main.StyleDefinitionsPart is { } sp ? Xml(sp) : null;
        NumberingXml = Main.NumberingDefinitionsPart is { } np ? Xml(np) : null;
        Styles = StyleChain.FromStyles(StylesXml?.Root);
        _numIds = [.. NumberingXml?.Root?.Elements(W + "num")
            .Select(n => int.TryParse((string?)n.Attribute(W + "numId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 0)
            .Where(v => v > 0) ?? []];
    }

    /// <summary>Style resolution: basedOn chain, effective outline level and numId.</summary>
    public StyleChain Styles { get; }

    /// <summary>The open package.</summary>
    public WordprocessingDocument Document { get; }

    /// <summary>The main document part (/word/document.xml).</summary>
    public MainDocumentPart Main { get; }

    /// <summary>styles.xml, or null when the package has none.</summary>
    public XDocument? StylesXml { get; }

    /// <summary>numbering.xml, or null when the package has none.</summary>
    public XDocument? NumberingXml { get; }

    /// <summary>The part's XML, loaded once per run (whitespace preserved).</summary>
    public XDocument Xml(OpenXmlPart part)
    {
        ArgumentNullException.ThrowIfNull(part);
        if (!_xml.TryGetValue(part.Uri, out var x)) _xml[part.Uri] = x = Docx.Load(part);
        return x;
    }

    /// <summary>The part's paragraphs in document order with field-aware visible text.</summary>
    public IReadOnlyList<ParagraphText> Paragraphs(OpenXmlPart part)
    {
        ArgumentNullException.ThrowIfNull(part);
        if (!_paragraphs.TryGetValue(part.Uri, out var list))
            _paragraphs[part.Uri] = list = Xml(part).Root is { } root ? FieldWalker.Walk(root) : [];
        return list;
    }

    /// <summary>The part URI as findings report it ("/word/document.xml").</summary>
    public static string PartName(OpenXmlPart part)
    {
        ArgumentNullException.ThrowIfNull(part);
        return part.Uri.ToString();
    }

    /// <summary>True when styles.xml defines <paramref name="styleId"/>.</summary>
    public bool IsStyleDefined(string styleId) => Styles.Style(styleId) is not null;

    /// <summary>
    /// True when Word shows a list number on the paragraph: an effective numId (paragraph or
    /// style chain, 0 = off) that numbering.xml actually defines (an undefined numId renders nothing).
    /// </summary>
    public bool HasNumbering(XElement paragraph) => Styles.NumId(paragraph) is { } id && _numIds.Contains(id);

    /// <summary>
    /// The run-property levels a paragraph style contributes, highest priority first: each style's
    /// w:rPr along the basedOn chain, then docDefaults/rPrDefault. Entries may be null. An
    /// undefined style id falls back to the default paragraph style, as Word renders it.
    /// </summary>
    public IEnumerable<XElement?> StyleRunLevels(string? styleId) =>
        Styles.Chain(styleId is not null && IsStyleDefined(styleId) ? styleId : Styles.DefaultParagraphStyle)
            .Select(s => s.Element(W + "rPr"))
            .Append(StylesXml?.Root?.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr"));

    /// <summary>
    /// Effective run property a paragraph style gives: the first level in <see cref="StyleRunLevels"/>
    /// whose rPr child <paramref name="element"/> carries <paramref name="attribute"/>. Each attribute
    /// inherits independently, as w:rFonts attributes do.
    /// </summary>
    public string? StyleRunProperty(string? styleId, string element, string attribute)
    {
        foreach (var rPr in StyleRunLevels(styleId))
            if (rPr?.Element(W + element)?.Attribute(W + attribute) is { } a) return a.Value;
        return null;
    }

    /// <summary>
    /// Locator for any element of a story part: the paragraph containing it (for a table
    /// property, the table's first paragraph), as "¶N \"excerpt\"". Null when there is none.
    /// </summary>
    public string? Locate(OpenXmlPart part, XElement element)
    {
        ArgumentNullException.ThrowIfNull(element);
        var p = element.AncestorsAndSelf(W + "p").FirstOrDefault()
                ?? element.AncestorsAndSelf(W + "tbl").FirstOrDefault()?.Descendants(W + "p").FirstOrDefault();
        if (p is null) return null;
        if (!_byElement.TryGetValue(part.Uri, out var map))
            _byElement[part.Uri] = map = Paragraphs(part).ToDictionary(t => t.Paragraph);
        return map.TryGetValue(p, out var text) ? text.Location : null;
    }
}
