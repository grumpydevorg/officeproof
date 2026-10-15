using System.Globalization;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Redline;

/// <summary>
/// Table-of-contents styles. TOC Update Field rebuilds every line from the TOC1..9 styles and
/// discards direct run formatting (a client report, 2026-10-15: Arial lines became Times New Roman from
/// Normal). Putting the font in the styles keeps it; style edits cannot be tracked.
/// </summary>
public static class TocStyles
{
    private static readonly XNamespace W = Ns.W;
    private static readonly string[] ThemeFontAttributes = ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"];

    /// <summary>
    /// Sets <paramref name="font"/> (ascii, hAnsi, cs) in each existing "toc 1".."toc 9" style.
    /// Levels without a style are listed in the notes (Word creates them from Normal on update).
    /// </summary>
    public static TocFontResult SetFont(WordprocessingDocument document, string font)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentException.ThrowIfNullOrWhiteSpace(font);
        var part = document.MainDocumentPart?.StyleDefinitionsPart ?? throw new InvalidOperationException("document has no styles part");
        var styles = Docx.Load(part);
        List<string> changed = [];
        List<string> untracked = [];
        List<string> missing = [];
        for (var level = 1; level <= 9; level++)
        {
            var name = "toc " + level.ToString(CultureInfo.InvariantCulture);
            var style = styles.Root!.Elements(W + "style").FirstOrDefault(s =>
                string.Equals((string?)s.Element(W + "name")?.Attribute(W + "val"), name, StringComparison.OrdinalIgnoreCase));
            if (style is null)
            {
                missing.Add(name);
                continue;
            }
            var rPr = style.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                // CT_Style: … pPr, rPr, tblPr, trPr, tcPr, tblStylePr
                var after = style.Elements().FirstOrDefault(e => e.Name.LocalName is "tblPr" or "trPr" or "tcPr" or "tblStylePr");
                if (after is not null) after.AddBeforeSelf(rPr);
                else style.Add(rPr);
            }
            var fonts = rPr.Element(W + "rFonts");
            if (fonts is null)
            {
                fonts = new XElement(W + "rFonts");
                // rStyle never appears in a style's rPr, so rFonts is first.
                rPr.AddFirst(fonts);
            }
            // Theme font attributes win over explicit names; drop them.
            foreach (var a in ThemeFontAttributes) fonts.Attribute(W + a)?.Remove();
            foreach (var a in new[] { "ascii", "hAnsi", "cs" }) fonts.SetAttributeValue(W + a, font);
            var id = (string?)style.Attribute(W + "styleId") ?? name;
            changed.Add(id);
            untracked.Add($"styles.xml: {id} (\"{name}\") font set to {font} (style definition; not a tracked change)");
        }
        Docx.Save(part, styles);
        List<string> notes = missing.Count == 0 ? [] : [$"no style for {string.Join(", ", missing)}; Word creates those from Normal if the TOC uses them"];
        return new TocFontResult(changed, untracked, notes);
    }
}
