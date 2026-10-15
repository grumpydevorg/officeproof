using System.Xml.Linq;

namespace OfficeProof.Core;

/// <summary>XML namespaces used across WordprocessingML parts.</summary>
public static class Ns
{
    public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    public static readonly XNamespace Xml = XNamespace.Xml;
}

/// <summary>Read-only helpers over the WordprocessingML run model.</summary>
public static class Wml
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>The paragraph's style id (w:pPr/w:pStyle/@w:val), or null.</summary>
    public static string? StyleId(XElement paragraph) =>
        (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

    /// <summary>
    /// The text a reader sees with all revisions shown as Word displays them before any
    /// accept or reject: inserted text included, deleted and moved-from text excluded,
    /// field results included, field instructions excluded. Tabs become '\t'.
    /// </summary>
    public static string VisibleText(XElement paragraph) =>
        string.Concat(VisibleTextNodes(paragraph).Select(t => t.Name == W + "tab" ? "\t" : t.Value));

    /// <summary>
    /// The w:t (and w:tab) elements behind <see cref="VisibleText"/>, in document order.
    /// Offsets into VisibleText map onto these nodes in sequence (w:tab counts as one char).
    /// </summary>
    public static IEnumerable<XElement> VisibleTextNodes(XElement paragraph) =>
        paragraph.Descendants()
            .Where(e => e.Name == W + "t" || (e.Name == W + "tab" && e.Parent?.Name == W + "r"))
            .Where(e => !e.Ancestors().Any(a => a.Name == W + "del" || a.Name == W + "moveFrom"))
            .Where(e => e.Ancestors(W + "p").First() == paragraph);

    /// <summary>True when the paragraph contains a complex or simple field whose instruction starts with <paramref name="fieldType"/> (e.g. "SEQ", "REF", "TOC").</summary>
    public static bool HasField(XElement paragraph, string fieldType) =>
        FieldInstructions(paragraph).Any(i => i.StartsWith(fieldType + " ", StringComparison.OrdinalIgnoreCase)
                                            || i.Equals(fieldType, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Trimmed field instructions in the paragraph: w:fldSimple/@w:instr and complex-field
    /// w:instrText runs. Nested fields (a REF inside an IF) each yield their own instruction;
    /// the outer instruction does not include the inner one.
    /// </summary>
    public static IEnumerable<string> FieldInstructions(XElement paragraph)
    {
        foreach (var s in paragraph.Descendants(W + "fldSimple"))
            yield return ((string?)s.Attribute(W + "instr") ?? "").Trim();
        // Open fields, innermost last. A field stays open until its "end"; its instruction
        // is complete at "separate" (or at "end" when it has no result).
        var open = new Stack<(System.Text.StringBuilder Instr, bool InResult)>();
        foreach (var e in paragraph.Descendants())
        {
            if (e.Name == W + "fldChar")
            {
                var kind = (string?)e.Attribute(W + "fldCharType");
                if (kind == "begin") open.Push((new System.Text.StringBuilder(), false));
                else if (kind == "separate" && open.Count > 0 && !open.Peek().InResult)
                {
                    var top = open.Pop();
                    open.Push((top.Instr, true));
                    yield return top.Instr.ToString().Trim();
                }
                else if (kind == "end" && open.Count > 0)
                {
                    var top = open.Pop();
                    if (!top.InResult) yield return top.Instr.ToString().Trim();
                }
            }
            else if (e.Name == W + "instrText" && open.Count > 0 && !open.Peek().InResult) open.Peek().Instr.Append(e.Value);
        }
    }

    /// <summary>True when any tracked-change markup is present under <paramref name="root"/>.</summary>
    public static bool HasRevisions(XElement root) => Revisions(root).Any();

    /// <summary>Tracked-change elements: insertions, deletions, moves and property changes.</summary>
    public static IEnumerable<XElement> Revisions(XElement root) =>
        root.Descendants().Where(e => e.Name.Namespace == W && RevisionNames.Contains(e.Name.LocalName));

    private static readonly HashSet<string> RevisionNames =
    [
        "ins", "del", "moveFrom", "moveTo",
        "pPrChange", "rPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
        "tblGridChange", "numberingChange", "tblPrExChange",
    ];
}
