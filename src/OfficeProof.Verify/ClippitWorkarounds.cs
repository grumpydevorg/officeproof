using System.Globalization;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// Prepares a package for Clippit 3.9.1's RevisionProcessor and repairs its output where it
/// differs from Word, each workaround proven by a test:
/// <list type="bullet">
/// <item>Wrappers. Clippit treats w:hyperlink and w:fldSimple as atomic run content
/// (RevisionProcessor's content-element list) and mishandles revisions inside them: Reject All
/// drops the w:del text and keeps the w:ins text (review of 2026-10-15, finding 3; pandoc and Word
/// reject correctly; PlantedDefectTests "fldSimple" showed the same for simple fields). Each
/// wrapper becomes a pair of bookmark markers around its content, which Clippit passes through,
/// and is rebuilt with its original attributes (r:id, w:anchor, w:docLocation, w:tgtFrame,
/// w:instr...) afterwards. A hyperlink whose content the view removed entirely (all inserted, on
/// Reject All; all deleted, on Accept All) is not rebuilt, as in Word; a simple field is, since
/// Word keeps a field with an empty result. A marker pair Clippit separated is rebuilt as an empty
/// wrapper flagged <see cref="Unrestored"/>, which the story view reports as unverifiable.</item>
/// <item>Section header references. Clippit rejects a w:sectPrChange by replacing the whole
/// w:sectPr with the recorded one, which (CT_SectPrBase) never holds header or footer
/// references, so every header and footer of the section vanished (PlantedDefectTests
/// "section-margins"). Word's Reject All restores the section properties and keeps the
/// references; the recorded w:sectPr gets the current references before Clippit runs.</item>
/// <item>Moved paragraph marks. Clippit's Reject All left the paragraph whose mark was moved in
/// (w:pPr/w:rPr/w:moveTo) as an empty paragraph (PlantedDefectTests "move"). For accept and
/// reject a moved-from mark behaves as a deleted mark and a moved-to mark as an inserted one, so
/// they are renamed w:del and w:ins before Clippit runs.</item>
/// </list>
/// </summary>
internal static class ClippitWorkarounds
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>Namespace of officeproof's own annotations in a view package.</summary>
    public static readonly XNamespace Dw = "urn:officeproof:verify";

    /// <summary>Attribute on a wrapper whose marker pair did not survive the view.</summary>
    public static readonly XName Unrestored = Dw + "unrestored";

    private static readonly XName[] Wrappers = [W + "hyperlink", W + "fldSimple"];

    private const string Prefix = "_officeproofWrapper";
    private const int FirstId = 1_900_000_000;

    /// <summary>A wrapper replaced by markers: its name and attributes.</summary>
    internal sealed record Wrapper(XName Name, XAttribute[] Attributes);

    /// <summary>Applies every workaround to every story part; returns what <see cref="Restore"/> needs.</summary>
    public static (byte[] Package, IReadOnlyList<Wrapper> Wrappers) Prepare(byte[] packageBytes)
    {
        var wrappers = new List<Wrapper>();
        var changed = false;
        using var output = new MemoryStream();
        output.Write(packageBytes);
        using (var doc = WordprocessingDocument.Open(output, isEditable: true))
        {
            foreach (var part in Docx.StoryParts(doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part")))
            {
                var xml = Docx.Load(part);
                var touched = ReplaceWrappers(xml.Root!, wrappers) | KeepSectionReferences(xml.Root!) | MovedMarksAsInsertions(xml.Root!);
                if (!touched) continue;
                Docx.Save(part, xml);
                changed = true;
            }
        }
        return (changed ? output.ToArray() : packageBytes, wrappers);
    }

    private static bool ReplaceWrappers(XElement root, List<Wrapper> wrappers)
    {
        // Innermost first, so a nested wrapper is replaced before its parent moves it.
        var found = root.Descendants().Where(e => Wrappers.Contains(e.Name)).Reverse().ToList();
        foreach (var element in found)
        {
            var n = wrappers.Count;
            wrappers.Add(new Wrapper(element.Name, [.. element.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => new XAttribute(a))]));
            var id = (FirstId + n).ToString(CultureInfo.InvariantCulture);
            element.ReplaceWith(
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", Prefix + n.ToString(CultureInfo.InvariantCulture))),
                element.Nodes(),
                new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id)));
        }
        return found.Count > 0;
    }

    private static bool KeepSectionReferences(XElement root)
    {
        var touched = false;
        foreach (var recorded in root.Descendants(W + "sectPrChange").Select(c => c.Element(W + "sectPr")).OfType<XElement>())
        {
            var current = recorded.Parent!.Parent!;
            var references = current.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference").ToList();
            if (references.Count == 0 || recorded.Elements().Any(e => e.Name == W + "headerReference" || e.Name == W + "footerReference")) continue;
            recorded.AddFirst(references.Select(r => new XElement(r)));
            touched = true;
        }
        return touched;
    }

    private static bool MovedMarksAsInsertions(XElement root)
    {
        var marks = root.Descendants(W + "pPr").Elements(W + "rPr").Elements()
            .Where(e => e.Name == W + "moveFrom" || e.Name == W + "moveTo").ToList();
        foreach (var mark in marks) mark.Name = mark.Name == W + "moveFrom" ? W + "del" : W + "ins";
        return marks.Count > 0;
    }

    /// <summary>Rebuilds the wrappers of <see cref="Prepare"/> in an open view package.</summary>
    public static void Restore(WordprocessingDocument view, IReadOnlyList<Wrapper> wrappers)
    {
        if (wrappers.Count == 0) return;
        foreach (var part in Docx.StoryParts(view.MainDocumentPart!))
        {
            var xml = Docx.Load(part);
            var starts = xml.Descendants(W + "bookmarkStart")
                .Where(b => ((string?)b.Attribute(W + "name"))?.StartsWith(Prefix, StringComparison.Ordinal) == true)
                .ToList();
            if (starts.Count == 0) continue;
            foreach (var start in starts)
            {
                var wrapper = wrappers[int.Parse(((string)start.Attribute(W + "name")!)[Prefix.Length..], CultureInfo.InvariantCulture)];
                var id = (string?)start.Attribute(W + "id");
                bool IsEnd(XElement e) => e.Name == W + "bookmarkEnd" && (string?)e.Attribute(W + "id") == id;
                var end = start.ElementsAfterSelf().FirstOrDefault(IsEnd);
                if (end is null)
                {
                    xml.Descendants().FirstOrDefault(IsEnd)?.Remove();
                    start.ReplaceWith(new XElement(wrapper.Name, wrapper.Attributes, new XAttribute(Unrestored, "true")));
                    continue;
                }
                var content = start.NodesAfterSelf().TakeWhile(node => node != end).ToList();
                foreach (var node in content) node.Remove();
                end.Remove();
                var empty = !content.OfType<XElement>().Any(e => e.DescendantsAndSelf(W + "r").Any());
                if (empty && wrapper.Name == W + "hyperlink")
                    start.ReplaceWith(content);
                else
                    start.ReplaceWith(new XElement(wrapper.Name, wrapper.Attributes, content));
            }
            Docx.Save(part, xml);
        }
    }
}
