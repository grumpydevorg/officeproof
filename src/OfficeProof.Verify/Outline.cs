using System.IO.Compression;
using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>Which state of a tracked-change document to read.</summary>
public enum OutlineView
{
    /// <summary>The document as stored, revisions shown (insertions in, deletions out).</summary>
    AsIs,
    /// <summary>After Word's Accept All Changes.</summary>
    AcceptAll,
    /// <summary>After Word's Reject All Changes.</summary>
    RejectAll,
}

/// <summary>
/// One main-story paragraph as a reader sees it in a given <see cref="OutlineView"/>.
/// <paramref name="Index"/> is the 0-based position among main-story paragraphs (table-cell
/// paragraphs included, text boxes excluded); <paramref name="Level"/> is the effective
/// 0-based outline level (null for body text); <paramref name="ListNumber"/> is the rendered
/// list string ("2.1", "•"), empty when the paragraph is not numbered; <paramref name="Text"/>
/// is the visible text without the list number.
/// </summary>
public sealed record OutlineEntry(int Index, string? StyleId, int? Level, string ListNumber, string Text);

/// <summary>
/// Reads the heading outline of a document after Accept All / Reject All, using Clippit's
/// revision processor and list-number retriever as an in-process stand-in for Word
/// (verified identical to Word on the v3 redline of a 42-page client report, 2026-10-15).
/// </summary>
public static class Outline
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>
    /// Headings (paragraphs with an effective outline level) and numbered paragraphs, in
    /// document order. The input is never modified; all processing happens on an in-memory copy.
    /// </summary>
    public static IReadOnlyList<OutlineEntry> Read(WordprocessingDocument document, OutlineView view = OutlineView.AsIs) =>
        Filter(Paragraphs(PackageBytes(document), view));

    /// <inheritdoc cref="Read(WordprocessingDocument, OutlineView)"/>
    public static IReadOnlyList<OutlineEntry> Read(string path, OutlineView view = OutlineView.AsIs) =>
        Filter(Paragraphs(File.ReadAllBytes(path), view));

    internal static IReadOnlyList<OutlineEntry> Filter(IEnumerable<OutlineEntry> paragraphs) =>
        [.. paragraphs.Where(IsOutlineEntry)];

    internal static bool IsOutlineEntry(OutlineEntry e) => e.Level is not null || e.ListNumber.Length > 0;

    /// <summary>Every main-story paragraph of <paramref name="packageBytes"/> in <paramref name="view"/>.</summary>
    internal static IReadOnlyList<OutlineEntry> Paragraphs(byte[] packageBytes, OutlineView view) =>
        ParagraphsOfView(ViewBytes(packageBytes, view));

    /// <summary>Every main-story paragraph of a package already in the wanted state.</summary>
    internal static IReadOnlyList<OutlineEntry> ParagraphsOfView(byte[] viewBytes)
    {
        using var stream = new MemoryStream(viewBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var main = doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part");
        // ListItemRetriever annotates the XDocument Clippit caches on the part, so the
        // paragraphs must come from GetXDocument(), not from a separately loaded copy.
        var xdoc = main.GetXDocument();
        var styles = StyleChain.Load(doc);
        var result = new List<OutlineEntry>();
        var index = 0;
        foreach (var p in MainStoryParagraphs(xdoc.Root!))
        {
            var number = ListItemRetriever.RetrieveListItem(doc, p) ?? "";
            result.Add(new OutlineEntry(index++, Wml.StyleId(p), styles.OutlineLevel(p), number, Wml.VisibleText(p)));
        }
        return result;
    }

    /// <summary>Body paragraphs in document order, excluding text-box content (a separate story in Word).</summary>
    internal static IEnumerable<XElement> MainStoryParagraphs(XElement documentRoot) =>
        documentRoot.Element(W + "body") is { } body
            ? body.Descendants(W + "p").Where(p => !p.Ancestors(W + "txbxContent").Any())
            : [];

    /// <summary>
    /// The package bytes after applying <paramref name="view"/> with Clippit's RevisionProcessor,
    /// corrected where it is known to differ from Word (<see cref="ClippitWorkarounds"/>):
    /// hyperlinks and simple fields come back with their identity, section header references
    /// survive a rejected section change, and the reject view keeps the redline's styles and
    /// numbering.
    /// </summary>
    internal static byte[] ViewBytes(byte[] packageBytes, OutlineView view)
    {
        if (view == OutlineView.AsIs) return packageBytes;
        if (view is not (OutlineView.AcceptAll or OutlineView.RejectAll)) throw new ArgumentOutOfRangeException(nameof(view));
        var (marked, wrappers) = ClippitWorkarounds.Prepare(packageBytes);
        var document = new WmlDocument("in.docx", marked);
        var processed = view == OutlineView.AcceptAll
            ? RevisionProcessor.AcceptRevisions(document).DocumentByteArray
            : RevisionProcessor.RejectRevisions(document).DocumentByteArray;
        if (view == OutlineView.AcceptAll && wrappers.Count == 0) return processed;
        using var output = new MemoryStream();
        output.Write(processed);
        using (var doc = WordprocessingDocument.Open(output, isEditable: true))
        {
            ClippitWorkarounds.Restore(doc, wrappers);
            if (view == OutlineView.RejectAll) RestoreDefinitionParts(packageBytes, doc);
        }
        return output.ToArray();
    }

    // Word's Reject All leaves tracked changes inside style and numbering definitions in place
    // (2026-10-15, v2 client redline: a w:pPrChange on Head1 survived Reject All and Word showed
    // "2.0 1.0 Summary"). Clippit may undo them, so the reject view puts the redline's own
    // styles.xml and numbering.xml back to show what Word shows.
    private static void RestoreDefinitionParts(byte[] source, WordprocessingDocument rejected)
    {
        using var src = new MemoryStream(source);
        using var srcZip = new ZipArchive(src, ZipArchiveMode.Read);
        var main = rejected.MainDocumentPart!;
        Restore(main.StyleDefinitionsPart, srcZip);
        Restore(main.NumberingDefinitionsPart, srcZip);

        static void Restore(OpenXmlPart? part, ZipArchive srcZip)
        {
            if (part is null) return;
            var entry = srcZip.GetEntry(part.Uri.OriginalString.TrimStart('/'));
            if (entry is null) return;
            using var s = entry.Open();
            using var buffer = new MemoryStream();
            s.CopyTo(buffer);
            Docx.Write(part, buffer.ToArray());
        }
    }

    internal static byte[] PackageBytes(WordprocessingDocument document)
    {
        using var stream = new MemoryStream();
        document.Clone(stream).Dispose();
        return stream.ToArray();
    }
}

