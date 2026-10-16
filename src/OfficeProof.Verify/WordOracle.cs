using System.Diagnostics;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>Microsoft Word's result after updating every field and table of contents.</summary>
/// <param name="TocLines">Each table of contents' lines as Word rebuilt them, in order.</param>
/// <param name="ErrorResults">Main-story paragraphs whose text contains "Error!" after the update (broken REF/SEQ results).</param>
public sealed record WordFieldReport(IReadOnlyList<string> TocLines, IReadOnlyList<string> ErrorResults);

/// <summary>
/// Word failed or timed out. <see cref="Exception.Message"/> carries what Word reported
/// (osascript's stderr verbatim on macOS, the COM error on Windows), or what the timeout left
/// behind.
/// </summary>
public sealed class WordOracleException : Exception
{
    public WordOracleException() { }
    public WordOracleException(string message) : base(message) { }
    public WordOracleException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>
/// Drives the real Microsoft Word as the ground truth behind <see cref="Outline"/> and the
/// verify engine: Word for Mac through AppleScript (<see cref="MacWord"/>), Word for Windows
/// through COM automation (<see cref="WindowsWord"/>). Both answer the same questions the same
/// way. Every call opens one temporary copy, closes it without saving and deletes it; the input
/// is only read. Calls are serialised across processes, so only one officeproof document is
/// open in Word at a time. <see cref="Outline"/> reads only heading candidates (Clippit's
/// outline entries for the same view): Word for Mac answers about 100 paragraphs per second.
/// </summary>
public static class WordOracle
{
    private static readonly TimeSpan DefaultTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// True with Microsoft Word for Mac (/Applications/Microsoft Word.app and osascript) or
    /// Microsoft Word for Windows (the Word.Application COM class registered).
    /// </summary>
    public static bool IsAvailable => Backend() is not null;

    /// <summary>
    /// Word's own copy of <paramref name="path"/> in <paramref name="view"/>, saved as .docx to
    /// <paramref name="outputPath"/> (overwritten). Word opens a temporary copy, runs Accept All
    /// or Reject All Changes (<see cref="OutlineView.AsIs"/>: neither, a plain resave), saves it
    /// in Word's format and closes it; <paramref name="path"/> is only read.
    /// </summary>
    /// <remarks>
    /// Word's document-wide Accept/Reject All covers every story, not only the main text:
    /// headers, footers, footnotes, endnotes, comments, text boxes (ContainerProbeTests: Word
    /// for Mac 16.113, 2026-10-15; Word for Windows 16.0.20430, 2026-10-16). Word rewrites the
    /// whole package on save (rsids, paragraph ids, run splits, added parts and settings), so
    /// compare a Word save only with another Word save, never with the input bytes.
    /// </remarks>
    public static void SaveView(string path, OutlineView view, string outputPath, TimeSpan? timeout = null)
    {
        RefuseInPlace(path, outputPath);
        var (backend, input) = Require(path);
        backend.SaveView(input, view, outputPath, timeout ?? DefaultTimeout);
    }

    /// <summary>
    /// Word's PDF of <paramref name="path"/> in <paramref name="view"/>, written to
    /// <paramref name="pdfPath"/> (overwritten). The copy's comments are deleted and markup is
    /// hidden first, so the pages show the text of the view as a reader prints it; for
    /// <see cref="OutlineView.AsIs"/> that is Word's "No Markup" display (insertions in,
    /// deletions out).
    /// </summary>
    /// <remarks>
    /// Comments are deleted because their anchors move text even though Word never prints the
    /// comments: on a 42-page client report (2026-10-15) the comment reference a redline anchored after
    /// "69" shifted the rest of a justified line 0.18 pt right, a page difference with identical
    /// text. Comments are compared as text, not as pages.
    /// </remarks>
    public static void ExportPdf(string path, OutlineView view, string pdfPath, TimeSpan? timeout = null)
    {
        RefuseInPlace(path, pdfPath);
        var (backend, input) = Require(path);
        backend.ExportPdf(input, view, pdfPath, timeout ?? DefaultTimeout);
    }

    /// <summary>
    /// Word's outline of <paramref name="path"/> in <paramref name="view"/>: list string, text and
    /// outline level of each paragraph Clippit reports as a heading or numbered paragraph, with
    /// the same <see cref="OutlineEntry.Index"/> so the two can be compared entry by entry.
    /// </summary>
    public static IReadOnlyList<OutlineEntry> Outline(string path, OutlineView view, TimeSpan? timeout = null)
    {
        var (backend, input) = Require(path);
        var viewBytes = Verify.Outline.ViewBytes(File.ReadAllBytes(path), view);
        var candidates = Verify.Outline.Filter(Verify.Outline.ParagraphsOfView(viewBytes));
        if (candidates.Count == 0) return [];
        var positions = WordParagraphPositions(viewBytes);
        var byPosition = backend.Paragraphs(input, view, [.. candidates.Select(c => positions[c.Index])], timeout ?? DefaultTimeout);
        return [.. candidates.Select(c => byPosition.TryGetValue(positions[c.Index], out var w)
            ? c with { ListNumber = w.ListString, Level = w.Level, Text = CleanText(w.Text) }
            : c with { ListNumber = "", Level = null, Text = "<missing in Word>" })];
    }

    /// <summary>
    /// Opens a copy in <paramref name="view"/>, updates every field and every table of contents,
    /// and returns the rebuilt TOC lines and any "Error!" field results.
    /// </summary>
    public static WordFieldReport UpdateFields(string path, OutlineView view = OutlineView.AcceptAll, TimeSpan? timeout = null)
    {
        var (backend, input) = Require(path);
        var (toc, text) = backend.UpdateFields(input, view, timeout ?? DefaultTimeout);
        static IEnumerable<string> Lines(string s) =>
            s.Split('\r').Select(CleanText).Where(l => l.Length > 0);
        return new WordFieldReport(
            [.. Lines(toc)],
            [.. Lines(text).Where(l => l.Contains("Error!", StringComparison.Ordinal))]);
    }

    /// <summary>
    /// 1-based Word paragraph number for each <see cref="OutlineEntry.Index"/>. Word's paragraph
    /// collection also counts each table row's end-of-row mark as a paragraph (Word for Mac and
    /// Word for Windows alike).
    /// </summary>
    internal static int[] WordParagraphPositions(byte[] viewBytes)
    {
        using var stream = new MemoryStream(viewBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var body = Docx.Load(doc.MainDocumentPart!).Root?.Element(Ns.W + "body");
        var positions = new List<int>();
        var word = 0;
        if (body is not null) Walk(body);
        return [.. positions];

        void Walk(XElement element)
        {
            foreach (var child in element.Elements())
            {
                if (child.Name == Ns.W + "p")
                {
                    positions.Add(++word);
                }
                else if (child.Name == Ns.W + "txbxContent")
                {
                    // a separate story in Word
                }
                else
                {
                    Walk(child);
                    if (child.Name == Ns.W + "tr") word++;
                }
            }
        }
    }

    /// <summary>
    /// The installed Word's evidence engine (the build officeproof's revision evidence was
    /// measured on) and its version (null when it cannot be read). Needs <see cref="IsAvailable"/>.
    /// </summary>
    internal static ((string Name, string Version, string Measured) Measured, string? Installed) Version()
    {
        var backend = Backend() ?? throw new WordOracleException("Microsoft Word is not available");
        return (RevisionEvidence.Word(backend.EngineKey), backend.Version());
    }

    /// <summary>True when <paramref name="installed"/> is the measured version or a build of it (16.113.3 of 16.113).</summary>
    internal static bool IsMeasured(string measured, string installed) =>
        installed == measured || installed.StartsWith(measured + ".", StringComparison.Ordinal);

    private static IWordBackend? Backend()
    {
        if (OperatingSystem.IsMacOS() && MacWord.Installed) return MacWord.Instance;
        if (OperatingSystem.IsWindows() && WindowsWord.Installed) return WindowsWord.Instance;
        return null;
    }

    // The checks every call makes before Word sees the document, and the copy Word is given.
    private static (IWordBackend Backend, WordInput Input) Require(string path)
    {
        var backend = Backend() ?? throw new WordOracleException(
            "Microsoft Word is not available: needs Word for Mac (/Applications/Microsoft Word.app and /usr/bin/osascript) or Word for Windows (the Word.Application COM class)");
        if (!File.Exists(path)) throw new FileNotFoundException("document not found", path);
        return (backend, new WordInput(CopyForWord(path).Bytes, Path.GetExtension(path)));
    }

    private static void RefuseInPlace(string path, string outputPath)
    {
        if (Path.GetFullPath(path) == Path.GetFullPath(outputPath))
            throw new ArgumentException("output must differ from input; officeproof never edits in place", nameof(outputPath));
    }

    /// <summary>
    /// The bytes Word is given for <paramref name="path"/>: the document itself, refused when
    /// Word would block on it, without <c>w:updateFields</c>, and with an explicit
    /// <c>w:tblLook</c> on every table that has none; with what was adjusted, which
    /// <see cref="WordCheck"/> reports, since Word's verdict then covers the adjusted copy.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Word for Mac 16.113 answers custom XML markup (w:customXml, unsupported since Word 2010)
    /// with a modal dialog that "display alerts" does not suppress; it blocks every later
    /// AppleScript call, even "close" and "quit", until someone answers it (2026-10-15,
    /// ContainerProbeTests: with the screen locked Word had to be killed). Such a document is
    /// refused before Word sees it, on Windows too, so both platforms answer the same documents.
    /// </para>
    /// <para>
    /// <c>w:updateFields</c> (which <c>generate</c> and <c>retemplate</c> set) makes Word ask on
    /// open whether to update the fields. Word for Mac 16.113 asks in a modal dialog that
    /// "display alerts" does not suppress and that may open off screen; <c>verify --word</c> on
    /// generate's sample waited out its 5-minute timeout, and the copy it left open broke every
    /// later open until Word was restarted (2026-10-16). The copy goes without it on both
    /// platforms, so Word shows the stored field results; <see cref="UpdateFields"/> updates
    /// them explicitly.
    /// </para>
    /// <para>
    /// Microsoft documents that Word reads an omitted w:tblLook as 0x04A0 ([MS-OE376] 2.1.160), but Word
    /// for Mac 16.113 and Word for Windows 16.0.20430 read such a table as 0x0000 on some opens
    /// and 0x04A0 on others (measured 2026-10-16: one in three to one in two opens of the same
    /// file, within one Word instance too; visible as Table.ApplyStyleHeadingRows right after
    /// opening and as the tblLook Word saves). Two Word saves or PDFs of one such document then
    /// differ. The copy carries the documented 0x04A0, written as Word writes it.
    /// </para>
    /// </remarks>
    internal static WordCopy CopyForWord(string path)
    {
        var bytes = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(bytes);
        var updateFieldsRemoved = false;
        var tablesGivenLook = 0;
        using (var doc = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var main = doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part");
            if (main.DocumentSettingsPart is { } settingsPart)
            {
                var settings = Docx.Load(settingsPart);
                var update = settings.Root?.Elements(Ns.W + "updateFields").ToList() ?? [];
                if (update.Count > 0)
                {
                    update.Remove();
                    Docx.Save(settingsPart, settings);
                    updateFieldsRemoved = true;
                }
            }
            foreach (var part in Docx.StoryParts(main))
            {
                var xml = Docx.Load(part);
                if (xml.Descendants(Ns.W + "customXml").Any())
                    throw new WordOracleException($"{part.Uri} holds custom XML markup (w:customXml), which Word answers with a modal dialog that blocks automation; Word cannot be asked about this document");
                var missing = xml.Descendants(Ns.W + "tblPr").Where(p => p.Parent?.Name == Ns.W + "tbl" && p.Element(Ns.W + "tblLook") is null).ToList();
                foreach (var tblPr in missing)
                {
                    var look = new XElement(Ns.W + "tblLook",
                        new XAttribute(Ns.W + "val", "04A0"), new XAttribute(Ns.W + "firstRow", "1"), new XAttribute(Ns.W + "lastRow", "0"),
                        new XAttribute(Ns.W + "firstColumn", "1"), new XAttribute(Ns.W + "lastColumn", "0"),
                        new XAttribute(Ns.W + "noHBand", "0"), new XAttribute(Ns.W + "noVBand", "1"));
                    // CT_TblPr sequence: tblLook comes before tblCaption, tblDescription and tblPrChange.
                    var after = tblPr.Elements().FirstOrDefault(e => e.Name == Ns.W + "tblCaption" || e.Name == Ns.W + "tblDescription" || e.Name == Ns.W + "tblPrChange");
                    if (after is null) tblPr.Add(look);
                    else after.AddBeforeSelf(look);
                }
                if (missing.Count == 0) continue;
                Docx.Save(part, xml);
                tablesGivenLook += missing.Count;
            }
        }
        return updateFieldsRemoved || tablesGivenLook > 0
            ? new WordCopy(stream.ToArray(), updateFieldsRemoved, tablesGivenLook)
            : new WordCopy(bytes, false, 0);
    }

    /// <summary>
    /// One officeproof document in Word at a time, across threads and processes: an exclusive
    /// lock file in <paramref name="workRoot"/> (.NET maps FileShare.None to flock on Unix and
    /// to a sharing-mode lock on Windows).
    /// </summary>
    internal static FileStream AcquireWord(string workRoot, TimeSpan timeout)
    {
        Directory.CreateDirectory(workRoot);
        var lockPath = Path.Combine(workRoot, "officeproof-word.lock");
        var clock = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) when (clock.Elapsed < timeout)
            {
                Thread.Sleep(200);
            }
            catch (IOException e)
            {
                throw new WordOracleException($"another officeproof call held Microsoft Word for longer than {timeout.TotalSeconds:0} s ({lockPath})", e);
            }
        }
    }

    // Word's paragraph content ends with the paragraph mark (\r) and, in tables, the cell mark (\a);
    // field braces show as \u0013-\u0015 only with field codes visible.
    private static string CleanText(string text) => text.TrimEnd('\r', '\n', '\a');
}

/// <summary>One paragraph as Word reads it: list string, outline level (0-8, null for body text) and text.</summary>
internal readonly record struct WordParagraph(string ListString, int? Level, string Text);

/// <summary>The document Word is given (<see cref="WordOracle.CopyForWord"/>) and its file extension (.docx, .dotx, ...).</summary>
internal sealed record WordInput(byte[] Bytes, string Extension);

/// <summary>
/// <see cref="WordOracle.CopyForWord"/>'s result: the bytes Word is given, and what differs from
/// the file: <c>w:updateFields</c> removed, <c>w:tblLook</c> added to that many tables.
/// </summary>
internal sealed record WordCopy(byte[] Bytes, bool UpdateFieldsRemoved, int TablesGivenLook)
{
    public bool Adjusted => UpdateFieldsRemoved || TablesGivenLook > 0;
}

/// <summary>
/// One platform's Word. Each call writes <c>input</c> to a fresh temporary copy, opens it,
/// applies <c>view</c> (Accept All, Reject All or neither) and closes the copy without saving,
/// also on error and timeout.
/// </summary>
internal interface IWordBackend
{
    /// <summary>This Word's key in <see cref="RevisionEvidence.Engines"/>.</summary>
    string EngineKey { get; }

    /// <summary>The installed Word's version, read without starting Word; null when it cannot be read.</summary>
    string? Version();

    void SaveView(WordInput input, OutlineView view, string outputPath, TimeSpan timeout);

    void ExportPdf(WordInput input, OutlineView view, string pdfPath, TimeSpan timeout);

    /// <summary>The paragraphs at the given 1-based positions of the main story; positions past the end are left out.</summary>
    IReadOnlyDictionary<int, WordParagraph> Paragraphs(WordInput input, OutlineView view, IReadOnlyList<int> positions, TimeSpan timeout);

    /// <summary>After updating every field and TOC: the TOCs' text, then the main story's text.</summary>
    (string Toc, string Text) UpdateFields(WordInput input, OutlineView view, TimeSpan timeout);
}
