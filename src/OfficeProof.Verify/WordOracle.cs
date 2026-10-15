using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>Microsoft Word's result after updating every field and table of contents.</summary>
/// <param name="TocLines">Each table of contents' lines as Word rebuilt them, in order.</param>
/// <param name="ErrorResults">Main-story paragraphs whose text contains "Error!" after the update (broken REF/SEQ results).</param>
public sealed record WordFieldReport(IReadOnlyList<string> TocLines, IReadOnlyList<string> ErrorResults);

/// <summary>
/// osascript or Word failed or timed out; <see cref="Exception.Message"/> carries osascript's
/// stderr verbatim, or what the timeout left behind.
/// </summary>
public sealed class WordOracleException : Exception
{
    public WordOracleException() { }
    public WordOracleException(string message) : base(message) { }
    public WordOracleException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>
/// Drives the real Microsoft Word for Mac through AppleScript (osascript) as the ground truth
/// behind <see cref="Outline"/> and the verify engine. Every call opens one temporary copy,
/// closes it without saving and deletes it; calls are serialised across processes, so only one
/// officeproof document is open in Word at a time. Word answers about 100 paragraphs per second
/// over AppleScript, so <see cref="Outline"/> reads only heading candidates (Clippit's outline
/// entries for the same view).
/// </summary>
/// <remarks>
/// Word is sandboxed: it may write only inside its container unless the user has granted a
/// folder. A "save as" anywhere else raises the "Grant File Access" dialog, which blocks
/// AppleScript until someone answers it; on 2026-10-15, with the screen locked, Word then
/// refused both "close" and "quit" and had to be killed. So every copy Word opens or writes
/// lives under Word's container (<c>~/Library/Containers/com.microsoft.Word/Data/tmp</c>) and
/// results are copied out from there.
/// </remarks>
public static class WordOracle
{
    private const string WordApp = "/Applications/Microsoft Word.app";
    private const string Osascript = "/usr/bin/osascript";
    private static readonly TimeSpan DefaultTimeout = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan CleanupTimeout = TimeSpan.FromSeconds(20);

    private static string WorkRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Containers", "com.microsoft.Word", "Data", "tmp");

    /// <summary>True on macOS with Microsoft Word installed in /Applications and osascript present.</summary>
    public static bool IsAvailable =>
        OperatingSystem.IsMacOS() && Directory.Exists(WordApp) && File.Exists(Osascript);

    /// <summary>
    /// Word's own copy of <paramref name="path"/> in <paramref name="view"/>, saved as .docx to
    /// <paramref name="outputPath"/> (overwritten). Word opens a temporary copy, runs Accept All
    /// or Reject All Changes (<see cref="OutlineView.AsIs"/>: neither, a plain resave), saves it
    /// in Word's format and closes it; <paramref name="path"/> is only read.
    /// </summary>
    /// <remarks>
    /// Word's document-wide Accept/Reject All covers every story, not only the main text:
    /// headers, footers, footnotes, endnotes, comments, text boxes (ContainerProbeTests,
    /// Word 16.113, 2026-10-15). Word rewrites the whole package on save (rsids, paragraph ids,
    /// run splits, added parts and settings), so compare a Word save only with another Word
    /// save, never with the input bytes.
    /// </remarks>
    public static void SaveView(string path, OutlineView view, string outputPath, TimeSpan? timeout = null)
    {
        RefuseInPlace(path, outputPath);
        // "save as" renames the open document, so it is addressed by its new name afterwards.
        Run(path, Body(view, """
                    save as d file name "@@OUTPATH@@" file format format document
                    set d to document "@@OUTNAME@@"
            """), timeout, outputPath);
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
        Run(path, Body(view, """
                    set track revisions of d to false
                    if (count of Word comments of d) > 0 then delete all comments d
                    set v to view of active window of d
                    set show revisions and comments of v to false
                    set revisions view of v to revisions view final
                    save as d file name "@@OUTPATH@@" file format format PDF
            """), timeout, pdfPath);
    }

    /// <summary>
    /// Word's outline of <paramref name="path"/> in <paramref name="view"/>: list string, text and
    /// outline level of each paragraph Clippit reports as a heading or numbered paragraph, with
    /// the same <see cref="OutlineEntry.Index"/> so the two can be compared entry by entry.
    /// </summary>
    public static IReadOnlyList<OutlineEntry> Outline(string path, OutlineView view, TimeSpan? timeout = null)
    {
        var viewBytes = Verify.Outline.ViewBytes(File.ReadAllBytes(path), view);
        var candidates = Verify.Outline.Filter(Verify.Outline.ParagraphsOfView(viewBytes));
        if (candidates.Count == 0) return [];
        var positions = WordParagraphPositions(viewBytes);

        var script = new StringBuilder();
        script.AppendLine(OutlineLevelHandler);
        // Variable names must not collide with Word's dictionary: inside its tell block "us"
        // resolves to a Word term and prints as the literal "us".
        script.AppendLine("set fieldSep to character id 31");
        script.AppendLine("set recordSep to character id 30");
        script.AppendLine("set out to \"\"");
        script.AppendLine(Body(view, """
                    set n to count of paragraphs of d
                    repeat with i in {INDICES}
                        set k to contents of i
                        if k <= n then
                            set p to paragraph k of d
                            set r to text object of p
                            set out to out & k & fieldSep & (list string of list format of r) & fieldSep & my levelOf(outline level of p) & fieldSep & (content of r) & recordSep
                        end if
                    end repeat
            """.Replace("{INDICES}", "{" + string.Join(", ", candidates.Select(c => positions[c.Index])) + "}", StringComparison.Ordinal)));
        script.AppendLine("return out");

        var byPosition = new Dictionary<int, (string List, int? Level, string Text)>();
        foreach (var record in Run(path, script.ToString(), timeout).Split('\u001e', StringSplitOptions.RemoveEmptyEntries))
        {
            var f = record.Split('\u001f');
            if (f.Length < 4) continue;
            var level = int.TryParse(f[2], CultureInfo.InvariantCulture, out var l) ? l : (int?)null;
            byPosition[int.Parse(f[0], CultureInfo.InvariantCulture)] = (f[1], level, CleanText(string.Join('\u001f', f[3..])));
        }
        return [.. candidates.Select(c => byPosition.TryGetValue(positions[c.Index], out var w)
            ? c with { ListNumber = w.List, Level = w.Level, Text = w.Text }
            : c with { ListNumber = "", Level = null, Text = "<missing in Word>" })];
    }

    /// <summary>
    /// Opens a copy in <paramref name="view"/>, updates every field and every table of contents,
    /// and returns the rebuilt TOC lines and any "Error!" field results.
    /// </summary>
    public static WordFieldReport UpdateFields(string path, OutlineView view = OutlineView.AcceptAll, TimeSpan? timeout = null)
    {
        var script = new StringBuilder();
        script.AppendLine("set recordSep to character id 30");
        script.AppendLine("set out to \"\"");
        script.AppendLine(Body(view, """
                    set i to 1
                    repeat
                        try
                            set f to field i of d
                        on error
                            exit repeat
                        end try
                        update field f
                        set i to i + 1
                    end repeat
                    set toc to ""
                    repeat with t from 1 to (count of tables of contents of d)
                        update (table of contents t of d)
                        set toc to toc & (content of text object of table of contents t of d)
                    end repeat
                    set out to toc & recordSep & (content of text object of d)
            """));
        script.AppendLine("return out");

        var output = Run(path, script.ToString(), timeout);
        var split = output.IndexOf('\u001e', StringComparison.Ordinal);
        var toc = split < 0 ? "" : output[..split];
        var text = split < 0 ? output : output[(split + 1)..];
        static IEnumerable<string> Lines(string s) =>
            s.Split('\r').Select(CleanText).Where(l => l.Length > 0);
        return new WordFieldReport(
            [.. Lines(toc)],
            [.. Lines(text).Where(l => l.Contains("Error!", StringComparison.Ordinal))]);
    }

    /// <summary>
    /// 1-based Word paragraph number for each <see cref="OutlineEntry.Index"/>. Word's paragraph
    /// collection also counts each table row's end-of-row mark as a paragraph.
    /// </summary>
    internal static int[] WordParagraphPositions(byte[] viewBytes)
    {
        using var stream = new MemoryStream(viewBytes);
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, isEditable: false);
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

    private const string OutlineLevelHandler = """
        on levelOf(ol)
            tell application "Microsoft Word"
                if ol is outline level1 then return "0"
                if ol is outline level2 then return "1"
                if ol is outline level3 then return "2"
                if ol is outline level4 then return "3"
                if ol is outline level5 then return "4"
                if ol is outline level6 then return "5"
                if ol is outline level7 then return "6"
                if ol is outline level8 then return "7"
                if ol is outline level9 then return "8"
            end tell
            return ""
        end levelOf
        """;

    // Opens the copy, applies the view, runs the work, and closes without saving even on error.
    // The copy is addressed by its unique file name rather than "active document", so a
    // document the user brings to the front meanwhile is never touched. The AppleScript event
    // timeout (default 2 minutes per event) is raised to the call's timeout: a Reject All or a
    // PDF export of a long report is one event. "display alerts" lets Word answer the alerts it
    // can answer itself while the script runs; it does not stop every dialog (see RefuseBlocking).
    private static string Body(OutlineView view, string work) => $$"""
        with timeout of {{TimeoutToken}} seconds
            tell application "Microsoft Word"
                set display alerts to alerts none
                open POSIX file "{{DocPathToken}}"
                try
                    set d to document "{{DocNameToken}}"
                    {{view switch
                    {
                        OutlineView.AcceptAll => "accept all revisions d",
                        OutlineView.RejectAll => "reject all revisions d",
                        _ => "",
                    }}}
            {{work}}
                    close d saving no
                on error m number k
                    try
                        close document "{{DocNameToken}}" saving no
                    end try
                    try
                        close document "{{OutNameToken}}" saving no
                    end try
                    error m number k
                end try
            end tell
        end timeout
        """;

    private const string DocPathToken = "@@DOCPATH@@";
    private const string DocNameToken = "@@DOCNAME@@";
    private const string OutPathToken = "@@OUTPATH@@";
    private const string OutNameToken = "@@OUTNAME@@";
    private const string TimeoutToken = "@@TIMEOUT@@";

    private static void RefuseInPlace(string path, string outputPath)
    {
        if (Path.GetFullPath(path) == Path.GetFullPath(outputPath))
            throw new ArgumentException("output must differ from input; officeproof never edits in place", nameof(outputPath));
    }

    // Word 16.113 answers custom XML markup (w:customXml, unsupported since Word 2010) with a
    // modal dialog that "display alerts" does not suppress; it blocks every later AppleScript
    // call, even "close" and "quit", until someone answers it (2026-10-15, ContainerProbeTests:
    // with the screen locked Word had to be killed). Such a document is refused before Word
    // sees it.
    private static void RefuseBlocking(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        foreach (var part in Docx.StoryParts(doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part")))
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "customXml" && reader.NamespaceURI == Ns.W.NamespaceName)
                    throw new WordOracleException($"{part.Uri} holds custom XML markup (w:customXml), which Word answers with a modal dialog that blocks automation; Word cannot be asked about this document");
            }
        }
    }

    // Runs one script against a fresh copy of path inside Word's container. When outputPath is
    // given, the script saves to @@OUTPATH@@ and that file is copied to outputPath.
    private static string Run(string path, string script, TimeSpan? timeout, string? outputPath = null)
    {
        if (!IsAvailable) throw new WordOracleException("Microsoft Word for Mac is not available (macOS, /Applications/Microsoft Word.app, /usr/bin/osascript)");
        if (!File.Exists(path)) throw new FileNotFoundException("document not found", path);
        RefuseBlocking(path);
        var limit = timeout ?? DefaultTimeout;
        var clock = Stopwatch.StartNew();
        using var gate = AcquireWord(limit);

        var id = "officeproof-" + Guid.NewGuid().ToString("N");
        var dir = Path.Combine(WorkRoot, id);
        try
        {
            Directory.CreateDirectory(dir);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            throw new WordOracleException($"cannot create the work folder {dir} inside Word's sandbox container (does this process have access to other apps' data?): {e.Message}", e);
        }
        var docName = id + Path.GetExtension(path);
        var outName = outputPath is null ? "" : id + "-out" + Path.GetExtension(outputPath);
        try
        {
            var copy = Path.Combine(dir, docName);
            var output = Path.Combine(dir, outName);
            File.Copy(path, copy);
            var scriptPath = Path.Combine(dir, "oracle.applescript");
            File.WriteAllText(scriptPath, script
                .Replace(DocPathToken, Quote(copy), StringComparison.Ordinal)
                .Replace(DocNameToken, Quote(docName), StringComparison.Ordinal)
                .Replace(OutPathToken, Quote(output), StringComparison.Ordinal)
                .Replace(OutNameToken, Quote(outName), StringComparison.Ordinal)
                .Replace(TimeoutToken, ((int)Math.Ceiling(limit.TotalSeconds) + 30).ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal),
                new UTF8Encoding(false));

            var remaining = limit - clock.Elapsed;
            var result = Osa(scriptPath, remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero);
            if (result is null)
            {
                var cleanup = CloseLeftovers(scriptDir: dir, docName, outName);
                throw new WordOracleException($"Word did not finish within {limit.TotalSeconds:0} s (busy, showing a dialog, or the GUI session is locked); {cleanup}");
            }
            var (exitCode, stdout, stderr) = result.Value;
            if (exitCode != 0)
                throw new WordOracleException($"osascript exited {exitCode}: {stderr.Trim()}");
            if (outputPath is not null)
            {
                if (!File.Exists(output))
                    throw new WordOracleException($"Word reported success but wrote no {Path.GetExtension(outputPath)} file");
                File.Copy(output, outputPath, overwrite: true);
            }
            // osascript prints the result followed by a newline.
            return stdout.EndsWith('\n') ? stdout[..^1] : stdout;
        }
        finally
        {
            try { Directory.Delete(dir, recursive: true); }
            catch (IOException) { /* Word may still hold the copy after a timeout */ }
        }
    }

    // Exit code and output of osascript, or null when it was killed at the timeout.
    private static (int ExitCode, string Stdout, string Stderr)? Osa(string scriptPath, TimeSpan timeout)
    {
        var psi = new ProcessStartInfo(Osascript, [scriptPath])
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        using var process = Process.Start(psi) ?? throw new WordOracleException("could not start osascript");
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(timeout))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            return null;
        }
        process.WaitForExit();
        return (process.ExitCode, stdout.Result, stderr.Result);
    }

    // After a timeout the copy may still be open in Word; try to close it so the next call
    // starts clean. Fails when Word is blocked by a dialog, which the message then says.
    private static string CloseLeftovers(string scriptDir, string docName, string outName)
    {
        var script = Path.Combine(scriptDir, "cleanup.applescript");
        var names = outName.Length == 0 ? new[] { docName } : [docName, outName];
        File.WriteAllText(script, $$"""
            with timeout of {{(int)CleanupTimeout.TotalSeconds}} seconds
                tell application "Microsoft Word"
                    set n to 0
                    repeat with x in {{{string.Join(", ", names.Select(n => "\"" + Quote(n) + "\""))}}}
                        if exists document (contents of x) then
                            close document (contents of x) saving no
                            set n to n + 1
                        end if
                    end repeat
                    return n
                end tell
            end timeout
            """, new UTF8Encoding(false));
        var result = Osa(script, CleanupTimeout + TimeSpan.FromSeconds(5));
        return result switch
        {
            null => $"closing the copy also timed out: Word is blocked (a modal dialog?) and may need to be dismissed or quit by hand; open copy: {docName}",
            { ExitCode: 0 } r => $"closed {r.Stdout.Trim()} leftover copy(ies) in Word",
            var r => $"closing the copy failed ({r.Value.Stderr.Trim()}); Word may still hold {docName}",
        };
    }

    // One officeproof document in Word at a time, across threads and processes: an exclusive
    // lock file next to the copies (.NET maps FileShare.None to flock on Unix).
    private static FileStream AcquireWord(TimeSpan timeout)
    {
        Directory.CreateDirectory(WorkRoot);
        var lockPath = Path.Combine(WorkRoot, "officeproof-word.lock");
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

    private static string Quote(string s) =>
        s.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);

    // Word's paragraph content ends with the paragraph mark (\r) and, in tables, the cell mark (\a);
    // field braces show as \u0013-\u0015 only with field codes visible.
    private static string CleanText(string text) => text.TrimEnd('\r', '\n', '\a');
}
