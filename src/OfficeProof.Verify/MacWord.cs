using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace OfficeProof.Verify;

/// <summary>
/// Microsoft Word for Mac through AppleScript (osascript): each call writes one script against
/// one copy and runs it in a child osascript process, which is killed at the timeout.
/// </summary>
/// <remarks>
/// Word is sandboxed: it may write only inside its container unless the user has granted a
/// folder. A "save as" anywhere else raises the "Grant File Access" dialog, which blocks
/// AppleScript until someone answers it; on 2026-10-15, with the screen locked, Word then
/// refused both "close" and "quit" and had to be killed. So every copy Word opens or writes
/// lives under Word's container (<c>~/Library/Containers/com.microsoft.Word/Data/tmp</c>) and
/// results are copied out from there. Word for Mac stays running between calls; it is the
/// user's own application instance.
/// </remarks>
internal sealed class MacWord : IWordBackend
{
    private const string WordApp = "/Applications/Microsoft Word.app";
    private const string Osascript = "/usr/bin/osascript";
    private static readonly TimeSpan CleanupTimeout = TimeSpan.FromSeconds(20);

    public static readonly MacWord Instance = new();

    private MacWord() { }

    public static bool Installed => Directory.Exists(WordApp) && File.Exists(Osascript);

    public string EngineKey => "word-mac";

    /// <summary>CFBundleShortVersionString of Word's Info.plist (an XML plist), e.g. 16.113.3.</summary>
    public string? Version()
    {
        try
        {
            var dict = System.Xml.Linq.XDocument.Load(Path.Combine(WordApp, "Contents", "Info.plist")).Root?.Element("dict");
            var key = dict?.Elements("key").FirstOrDefault(k => k.Value == "CFBundleShortVersionString");
            return (key?.ElementsAfterSelf().FirstOrDefault() is { Name.LocalName: "string" } value) ? value.Value : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or System.Xml.XmlException)
        {
            return null;
        }
    }

    private static string WorkRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Containers", "com.microsoft.Word", "Data", "tmp");

    public void SaveView(WordInput input, OutlineView view, string outputPath, TimeSpan timeout) =>
        // "save as" renames the open document, so it is addressed by its new name afterwards.
        Run(input, Body(view, """
                    save as d file name "@@OUTPATH@@" file format format document
                    set d to document "@@OUTNAME@@"
            """), timeout, outputPath);

    public void ExportPdf(WordInput input, OutlineView view, string pdfPath, TimeSpan timeout) =>
        Run(input, Body(view, """
                    set track revisions of d to false
                    if (count of Word comments of d) > 0 then delete all comments d
                    set v to view of active window of d
                    set show revisions and comments of v to false
                    set revisions view of v to revisions view final
                    save as d file name "@@OUTPATH@@" file format format PDF
            """), timeout, pdfPath);

    public IReadOnlyDictionary<int, WordParagraph> Paragraphs(WordInput input, OutlineView view, IReadOnlyList<int> positions, TimeSpan timeout)
    {
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
            """.Replace("{INDICES}", "{" + string.Join(", ", positions) + "}", StringComparison.Ordinal)));
        script.AppendLine("return out");

        var byPosition = new Dictionary<int, WordParagraph>();
        foreach (var record in Run(input, script.ToString(), timeout).Split('\u001e', StringSplitOptions.RemoveEmptyEntries))
        {
            var f = record.Split('\u001f');
            if (f.Length < 4) continue;
            var level = int.TryParse(f[2], CultureInfo.InvariantCulture, out var l) ? l : (int?)null;
            byPosition[int.Parse(f[0], CultureInfo.InvariantCulture)] = new WordParagraph(f[1], level, string.Join('\u001f', f[3..]));
        }
        return byPosition;
    }

    public (string Toc, string Text) UpdateFields(WordInput input, OutlineView view, TimeSpan timeout)
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

        var output = Run(input, script.ToString(), timeout);
        var split = output.IndexOf('\u001e', StringComparison.Ordinal);
        return split < 0 ? ("", output) : (output[..split], output[(split + 1)..]);
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
    // can answer itself while the script runs; it does not stop every dialog (see
    // WordOracle.RefuseBlocking).
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

    // Runs one script against a fresh copy of input inside Word's container. When outputPath is
    // given, the script saves to @@OUTPATH@@ and that file is copied to outputPath.
    private static string Run(WordInput input, string script, TimeSpan limit, string? outputPath = null)
    {
        var clock = Stopwatch.StartNew();
        using var gate = WordOracle.AcquireWord(WorkRoot, limit);

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
        var docName = id + input.Extension;
        var outName = outputPath is null ? "" : id + "-out" + Path.GetExtension(outputPath);
        try
        {
            var copy = Path.Combine(dir, docName);
            var output = Path.Combine(dir, outName);
            File.WriteAllBytes(copy, input.Bytes);
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

    private static string Quote(string s) =>
        s.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
}
