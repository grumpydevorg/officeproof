using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OfficeProof.Verify;

/// <summary>
/// Microsoft Word for Windows through COM automation, in process: late-bound
/// (<c>dynamic</c>) calls on <c>Word.Application</c>, so no Office interop assembly is
/// referenced and any installed Word version answers.
/// </summary>
/// <remarks>
/// <para>
/// Each call starts its own invisible Word instance on a dedicated STA thread, opens the copy,
/// and quits that instance afterwards; <c>Word.Application</c> starts a new WINWORD process for
/// every activation (measured 2026-10-16, Word 16.0.20430), so a Word session the user has open
/// is never touched. The instance is identified as the one WINWORD process that appeared during
/// activation; at the timeout that process is killed, which also ends the blocked COM call. If
/// it cannot be identified (no new process, or more than one), it is never killed, and quit
/// only when a new process appeared.
/// </para>
/// <para>
/// In-process COM rather than a PowerShell script (the AppleScript design's counterpart): a
/// PowerShell child would add a script interpreter whose execution policy Group Policy can
/// forbid, a text protocol for results and errors, and would still leave the COM-launched Word
/// running when the child is killed. In process the COM error arrives as an exception and the
/// Word process is known. Needs built-in COM interop: the CLI must not be published trimmed or
/// AOT-compiled.
/// </para>
/// </remarks>
[SupportedOSPlatform("windows")]
internal sealed class WindowsWord : IWordBackend
{
    private const string ProgId = "Word.Application";
    private const string WordProcess = "WINWORD";
    private static readonly TimeSpan ExitTimeout = TimeSpan.FromSeconds(30);

    // Word constants (Word object model enumerations).
    private const int WdAlertsNone = 0;
    private const int MsoAutomationSecurityForceDisable = 3;
    private const int WdDoNotSaveChanges = 0;
    private const int WdFormatXmlDocument = 12;
    private const int WdExportFormatPdf = 17;
    private const int WdExportOptimizeForPrint = 0;
    private const int WdExportAllDocument = 0;
    private const int WdExportDocumentContent = 0;
    private const int WdExportCreateNoBookmarks = 0;
    private const int WdRevisionsViewFinal = 0;
    private const int WdOutlineLevel9 = 9;
    private const int WdStoryTypeLast = 17;

    // Given for every open, so Word fails instead of prompting for a password: a document with a
    // write-reservation password failed with 0x800A1066 "Command failed" within seconds
    // (2026-10-16); an unprotected document ignores it. An encrypted document never reaches
    // Word: it is not an Open XML package, and WordOracle cannot read it.
    private const string NoPassword = "officeproof-no-password";

    public static readonly WindowsWord Instance = new();

    private WindowsWord() { }

    public static bool Installed => Type.GetTypeFromProgID(ProgId) is not null;

    public string EngineKey => "word-windows";

    /// <summary>
    /// The file version of the executable registered as the <c>Word.Application</c> COM server
    /// (<c>CLSID\{…}\LocalServer32</c>), e.g. 16.0.20430.20146.
    /// </summary>
    public string? Version()
    {
        if (Type.GetTypeFromProgID(ProgId) is not { } type) return null;
        using var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey($@"CLSID\{{{type.GUID}}}\LocalServer32");
        if (key?.GetValue(null) is not string command) return null;
        // "C:\...\WINWORD.EXE" /Automation, quoted or not.
        var exe = command.StartsWith('"')
            ? command[1..Math.Max(1, command.IndexOf('"', 1))]
            : command[..(command.IndexOf(".exe", StringComparison.OrdinalIgnoreCase) is var i and >= 0 ? i + 4 : command.Length)];
        return File.Exists(exe) ? FileVersionInfo.GetVersionInfo(exe).FileVersion : null;
    }

    private static string WorkRoot => Path.Combine(Path.GetTempPath(), "officeproof-word");

    public void SaveView(WordInput input, OutlineView view, string outputPath, TimeSpan timeout) =>
        Run(input, view, timeout, outputPath, (doc, output) =>
        {
            doc.SaveAs2(FileName: output, FileFormat: WdFormatXmlDocument, AddToRecentFiles: false);
            return 0;
        });

    // Exported from a fresh open of Word's save of the view: after a live Accept/Reject All, Word
    // for Windows 16.0.20430 laid the edited line out 0.12 pt off (the rest of the line after a
    // rejected replacement on generate's sample; 2026-10-16), while its save, reopened, prints
    // exactly as the original. Word for Mac prints the live view identically.
    public void ExportPdf(WordInput input, OutlineView view, string pdfPath, TimeSpan timeout) =>
        Run(input, view, timeout, pdfPath, reopenView: true, work: (doc, output) =>
        {
            doc.TrackRevisions = false;
            if ((int)doc.Comments.Count > 0) doc.DeleteAllComments();
            HideMarkup(doc);
            doc.ExportAsFixedFormat(OutputFileName: output, ExportFormat: WdExportFormatPdf, OpenAfterExport: false,
                OptimizeFor: WdExportOptimizeForPrint, Range: WdExportAllDocument, Item: WdExportDocumentContent,
                IncludeDocProps: true, KeepIRM: true, CreateBookmarks: WdExportCreateNoBookmarks, DocStructureTags: true,
                BitmapMissingFonts: true, UseISO19005_1: false);
            return 0;
        });

    public IReadOnlyDictionary<int, WordParagraph> Paragraphs(WordInput input, OutlineView view, IReadOnlyList<int> positions, TimeSpan timeout) =>
        Run(input, view, timeout, outputPath: null, (doc, _) =>
        {
            HideMarkup(doc);
            var result = new Dictionary<int, WordParagraph>();
            var paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            foreach (var k in positions)
            {
                if (k > count) continue;
                var p = paragraphs.Item(k);
                var r = p.Range;
                int level = p.OutlineLevel;
                result[k] = new WordParagraph((string?)r.ListFormat.ListString ?? "", level <= WdOutlineLevel9 ? level - 1 : null, (string?)r.Text ?? "");
            }
            return (IReadOnlyDictionary<int, WordParagraph>)result;
        });

    public (string Toc, string Text) UpdateFields(WordInput input, OutlineView view, TimeSpan timeout) =>
        Run(input, view, timeout, outputPath: null, (doc, _) =>
        {
            HideMarkup(doc);
            // Every story: the main text, headers, footers, notes, comments and text boxes;
            // StoryRanges holds the first range of each story type and NextStoryRange links the rest.
            var stories = doc.StoryRanges;
            for (var type = 1; type <= WdStoryTypeLast; type++)
            {
                dynamic? story;
                try
                {
                    story = stories.Item(type);
                }
                catch (COMException)
                {
                    continue; // the document has no story of this type
                }
                for (; story is not null; story = story.NextStoryRange)
                    story.Fields.Update();
            }
            var toc = "";
            var tocs = doc.TablesOfContents;
            for (var t = 1; t <= (int)tocs.Count; t++)
            {
                var table = tocs.Item(t);
                table.Update();
                toc += (string?)table.Range.Text ?? "";
            }
            return (toc, (string?)doc.Content.Text ?? "");
        });

    // Word's "No Markup" display: insertions shown as text, deletions hidden. Range.Text follows
    // the display, so without this an as-is read includes tracked-deleted text ("1.0 Summary"
    // for a heading whose typed number was deleted; Word 16.0.20430, 2026-10-16).
    private static void HideMarkup(dynamic doc)
    {
        var v = doc.Windows.Item(1).View;
        v.ShowRevisionsAndComments = false;
        v.RevisionsView = WdRevisionsViewFinal;
    }

    // One copy, one Word instance, one STA thread; the thread is abandoned (its Word killed) at
    // the timeout. When outputPath is given, work saves to its second argument and that file is
    // copied to outputPath.
    private static T Run<T>(WordInput input, OutlineView view, TimeSpan limit, string? outputPath, Func<dynamic, string, T> work, bool reopenView = false)
    {
        // Measured 2026-10-16 over OpenSSH (session 0): Documents.Open returned no document and
        // SaveAs2 failed with 0x800A1066 "Command failed"; the same calls work in a desktop session.
        using (var self = Process.GetCurrentProcess())
        {
            if (self.SessionId == 0)
                throw new WordOracleException("Word for Windows cannot open or save documents from session 0 (a Windows service or an SSH logon); run officeproof in a desktop session");
        }
        var clock = Stopwatch.StartNew();
        using var gate = WordOracle.AcquireWord(WorkRoot, limit);
        var id = "officeproof-" + Guid.NewGuid().ToString("N");
        var dir = Path.Combine(WorkRoot, id);
        Directory.CreateDirectory(dir);
        try
        {
            var copy = Path.Combine(dir, id + input.Extension);
            var output = outputPath is null ? "" : Path.Combine(dir, id + "-out" + Path.GetExtension(outputPath));
            // A new file, so no Zone.Identifier stream (File.Copy would copy a downloaded
            // document's mark of the web, and Word could open the copy in Protected View).
            File.WriteAllBytes(copy, input.Bytes);

            var session = new Session<T>(copy, view, output, work, reopenView);
            var thread = new Thread(session.Run) { IsBackground = true, Name = "officeproof-word" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            var remaining = limit - clock.Elapsed;
            if (!thread.Join(remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero))
            {
                var cleanup = session.Abandon();
                thread.Join(ExitTimeout);
                throw new WordOracleException($"Word did not finish within {limit.TotalSeconds:0} s (busy or showing a dialog); {cleanup}");
            }
            session.AwaitExit();
            if (session.Error is { } error)
                throw error as WordOracleException ?? new WordOracleException(Describe(error), error);
            if (outputPath is not null)
            {
                if (!File.Exists(output))
                    throw new WordOracleException($"Word reported success but wrote no {Path.GetExtension(outputPath)} file");
                File.Copy(output, outputPath, overwrite: true);
            }
            return session.Result!;
        }
        finally
        {
            try { Directory.Delete(dir, recursive: true); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* Word may still hold the copy after a timeout */ }
        }
    }

    private static string Describe(Exception e) => e is COMException com
        ? $"Word failed: {com.Message.Trim()} (HRESULT 0x{com.HResult:X8})"
        : $"Word failed: {e.Message}";

    private static HashSet<int> WordPids()
    {
        var processes = Process.GetProcessesByName(WordProcess);
        try
        {
            return [.. processes.Select(p => p.Id)];
        }
        finally
        {
            foreach (var p in processes) p.Dispose();
        }
    }

    // reopenView: after Accept/Reject All, save the view and work on a fresh open of that save.
    private sealed class Session<T>(string copy, OutlineView view, string output, Func<dynamic, string, T> work, bool reopenView)
    {
        private int _pid;
        private bool _quit;

        public T? Result { get; private set; }

        public Exception? Error { get; private set; }

        // The thread body: everything Word-side happens here, and nothing escapes it.
        public void Run()
        {
            dynamic? app = null;
            dynamic? doc = null;
            try
            {
                Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(new RetryWhenBusy(), out _));
                var before = WordPids();
                app = Activator.CreateInstance(Type.GetTypeFromProgID(ProgId, throwOnError: true)!)!;
                var started = WordPids();
                started.ExceptWith(before);
                _quit = started.Count > 0;
                if (started.Count == 1) Volatile.Write(ref _pid, started.Single());
                app.DisplayAlerts = WdAlertsNone;
                app.AutomationSecurity = MsoAutomationSecurityForceDisable;
                doc = Open(app, copy);
                if (view == OutlineView.AcceptAll) doc.AcceptAllRevisions();
                else if (view == OutlineView.RejectAll) doc.RejectAllRevisions();
                if (reopenView && view != OutlineView.AsIs)
                {
                    var saved = Path.Combine(Path.GetDirectoryName(copy)!, Path.GetFileNameWithoutExtension(copy) + "-view.docx");
                    doc.SaveAs2(FileName: saved, FileFormat: WdFormatXmlDocument, AddToRecentFiles: false);
                    doc.Close(SaveChanges: WdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(doc);
                    doc = null;
                    doc = Open(app, saved);
                }
                Result = work(doc, output);
            }
            // Rethrown on the caller's thread: an exception escaping this thread would end the process.
            catch (Exception e)
            {
                Error = e;
            }
            finally
            {
                Close(doc, app);
                _ = CoRegisterMessageFilter(null, out _); // revoking only fails when none is registered
            }
        }

        // Visible: the document gets a window (hidden with the invisible application); Accept/Reject
        // All and the view settings need one (0x800A11FD without).
        private static dynamic Open(dynamic app, string path) =>
            app.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false,
                PasswordDocument: NoPassword, WritePasswordDocument: NoPassword, Revert: false, Visible: true,
                OpenAndRepair: false, NoEncodingDialog: true)
            ?? throw new WordOracleException("Word opened the copy but returned no document");

        private void Close(dynamic? doc, dynamic? app)
        {
            try
            {
                doc?.Close(SaveChanges: WdDoNotSaveChanges);
            }
            catch (COMException) { /* the copy is deleted anyway, and the instance quits */ }
            if (app is not null && _quit)
            {
                try
                {
                    app.NormalTemplate.Saved = true;
                    app.Quit(SaveChanges: WdDoNotSaveChanges);
                }
                catch (COMException) { /* AwaitExit kills an identified instance that stays */ }
            }
            if (doc is not null) Marshal.FinalReleaseComObject(doc);
            if (app is not null) Marshal.FinalReleaseComObject(app);
            // The intermediate objects (collections, ranges) hold the instance alive until released.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // After a normal finish: the instance officeproof started must be gone before the next
        // call starts one, and one that ignores Quit is ended.
        public void AwaitExit()
        {
            var pid = Volatile.Read(ref _pid);
            if (pid == 0 || !_quit) return;
            using var process = Find(pid);
            if (process is null || process.WaitForExit(ExitTimeout)) return;
            process.Kill();
            process.WaitForExit(ExitTimeout);
        }

        // At the timeout: kill the instance officeproof started, which ends the blocked call.
        public string Abandon()
        {
            var pid = Volatile.Read(ref _pid);
            if (pid == 0)
                return "the Word instance could not be identified, so it was left running; end its WINWORD process in Task Manager";
            using var process = Find(pid);
            if (process is null) return "its Word instance had already exited";
            process.Kill();
            return process.WaitForExit(ExitTimeout)
                ? $"killed the Word instance officeproof started (WINWORD pid {pid})"
                : $"killing the Word instance officeproof started (WINWORD pid {pid}) did not end it within {ExitTimeout.TotalSeconds:0} s";
        }

        private static Process? Find(int pid)
        {
            try
            {
                return Process.GetProcessById(pid);
            }
            catch (ArgumentException)
            {
                return null; // exited
            }
        }
    }

    // Word rejects calls while it is busy (starting, opening, laying out) with
    // RPC_E_CALL_REJECTED; a COM message filter retries them until the call's timeout ends the
    // session instead of failing the call.
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

        [PreserveSig] int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

        [PreserveSig] int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
    }

    private sealed class RetryWhenBusy : IOleMessageFilter
    {
        private const int ServerCallIsHandled = 0;
        private const int ServerCallRetryLater = 2;
        private const int RetryAfterMilliseconds = 100;
        private const int CancelCall = -1;
        private const int PendingMessageWaitDefaultProcess = 2;

        public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo) => ServerCallIsHandled;

        public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType) =>
            rejectType == ServerCallRetryLater ? RetryAfterMilliseconds : CancelCall;

        public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType) => PendingMessageWaitDefaultProcess;
    }

    [DllImport("ole32.dll")]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}
