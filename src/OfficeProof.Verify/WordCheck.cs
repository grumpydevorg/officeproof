using System.Globalization;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// The Reject-All invariant checked by Microsoft Word itself rather than by Clippit: Word's
/// Reject All of the redline against Word's Reject All of the original, story by story (the
/// same <see cref="StoryComparison"/> the in-process check uses) and page by page (Word's PDFs
/// through <see cref="PageDiff"/>). Each side is one Word save of its own input, so Word's save
/// normalisation (rsids, paragraph ids, run splits, added parts) is the same on both.
/// </summary>
public static class WordCheck
{
    /// <summary>How many differing pages are reported one by one before summarising.</summary>
    private const int MaxReportedPages = 5;

    /// <summary>
    /// Rules: <c>word-reject-mismatch</c> (Error, a story differs after Word's Reject All),
    /// <c>comment-changed</c> and <c>unverifiable-revision</c> (Error, as in
    /// <see cref="RevisionCheck"/>), <c>render-mismatch</c> (Error, a printed page differs),
    /// <c>word-render</c> (Info, the pages that were compared), <c>word-version</c> (Info when
    /// officeproof's revision evidence was measured on this Word's version, Warning otherwise) and
    /// <c>word-adjusted</c> (Warning, Word was given adjusted copies, so its verdict covers the
    /// adjusted copies). Differing page images are kept in <paramref name="keepDir"/> when it is
    /// given. <paramref name="options"/> declares the style changes the redline intends (as for
    /// <see cref="RevisionCheck.Run"/>). Needs <see cref="WordOracle.IsAvailable"/>.
    /// </summary>
    public static IReadOnlyList<Finding> RejectAll(string originalPath, string redlinePath, string? keepDir = null, TimeSpan? timeout = null,
        VerifyOptions? options = null)
    {
        options ??= new VerifyOptions();
        var qualifications = new List<Finding> { VersionFinding() };
        if (AdjustedFinding(originalPath, redlinePath) is { } adjusted) qualifications.Add(adjusted);
        var dir = Path.Combine(Path.GetTempPath(), "officeproof-wordcheck-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            // Word's save drops direct formatting that equals the paragraph's style, so a style
            // the redline changed alters how Word saves untouched paragraphs: a client report, 2026-10-15,
            // --toc-font put Arial into TOC1-3 and Word's Reject-All save of the redline lost the
            // (identical) direct Arial on 34 TOC lines while every page printed the same. The
            // stories are therefore compared with the original carrying the redline's added
            // styles and lists and its declared style changes; any other definition stays the
            // original's, so an undeclared change still shows. The pages use the real original.
            var aligned = Path.Combine(dir, "original-with-intended-definitions.docx");
            AlignDefinitions(originalPath, redlinePath, options.IntendedStyles, aligned);
            var original = Path.Combine(dir, "original-rejected.docx");
            var redline = Path.Combine(dir, "redline-rejected.docx");
            WordOracle.SaveView(aligned, OutlineView.RejectAll, original, timeout);
            WordOracle.SaveView(redlinePath, OutlineView.RejectAll, redline, timeout);
            var originalView = StoryView.Load(File.ReadAllBytes(original));
            var findings = StoryComparison.Compare(originalView, StoryView.Load(File.ReadAllBytes(redline), originalView))
                // Added comments are already reported once by the in-process check.
                .Where(f => f.Rule != "comment-added")
                .Select(f => f.Rule == "reject-mismatch"
                    ? f with { Rule = "word-reject-mismatch", Message = "Word's Reject All: " + f.Message }
                    : f)
                .ToList();

            var originalPdf = Path.Combine(dir, "original-rejected.pdf");
            var redlinePdf = Path.Combine(dir, "redline-rejected.pdf");
            WordOracle.ExportPdf(originalPath, OutlineView.RejectAll, originalPdf, timeout);
            WordOracle.ExportPdf(redlinePath, OutlineView.RejectAll, redlinePdf, timeout);
            findings.AddRange(Pages(PageDiff.Compare(originalPdf, redlinePdf, keepDir), keepDir));
            return [.. qualifications, .. findings];
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    /// <summary>Whether officeproof's revision evidence was measured on the installed Word's version.</summary>
    internal static Finding VersionFinding()
    {
        var (measured, installed) = WordOracle.Version();
        var evidence = $"officeproof's revision evidence was measured on {measured.Name} {measured.Version} ({measured.Measured})";
        if (installed is null)
            return new Finding("word-version", Severity.Warning, $"could not read the version of {measured.Name}; {evidence}");
        return WordOracle.IsMeasured(measured.Version, installed)
            ? new Finding("word-version", Severity.Info, $"{measured.Name} {installed}: {evidence}")
            : new Finding("word-version", Severity.Warning,
                $"{measured.Name} {installed} is not the version {evidence}: Word's own check below ran on this version, but which " +
                "revisions Word and Clippit treat alike was not re-measured on it (maintainers: run the Word tests, OFFICEPROOF_WORD_TESTS=1, " +
                "and record the version in revision-evidence.json)");
    }

    /// <summary>
    /// What <see cref="WordOracle.CopyForWord"/> changed in either document before Word saw it,
    /// and what that leaves unproven for the recipient's Word, which opens the files as they are.
    /// </summary>
    internal static Finding? AdjustedFinding(string originalPath, string redlinePath)
    {
        var original = WordOracle.CopyForWord(originalPath);
        var redline = WordOracle.CopyForWord(redlinePath);
        if (!original.Adjusted && !redline.Adjusted) return null;
        var parts = new List<string>();
        if (original.TablesGivenLook + redline.TablesGivenLook > 0)
            parts.Add($"w:tblLook 04A0 added to tables that have none (original {original.TablesGivenLook}, redline {redline.TablesGivenLook}); " +
                      "as the files are, Word styles those tables' header row and first column at random on each open, " +
                      "in the original as in the redline");
        if (original.UpdateFieldsRemoved || redline.UpdateFieldsRemoved)
            parts.Add($"w:updateFields removed ({Which(original.UpdateFieldsRemoved, redline.UpdateFieldsRemoved)}); " +
                      "as the files are, Word offers to update every field on open, and the updated results were not compared");
        return new Finding("word-adjusted", Severity.Warning,
            "Word was given adjusted copies, so its verdict covers the adjusted copies, not the files as they are: " + string.Join("; ", parts));

        static string Which(bool original, bool redline) => original && redline ? "both" : original ? "original" : "redline";
    }

    /// <summary>
    /// A copy of <paramref name="originalPath"/> carrying the styles and lists
    /// <paramref name="redlinePath"/> adds and its changes to <paramref name="intendedStyles"/>;
    /// every other style, the document defaults and every existing list stay the original's.
    /// </summary>
    internal static void AlignDefinitions(string originalPath, string redlinePath, IReadOnlySet<string> intendedStyles, string outputPath)
    {
        using var source = Docx.OpenRead(redlinePath);
        using var target = Docx.OpenCopy(originalPath, outputPath);
        var from = source.MainDocumentPart ?? throw new InvalidDataException($"{redlinePath} has no main document part");
        var to = target.MainDocumentPart ?? throw new InvalidDataException($"{originalPath} has no main document part");
        if (from.StyleDefinitionsPart is { } styles)
            Align(styles, to.StyleDefinitionsPart ?? to.AddNewPart<StyleDefinitionsPart>(), d => d.Kind == "style" && d.Id is not null && intendedStyles.Contains(d.Id));
        if (from.NumberingDefinitionsPart is { } numbering)
            Align(numbering, to.NumberingDefinitionsPart ?? to.AddNewPart<NumberingDefinitionsPart>(), _ => false);

        static void Align(OpenXmlPart redline, OpenXmlPart original, Func<PartDiff.Definition, bool> replace)
        {
            var theirs = Docx.Load(redline).Root ?? throw new InvalidDataException($"{redline.Uri} has no root element");
            var ours = Load(original) ?? new XElement(theirs.Name, theirs.Attributes().Where(a => a.IsNamespaceDeclaration));
            Docx.Save(original, new XDocument(PartDiff.Align(ours, theirs, replace)));
        }

        // A part just added has no content yet.
        static XElement? Load(OpenXmlPart part)
        {
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            return s.Length == 0 ? null : XDocument.Load(s, LoadOptions.PreserveWhitespace).Root;
        }
    }

    private static IEnumerable<Finding> Pages(PageDiffResult result, string? keepDir)
    {
        if (result.PagesA != result.PagesB)
            yield return new Finding("render-mismatch", Severity.Error,
                $"Word prints {result.PagesB} page(s) after Reject All of the redline, {result.PagesA} for the original");
        var differing = result.Differing.ToList();
        foreach (var page in differing.Take(MaxReportedPages))
            yield return new Finding("render-mismatch", Severity.Error,
                string.Create(CultureInfo.InvariantCulture,
                    $"after Reject All, Word prints page {page.Page} differently from the original ({page.ChangedPixels} pixels changed{Box(page.Box)})"),
                Location: $"page {page.Page}");
        if (differing.Count > MaxReportedPages)
            yield return new Finding("render-mismatch", Severity.Error, $"{differing.Count - MaxReportedPages} further differing page(s) not shown");
        yield return new Finding("word-render", Severity.Info,
            $"Word's PDFs after Reject All: {result.Pages.Count - differing.Count} of {result.Pages.Count} page(s) identical at {PageDiff.Dpi} dpi" +
            (differing.Count > 0 && keepDir is not null ? $"; differing pages kept in {keepDir}" : ""));

        static string Box(PixelBox? box) => box is { } b
            ? string.Create(CultureInfo.InvariantCulture, $", within x {b.Left}-{b.Right}, y {b.Top}-{b.Bottom} at {PageDiff.Dpi} dpi")
            : "";
    }
}
