using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;
using static OfficeProof.TestKit.DocxBuilder;
using static OfficeProof.Verify.Tests.Redlines;

namespace OfficeProof.Verify.Tests;

/// <summary>
/// <see cref="WordCheck"/> against the real Microsoft Word. Opt-in: OFFICEPROOF_WORD_TESTS=1 with
/// Word installed (Mac: an unlocked GUI session; Windows: Word for Windows) and poppler's pdftoppm on PATH.
/// </summary>
[Collection(WordTests.Name)]
public sealed class WordCheckTests : IDisposable
{
    private readonly Redlines _r = new();

    public void Dispose() => _r.Dispose();

    private static void RequireWord() =>
        Assert.SkipUnless(Environment.GetEnvironmentVariable("OFFICEPROOF_WORD_TESTS") == "1" && WordOracle.IsAvailable,
            "set OFFICEPROOF_WORD_TESTS=1 with Microsoft Word installed (Mac or Windows)");

    private static readonly TimeSpan Timeout = TimeSpan.FromMinutes(2);

    private static string[] Errors(IEnumerable<Finding> findings) =>
        [.. findings.Where(f => f.Severity == Severity.Error).Select(f => $"{f.Rule}: {f.Message}")];

    private string Original() => new DocxBuilder()
        .Header("rIdH1", P(R("Project P042 header")))
        .Heading("Head1", "Summary")
        .Raw(P(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "TOC1"))),
            new XElement(W + "r",
                new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial"))),
                new XElement(W + "t", "Summary line in Arial"))))
        .Para("Body text that the redline edits.")
        .Save(Path.Combine(_r.Dir, "word-original.docx"));

    private string RestyledTocRedline(string name, string original) => _r.Make(name, from: original,
        body: b => _r.TrackedReplacePrefix(Paragraph(b, "Body text that the redline edits."), "Body text", "Body copy"),
        styles: s => Style(s, "TOC1").Add(TocArial()));

    private static XElement TocArial() => new(W + "rPr",
        new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial")));

    private static VerifyOptions Declared(params string[] styles) => new() { IntendedStyles = new HashSet<string>(styles, StringComparer.Ordinal) };

    [Fact]
    public void Tracked_edit_with_a_declared_restyled_toc_passes_in_word()
    {
        RequireWord();
        var original = Original();
        // A tracked body edit plus the --toc-font change: Arial moved into the TOC1 style
        // (untrackable, declared), which makes Word drop the identical direct Arial when it
        // saves. Word's stories still match, and every page prints the same (a client report, 2026-10-15).
        var redline = RestyledTocRedline("word-redline", original);

        var findings = WordCheck.RejectAll(original, redline, timeout: Timeout, options: Declared("TOC1"));

        Assert.Empty(Errors(findings));
        Assert.Contains(findings, f => f.Rule == "word-render");
    }

    [Fact]
    public void Undeclared_restyled_toc_fails_in_words_stories()
    {
        RequireWord();
        var original = Original();
        var redline = RestyledTocRedline("word-redline-undeclared", original);

        // The original keeps its own TOC1, so Word's save keeps the direct Arial there and drops
        // it in the redline: the change to the style shows instead of being aligned away.
        var errors = Errors(WordCheck.RejectAll(original, redline, timeout: Timeout));

        Assert.Contains(errors, e => e.StartsWith("word-reject-mismatch:", StringComparison.Ordinal));
    }

    [Fact]
    public void Alignment_takes_added_and_declared_definitions_only()
    {
        var original = _r.Make("align-original", package: doc => Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(Numbering())));
        var redline = _r.Make("align-redline", from: original,
            styles: s =>
            {
                Style(s, "TOC1").Add(TocArial());
                Style(s, "TOC2").Add(TocArial());
                s.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "Head1num"),
                    new XElement(W + "name", new XAttribute(W + "val", "Head1 numbered")), new XElement(W + "basedOn", new XAttribute(W + "val", "Head1"))));
            },
            package: doc =>
            {
                var part = doc.MainDocumentPart!.NumberingDefinitionsPart!;
                var xml = Docx.Load(part);
                var root = xml.Root!;
                var abstractNum = new XElement(root.Element(W + "abstractNum")!);
                abstractNum.SetAttributeValue(W + "abstractNumId", 1);
                root.Element(W + "abstractNum")!.AddAfterSelf(abstractNum);
                root.Add(new XElement(W + "num", new XAttribute(W + "numId", 2), new XElement(W + "abstractNumId", new XAttribute(W + "val", 1))));
                root.Element(W + "abstractNum")!.Descendants(W + "lvlText").First().SetAttributeValue(W + "val", "%1.");
                Docx.Save(part, xml);
            });
        var aligned = Path.Combine(_r.Dir, "aligned.docx");

        WordCheck.AlignDefinitions(original, redline, new HashSet<string>(["TOC1"], StringComparer.Ordinal), aligned);

        using var doc = Docx.OpenRead(aligned);
        var styles = Docx.Load(doc.MainDocumentPart!.StyleDefinitionsPart!).Root!;
        var numbering = Docx.Load(doc.MainDocumentPart!.NumberingDefinitionsPart!).Root!;
        Assert.NotNull(Style(styles, "TOC1").Element(W + "rPr"));
        Assert.Null(Style(styles, "TOC2").Element(W + "rPr"));
        Assert.Equal("Head1num", (string?)styles.Elements(W + "style").Last().Attribute(W + "styleId"));
        Assert.Equal(["abstractNum 0", "abstractNum 1", "num 1", "num 2"],
            numbering.Elements().Select(e => $"{e.Name.LocalName} {(string?)e.Attribute(W + "abstractNumId") ?? (string?)e.Attribute(W + "numId")}"));
        Assert.Equal("%1.0", (string?)numbering.Element(W + "abstractNum")!.Descendants(W + "lvlText").First().Attribute(W + "val"));
        Assert.Empty(Docx.Validate(doc));
    }

    [Fact]
    public void Untracked_header_edit_fails_in_words_stories_and_pages()
    {
        RequireWord();
        var original = Original();
        var redline = _r.Make("word-header", from: original, package: doc =>
        {
            var header = doc.MainDocumentPart!.HeaderParts.Single();
            var x = Docx.Load(header);
            x.Descendants(W + "t").Single().Value = "PROJECT P042 HEADER";
            Docx.Save(header, x);
        });

        var errors = Errors(WordCheck.RejectAll(original, redline, timeout: Timeout));

        Assert.Contains(errors, e => e.StartsWith("word-reject-mismatch: Word's Reject All: header", StringComparison.Ordinal));
        Assert.Contains(errors, e => e.StartsWith("render-mismatch:", StringComparison.Ordinal));
    }
}
