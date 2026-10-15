using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using OfficeProof.Core;
using OfficeProof.Redline;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Redline.Tests;

internal static class Support
{
    public static readonly XNamespace W = Ns.W;
    public static readonly RevisionAuthor Author = new("Claude", "C", new DateTimeOffset(2026, 10, 15, 1, 2, 3, TimeSpan.Zero));

    public static readonly Dictionary<string, int> ReportLevels = new() { ["Head1"] = 1, ["Head2"] = 2, ["Head3"] = 3 };

    public static string TempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "officeproof-redline-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string TypedNumberReport() => DocxBuilder.TypedNumberReport(Path.Combine(TempDir(), "report.docx"));

    /// <summary>Copies <paramref name="input"/>, edits it through a TrackedEditor, saves, returns the output path.</summary>
    public static string Edit(string input, Action<TrackedEditor> edit)
    {
        var output = Path.Combine(Path.GetDirectoryName(input)!, "out-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = Docx.OpenCopy(input, output))
        {
            var editor = new TrackedEditor(doc, Author);
            edit(editor);
            editor.Save();
        }
        return output;
    }

    public static XDocument MainXml(string path)
    {
        using var doc = Docx.OpenRead(path);
        return Docx.Load(doc.MainDocumentPart!);
    }

    public static XDocument PartXml(string path, Func<MainDocumentPart, OpenXmlPart?> select)
    {
        using var doc = Docx.OpenRead(path);
        return Docx.Load(select(doc.MainDocumentPart!)!);
    }

    public static List<string> BodyTexts(string path) =>
        MainXml(path).Descendants(W + "p").Select(Wml.VisibleText).ToList();

    /// <summary>Paragraph text and Word list number after Clippit accept or reject (Word's behaviour, verified on a client report).</summary>
    public static List<(string Text, string Number, string? Style)> Outline(string path, bool accept)
    {
        var wml = new WmlDocument(path);
        var result = accept ? RevisionProcessor.AcceptRevisions(wml) : RevisionProcessor.RejectRevisions(wml);
        using var ms = new MemoryStream();
        ms.Write(result.DocumentByteArray, 0, result.DocumentByteArray.Length);
        using var doc = WordprocessingDocument.Open(ms, true);
        var xml = doc.MainDocumentPart!.GetXDocument();
        return xml.Descendants(W + "p")
            .Select(p => (Wml.VisibleText(p), ListItemRetriever.RetrieveListItem(doc, p, (ListItemRetrieverSettings?)null!) ?? "", Wml.StyleId(p)))
            .ToList();
    }

}
