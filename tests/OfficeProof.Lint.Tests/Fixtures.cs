using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Lint.Tests;

/// <summary>Temp paths, WordprocessingML snippets and package patches for lint fixtures.</summary>
internal static class Fixtures
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>A fresh .docx path under the temp directory.</summary>
    public static string NewPath(string name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "officeproof-lint-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name + ".docx");
    }

    public static string LegacyReport() => DocxBuilder.TypedNumberReport(NewPath("legacy"));

    /// <summary>Saves the builder and runs one rule.</summary>
    public static IReadOnlyList<Finding> RunRule(DocxBuilder builder, string ruleId, Action<string>? patch = null)
    {
        var path = builder.Save(NewPath(ruleId));
        patch?.Invoke(path);
        return RunRule(path, ruleId);
    }

    public static IReadOnlyList<Finding> RunRule(string path, string ruleId) => Linter.Run(path, Linter.Rules([ruleId]));

    // --- WordprocessingML snippets -------------------------------------------------------

    public static XElement P(string? style, params object?[] content) =>
        new(W + "p", style is null ? null : new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", style))), content);

    public static XElement PNum(string style, int numId, params object?[] content) =>
        new(W + "p", new XElement(W + "pPr",
                new XElement(W + "pStyle", new XAttribute(W + "val", style)),
                new XElement(W + "numPr", new XElement(W + "ilvl", new XAttribute(W + "val", 0)), new XElement(W + "numId", new XAttribute(W + "val", numId)))),
            content);

    public static XElement R(string text, XElement? rPr = null) =>
        new(W + "r", rPr, new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text));

    public static XElement FldChar(string kind) => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", kind)));

    /// <summary>A complex field: begin, instruction, separate, cached result, end.</summary>
    public static XElement[] Field(string instruction, string result) =>
    [
        FldChar("begin"),
        new(W + "r", new XElement(W + "instrText", new XAttribute(Ns.Xml + "space", "preserve"), $" {instruction} ")),
        FldChar("separate"),
        R(result),
        FldChar("end"),
    ];

    public static XElement BookmarkStart(int id, string name) =>
        new(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name));

    public static XElement BookmarkEnd(int id) => new(W + "bookmarkEnd", new XAttribute(W + "id", id));

    public static XElement Ins(int id, params XElement[] runs) =>
        new(W + "ins", new XAttribute(W + "id", id), new XAttribute(W + "author", "Test"), new XAttribute(W + "date", "2026-10-15T00:00:00Z"), runs);

    public static XElement Del(int id, string text) =>
        new(W + "del", new XAttribute(W + "id", id), new XAttribute(W + "author", "Test"), new XAttribute(W + "date", "2026-10-15T00:00:00Z"),
            new XElement(W + "r", new XElement(W + "delText", new XAttribute(Ns.Xml + "space", "preserve"), text)));

    public static XElement RFonts(string face) =>
        new(W + "rFonts", new XAttribute(W + "ascii", face), new XAttribute(W + "hAnsi", face));

    /// <summary>A paragraph style; <paramref name="pPr"/> children go in its w:pPr.</summary>
    public static XElement ParaStyle(string id, string? basedOn, params XElement[] pPr) =>
        new(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", id),
            new XElement(W + "name", new XAttribute(W + "val", id)),
            basedOn is null ? null : new XElement(W + "basedOn", new XAttribute(W + "val", basedOn)),
            pPr.Length == 0 ? null : new XElement(W + "pPr", pPr));

    public static XElement NumPr(int numId) =>
        new(W + "numPr", new XElement(W + "ilvl", new XAttribute(W + "val", 0)), new XElement(W + "numId", new XAttribute(W + "val", numId)));

    public static XElement OutlineLvl(int level) => new(W + "outlineLvl", new XAttribute(W + "val", level));

    // --- package patches -------------------------------------------------------------------

    public static void Patch(string path, Action<MainDocumentPart> edit)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: true);
        edit(doc.MainDocumentPart!);
    }

    public static void EditStyles(string path, Action<XElement> edit) => Patch(path, main =>
    {
        var x = Docx.Load(main.StyleDefinitionsPart!);
        edit(x.Root!);
        Docx.Save(main.StyleDefinitionsPart!, x);
    });

    /// <summary>The w:style element with <paramref name="id"/> inside a styles root.</summary>
    public static XElement StyleIn(XElement styles, string id) =>
        styles.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);

    /// <summary>Adds numbering.xml defining numId 1 (a "%1.0" decimal list); <paramref name="extra"/> goes inside the abstractNum.</summary>
    public static void AddNumbering(string path, XElement? extra = null) => Patch(path, main =>
        Docx.Save(main.AddNewPart<NumberingDefinitionsPart>(), new XDocument(new XElement(W + "numbering",
            new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 0),
                new XElement(W + "multiLevelType", new XAttribute(W + "val", "multilevel")),
                new XElement(W + "lvl", new XAttribute(W + "ilvl", 0),
                    new XElement(W + "start", new XAttribute(W + "val", 1)),
                    new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
                    new XElement(W + "lvlText", new XAttribute(W + "val", "%1.0")),
                    extra)),
            new XElement(W + "num", new XAttribute(W + "numId", 1), new XElement(W + "abstractNumId", new XAttribute(W + "val", 0)))))));

    public static void AddComments(string path, int count) => Patch(path, main =>
        Docx.Save(main.AddNewPart<WordprocessingCommentsPart>(), new XDocument(new XElement(W + "comments",
            new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            Enumerable.Range(0, count).Select(i => new XElement(W + "comment",
                new XAttribute(W + "id", i), new XAttribute(W + "author", "Test"), new XAttribute(W + "date", "2026-10-15T00:00:00Z"),
                P(null, R($"note {i}"))))))));
}
