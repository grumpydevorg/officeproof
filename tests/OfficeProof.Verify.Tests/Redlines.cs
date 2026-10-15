using System.IO.Compression;
using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify.Tests;

/// <summary>Builds redlines of the TypedNumberReport fixture with raw WordprocessingML, as a redline tool would.</summary>
internal sealed class Redlines : IDisposable
{
    public static readonly XNamespace W = Ns.W;
    private const string Author = "officeproof-test";
    private const string Date = "2026-10-15T00:00:00Z";
    private int _id = 1000;

    public string Dir { get; } = Path.Combine(Path.GetTempPath(), "officeproof-verify-tests", Guid.NewGuid().ToString("N"));
    public string Original { get; }

    public Redlines()
    {
        Directory.CreateDirectory(Dir);
        Original = DocxBuilder.TypedNumberReport(Path.Combine(Dir, "original.docx"));
    }

    public void Dispose()
    {
        try { Directory.Delete(Dir, recursive: true); }
        catch (IOException) { }
    }

    /// <summary>Copies the original (or <paramref name="from"/>) and lets the test edit document.xml and styles.xml.</summary>
    public string Make(string name, Action<XElement>? body = null, Action<XElement>? styles = null, Action<WordprocessingDocument>? package = null, string? from = null)
    {
        var path = Path.Combine(Dir, name + ".docx");
        using var doc = Docx.OpenCopy(from ?? Original, path);
        var main = doc.MainDocumentPart!;
        if (body is not null)
        {
            var x = Docx.Load(main);
            body(x.Root!.Element(W + "body")!);
            Docx.Save(main, x);
        }
        if (styles is not null)
        {
            var x = Docx.Load(main.StyleDefinitionsPart!);
            styles(x.Root!);
            Docx.Save(main.StyleDefinitionsPart!, x);
        }
        package?.Invoke(doc);
        return path;
    }

    /// <summary>Copies <paramref name="from"/> (default: the original) and lets the test edit any part through <see cref="PackageEdit"/>.</summary>
    public string Edit(string name, Action<PackageEdit> edit, string? from = null)
    {
        PackageEdit? e = null;
        var path = Make(name, from: from, package: doc =>
        {
            e = new PackageEdit(this, doc);
            edit(e);
            e.Save();
        });
        e!.ApplyEntryRewrites(path);
        return path;
    }

    /// <summary>The body paragraph with exactly <paramref name="text"/>, skipping TOC lines that repeat heading text.</summary>
    public static XElement Paragraph(XElement body, string text) =>
        body.Descendants(W + "p").First(p => Wml.VisibleText(p) == text && Wml.StyleId(p)?.StartsWith("TOC", StringComparison.Ordinal) != true);

    public static XElement Style(XElement styles, string id) =>
        styles.Elements(W + "style").First(s => (string?)s.Attribute(W + "styleId") == id);

    public static XElement Run(string text) =>
        new(W + "r", new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text));

    public XElement Ins(params object[] runs) => new(W + "ins", Mark(), runs);

    public XElement Del(string text) => new(W + "del", Mark(),
        new XElement(W + "r", new XElement(W + "delText", new XAttribute(Ns.Xml + "space", "preserve"), text)));

    public object[] Mark() => [new XAttribute(W + "id", _id++), new XAttribute(W + "author", Author), new XAttribute(W + "date", Date)];

    /// <summary>Replaces the leading <paramref name="oldPrefix"/> of a single-run paragraph as a tracked del + ins.</summary>
    public void TrackedReplacePrefix(XElement paragraph, string oldPrefix, string newPrefix)
    {
        var text = Wml.VisibleText(paragraph);
        paragraph.Elements(W + "r").Remove();
        paragraph.Add(Del(oldPrefix));
        if (newPrefix.Length > 0) paragraph.Add(Ins(Run(newPrefix)));
        paragraph.Add(Run(text[oldPrefix.Length..]));
    }

    /// <summary>
    /// The v3 approach from a client redline (2026-10-15): new styles "Head1num"/"Head2num" based on Head1/Head2 with
    /// their own numPr into a new multilevel list ("%1.0", "%1.%2"), applied to each typed-number
    /// heading as a tracked pPrChange, with the typed number deleted as w:del.
    /// "2.0 Aims" becomes a level-2 heading, so Accept All reads 2.1 Aims.
    /// </summary>
    public string NumberedHeadings(string name, Action<XElement>? body = null) => Make(name,
        body: b =>
        {
            Number(b, "1.0 Summary", "1.0 ", "Head1", "Head1num");
            Number(b, "2.0 Background", "2.0 ", "Head1", "Head1num");
            Number(b, "2.0 Aims", "2.0 ", "Head1", "Head2num");
            Number(b, "3.0 Data", "3.0 ", "Head1", "Head1num");
            Number(b, "3.1 Frog Pond", "3.1 ", "Head2", "Head2num");
            Number(b, "3.2 Birds", "3.2 ", "Head2", "Head2num");
            body?.Invoke(b);
        },
        styles: s => s.Add(
            NumberedStyle("Head1num", "Head1", 0),
            NumberedStyle("Head2num", "Head2", 1)),
        package: doc => Docx.Save(doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>(), new XDocument(Numbering())));

    private void Number(XElement body, string text, string typed, string oldStyle, string newStyle)
    {
        var p = Paragraph(body, text);
        p.Element(W + "pPr")!.ReplaceWith(new XElement(W + "pPr",
            new XElement(W + "pStyle", new XAttribute(W + "val", newStyle)),
            new XElement(W + "pPrChange", Mark(), new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", oldStyle))))));
        TrackedReplacePrefix(p, typed, "");
    }

    private static XElement NumberedStyle(string id, string basedOn, int ilvl) =>
        new(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", id),
            new XElement(W + "name", new XAttribute(W + "val", id)),
            new XElement(W + "basedOn", new XAttribute(W + "val", basedOn)),
            new XElement(W + "pPr", new XElement(W + "numPr",
                new XElement(W + "ilvl", new XAttribute(W + "val", ilvl)),
                new XElement(W + "numId", new XAttribute(W + "val", 1)))));

    internal static XElement Numbering()
    {
        XElement Lvl(int ilvl, string style, string text) => new(W + "lvl", new XAttribute(W + "ilvl", ilvl),
            new XElement(W + "start", new XAttribute(W + "val", 1)),
            new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
            new XElement(W + "pStyle", new XAttribute(W + "val", style)),
            new XElement(W + "lvlText", new XAttribute(W + "val", text)),
            new XElement(W + "lvlJc", new XAttribute(W + "val", "left")));
        return new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 0),
                new XElement(W + "multiLevelType", new XAttribute(W + "val", "multilevel")),
                Lvl(0, "Head1num", "%1.0"),
                Lvl(1, "Head2num", "%1.%2")),
            new XElement(W + "num", new XAttribute(W + "numId", 1),
                new XElement(W + "abstractNumId", new XAttribute(W + "val", 0))));
    }
}

/// <summary>Edits any part of a package copy; every part loaded through it is saved once at the end.</summary>
internal sealed class PackageEdit(Redlines redlines, WordprocessingDocument doc)
{
    private static readonly XNamespace W = Ns.W;
    private readonly Dictionary<OpenXmlPart, XDocument> _loaded = [];

    /// <summary>The redline helper, for <see cref="Redlines.Ins"/>, <see cref="Redlines.Del"/> and <see cref="Redlines.Mark"/>.</summary>
    public Redlines R => redlines;

    public MainDocumentPart Main => doc.MainDocumentPart!;

    public XElement Root(OpenXmlPart part)
    {
        if (!_loaded.TryGetValue(part, out var xml)) _loaded[part] = xml = Docx.Load(part);
        return xml.Root!;
    }

    public XElement Body => Root(Main).Element(W + "body")!;
    public XElement Header => Root(Main.HeaderParts.Single());
    public XElement Footer => Root(Main.FooterParts.Single());
    public XElement Footnotes => Root(Main.FootnotesPart!);
    public XElement Endnotes => Root(Main.EndnotesPart!);
    public XElement Comments => Root(Main.WordprocessingCommentsPart!);

    public void Save()
    {
        foreach (var (part, xml) in _loaded) Docx.Save(part, xml);
    }

    private readonly List<(string Entry, Func<string, string> Edit)> _rewrites = [];

    /// <summary>
    /// Rewrites a zip entry as text after the package is closed, for edits the SDK refuses
    /// (e.g. pointing a relationship at a part another relationship already targets).
    /// </summary>
    public void Rewrite(string entry, Func<string, string> edit) => _rewrites.Add((entry, edit));

    public void ApplyEntryRewrites(string path)
    {
        if (_rewrites.Count == 0) return;
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        foreach (var (name, edit) in _rewrites)
        {
            var entry = zip.GetEntry(name) ?? throw new InvalidOperationException($"no entry {name}");
            string text;
            using (var reader = new StreamReader(entry.Open())) text = reader.ReadToEnd();
            var edited = edit(text);
            if (edited == text) throw new InvalidOperationException($"rewrite of {name} changed nothing");
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(name).Open());
            writer.Write(edited);
        }
    }
}
