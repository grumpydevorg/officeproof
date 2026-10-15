using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.TestKit;

/// <summary>
/// Builds small, deterministic .docx fixtures in code, so tests never depend on client files.
/// The styles mirror a real client report's house set: Normal is Times New Roman 12 pt; Head1, Head2
/// and Head3 are custom heading styles with outline levels 0-2 and no numbering; TOC1-3 are
/// based on Normal with no font of their own.
/// </summary>
public sealed class DocxBuilder
{
    private static readonly XNamespace W = Ns.W;
    private readonly List<XElement> _body = [];
    private readonly List<XElement> _extraStyles = [];
    private readonly List<(string RelId, bool IsHeader, XElement[] Paragraphs)> _headerFooters = [];
    private readonly List<(int Id, XElement[] Paragraphs)> _footnotes = [];
    private readonly List<(int Id, XElement[] Paragraphs)> _endnotes = [];
    private readonly List<(string RelId, string Url)> _links = [];
    private readonly List<(int Id, string Text)> _comments = [];
    private readonly List<(string RelId, byte[] Png)> _images = [];

    /// <summary>A heading paragraph in <paramref name="styleId"/> with literal text (typed numbers included).</summary>
    public DocxBuilder Heading(string styleId, string text) => Para(text, styleId);

    /// <summary>A plain paragraph, optionally styled.</summary>
    public DocxBuilder Para(string text, string? styleId = null)
    {
        var p = new XElement(W + "p");
        if (styleId is not null) p.Add(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", styleId))));
        p.Add(Run(text));
        _body.Add(p);
        return this;
    }

    /// <summary>
    /// A table of contents: a TOC \o "1-3" field whose cached entries are TOC1..3 paragraphs
    /// with Arial applied as direct run formatting (as in that report).
    /// </summary>
    public DocxBuilder Toc(params (int Level, string Text)[] entries)
    {
        for (var i = 0; i < entries.Length; i++)
        {
            var (level, text) = entries[i];
            var p = new XElement(W + "p",
                new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", $"TOC{level}"))));
            if (i == 0)
            {
                p.Add(FieldRun("begin"), InstrRun(" TOC \\o \"1-3\" \\h \\z \\u "), FieldRun("separate"));
            }
            p.Add(new XElement(W + "r",
                new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial"))),
                new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text)));
            if (i == entries.Length - 1) p.Add(FieldRun("end"));
            _body.Add(p);
        }
        return this;
    }

    /// <summary>Appends an arbitrary body element (paragraph or table).</summary>
    public DocxBuilder Raw(XElement element)
    {
        _body.Add(element);
        return this;
    }

    /// <summary>Adds a style definition (a complete w:style element).</summary>
    public DocxBuilder Style(XElement style)
    {
        _extraStyles.Add(style);
        return this;
    }

    /// <summary>
    /// A default-type header holding <paramref name="paragraphs"/>, referenced from the section
    /// properties as <paramref name="relId"/>.
    /// </summary>
    public DocxBuilder Header(string relId, params XElement[] paragraphs)
    {
        _headerFooters.Add((relId, true, paragraphs));
        return this;
    }

    /// <summary>A default-type footer, as <see cref="Header"/>.</summary>
    public DocxBuilder Footer(string relId, params XElement[] paragraphs)
    {
        _headerFooters.Add((relId, false, paragraphs));
        return this;
    }

    /// <summary>
    /// A footnote with id <paramref name="id"/> (1 or more), referenced from the body with
    /// <see cref="NoteReference"/>. The part also gets Word's separator (-1) and continuation
    /// separator (0).
    /// </summary>
    public DocxBuilder Footnote(int id, params XElement[] paragraphs)
    {
        _footnotes.Add((id, paragraphs));
        return this;
    }

    /// <summary>An endnote, as <see cref="Footnote"/>.</summary>
    public DocxBuilder Endnote(int id, params XElement[] paragraphs)
    {
        _endnotes.Add((id, paragraphs));
        return this;
    }

    /// <summary>
    /// An external hyperlink relationship <paramref name="relId"/> to <paramref name="url"/> on
    /// the main document part, for <c>&lt;w:hyperlink r:id="relId"&gt;</c> in the body.
    /// </summary>
    public DocxBuilder Link(string relId, string url)
    {
        _links.Add((relId, url));
        return this;
    }

    /// <summary>
    /// A comment with id <paramref name="id"/>; the body anchors it with commentRangeStart/End
    /// and a commentReference run carrying the same id.
    /// </summary>
    public DocxBuilder Comment(int id, string text)
    {
        _comments.Add((id, text));
        return this;
    }

    /// <summary>A paragraph of the given content (pPr, runs, hyperlinks, fields...).</summary>
    public static XElement P(params object?[] content) => new(W + "p", content);

    /// <summary>A plain run.</summary>
    public static XElement R(string text) => Run(text);

    /// <summary>A run holding a footnote (or endnote) reference.</summary>
    public static XElement NoteReference(int id, bool endnote = false) =>
        new(W + "r", new XElement(W + (endnote ? "endnoteReference" : "footnoteReference"), new XAttribute(W + "id", id)));

    /// <summary>Two distinct, valid 1x1 PNG images for image-part fixtures.</summary>
    public static readonly byte[] PngA = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==");

    /// <inheritdoc cref="PngA"/>
    public static readonly byte[] PngB = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    /// <summary>A PNG image part on the main document part under relationship <paramref name="relId"/>, for <see cref="Picture"/>.</summary>
    public DocxBuilder Image(string relId, byte[] png)
    {
        _images.Add((relId, png));
        return this;
    }

    /// <summary>A run holding an inline picture (DrawingML) of image relationship <paramref name="relId"/>.</summary>
    public static XElement Picture(string relId, int id) => new(W + "r", new XElement(W + "drawing",
        Inline(id, $"Picture {id}", Pic + "pic", new XElement(Pic + "pic",
            new XAttribute(XNamespace.Xmlns + "pic", Pic.NamespaceName),
            new XElement(Pic + "nvPicPr",
                new XElement(Pic + "cNvPr", new XAttribute("id", 0), new XAttribute("name", $"image{id}.png")),
                new XElement(Pic + "cNvPicPr")),
            new XElement(Pic + "blipFill",
                new XElement(A + "blip", new XAttribute(Ns.R + "embed", relId)),
                new XElement(A + "stretch", new XElement(A + "fillRect"))),
            new XElement(Pic + "spPr", Frame())))));

    /// <summary>A run holding a VML text box (w:pict/v:shape/v:textbox/w:txbxContent).</summary>
    public static XElement VmlTextBox(params XElement[] paragraphs) => new(W + "r", VmlBox("TextBox1", paragraphs));

    /// <summary>
    /// A run holding a DrawingML text box as Word writes it: mc:AlternateContent with a wps shape
    /// in mc:Choice and the same text as a VML box in mc:Fallback.
    /// </summary>
    public static XElement ShapeTextBox(int id, params XElement[] paragraphs) => new(W + "r",
        new XElement(Mc + "AlternateContent", new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
            new XElement(Mc + "Choice", new XAttribute(XNamespace.Xmlns + "wps", Wps.NamespaceName), new XAttribute("Requires", "wps"),
                new XElement(W + "drawing", Inline(id, $"Text Box {id}", Wps + "wsp", new XElement(Wps + "wsp",
                    new XElement(Wps + "cNvSpPr", new XAttribute("txBox", 1)),
                    new XElement(Wps + "spPr", Frame()),
                    new XElement(Wps + "txbx", new XElement(W + "txbxContent", paragraphs)),
                    new XElement(Wps + "bodyPr"))))),
            new XElement(Mc + "Fallback", VmlBox($"TextBox{id}", paragraphs.Select(p => new XElement(p)).ToArray()))));

    /// <summary>A run-level content control holding <paramref name="content"/> (runs).</summary>
    public static XElement Sdt(int id, params object[] content) => new(W + "sdt",
        new XElement(W + "sdtPr", new XElement(W + "id", new XAttribute(W + "val", id))),
        new XElement(W + "sdtContent", content));

    /// <summary>A smart tag around <paramref name="content"/> (runs).</summary>
    public static XElement SmartTag(params object[] content) => new(W + "smartTag",
        new XAttribute(W + "uri", "urn:schemas-microsoft-com:office:smarttags"), new XAttribute(W + "element", "place"), content);

    /// <summary>A simple field with instruction <paramref name="instruction"/> and result <paramref name="content"/>.</summary>
    public static XElement FldSimple(string instruction, params object[] content) =>
        new(W + "fldSimple", new XAttribute(W + "instr", instruction), content);

    private static XElement Inline(int id, string name, XName graphicDataUri, XElement graphic) =>
        new(Wp + "inline", new XAttribute(XNamespace.Xmlns + "wp", Wp.NamespaceName), new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
            new XElement(Wp + "extent", new XAttribute("cx", 914400), new XAttribute("cy", 457200)),
            new XElement(Wp + "docPr", new XAttribute("id", id), new XAttribute("name", name)),
            new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", graphicDataUri.NamespaceName), graphic)));

    private static XElement[] Frame() =>
    [
        new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
            new XElement(A + "ext", new XAttribute("cx", 914400), new XAttribute("cy", 457200))),
        new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")),
    ];

    private static XElement VmlBox(string id, XElement[] paragraphs) => new(W + "pict",
        new XElement(V + "shape", new XAttribute(XNamespace.Xmlns + "v", V.NamespaceName),
            new XAttribute("id", id), new XAttribute("type", "#_x0000_t202"), new XAttribute("style", "width:72pt;height:36pt"),
            new XElement(V + "textbox", new XElement(W + "txbxContent", paragraphs))));

    /// <summary>Writes the package to <paramref name="path"/> and returns the path.</summary>
    public string Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var sectPr = new XElement(W + "sectPr");
            foreach (var (relId, isHeader, paragraphs) in _headerFooters)
            {
                OpenXmlPart part = isHeader ? main.AddNewPart<HeaderPart>(relId) : main.AddNewPart<FooterPart>(relId);
                Docx.Save(part, new XDocument(Root(isHeader ? "hdr" : "ftr", paragraphs)));
                sectPr.Add(new XElement(W + (isHeader ? "headerReference" : "footerReference"),
                    new XAttribute(W + "type", "default"), new XAttribute(Ns.R + "id", relId)));
            }
            sectPr.Add(new XElement(W + "pgSz", new XAttribute(W + "w", 11906), new XAttribute(W + "h", 16838)));
            Docx.Save(main, new XDocument(Root("document", new XElement(W + "body", _body, sectPr))));
            foreach (var (relId, url) in _links)
                main.AddHyperlinkRelationship(new Uri(url, UriKind.Absolute), isExternal: true, relId);
            foreach (var (relId, png) in _images)
            {
                var image = main.AddImagePart(ImagePartType.Png, relId);
                using var s = new MemoryStream(png);
                image.FeedData(s);
            }
            if (_footnotes.Count > 0)
                Docx.Save(main.AddNewPart<FootnotesPart>(), new XDocument(Notes("footnote", _footnotes)));
            if (_endnotes.Count > 0)
                Docx.Save(main.AddNewPart<EndnotesPart>(), new XDocument(Notes("endnote", _endnotes)));
            if (_comments.Count > 0)
                Docx.Save(main.AddNewPart<WordprocessingCommentsPart>(), new XDocument(Root("comments",
                    _comments.Select(c => new XElement(W + "comment", new XAttribute(W + "id", c.Id),
                        new XAttribute(W + "author", "fixture"), new XAttribute(W + "date", "2026-10-15T00:00:00Z"),
                        P(R(c.Text)))))));
            Docx.Save(main.AddNewPart<StyleDefinitionsPart>(), new XDocument(Styles()));
            Docx.Save(main.AddNewPart<DocumentSettingsPart>(), new XDocument(Root("settings")));
        }
        return path;
    }

    private static XElement Root(string name, params object?[] content) => new(W + name,
        new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "r", Ns.R.NamespaceName),
        content);

    private static XElement Notes(string kind, IEnumerable<(int Id, XElement[] Paragraphs)> notes) => Root(kind + "s",
        new XElement(W + kind, new XAttribute(W + "type", "separator"), new XAttribute(W + "id", -1),
            P(new XElement(W + "r", new XElement(W + "separator")))),
        new XElement(W + kind, new XAttribute(W + "type", "continuationSeparator"), new XAttribute(W + "id", 0),
            P(new XElement(W + "r", new XElement(W + "continuationSeparator")))),
        notes.Select(n => new XElement(W + kind, new XAttribute(W + "id", n.Id), n.Paragraphs)));

    /// <summary>
    /// The canonical regression fixture: a report shaped like that report, with typed heading
    /// numbers (including a duplicated "2.0"), typed caption numbers, typed cross-references,
    /// a "see section 3.2" reference and a TOC with direct Arial.
    /// </summary>
    public static string TypedNumberReport(string path) => new DocxBuilder()
        .Toc((1, "1.0 Summary"), (1, "2.0 Background"), (1, "2.0 Aims"), (1, "3.0 Data"), (2, "3.1 Frog Pond"), (2, "3.2 Birds"))
        .Heading("Head1", "Abbreviations")
        .Para("m asl  metres above sea level")
        .Heading("Head1", "1.0 Summary")
        .Para("Summary text.")
        .Heading("Head1", "2.0 Background")
        .Para("Background text (Figure 1).")
        .Heading("Head1", "2.0 Aims")
        .Para("Aims text.")
        .Heading("Head1", "3.0 Data")
        .Heading("Head2", "3.1 Frog Pond")
        .Para("Stations near Wellpad D (Figure 1). Counts in Table 1.")
        .Para("Table 1. Recordings per station.")
        .Para("Figure 1. Photographs of the stations.")
        .Heading("Head2", "3.2 Birds")
        .Para("Reference calls (Figure 2); see section 3.1 for stations.")
        .Para("Figure 2. Ordination of reference calls.")
        .Heading("Head3", "3.2.1 Detail")
        .Para("Detail text, compare Figure 1 and Figure 2.")
        .Heading("Head1", "Appendix 1 Species")
        .Save(path);

    private static XElement Run(string text) =>
        new(W + "r", new XElement(W + "t", new XAttribute(Ns.Xml + "space", "preserve"), text));

    private static XElement FieldRun(string kind) =>
        new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", kind)));

    private static XElement InstrRun(string code) =>
        new(W + "r", new XElement(W + "instrText", new XAttribute(Ns.Xml + "space", "preserve"), code));

    private XElement Styles()
    {
        XElement ParaStyle(string id, string name, string? basedOn, XElement? pPr, XElement? rPr, bool isDefault = false)
        {
            var s = new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", id));
            if (isDefault) s.Add(new XAttribute(W + "default", "1"));
            s.Add(new XElement(W + "name", new XAttribute(W + "val", name)));
            if (basedOn is not null) s.Add(new XElement(W + "basedOn", new XAttribute(W + "val", basedOn)));
            if (pPr is not null) s.Add(pPr);
            if (rPr is not null) s.Add(rPr);
            return s;
        }
        XElement Outline(int level) => new(W + "pPr", new XElement(W + "keepNext"), new XElement(W + "outlineLvl", new XAttribute(W + "val", level)));
        XElement Font(string face, int halfPoints, bool bold) => new(W + "rPr",
            new XElement(W + "rFonts", new XAttribute(W + "ascii", face), new XAttribute(W + "hAnsi", face), new XAttribute(W + "cs", face)),
            bold ? new XElement(W + "b") : null,
            new XElement(W + "sz", new XAttribute(W + "val", halfPoints)));

        return new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            new XElement(W + "docDefaults", new XElement(W + "rPrDefault", new XElement(W + "rPr",
                new XElement(W + "rFonts", new XAttribute(W + "ascii", "Century"), new XAttribute(W + "hAnsi", "Century")))),
                new XElement(W + "pPrDefault")),
            ParaStyle("Normal", "Normal", null, null, Font("Times New Roman", 24, false), isDefault: true),
            ParaStyle("Head1", "Head1", "Normal", Outline(0), Font("Arial", 32, true)),
            ParaStyle("Head2", "Head2", "Normal", Outline(1), Font("Arial", 24, true)),
            ParaStyle("Head3", "Head3", "Normal", Outline(2), Font("Arial", 22, false)),
            ParaStyle("TOC1", "toc 1", "Normal", null, null),
            ParaStyle("TOC2", "toc 2", "Normal", new XElement(W + "pPr", new XElement(W + "ind", new XAttribute(W + "left", 240))), null),
            ParaStyle("TOC3", "toc 3", "Normal", new XElement(W + "pPr", new XElement(W + "ind", new XAttribute(W + "left", 480))), null),
            _extraStyles);
    }
}
