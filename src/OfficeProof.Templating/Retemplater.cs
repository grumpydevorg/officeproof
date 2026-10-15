using System.Collections.Frozen;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>Outcome of <see cref="Retemplater.Retemplate"/>.</summary>
/// <param name="OutputPath">The retemplated copy.</param>
/// <param name="HeadingsRemapped">Paragraphs moved from a legacy heading style (<see cref="TemplateManifest.HeadingStyles"/>) to its canonical heading.</param>
/// <param name="NumbersDropped">Typed heading numbers ("2.1 ") removed; the template's numbering supplies them.</param>
/// <param name="DirectFormattingRemoved">Direct run-property groups (font, size, colour, language) removed.</param>
/// <param name="CoverParagraphsReplaced">Hand-formatted cover paragraphs replaced by the data-bound cover block.</param>
/// <param name="CoverBlockInserted">False when the manifest has no cover anchor or the document has no paragraph containing it, so its cover was left alone.</param>
/// <param name="Sections">Sections retargeted to the template header/footer.</param>
/// <param name="StaleHeaderFooterParts">The document's own header/footer parts, unreferenced after retargeting and removed.</param>
/// <param name="IdentitySources">Where the document's own identity came from; <see cref="IdentitySources.None"/> means the seed supplied it all.</param>
/// <param name="IdentityKeys">Metadata keys the document supplied (these override the seed).</param>
/// <param name="Metadata">The metadata written: the seed overridden by the document's identity.</param>
/// <param name="Fonts">Font embedding outcome, including missing faces; null when the template embeds no fonts.</param>
/// <param name="Cover">Cover image outcome.</param>
/// <param name="Settings">Settings added.</param>
public sealed record RetemplateResult(
    string OutputPath,
    int HeadingsRemapped,
    int NumbersDropped,
    int DirectFormattingRemoved,
    int CoverParagraphsReplaced,
    bool CoverBlockInserted,
    int Sections,
    int StaleHeaderFooterParts,
    IdentitySources IdentitySources,
    IReadOnlyList<string> IdentityKeys,
    ReportMetadata Metadata,
    FontEmbedResult? Fonts,
    CoverHeaderResult Cover,
    SettingsChange Settings);

/// <summary>
/// Applies the template to a real report: swaps the canonical parts, remaps the manifest's
/// legacy heading styles to the canonical headings, drops typed heading numbers, strips direct
/// font/size/colour so the body and headings inherit the template's styles, binds the cover,
/// adds the template footer and a clean first page, embeds fonts and turns on updateFields.
/// </summary>
/// <remarks>
/// Schema validation and lint are separate steps; the CLI runs them on the output.
/// </remarks>
public static class Retemplater
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace R = Ns.R;

    private static readonly FrozenSet<string> TocStyles = new[] { "TOC1", "TOC2", "TOC3", "TOCHeading" }.ToFrozenSet();

    // CT_SectPr children that follow w:titlePg; it goes before the first one present.
    private static readonly FrozenSet<string> AfterTitlePg =
        new[] { "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange" }.ToFrozenSet();

    /// <summary>
    /// Retemplates a copy of <paramref name="input"/> at <paramref name="output"/> (opened with
    /// <see cref="Docx.OpenCopy"/>; the input is never modified).
    /// </summary>
    /// <param name="source">The canonical parts.</param>
    /// <param name="input">The report to retemplate.</param>
    /// <param name="output">Where to write the result; must differ from <paramref name="input"/>.</param>
    /// <param name="seed">Metadata for whatever the document cannot say about itself.</param>
    /// <param name="fontDirectory">Where the manifest's font faces are; default <see cref="FontEmbedder.UserFontDirectory"/>.</param>
    public static RetemplateResult Retemplate(TemplateSource source, string input, string output, ReportMetadata seed, string? fontDirectory = null)
    {
        var fonts = source.FontEmbedder(fontDirectory);
        var manifest = source.Manifest;
        using var doc = Docx.OpenCopy(input, output);
        var main = doc.MainDocumentPart ?? throw new InvalidDataException($"{input}: no main document part");

        Docx.Write(main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>(), source.ReadBytes(TemplateSource.Styles));
        Docx.Write(main.ThemePart ?? main.AddNewPart<ThemePart>(), source.ReadBytes(TemplateSource.Theme));
        Docx.Write(main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>(), source.ReadBytes(TemplateSource.Numbering));

        // The DOCUMENT is the source of truth for its own identity; the metadata file is only a
        // seed for whatever the document cannot tell us. Before 2026-08-27 it was the other way
        // round, and because report.yaml was a single hand-edited file with no --metadata flag, a
        // retemplated report came out headed with the previous run's title, inheriting its
        // identity. A global file that stamps the last document's name onto the next one is the
        // same failure this toolkit exists to remove.
        var identity = IdentityReader.Read(doc, manifest);
        var metadata = seed.With(identity.Values);
        DocPropsWriter.Write(doc, metadata, manifest, newDocument: false);

        var stale = main.HeaderParts.Cast<OpenXmlPart>().Concat(main.FooterParts).ToList();
        var storeId = DataStoreWriter.Write(main, metadata, manifest.DataStore);
        var headerId = HeaderFooterBuilder.AddRunningHeader(main, source, metadata, storeId);
        var footerId = HeaderFooterBuilder.AddFooter(main, source);
        var cover = HeaderFooterBuilder.AddCoverHeader(main, source, metadata);

        var xml = Docx.Load(main);
        var (headings, numbers, stripped) = Restyle(xml, manifest.HeadingStyles);
        var sections = Retarget(xml, headerId, cover.RelationshipId, footerId);
        var (coverReplaced, coverInserted) = ReplaceCoverBlock(xml, source, metadata, storeId);
        Docx.Save(main, xml);

        // Every section now points at the template's header/footer, so the document's own are
        // unreferenced. Removing them keeps later lints (which walk every header/footer part)
        // from reporting on content nobody can see.
        foreach (var part in stale) main.DeletePart(part);

        var settings = SettingsInjector.Apply(main);
        var embedded = fonts?.Embed(main);

        return new RetemplateResult(Path.GetFullPath(output), headings, numbers, stripped, coverReplaced, coverInserted,
            sections, stale.Count, identity.Sources, [.. identity.Values.Keys], metadata, embedded, cover, settings);
    }

    private static (int Headings, int Numbers, int Stripped) Restyle(XDocument xml, FrozenDictionary<string, string> headingMap)
    {
        int headings = 0, numbers = 0, stripped = 0;
        var inCover = true;   // preserve the cover's own sizing; normalise everything after it

        foreach (var p in xml.Descendants(W + "p").ToList())
        {
            var pStyle = p.Element(W + "pPr")?.Element(W + "pStyle");
            var sid = (string?)pStyle?.Attribute(W + "val");
            var heading = sid is not null && headingMap.TryGetValue(sid, out var mapped) ? mapped : null;
            if (sid is not null && (heading is not null || TocStyles.Contains(sid))) inCover = false;

            if (heading is not null)
            {
                pStyle!.SetAttributeValue(W + "val", heading);
                headings++;
                if (DropTypedNumber(p)) numbers++;
            }

            // Strip direct font/colour everywhere; keep size on the cover so its title block survives.
            string[] names = inCover ? ["rFonts", "color", "lang"] : ["rFonts", "sz", "szCs", "color", "lang"];
            foreach (var rPr in p.Descendants(W + "rPr"))
                foreach (var name in names)
                {
                    var els = rPr.Elements(W + name).ToList();
                    if (els.Count == 0) continue;
                    els.Remove();
                    stripped++;
                }
        }
        return (headings, numbers, stripped);
    }

    // The template's Heading styles number themselves, so a typed "2.1 " would print twice.
    // Matched on the text Word displays (tabs count as whitespace) and removed node by node,
    // since Word splits a heading's text across runs arbitrarily.
    //
    // "Typed number" is Core's TypedNumbers.HeadingNumber, shared with Lint and Redline
    // (2026-10-15 review: the three disagreed). Behaviour change from the legacy CLI, whose
    // regex `^\d+(\.\d+)+\.?\s+` required a dot inside the number: "4. Results" and
    // "1 Introduction" kept their typed numbers and printed them twice under the template's
    // numbering; both are now stripped. A year ("2025 Field season") is still not a number,
    // because the first component is capped at three digits.
    private static bool DropTypedNumber(XElement paragraph)
    {
        var nodes = Wml.VisibleTextNodes(paragraph).ToList();
        var m = TypedNumbers.HeadingNumber().Match(Wml.VisibleText(paragraph));
        if (!m.Success) return false;
        var remove = m.Length;
        foreach (var node in nodes)
        {
            if (remove <= 0) break;
            if (node.Name == W + "tab")
            {
                node.Remove();
                remove--;
                continue;
            }
            var text = node.Value;
            if (text.Length <= remove)
            {
                remove -= text.Length;
                node.Value = "";
            }
            else
            {
                node.Value = text[remove..];
                remove = 0;
            }
        }
        return true;
    }

    // Retarget every section to the one consistent header/footer; the first section also gets
    // the cover header on its first page (w:titlePg).
    private static int Retarget(XDocument xml, string headerId, string firstHeaderId, string footerId)
    {
        // A w:sectPr inside w:sectPrChange is the tracked previous version; it cannot carry
        // header/footer references.
        var sections = xml.Descendants(W + "sectPr").Where(s => s.Parent?.Name != W + "sectPrChange").ToList();
        for (var i = 0; i < sections.Count; i++)
        {
            var sect = sections[i];
            sect.Elements(W + "headerReference").Remove();
            sect.Elements(W + "footerReference").Remove();
            sect.Elements(W + "titlePg").Remove();

            List<XElement> refs = [Reference("headerReference", "default", headerId)];
            if (i == 0) refs.Add(Reference("headerReference", "first", firstHeaderId));
            refs.Add(Reference("footerReference", "default", footerId));
            sect.AddFirst(refs);

            if (i == 0)
            {
                var titlePg = new XElement(W + "titlePg");
                var follower = sect.Elements().FirstOrDefault(e => e.Name.Namespace == W && AfterTitlePg.Contains(e.Name.LocalName));
                if (follower is not null) follower.AddBeforeSelf(titlePg);
                else sect.Add(titlePg);
            }
        }
        return sections.Count;
    }

    private static XElement Reference(string kind, string type, string id) =>
        new(W + kind, new XAttribute(W + "type", type), new XAttribute(R + "id", id));

    // Replace a hand-formatted cover metadata block with the styled, data-bound block. The typed
    // block sits directly above the paragraph containing the manifest's cover anchor (in the
    // reports this was written against, the company registration line): every body paragraph
    // before it is replaced and the anchor line is kept. Without an anchor, or a paragraph
    // containing it, the cover is left untouched.
    private static (int Replaced, bool Inserted) ReplaceCoverBlock(XDocument xml, TemplateSource source, ReportMetadata metadata, string storeId)
    {
        if (source.Manifest.CoverAnchor is not { } text) return (0, false);
        var anchor = xml.Root?.Element(W + "body")?.Elements(W + "p").FirstOrDefault(p => Wml.VisibleText(p).Contains(text, StringComparison.Ordinal));
        if (anchor is null) return (0, false);

        var before = anchor.ElementsBeforeSelf(W + "p").ToList();
        before.ForEach(e => e.Remove());
        var block = PartTemplate.Fill(source.ReadText(TemplateSource.CoverBlock), metadata, ("STOREID", storeId));
        var wrapped = XElement.Parse($"<block xmlns:w=\"{W.NamespaceName}\">{block}</block>", LoadOptions.PreserveWhitespace);
        anchor.AddBeforeSelf(wrapped.Elements());
        return (before.Count, true);
    }
}
