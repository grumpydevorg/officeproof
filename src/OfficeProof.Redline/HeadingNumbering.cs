using System.Globalization;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Redline;

/// <summary>Options for <see cref="HeadingNumbering.Apply"/>.</summary>
public sealed record HeadingNumberingOptions
{
    /// <summary>
    /// Heading title prefix (text after the typed number) → outline level (1-based) to use instead
    /// of its style's. A client report (2026-10-15): "Aims of the Pilot Study" was a second "2.0" → level 2.
    /// </summary>
    public IReadOnlyDictionary<string, int> LevelOverrides { get; init; } = new Dictionary<string, int>();

    /// <summary>Turn typed "section N.N" references into REF \r \h fields to the heading.</summary>
    public bool LinkSectionReferences { get; init; } = true;

    /// <summary>Add comments explaining the new styles (first heading) and each level override.</summary>
    public bool ExplainInComment { get; init; } = true;
}

/// <summary>
/// Replaces typed heading numbers with real multilevel numbering that survives Reject All
/// (v3 redline of a client report, 2026-10-15, verified in Word): new "&lt;style&gt; numbered" styles based on
/// the originals carry the numbering; each numbered heading switches style as a tracked
/// w:pPrChange and its typed number is a tracked deletion. Original styles and unnumbered
/// headings are untouched. Style and numbering definitions cannot be tracked (Word ignores
/// revisions inside styles.xml on Reject All), so they are reported as untracked changes.
/// </summary>
public static class HeadingNumbering
{
    private static readonly XNamespace W = Ns.W;

    private sealed record Heading(XElement Paragraph, string Typed, string Title, int Level, int TypedLength)
    {
        public string NewNumber { get; set; } = "";
        public string? Bookmark { get; set; }
    }

    /// <summary>
    /// Numbers headings in the styles of <paramref name="styleLevels"/> (style id → outline
    /// level, 1-based). Throws if a style is missing, the numbered styles already exist, or a
    /// level override matches no numbered heading.
    /// </summary>
    public static HeadingNumberingResult Apply(TrackedEditor editor, IReadOnlyDictionary<string, int> styleLevels, HeadingNumberingOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(editor);
        ArgumentNullException.ThrowIfNull(styleLevels);
        options ??= new HeadingNumberingOptions();
        if (styleLevels.Count == 0) throw new ArgumentException("no heading styles given", nameof(styleLevels));
        if (styleLevels.Values.Any(l => l is < 1 or > 9)) throw new ArgumentException("levels are 1..9", nameof(styleLevels));

        var main = editor.Document.MainDocumentPart!;
        var stylesPart = main.StyleDefinitionsPart ?? throw new InvalidOperationException("document has no styles part");
        var styles = Docx.Load(stylesPart);
        var chain = StyleChain.FromStyles(styles.Root);
        foreach (var id in styleLevels.Keys)
        {
            if (chain.Style(id) is null) throw new InvalidOperationException($"style '{id}' not found");
            if (chain.Style(NumberedId(id)) is not null) throw new InvalidOperationException($"style '{NumberedId(id)}' already exists; numbering was already applied");
        }
        // The numbered style for each level: the first style mapped to it.
        var styleForLevel = styleLevels.GroupBy(kv => kv.Value).ToDictionary(g => g.Key, g => g.First().Key);
        foreach (var level in options.LevelOverrides.Values)
            if (!styleForLevel.ContainsKey(level)) throw new ArgumentException($"no heading style has level {level}", nameof(options));

        // ---------------------------------------------------------------- find headings
        List<Heading> numbered = [];
        List<string> unnumbered = [];
        List<string> notes = [];
        var usedOverrides = new HashSet<string>();
        foreach (var p in editor.Paragraphs)
        {
            if (Wml.StyleId(p) is not { } sid || !styleLevels.TryGetValue(sid, out var styleLevel)) continue;
            var text = TrackedEditor.Text(p);
            if (text.Trim().Length == 0) continue;
            var m = TypedNumbers.HeadingNumber().Match(text);
            if (!m.Success || !editor.IsPlainText(p, 0, m.Length))
            {
                unnumbered.Add(text.Trim());
                continue;
            }
            var title = text[m.Length..];
            var level = styleLevel;
            var key = options.LevelOverrides.Keys.FirstOrDefault(k => title.StartsWith(k, StringComparison.Ordinal));
            if (key is not null)
            {
                level = options.LevelOverrides[key];
                usedOverrides.Add(key);
            }
            else if (TypedDepth(m.Groups["n"].Value) != styleLevel)
            {
                notes.Add($"\"{text.Trim()}\": typed number {m.Groups["n"].Value} is level {TypedDepth(m.Groups["n"].Value)} but style {sid} is level {styleLevel}; numbered by its style");
            }
            numbered.Add(new Heading(p, m.Groups["n"].Value, title, level, m.Length));
        }
        foreach (var key in options.LevelOverrides.Keys.Where(k => !usedOverrides.Contains(k)))
            throw new InvalidOperationException($"level override \"{key}\" matches no numbered heading");

        // What Word will show after accepting: %1.0 at level 1, %1.%2[.%3…] below.
        var counters = new int[9];
        foreach (var h in numbered)
        {
            counters[h.Level - 1]++;
            for (var i = h.Level; i < 9; i++) counters[i] = 0;
            h.NewNumber = h.Level == 1
                ? counters[0].ToString(CultureInfo.InvariantCulture) + ".0"
                : string.Join(".", counters.Take(h.Level).Select(c => c.ToString(CultureInfo.InvariantCulture)));
        }

        // ---------------------------------------------------------------- definitions (untracked)
        var numbering = AddNumbering(main, styles, chain, styleForLevel, out var abstractId, out var numId);
        foreach (var (baseId, level) in styleLevels)
        {
            var baseName = StyleName(chain, baseId);
            var next = chain.DefaultParagraphStyle;
            styles.Root!.Add(new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"), new XAttribute(W + "customStyle", "1"), new XAttribute(W + "styleId", NumberedId(baseId)),
                new XElement(W + "name", new XAttribute(W + "val", baseName + " numbered")),
                new XElement(W + "basedOn", new XAttribute(W + "val", baseId)),
                next is null ? null : new XElement(W + "next", new XAttribute(W + "val", next)),
                new XElement(W + "qFormat"),
                new XElement(W + "pPr", new XElement(W + "numPr",
                    level > 1 ? new XElement(W + "ilvl", new XAttribute(W + "val", level - 1)) : null,
                    new XElement(W + "numId", new XAttribute(W + "val", numId))))));
        }
        Docx.Save(stylesPart, styles);
        Docx.Save(numbering.Part, numbering.Xml);
        List<string> untracked =
        [
            .. styleLevels.Select(kv => $"styles.xml: added paragraph style '{NumberedId(kv.Key)}' (\"{StyleName(chain, kv.Key)} numbered\"), based on {kv.Key}, numbered at level {kv.Value}"),
            $"numbering.xml: added multilevel list abstractNum {abstractId} / num {numId} (\"%1.0\", \"%1.%2\", \"%1.%2.%3\")",
        ];

        // ---------------------------------------------------------------- tracked heading edits
        List<AddedComment> comments = [];
        foreach (var h in numbered)
        {
            editor.Replace(h.Paragraph, 0, h.TypedLength, []);
            var style = NumberedId(styleForLevel[h.Level]);
            editor.ChangeParagraph(h.Paragraph, pPr => TrackedEditor.SetStyle(pPr, style));
        }
        if (options.ExplainInComment)
        {
            foreach (var h in numbered.Where(h => options.LevelOverrides.Keys.Any(k => h.Title.StartsWith(k, StringComparison.Ordinal))))
            {
                comments.Add(editor.Comment(h.Paragraph, h.Paragraph,
                    $"This was typed as {h.Typed}. It now has the level-{h.Level} style ({StyleName(chain, styleForLevel[h.Level])} numbered) and numbers as {h.NewNumber}. "
                    + $"To keep its own level instead, apply the other numbered heading style; later sections renumber (update the TOC afterwards)."));
            }
            if (numbered.Count > 0)
            {
                var names = string.Join(", ", styleLevels.OrderBy(kv => kv.Value)
                    .Select(kv => $"'{StyleName(chain, kv.Key)} numbered' = {Example(kv.Value)}"));
                comments.Insert(0, editor.Comment(numbered[0].Paragraph, numbered[0].Paragraph,
                    "Heading numbers were typed, so they could not realign. Numbered headings now use new styles with their "
                    + $"own numbering, each based on your existing style so they look the same: {names}. Apply the style and the "
                    + "heading numbers itself; move or delete a section and the rest renumber. Unnumbered headings keep their "
                    + "original style. Accept the heading changes together: rejecting only the deleted typed numbers would show "
                    + "each number twice. After accepting, update the table of contents (Update Field > Update entire table)."));
            }
        }

        // ---------------------------------------------------------------- "section N.N" → REF \r \h
        var sectionRefs = 0;
        if (options.LinkSectionReferences)
        {
            var byTyped = numbered.ToLookup(h => h.Typed);
            foreach (var p in editor.Paragraphs.ToList())
            {
                if (editor.StartsInsideField(p)) continue;
                var text = TrackedEditor.Text(p);
                foreach (var m in TypedNumbers.SectionReference().Matches(text).Reverse())
                {
                    var g = m.Groups["n"];
                    if (!editor.IsPlainText(p, g.Index, g.Index + g.Length)) continue;
                    var targets = byTyped[g.Value].ToList();
                    if (targets.Count != 1)
                    {
                        var runs = editor.IsolateRuns(p, g.Index, g.Index + g.Length);
                        comments.Add(editor.Comment(runs[0], runs[^1], targets.Count == 0
                            ? $"No heading was numbered {g.Value}; this section reference is left as typed text."
                            : $"{targets.Count} headings were numbered {g.Value}; this section reference is left as typed text. Insert > Cross-reference > Heading to link the intended one."));
                        notes.Add($"section {g.Value} in \"{ParagraphText.Excerpt(text)}\": {(targets.Count == 0 ? "no such heading" : "ambiguous")}");
                        continue;
                    }
                    var h = targets[0];
                    if (h.Bookmark is null)
                    {
                        h.Bookmark = editor.NewRefBookmarkName();
                        editor.Bookmark(h.Paragraph, h.Paragraph, h.Paragraph, h.Bookmark);
                    }
                    editor.Replace(p, g.Index, g.Index + g.Length, [Inline.Field($@"REF {h.Bookmark} \r \h", h.NewNumber)]);
                    sectionRefs++;
                }
            }
        }

        return new HeadingNumberingResult(numbered.Count, unnumbered, sectionRefs, untracked, comments, notes);
    }

    private static string NumberedId(string styleId) => styleId.Replace(" ", "", StringComparison.Ordinal) + "numbered";

    private static string StyleName(StyleChain chain, string id) =>
        (string?)chain.Style(id)?.Element(W + "name")?.Attribute(W + "val") ?? id;

    private static string Example(int level) => level == 1 ? "1.0" : string.Join(".", Enumerable.Repeat("1", level));

    /// <summary>"1.0" and "2" are level 1, "3.1" level 2, "3.2.1" level 3.</summary>
    private static int TypedDepth(string typed)
    {
        var parts = typed.Split('.');
        return parts.Length == 2 && parts[1] == "0" ? 1 : parts.Length;
    }

    private sealed record NumberingXml(NumberingDefinitionsPart Part, XDocument Xml);

    private static NumberingXml AddNumbering(MainDocumentPart main, XDocument styles, StyleChain chain,
        Dictionary<int, string> styleForLevel, out int abstractId, out int numId)
    {
        var part = main.NumberingDefinitionsPart;
        XDocument xml;
        if (part is null)
        {
            part = main.AddNewPart<NumberingDefinitionsPart>();
            xml = new XDocument(new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName)));
        }
        else
        {
            xml = Docx.Load(part);
        }
        var root = xml.Root!;
        abstractId = root.Elements(W + "abstractNum").Select(a => (int?)a.Attribute(W + "abstractNumId") ?? 0).DefaultIfEmpty(0).Max() + 1;
        numId = root.Elements(W + "num").Select(n => (int?)n.Attribute(W + "numId") ?? 0).DefaultIfEmpty(0).Max() + 1;

        var levels = Enumerable.Range(0, 9).Select(i =>
        {
            var styleId = styleForLevel.GetValueOrDefault(i + 1);
            // Number size matches the heading: without it a heading whose paragraph mark carries
            // direct smaller formatting shows a shrunken number (a client report, 2026-10-15).
            var size = styleId is null ? null : FontSize(styles, chain, styleId);
            return new XElement(W + "lvl", new XAttribute(W + "ilvl", i),
                new XElement(W + "start", new XAttribute(W + "val", 1)),
                new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
                styleId is null ? null : new XElement(W + "pStyle", new XAttribute(W + "val", NumberedId(styleId))),
                new XElement(W + "suff", new XAttribute(W + "val", "space")),
                new XElement(W + "lvlText", new XAttribute(W + "val", i == 0 ? "%1.0" : string.Join(".", Enumerable.Range(1, i + 1).Select(k => $"%{k}")))),
                new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                new XElement(W + "pPr", new XElement(W + "ind", new XAttribute(W + "left", 0), new XAttribute(W + "firstLine", 0))),
                size is null ? null : new XElement(W + "rPr",
                    new XElement(W + "sz", new XAttribute(W + "val", size)),
                    new XElement(W + "szCs", new XAttribute(W + "val", size))));
        });
        var abstractNum = new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", abstractId),
            new XElement(W + "multiLevelType", new XAttribute(W + "val", "multilevel")),
            new XElement(W + "name", new XAttribute(W + "val", "Numbered headings")),
            levels);
        var num = new XElement(W + "num", new XAttribute(W + "numId", numId),
            new XElement(W + "abstractNumId", new XAttribute(W + "val", abstractId)));

        // CT_Numbering order: numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?
        if (root.Elements(W + "abstractNum").LastOrDefault() is { } lastAbstract) lastAbstract.AddAfterSelf(abstractNum);
        else if (root.Elements(W + "numPicBullet").LastOrDefault() is { } lastPic) lastPic.AddAfterSelf(abstractNum);
        else root.AddFirst(abstractNum);
        if (root.Elements(W + "num").LastOrDefault() is { } lastNum) lastNum.AddAfterSelf(num);
        else abstractNum.AddAfterSelf(num);
        return new NumberingXml(part, xml);
    }

    /// <summary>The style's font size in half-points, following basedOn, then docDefaults.</summary>
    private static string? FontSize(XDocument styles, StyleChain chain, string styleId) =>
        chain.Chain(styleId).Select(s => (string?)s.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val")).FirstOrDefault(sz => sz is not null)
        ?? (string?)styles.Root!.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val");
}
