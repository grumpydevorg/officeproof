using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// Finds package parts that changed between original and redline where no story comparison
/// looks. Story parts (main document, headers, footers, footnotes, endnotes, comments) and every
/// part their markup references (images, embedded objects) are verified by
/// <see cref="StoryComparison"/>. Nothing else can be tracked, and Word's Reject All keeps it.
/// <para>
/// The definitions every story is drawn with are gated (rule <c>definition-changed</c>, Error):
/// changing or removing an existing style (unless declared in
/// <see cref="VerifyOptions.IntendedStyles"/>), the document defaults, an existing list
/// (abstractNum, num, numPicBullet), a settings element that can change layout, or the theme.
/// Adding styles and lists, user-interface flags on a style, latent styles and the
/// review-state settings are <c>untracked-change</c> Warnings, as is every other changed part
/// (fonts, web settings, custom XML).
/// </para>
/// </summary>
internal static class PartDiff
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    // Parts a redline legitimately rewrites besides the stories: comment metadata and authors,
    // package plumbing and document properties.
    private static readonly HashSet<string> ExpectedParts = new(StringComparer.OrdinalIgnoreCase)
    {
        "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml", "word/people.xml", "[Content_Types].xml",
    };

    // Style children that only steer Word's user interface (style gallery, styles pane, editing
    // restrictions) and bookkeeping: they change nothing a reader sees (ECMA-376 Part 1, 17.7.4).
    private static readonly HashSet<XName> StyleInterfaceFlags =
        [W + "rsid", W + "semiHidden", W + "unhideWhenUsed", W + "qFormat", W + "uiPriority", W + "locked", W + "hidden"];

    // Settings that record review state and bookkeeping (revision tracking, proofing, zoom, ids,
    // protection, variables, the next shape id) rather than how the document is laid out. Every
    // other settings element (compatibility options, tab stops, even/odd headers, note
    // properties, hyphenation, ...) can change the printed pages and is gated.
    private static readonly HashSet<XName> ReviewSettings =
    [
        W + "rsids", W + "proofState", W + "zoom", W + "view", W + "removePersonalInformation", W + "removeDateAndTime",
        W + "trackRevisions", W + "doNotTrackMoves", W + "doNotTrackFormatting", W + "revisionView",
        W + "documentProtection", W + "writeProtection", W + "updateFields", W + "attachedTemplate", W + "docVars",
        W + "activeWritingStyle", W + "hdrShapeDefaults", W + "shapeDefaults", Ns.W14 + "docId", W15 + "docId",
    ];

    private enum Role { Styles, Numbering, Settings, Theme }

    /// <summary>One top-level child of a definition part: a style, list, setting or defaults block.</summary>
    internal sealed record Definition(string Key, string Kind, string? Id, XElement Element);

    public static IEnumerable<Finding> Changes(byte[] original, byte[] redline, IReadOnlySet<string> intendedStyles)
    {
        var before = DefinitionParts(original);
        var after = DefinitionParts(redline);
        var findings = new List<Finding>();
        foreach (var role in Enum.GetValues<Role>())
        {
            var x = before.GetValueOrDefault(role);
            var y = after.GetValueOrDefault(role);
            if (x is not null || y is not null) findings.AddRange(CompareDefinitions(role, x, y, intendedStyles));
        }
        var definitionParts = before.Values.Concat(after.Values).Select(p => p.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var covered = Relationships.CoveredParts(original).Union(Relationships.CoveredParts(redline), StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        using var a = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        using var b = new ZipArchive(new MemoryStream(redline), ZipArchiveMode.Read);
        var names = a.Entries.Select(e => e.FullName).Union(b.Entries.Select(e => e.FullName), StringComparer.OrdinalIgnoreCase)
            .Where(n => !Expected(n) && !covered.Contains(n) && !definitionParts.Contains(n))
            .Order(StringComparer.Ordinal);
        foreach (var name in names)
        {
            var x = Read(a, name);
            var y = Read(b, name);
            var part = "/" + name;
            if (x is null)
                findings.Add(Untracked($"{name} added by the redline (not a tracked change)", part));
            else if (y is null)
                findings.Add(Untracked($"{name} removed by the redline (not a tracked change)", part));
            else if (x.AsSpan().SequenceEqual(y))
                continue;
            else if (!IsXml(name))
                findings.Add(Untracked($"{name} content changed (not a tracked change)", part));
            else if (Parse(x) is var px && Parse(y) is var py && !XNode.DeepEquals(px, py))
                findings.Add(Untracked($"{name} changed outside tracked changes (not a tracked change): {Summary(Keyed(px), Keyed(py))}", part));
        }
        return findings;
    }

    private static Finding Untracked(string message, string part) => new("untracked-change", Severity.Warning, message, part);

    private static Finding Gate(string message, string part, string location) => new("definition-changed", Severity.Error, message, part, location);

    private sealed record PartBytes(string Name, byte[] Bytes);

    private static Dictionary<Role, PartBytes> DefinitionParts(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var result = new Dictionary<Role, PartBytes>();
        if (doc.MainDocumentPart is not { } main) return result;
        Add(Role.Styles, main.StyleDefinitionsPart);
        Add(Role.Numbering, main.NumberingDefinitionsPart);
        Add(Role.Settings, main.DocumentSettingsPart);
        Add(Role.Theme, main.ThemePart);
        return result;

        void Add(Role role, OpenXmlPart? part)
        {
            if (part is null) return;
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            using var m = new MemoryStream();
            s.CopyTo(m);
            result[role] = new PartBytes(part.Uri.OriginalString.TrimStart('/'), m.ToArray());
        }
    }

    private static IEnumerable<Finding> CompareDefinitions(Role role, PartBytes? before, PartBytes? after, IReadOnlySet<string> intendedStyles)
    {
        var name = (after ?? before)!.Name;
        var part = "/" + name;
        if (role == Role.Theme)
        {
            if (before is not null && after is not null && (before.Bytes.AsSpan().SequenceEqual(after.Bytes) || XNode.DeepEquals(Parse(before.Bytes), Parse(after.Bytes))))
                yield break;
            var what = before is null ? "added" : after is null ? "removed" : "changed";
            yield return Gate($"{name} {what} outside tracked changes: the theme supplies the fonts and colours of every theme-font and " +
                "theme-colour run and Reject All keeps the redline's, so the document will not look as it did; leave the theme untouched",
                part, "theme");
            yield break;
        }

        var x = before is null ? null : Parse(before.Bytes);
        var y = after is null ? null : Parse(after.Bytes);
        if (x is not null && y is not null && XNode.DeepEquals(x, y)) yield break;
        var old = x is null ? new Dictionary<string, Definition>() : Keyed(x);
        var now = y is null ? new Dictionary<string, Definition>() : Keyed(y);

        var noted = new List<(string Change, string Key, string Note)>();
        foreach (var (key, d) in old.Concat(now.Where(kv => !old.ContainsKey(kv.Key))))
        {
            var o = old.GetValueOrDefault(key);
            var n = now.GetValueOrDefault(key);
            if (o is not null && n is not null && XNode.DeepEquals(o.Element, n.Element)) continue;
            var change = o is null ? "added" : n is null ? "removed" : "changed";
            switch (role, d.Kind)
            {
                case (Role.Styles, "style") when o is null:
                    noted.Add(("added", key, ""));
                    break;
                case (Role.Styles, "style") when d.Id is not null && intendedStyles.Contains(d.Id):
                    noted.Add((change, key, "declared intended"));
                    break;
                case (Role.Styles, "style") when n is not null && XNode.DeepEquals(WithoutInterfaceFlags(o!.Element), WithoutInterfaceFlags(n.Element)):
                    noted.Add((change, key, "user-interface flags only"));
                    break;
                case (Role.Styles, "style"):
                    var label = $"style {d.Id}{StyleName(o!.Element)}";
                    yield return Gate(n is null
                        ? $"{name}: existing {label} removed outside tracked changes: paragraphs and runs in it fall back to the defaults, and " +
                          $"Reject All cannot bring it back; keep the style, or declare the removal intended (--allow-style {d.Id})"
                        : $"{name}: existing {label} changed outside tracked changes: Word cannot track a style definition and its Reject All keeps " +
                          $"the redline's, so everything in this style looks different from the original; leave the definition untouched, or declare " +
                          $"the change intended (--allow-style {d.Id})",
                        part, key);
                    break;
                case (Role.Styles, "docDefaults"):
                    yield return Gate($"{name}: document defaults (w:docDefaults) {change} outside tracked changes: they are the base formatting of " +
                        "every paragraph and run, and Reject All keeps the redline's; leave them untouched", part, key);
                    break;
                case (Role.Styles, "latentStyles"):
                    noted.Add((change, key, "user interface only"));
                    break;
                case (Role.Numbering, "abstractNum" or "num" or "numPicBullet") when o is null:
                    noted.Add(("added", key, ""));
                    break;
                case (Role.Numbering, "abstractNum" or "num" or "numPicBullet"):
                    yield return Gate($"{name}: existing {key} {change} outside tracked changes: every paragraph in that list numbers, indents or " +
                        "bullets differently and Reject All keeps the redline's definition; add a new list instead of changing this one", part, key);
                    break;
                case (Role.Numbering, "numIdMacAtCleanup"):
                    noted.Add((change, key, ""));
                    break;
                case (Role.Settings, _) when ReviewSettings.Contains(d.Element.Name):
                    noted.Add((change, key, "review state"));
                    break;
                default:
                    yield return Gate($"{name}: {Describe(role)} {key} {change} outside tracked changes: Word cannot track it and Reject All keeps " +
                        "the redline's, so it can change how the document looks; leave it untouched", part, key);
                    break;
            }
        }
        if (x is not null && y is not null && !XNode.DeepEquals(new XElement(x.Name, x.Attributes()), new XElement(y.Name, y.Attributes())))
            noted.Add(("changed", "root attributes", ""));
        if (noted.Count > 0)
        {
            var origin = before is null ? " (part added by the redline)" : after is null ? " (part removed by the redline)" : "";
            var summary = noted.GroupBy(c => (c.Change, c.Note))
                .Select(g => $"{g.Key.Change} {string.Join(", ", g.Select(c => c.Key))}{(g.Key.Note.Length > 0 ? $" ({g.Key.Note})" : "")}");
            yield return Untracked($"{name} changed outside tracked changes{origin} ({Why(role)}): {string.Join("; ", summary)}", part);
        }
    }

    private static string Describe(Role role) => role switch
    {
        Role.Styles => "style-part element",
        Role.Numbering => "numbering element",
        _ => "setting",
    };

    private static string Why(Role role) => role switch
    {
        Role.Styles => "style definitions cannot be tracked and survive Reject All",
        Role.Numbering => "numbering definitions cannot be tracked and survive Reject All",
        _ => "settings cannot be tracked and survive Reject All",
    };

    private static string StyleName(XElement style) =>
        (string?)style.Element(W + "name")?.Attribute(W + "val") is { } n && n != (string?)style.Attribute(W + "styleId") ? $" (\"{n}\")" : "";

    private static XElement WithoutInterfaceFlags(XElement style) =>
        new(style.Name, style.Attributes(), style.Elements().Where(e => !StyleInterfaceFlags.Contains(e.Name)));

    private static string Summary(Dictionary<string, Definition> x, Dictionary<string, Definition> y)
    {
        var added = y.Keys.Where(k => !x.ContainsKey(k)).ToList();
        var removed = x.Keys.Where(k => !y.ContainsKey(k)).ToList();
        var changed = x.Keys.Where(k => y.TryGetValue(k, out var e) && !XNode.DeepEquals(x[k].Element, e.Element)).ToList();
        var parts = new List<string>();
        if (added.Count > 0) parts.Add("added " + string.Join(", ", added));
        if (removed.Count > 0) parts.Add("removed " + string.Join(", ", removed));
        if (changed.Count > 0) parts.Add("changed " + string.Join(", ", changed));
        if (parts.Count == 0) parts.Add("root attributes changed");
        return string.Join("; ", parts);
    }

    /// <summary>
    /// The top-level children of a part by identity: styles by styleId, lists by abstractNumId /
    /// numId / numPicBulletId, anything else by element name (prefixed outside the main namespace).
    /// Works on canonical and raw elements alike.
    /// </summary>
    internal static Dictionary<string, Definition> Keyed(XElement root)
    {
        var result = new Dictionary<string, Definition>(StringComparer.Ordinal);
        foreach (var e in root.Elements())
        {
            var kind = e.Name.Namespace == W ? e.Name.LocalName : $"{Prefix(e.Name.Namespace)}:{e.Name.LocalName}";
            var id = (string?)e.Attribute(W + "styleId") ?? (string?)e.Attribute(W + "abstractNumId") ?? (string?)e.Attribute(W + "numId")
                ?? (string?)e.Attribute(W + "numPicBulletId");
            var key = id is null ? kind : $"{kind} {id}";
            var unique = key;
            for (var n = 2; result.ContainsKey(unique); n++) unique = $"{key} #{n}";
            result.Add(unique, new Definition(unique, kind, id, e));
        }
        return result;
    }

    private static string Prefix(XNamespace ns) =>
        ns == Ns.W14 ? "w14" : ns == W15 ? "w15" : ns == Ns.R ? "r" : ns.NamespaceName.Length == 0 ? "" : ns.NamespaceName;

    /// <summary>
    /// <paramref name="original"/>'s definition part with <paramref name="redline"/>'s additions
    /// and the definitions <paramref name="replace"/> selects taken from the redline; everything
    /// else stays the original's. New definitions go after their predecessor in the redline, so
    /// the schema order (abstractNum before num, docDefaults before styles) is kept.
    /// </summary>
    internal static XElement Align(XElement original, XElement redline, Func<Definition, bool> replace)
    {
        var result = new XElement(original);
        var present = Keyed(result).ToDictionary(kv => kv.Key, kv => kv.Value.Element, StringComparer.Ordinal);
        XElement? anchor = null;
        foreach (var d in Keyed(redline).Values)
        {
            if (present.TryGetValue(d.Key, out var existing))
            {
                if (replace(d))
                {
                    var copy = new XElement(d.Element);
                    existing.ReplaceWith(copy);
                    existing = present[d.Key] = copy;
                }
                anchor = existing;
                continue;
            }
            var added = new XElement(d.Element);
            if (anchor is null) result.AddFirst(added);
            else anchor.AddAfterSelf(added);
            anchor = present[d.Key] = added;
        }
        return result;
    }

    private static bool Expected(string name) =>
        ExpectedParts.Contains(name)
        || name.StartsWith("docProps/", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith('/');

    private static bool IsXml(string name) =>
        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    private static byte[]? Read(ZipArchive zip, string name)
    {
        var entry = zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, name, StringComparison.OrdinalIgnoreCase));
        if (entry is null) return null;
        using var s = entry.Open();
        using var m = new MemoryStream();
        s.CopyTo(m);
        return m.ToArray();
    }

    /// <summary>
    /// Parses a part; the canonical form drops namespace declarations, comments and processing
    /// instructions and sorts attributes, so a re-serialised but unchanged part compares equal.
    /// </summary>
    private static XElement Parse(byte[] bytes)
    {
        using var reader = XmlReader.Create(new MemoryStream(bytes), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XDocument.Load(reader).Root ?? throw new InvalidDataException("part has no root element");
        return Canonical(root);
    }

    private static XElement Canonical(XElement e) => new(e.Name,
        e.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)
            .Select(a => new XAttribute(a.Name, a.Value)),
        e.Nodes().Select(n => n switch
        {
            XElement c => (object?)Canonical(c),
            XText t => new XText(t.Value),
            _ => null,
        }).OfType<object>());
}
