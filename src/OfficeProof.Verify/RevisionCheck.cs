using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// What a redline declares it changes on purpose outside tracked changes. Everything else that
/// Reject All cannot undo is an error.
/// </summary>
public sealed record VerifyOptions
{
    /// <summary>
    /// Style ids (w:styleId, e.g. "TOC1") whose existing definitions the redline intends to
    /// change or remove. Any other changed or removed existing style is rule
    /// <c>definition-changed</c> (Error); these are <c>untracked-change</c> warnings, and
    /// <see cref="WordCheck"/> gives the original their redline definitions before comparing.
    /// </summary>
    public IReadOnlySet<string> IntendedStyles { get; init; } = new HashSet<string>(StringComparer.Ordinal);
}

/// <summary>
/// Checks that a tracked-change redline does what it claims against the document it was made
/// from: rejecting every change gives the original back, accepting gives sound references and
/// the intended heading numbers, and nothing changed silently outside the tracked story.
/// </summary>
public static class RevisionCheck
{
    private static readonly XNamespace W = Ns.W;

    /// <summary>
    /// Runs every check. Rules: <c>reject-mismatch</c> (Error, any story: main, headers, footers,
    /// notes, text boxes), <c>comment-changed</c> (Error), <c>comment-added</c> (Info),
    /// <c>unverifiable-revision</c> (Error, fail-closed), <c>revision-in-styles</c> (Error),
    /// <c>broken-reference</c> (Error), <c>accept-outline</c> (Info), <c>definition-changed</c>
    /// (Error: an existing style not in <see cref="VerifyOptions.IntendedStyles"/>, the document
    /// defaults, an existing list, a layout setting or the theme changed), <c>untracked-change</c>
    /// (Warning: added styles and lists, declared style changes, review-state settings and other
    /// parts no story covers) and <c>schema</c> (Error, from <see cref="Docx.Validate(string)"/>
    /// on the redline).
    /// </summary>
    public static IReadOnlyList<Finding> Run(string originalPath, string redlinePath, VerifyOptions? options = null)
    {
        options ??= new VerifyOptions();
        var original = File.ReadAllBytes(originalPath);
        var redline = File.ReadAllBytes(redlinePath);
        var originalRejected = Outline.ViewBytes(original, OutlineView.RejectAll);
        var redlineRejected = Outline.ViewBytes(redline, OutlineView.RejectAll);
        var redlineAccepted = Outline.ViewBytes(redline, OutlineView.AcceptAll);

        var findings = new List<Finding>();
        var originalView = StoryView.Load(originalRejected);
        findings.AddRange(StoryComparison.Compare(originalView, StoryView.Load(redlineRejected, originalView)));
        findings.AddRange(RevisionAllowlist.Check(original, "original"));
        findings.AddRange(RevisionAllowlist.Check(redline, "redline"));
        findings.AddRange(RevisionsInDefinitions(redline));
        findings.AddRange(BrokenReferences(redlineAccepted));
        findings.Add(AcceptOutline(Outline.ParagraphsOfView(redlineAccepted)));
        findings.AddRange(PartDiff.Changes(original, redline, options.IntendedStyles));
        findings.AddRange(Docx.Validate(redlinePath));
        return findings;
    }

    // 2026-10-15, v2 client redline: a w:pPrChange on the Head1 style definition survived Word's Reject All,
    // so the restored typed numbers sat next to the style's list numbers ("2.0 1.0 Summary").
    // Word never undoes revisions inside styles.xml or numbering.xml.
    internal static IEnumerable<Finding> RevisionsInDefinitions(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var main = doc.MainDocumentPart;
        OpenXmlPart?[] parts = [main?.StyleDefinitionsPart, main?.StylesWithEffectsPart, main?.NumberingDefinitionsPart];
        foreach (var part in parts.OfType<OpenXmlPart>())
        {
            var root = Docx.Load(part).Root;
            if (root is null) continue;
            foreach (var group in Wml.Revisions(root).GroupBy(Owner))
            {
                var kinds = string.Join(", ", group.Select(r => "w:" + r.Name.LocalName).Distinct());
                yield return new Finding("revision-in-styles", Severity.Error,
                    $"tracked change ({kinds}) inside {group.Key}: Word's Reject All does not undo revisions in style or " +
                    "numbering definitions (2026-10-15, v2 client redline: Reject All produced '2.0 1.0 Summary'); " +
                    "leave the definition untouched and track the change on the paragraphs",
                    part.Uri.OriginalString, group.Key);
            }
        }

        static string Owner(XElement revision)
        {
            var owner = revision.Ancestors().FirstOrDefault(a => a.Name == W + "style" || a.Name == W + "abstractNum" || a.Name == W + "num");
            return owner is null ? "definitions"
                : owner.Name == W + "style" ? $"style {(string?)owner.Attribute(W + "styleId")}"
                : owner.Name == W + "abstractNum" ? $"abstractNum {(string?)owner.Attribute(W + "abstractNumId")}"
                : $"num {(string?)owner.Attribute(W + "numId")}";
        }
    }

    /// <summary>REF / PAGEREF / NOTEREF fields in the accepted document whose bookmark does not exist.</summary>
    internal static IEnumerable<Finding> BrokenReferences(byte[] acceptedPackage)
    {
        using var stream = new MemoryStream(acceptedPackage);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var main = doc.MainDocumentPart ?? throw new InvalidDataException("package has no main document part");
        var stories = Docx.StoryParts(main).Select(p => (Part: p, Root: Docx.Load(p).Root!)).ToList();
        var bookmarks = stories
            .SelectMany(s => s.Root.Descendants(W + "bookmarkStart"))
            .Select(b => (string?)b.Attribute(W + "name"))
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var (part, root) in stories)
        {
            var index = 0;
            foreach (var p in root.Descendants(W + "p"))
            {
                foreach (var instr in Wml.FieldInstructions(p))
                {
                    var name = ReferencedBookmark(instr);
                    // Word regenerates _Toc bookmarks and the TOC's PAGEREFs together on Update Table,
                    // so a stale TOC entry is not a broken reference.
                    if (name is null || name.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase) || bookmarks.Contains(name))
                        continue;
                    yield return new Finding("broken-reference", Severity.Error,
                        $"after Accept All, field {{ {instr} }} points at bookmark '{name}', which does not exist; Word shows 'Error! Reference source not found.' on update",
                        part.Uri.OriginalString, $"paragraph {index}: \"{ParagraphText.Excerpt(Wml.VisibleText(p))}\"");
                }
                index++;
            }
        }
    }

    /// <summary>The bookmark a REF, PAGEREF or NOTEREF instruction names, else null.</summary>
    internal static string? ReferencedBookmark(string instruction)
    {
        var tokens = Tokenize(instruction).ToList();
        if (tokens.Count < 2) return null;
        return tokens[0].ToUpperInvariant() is "REF" or "PAGEREF" or "NOTEREF" && !tokens[1].StartsWith('\\') ? tokens[1] : null;
    }

    private static IEnumerable<string> Tokenize(string instruction)
    {
        var i = 0;
        while (i < instruction.Length)
        {
            if (char.IsWhiteSpace(instruction[i])) { i++; continue; }
            if (instruction[i] == '"')
            {
                var end = instruction.IndexOf('"', i + 1);
                if (end < 0) end = instruction.Length;
                yield return instruction[(i + 1)..end];
                i = end + 1;
                continue;
            }
            var start = i;
            while (i < instruction.Length && !char.IsWhiteSpace(instruction[i])) i++;
            yield return instruction[start..i];
        }
    }

    internal static Finding AcceptOutline(IReadOnlyList<OutlineEntry> accepted)
    {
        var headings = accepted.Where(e => e.Level is not null).Select(Describe).ToList();
        var message = headings.Count == 0
            ? "Accept All: no headings"
            : "Accept All heading sequence: " + string.Join(" | ", headings);
        return new Finding("accept-outline", Severity.Info, message, "/word/document.xml");
    }

    internal static string Describe(OutlineEntry e) =>
        (e.ListNumber.Length > 0 ? e.ListNumber + " " : "") + ParagraphText.Excerpt(e.Text);
}

/// <summary>An item present only in the original, only after Reject All, or changed between them.</summary>
internal sealed record Difference<T>(T? Original, T? Rejected) where T : class;

/// <summary>Minimal LCS diff over item keys, reporting changed / lost / extra items.</summary>
internal static class SequenceDiff
{
    // A full LCS table above this many cells (~4 M ints, 16 MB) falls back to positional pairing.
    private const long MaxCells = 4_000_000;

    public static IReadOnlyList<Difference<T>> Compare<T, TKey>(IReadOnlyList<T> a, IReadOnlyList<T> b, Func<T, TKey> key)
        where T : class
    {
        var cmp = EqualityComparer<TKey>.Default;
        var ka = a.Select(key).ToArray();
        var kb = b.Select(key).ToArray();
        var start = 0;
        while (start < a.Count && start < b.Count && cmp.Equals(ka[start], kb[start])) start++;
        int endA = a.Count, endB = b.Count;
        while (endA > start && endB > start && cmp.Equals(ka[endA - 1], kb[endB - 1])) { endA--; endB--; }
        int n = endA - start, m = endB - start;
        var result = new List<Difference<T>>();
        if (n == 0 && m == 0) return result;
        if ((long)(n + 1) * (m + 1) > MaxCells)
        {
            for (var i = 0; start + i < endA || start + i < endB; i++)
            {
                var o = start + i < endA ? a[start + i] : null;
                var r = start + i < endB ? b[start + i] : null;
                if (o is not null && r is not null && cmp.Equals(ka[start + i], kb[start + i])) continue;
                result.Add(new Difference<T>(o, r));
            }
            return result;
        }

        var lcs = new int[n + 1, m + 1];
        for (var i = n - 1; i >= 0; i--)
            for (var j = m - 1; j >= 0; j--)
                lcs[i, j] = cmp.Equals(ka[start + i], kb[start + j])
                    ? lcs[i + 1, j + 1] + 1
                    : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);

        // Walk the table into runs of removals/additions; pair them up as "changed" items.
        var lost = new List<T>();
        var extra = new List<T>();
        int x = 0, y = 0;
        while (x < n || y < m)
        {
            if (x < n && y < m && cmp.Equals(ka[start + x], kb[start + y]))
            {
                Flush();
                x++; y++;
            }
            else if (y < m && (x == n || lcs[x, y + 1] >= lcs[x + 1, y])) extra.Add(b[start + y++]);
            else lost.Add(a[start + x++]);
        }
        Flush();
        return result;

        void Flush()
        {
            var paired = Math.Min(lost.Count, extra.Count);
            for (var i = 0; i < paired; i++) result.Add(new Difference<T>(lost[i], extra[i]));
            foreach (var o in lost.Skip(paired)) result.Add(new Difference<T>(o, null));
            foreach (var r in extra.Skip(paired)) result.Add(new Difference<T>(null, r));
            lost.Clear();
            extra.Clear();
        }
    }
}
