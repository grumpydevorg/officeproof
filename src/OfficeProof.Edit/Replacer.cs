using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Edit;

/// <summary>
/// Untracked find/replace across the story parts of our own documents, on Clippit's
/// <see cref="OpenXmlRegex"/>: a paragraph is searched as one string, so text Word split over
/// several runs (spell-check, rsids, a bold word) still matches. The replacement takes the run
/// properties of the first matched character; the other matched characters' runs are removed, so
/// a match across differently formatted runs comes out in the first run's formatting. Matches
/// never cross paragraphs. Leading and trailing spaces in the replacement are kept
/// (<c>xml:space="preserve"</c>); a tab becomes <c>w:tab</c>, a carriage return <c>w:br</c>.
/// <para>
/// A matched paragraph is rewritten by Clippit: runs with identical properties are merged and run
/// rsid attributes dropped (paragraph attributes are kept). Paragraphs without a match and parts
/// without a match keep their bytes, as does every other part of the package.
/// </para>
/// <para>
/// Refused, writing nothing: a document with tracked changes (an untracked edit would merge
/// silently into someone's redline; that is <c>redline</c>'s job), a pair with no match (typo
/// protection), a total that differs from <see cref="ReplaceOptions.Expect"/>, and a match that
/// covers anything but text (a field code, drawing, note or comment reference), which the
/// replacement would delete.
/// </para>
/// </summary>
public static class Replacer
{
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    // Run children Clippit maps to characters it turns back into the same elements
    // (UnicodeMapper.RunToString / CharToRunChild); every other run child is not text.
    private static readonly HashSet<XName> TextLike =
        [W + "tab", W + "br", W + "cr", W + "noBreakHyphen", W + "softHyphen", W + "sym"];

    /// <summary>
    /// Applies <paramref name="pairs"/> in order to a copy of <paramref name="input"/> at
    /// <paramref name="output"/>; each pair searches the result of the pairs before it. Every pair
    /// is planned against the input first; any error leaves <paramref name="output"/> unwritten.
    /// </summary>
    public static ReplaceResult Apply(string input, string output, IReadOnlyList<ReplacePair> pairs, ReplaceOptions options)
    {
        ArgumentNullException.ThrowIfNull(pairs);
        ArgumentNullException.ThrowIfNull(options);
        if (pairs.Count == 0) throw new ArgumentException("nothing to replace: give at least one --find/--with pair", nameof(pairs));
        if (Path.GetFullPath(input) == Path.GetFullPath(output))
            throw new ArgumentException("output must differ from input; officeproof never edits in place", nameof(output));
        if ((options.Stories & (Stories.All | Stories.Comments)) == Stories.None)
            throw new ArgumentException("no story selected", nameof(options));
        var compiled = pairs.Select(p => Compile(p, options.Regex)).ToList();

        List<Story> stories;
        var errors = new List<Finding>();
        using (var doc = Docx.OpenRead(input))
        {
            var main = doc.MainDocumentPart ?? throw new InvalidDataException($"{input} has no main document part");
            errors.AddRange(RevisionErrors(main));
            if (errors.Count > 0) return Refused(errors);
            stories = [.. Selected(main, options.Stories).Select(p => new Story(p.Uri.ToString(), Docx.Load(p)))];
        }

        var results = new List<PairResult>();
        foreach (var (pair, regex, replacement) in compiled)
        {
            var parts = new List<PartMatches>();
            foreach (var story in stories)
            {
                var n = ReplaceInStory(story, pair, regex, replacement, errors);
                if (n > 0) parts.Add(new PartMatches(story.Uri, n));
            }
            var result = new PairResult(pair, parts);
            if (result.Total == 0)
                errors.Add(new Finding("replace-no-match", Severity.Error,
                    $"\"{pair.Find}\" matches nothing in the selected stories ({options.Stories}){(results.Count > 0 ? " after the pairs before it" : "")}"));
            results.Add(result);
        }
        var total = results.Sum(r => r.Total);
        if (options.Expect is { } expect && total != expect)
            errors.Add(new Finding("replace-expect", Severity.Error,
                $"{total} match(es) in total, --expect {expect}: " + string.Join("; ", results.Select(r => $"\"{r.Pair.Find}\" {r.Total}"))));
        if (errors.Count > 0) return new ReplaceResult(results, errors, [], []);

        var edited = stories.Where(s => s.Edited).ToList();
        // Copy the bytes, not the file: File.Copy also copies a read-only input's mode, and the
        // copy then cannot be opened for editing (2026-10-15 smoke on a read-only copy of minutes).
        using (var source = File.OpenRead(input))
        using (var target = File.Create(output))
            source.CopyTo(target);
        using (var doc = WordprocessingDocument.Open(output, isEditable: true))
        {
            var byUri = Docx.StoryParts(doc.MainDocumentPart!).ToDictionary(p => p.Uri.ToString());
            foreach (var story in edited) Docx.Save(byUri[story.Uri], story.Xml);
        }
        return new ReplaceResult(results, [], [.. edited.Select(s => s.Uri)], NewSchemaErrors(input, output));
    }

    private sealed class Story(string uri, XDocument xml)
    {
        public string Uri { get; } = uri;
        public XDocument Xml { get; } = xml;
        public bool Edited { get; set; }
    }

    private static ReplaceResult Refused(List<Finding> errors) => new([], errors, [], []);

    private static (ReplacePair Pair, Regex Regex, string Replacement) Compile(ReplacePair pair, bool isRegex)
    {
        ArgumentNullException.ThrowIfNull(pair);
        if (string.IsNullOrEmpty(pair.Find)) throw new ArgumentException("--find must not be empty");
        // Only the characters CharToRunChild turns into elements: tab, line break (\r), page break (\f).
        foreach (var c in pair.With)
        {
            if (char.IsControl(c) && c is not ('\t' or '\r' or '\f'))
                throw new ArgumentException($"--with \"{Visible(pair.With)}\" contains control character U+{(int)c:X4}; only tab, \\r (line break) and \\f (page break) have a Word form");
        }
        Regex regex;
        try
        {
            regex = new Regex(isRegex ? pair.Find : Regex.Escape(pair.Find), RegexOptions.CultureInvariant);
        }
        catch (RegexParseException e)
        {
            throw new ArgumentException($"--find \"{pair.Find}\" is not a valid regular expression: {e.Message}", e);
        }
        if (regex.IsMatch(string.Empty))
            throw new ArgumentException($"--find \"{pair.Find}\" matches the empty string");
        // Literal mode: '$' in the replacement is text, not a group reference.
        return (pair, regex, isRegex ? pair.With : pair.With.Replace("$", "$$", StringComparison.Ordinal));
    }

    private static IEnumerable<OpenXmlPart> Selected(MainDocumentPart main, Stories stories) =>
        Docx.StoryParts(main).Where(p => p switch
        {
            MainDocumentPart => stories.HasFlag(Stories.Main),
            HeaderPart => stories.HasFlag(Stories.Headers),
            FooterPart => stories.HasFlag(Stories.Footers),
            FootnotesPart or EndnotesPart => stories.HasFlag(Stories.Notes),
            WordprocessingCommentsPart => stories.HasFlag(Stories.Comments),
            _ => false,
        });

    /// <summary>Tracked-change markup (Clippit's full list, cell and move ranges included) in any story, styles or numbering.</summary>
    private static IEnumerable<Finding> RevisionErrors(MainDocumentPart main)
    {
        var parts = Docx.StoryParts(main).Concat(new OpenXmlPart?[] { main.StyleDefinitionsPart, main.NumberingDefinitionsPart }.OfType<OpenXmlPart>());
        foreach (var part in parts)
        {
            var root = Docx.Load(part).Root;
            if (root is null) continue;
            var kinds = root.Descendants().Where(e => RevisionProcessor.TrackedRevisionsElements.Contains(e.Name))
                .GroupBy(e => e.Name.LocalName).OrderByDescending(g => g.Count()).ToList();
            if (kinds.Count > 0)
                yield return new Finding("replace-revisions", Severity.Error,
                    $"{kinds.Sum(g => g.Count())} tracked change(s) ({string.Join(", ", kinds.Select(g => $"{g.Count()} {g.Key}"))}); "
                    + "an untracked replace would merge silently into the redline. Accept or reject them first, "
                    + "or make the change as a tracked edit with `officeproof redline`",
                    part.Uri.ToString());
        }
    }

    /// <summary>Replaces every match in <paramref name="story"/>; returns the matches counted (text-box fallbacks excluded).</summary>
    private static int ReplaceInStory(Story story, ReplacePair pair, Regex regex, string replacement, List<Finding> errors)
    {
        var root = story.Xml.Root;
        if (root is null) return 0;
        // Innermost first: a paragraph in a text box is done before the paragraph holding the
        // box, which Clippit then copies whole (it never recurses into a matched paragraph's runs).
        var paragraphs = root.Descendants(W + "p").OrderByDescending(p => p.Ancestors().Count()).ToList();
        var counted = 0;
        foreach (var paragraph in paragraphs)
        {
            var units = Units(paragraph);
            var text = string.Concat(units.Select(u => u.Ch));
            var matches = regex.Matches(text).Where(m => m.Length > 0).ToList();
            if (matches.Count == 0) continue;
            // Word writes a DrawingML text box twice (mc:Choice and the VML mc:Fallback); both are
            // replaced so they stay the same, but the reader sees one. A refused match still counts
            // (it is not a typo), so its pair does not also report no match.
            if (!paragraph.Ancestors(Mc + "Fallback").Any()) counted += matches.Count;

            var refused = false;
            foreach (var m in matches)
            {
                var foreign = units.Skip(m.Index).Take(m.Length).Where(u => u.NonText is not null).Select(u => u.NonText!).Distinct().ToList();
                if (foreign.Count == 0) continue;
                errors.Add(new Finding("replace-non-text", Severity.Error,
                    $"\"{pair.Find}\" matches \"{Visible(m.Value)}\", which covers {string.Join(", ", foreign.Select(f => "w:" + f))} "
                    + "(a field code, drawing, note or comment reference) that the replacement would delete; narrow the find to the text",
                    story.Uri, Excerpt(text, m)));
                refused = true;
            }
            if (refused) continue;

            // Clippit computes the same paragraph string after splitting runs; a match it sees
            // differently (or not at all) would edit something other than what was planned.
            var expected = matches.Select(m => (m.Index, m.Value)).ToHashSet();
            var seen = 0;
            var disagreed = false;
            OpenXmlRegex.Replace([paragraph], regex, replacement, (_, m) =>
            {
                if (expected.Contains((m.Index, m.Value)))
                {
                    seen++;
                    return true;
                }
                disagreed = true;
                return false;
            }, coalesceContent: true);
            if (disagreed || seen != matches.Count)
            {
                errors.Add(new Finding("replace-unaligned", Severity.Error,
                    $"Clippit replaced {seen} of {matches.Count} match(es) of \"{pair.Find}\"{(disagreed ? " and saw others" : "")} in one paragraph; refusing a partial edit",
                    story.Uri, Excerpt(text, matches[0])));
                continue;
            }
            story.Edited = true;
        }
        return counted;
    }

    /// <summary>One character of the paragraph string; <paramref name="NonText"/> names the run child behind a non-text character.</summary>
    private readonly record struct Unit(char Ch, string? NonText);

    /// <summary>
    /// The paragraph string OpenXmlRegex builds after splitting runs (one unit per character of
    /// w:t, one per other run child, text boxes excluded), with where each character came from.
    /// </summary>
    private static List<Unit> Units(XElement paragraph)
    {
        var units = new List<Unit>();
        foreach (var run in paragraph.DescendantsTrimmed(W + "txbxContent").Where(d => d.Name == W + "r" && d.Parent?.Name != W + "del"))
        {
            foreach (var child in run.Elements())
            {
                if (child.Name == W + "rPr" || child.Name == W + "lastRenderedPageBreak") continue;
                if (child.Name == W + "t")
                {
                    foreach (var c in child.Value) units.Add(new Unit(c, null));
                    continue;
                }
                var s = UnicodeMapper.RunToString(child);
                if (s.Length == 0) s = " ";
                var nonText = TextLike.Contains(child.Name) && !s.Contains(UnicodeMapper.StartOfHeading, StringComparison.Ordinal) ? null : child.Name.LocalName;
                foreach (var c in s) units.Add(new Unit(c, nonText));
            }
        }
        return units;
    }

    /// <summary>Schema errors in the output beyond the input's own, compared by part and message (locations shift when runs merge).</summary>
    private static List<Finding> NewSchemaErrors(string input, string output)
    {
        var before = Docx.Validate(input).GroupBy(f => (f.Part, f.Message)).ToDictionary(g => g.Key, g => g.Count());
        var fresh = new List<Finding>();
        foreach (var g in Docx.Validate(output).GroupBy(f => (f.Part, f.Message)))
            fresh.AddRange(g.Skip(before.GetValueOrDefault(g.Key)));
        return fresh;
    }

    private static string Excerpt(string text, Match m)
    {
        var start = Math.Max(0, m.Index - 30);
        var end = Math.Min(text.Length, m.Index + m.Length + 30);
        return $"\"{(start > 0 ? "…" : "")}{Visible(text[start..end])}{(end < text.Length ? "…" : "")}\"";
    }

    private static string Visible(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (c == '\t') sb.Append("\\t");
            else if (c == '\r') sb.Append("\\r");
            else if (c == '\f') sb.Append("\\f");
            else if (char.IsControl(c)) sb.Append('¤');
            else sb.Append(c);
        }
        return sb.ToString();
    }
}
