using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Redline;

/// <summary>
/// Tracked-change editing of one document's main story. Owns the id allocation for
/// revisions, bookmarks and comments so every w:id it writes is unique in the package.
/// Text offsets are into <see cref="Text"/> (Word's display text with revisions shown).
/// Edits are made in memory; call <see cref="Save"/> to write document.xml and comments.xml.
/// If an operation throws, the in-memory document may be partly edited: discard it.
/// </summary>
public sealed partial class TrackedEditor
{
    private static readonly XNamespace W = Ns.W;
    private readonly XDocument _main;
    private readonly Dictionary<XElement, int> _fieldDepthAtStart = [];
    private XDocument? _comments;
    private int _nextId;
    private long _nextRef;

    public TrackedEditor(WordprocessingDocument document, RevisionAuthor author)
    {
        ArgumentNullException.ThrowIfNull(document);
        Document = document;
        Author = author;
        var main = document.MainDocumentPart ?? throw new ArgumentException("document has no main part", nameof(document));
        _main = Docx.Load(main);
        if (main.WordprocessingCommentsPart is { } cp) _comments = Docx.Load(cp);

        // 2026-10-15 (a client redline): w:id must be unique across ins/del/pPrChange/comments; bookmarks
        // have their own id space but we keep every id distinct anyway, across all stories.
        var maxId = 0;
        foreach (var part in Docx.StoryParts(main))
        {
            var xml = part == main ? _main : Docx.Load(part);
            foreach (var a in xml.Descendants().Attributes(W + "id"))
                if (int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) && v > maxId) maxId = v;
        }
        _nextId = maxId + 1;

        var maxRef = 90000000L;
        foreach (var b in _main.Descendants(W + "bookmarkStart"))
        {
            var m = RefName().Match((string?)b.Attribute(W + "name") ?? "");
            if (m.Success && long.TryParse(m.Groups[1].Value, CultureInfo.InvariantCulture, out var v) && v > maxRef) maxRef = v;
        }
        _nextRef = maxRef + 1;

        // Complex fields can span paragraphs (a TOC's cached lines). Record the field depth at
        // each paragraph start once; our edits only add or delete balanced fields inside one
        // paragraph, so these depths never change.
        var depth = 0;
        foreach (var e in Body.Descendants())
        {
            if (InFallback(e)) continue;
            if (e.Name == W + "p") _fieldDepthAtStart[e] = depth;
            else if (e.Name == W + "fldChar")
            {
                var kind = (string?)e.Attribute(W + "fldCharType");
                if (kind == "begin") depth++;
                else if (kind == "end" && depth > 0) depth--;
            }
        }
    }

    /// <summary>The package being edited (styles and numbering are edited directly on it).</summary>
    public WordprocessingDocument Document { get; }

    /// <summary>Author and date of every change this editor writes.</summary>
    public RevisionAuthor Author { get; }

    /// <summary>The w:body of the main document.</summary>
    public XElement Body => _main.Root!.Element(W + "body")!;

    /// <summary>
    /// Every paragraph in the body (including table cells and text boxes), in document order.
    /// mc:Fallback copies of text boxes are excluded, as in <see cref="FieldWalker"/>: Word shows
    /// the mc:Choice content, and editing both copies would double every change.
    /// </summary>
    public IEnumerable<XElement> Paragraphs => Body.Descendants(W + "p").Where(p => !InFallback(p));

    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static bool InFallback(XElement e) => e.Ancestors(Mc + "Fallback").Any();

    /// <summary>Display text of <paramref name="paragraph"/>; offsets for all edit methods index into it.</summary>
    public static string Text(XElement paragraph) => Wml.VisibleText(paragraph);

    /// <summary>True when the paragraph begins inside a complex field opened in an earlier paragraph (e.g. a TOC line).</summary>
    public bool StartsInsideField(XElement paragraph) =>
        _fieldDepthAtStart.TryGetValue(paragraph, out var d) && d > 0;

    /// <summary>The single body paragraph whose text contains <paramref name="snippet"/>; throws naming the snippet otherwise.</summary>
    public XElement FindParagraph(string snippet)
    {
        ArgumentException.ThrowIfNullOrEmpty(snippet);
        var hits = Paragraphs.Where(p => Text(p).Contains(snippet, StringComparison.Ordinal)).Take(2).ToList();
        return hits.Count switch
        {
            1 => hits[0],
            0 => throw new InvalidOperationException($"snippet matches no paragraph: \"{snippet}\""),
            _ => throw new InvalidOperationException($"snippet matches more than one paragraph: \"{snippet}\""),
        };
    }

    /// <summary>A fresh bookmark name in Word's cross-reference form, "_Ref" + digits.</summary>
    public string NewRefBookmarkName() => "_Ref" + (_nextRef++).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// True when [start, end) is ordinary text: not inside any insertion, field result or
    /// field that spans paragraphs. Scanners use it to skip text a redline already produced.
    /// </summary>
    public bool IsPlainText(XElement paragraph, int start, int end) =>
        Map(paragraph).Segments.Where(s => s.Length > 0 && s.Start < end && s.Start + s.Length > start)
            .All(s => s.Insertion is null && s.Field is null && !s.Foreign);

    /// <summary>
    /// Splits runs (keeping their rPr) so that [start, end) is covered exactly by whole runs and
    /// returns them in order. For edits the range must not touch another revision or part of a
    /// field; a field whose whole result lies in the range is included entirely.
    /// </summary>
    public IReadOnlyList<XElement> IsolateRuns(XElement paragraph, int start, int end) => Isolate(paragraph, start, end, forEdit: false);

    /// <summary>
    /// Tracked replacement of [start, end) with <paramref name="content"/>: the old runs move into
    /// w:del (w:t → w:delText, w:instrText → w:delInstrText) and the new content goes into one
    /// w:ins after them, formatted like the first replaced run. start == end inserts only;
    /// empty content deletes only.
    /// </summary>
    public TrackedReplacement Replace(XElement paragraph, int start, int end, IReadOnlyList<Inline> content)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        ArgumentNullException.ThrowIfNull(content);
        var text = Text(paragraph);
        if (start < 0 || end < start || end > text.Length)
            throw new ArgumentOutOfRangeException(nameof(end), $"range [{start}, {end}) outside paragraph of length {text.Length}");

        List<XElement> deletions = [];
        XElement? rPr;
        Action<XElement> place;
        if (start == end)
        {
            (rPr, place) = InsertionPoint(paragraph, start);
        }
        else
        {
            var runs = Isolate(paragraph, start, end, forEdit: true);
            rPr = CleanRPr(runs[0]);
            foreach (var group in AdjacentGroups(runs))
            {
                var del = Tracked("del");
                group[0].AddBeforeSelf(del);
                foreach (var r in group)
                {
                    r.Remove();
                    foreach (var t in r.Elements(W + "t").ToList()) t.ReplaceWith(Preserved(W + "delText", t.Value));
                    foreach (var i in r.Elements(W + "instrText").ToList()) i.ReplaceWith(Preserved(W + "delInstrText", i.Value));
                    del.Add(r);
                }
                deletions.Add(del);
            }
            var last = deletions[^1];
            place = ins => last.AddAfterSelf(ins);
        }

        XElement? insertion = null;
        if (content.Count > 0 && content.Any(c => c is not TextInline { Value.Length: 0 }))
        {
            insertion = Tracked("ins");
            foreach (var item in content)
            {
                switch (item)
                {
                    case TextInline t:
                        foreach (var child in TextChildren(t.Value)) insertion.Add(NewRun(rPr, child));
                        break;
                    case FieldInline f:
                        insertion.Add(
                            NewRun(rPr, FldChar("begin")),
                            NewRun(rPr, Preserved(W + "instrText", $" {f.Code} ")),
                            NewRun(rPr, FldChar("separate")),
                            NewRun(rPr, Preserved(W + "t", f.CachedResult)),
                            NewRun(rPr, FldChar("end")));
                        break;
                    default:
                        throw new ArgumentException($"unknown inline {item.GetType().Name}", nameof(content));
                }
            }
            place(insertion);
        }
        return new TrackedReplacement(deletions, insertion);
    }

    /// <summary>
    /// <see cref="Replace"/> with plain text, tracking only the span that actually differs:
    /// the common prefix and suffix of old and new text (snapped to word boundaries) stay untouched.
    /// </summary>
    public TrackedReplacement ReplaceText(XElement paragraph, int start, int end, string newText)
    {
        ArgumentNullException.ThrowIfNull(newText);
        var old = Text(paragraph)[start..end];
        var (prefix, suffix) = CommonEnds(old, newText);
        var changed = newText[prefix..(newText.Length - suffix)];
        return Replace(paragraph, start + prefix, end - suffix, changed.Length == 0 ? [] : [Inline.Text(changed)]);
    }

    /// <summary>
    /// Adds a bookmark from before <paramref name="startNode"/> to after <paramref name="endNode"/>.
    /// Passing the paragraph itself as start/end bookmarks from its first to its last content.
    /// Returns the bookmark id.
    /// </summary>
    public int Bookmark(XElement paragraph, XElement startNode, XElement endNode, string name)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        ArgumentException.ThrowIfNullOrEmpty(name);
        if (_main.Descendants(W + "bookmarkStart").Any(b => (string?)b.Attribute(W + "name") == name))
            throw new InvalidOperationException($"bookmark '{name}' already exists");
        var id = _nextId++;
        var s = new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name));
        var e = new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id));
        AddAtStart(paragraph, startNode, s);
        AddAtEnd(paragraph, endNode, e);
        return id;
    }

    /// <summary>
    /// Changes paragraph properties as a tracked w:pPrChange holding the prior pPr, so Reject All
    /// restores it. <paramref name="mutatePPr"/> edits the live w:pPr (created if missing) and must
    /// keep schema order; <see cref="SetStyle"/> helps for the common case.
    /// </summary>
    public XElement ChangeParagraph(XElement paragraph, Action<XElement> mutatePPr)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        ArgumentNullException.ThrowIfNull(mutatePPr);
        var pPr = paragraph.Element(W + "pPr");
        if (pPr is null)
        {
            pPr = new XElement(W + "pPr");
            paragraph.AddFirst(pPr);
        }
        XElement old;
        if (pPr.Element(W + "pPrChange") is { } existing)
        {
            RefuseOtherAuthor(existing, "paragraph properties");
            // Already changed by this author: keep the original properties as the baseline.
            old = existing.Element(W + "pPr") ?? new XElement(W + "pPr");
            existing.Remove();
        }
        else
        {
            old = new XElement(W + "pPr", pPr.Elements()
                .Where(c => c.Name != W + "rPr" && c.Name != W + "sectPr")
                .Select(c => new XElement(c)));
        }
        mutatePPr(pPr);
        var change = Tracked("pPrChange");
        change.Add(old);
        pPr.Add(change); // last child of CT_PPr, after rPr and sectPr
        return change;
    }

    /// <summary>Sets w:pStyle on a pPr, inserting it first (its schema position) when missing.</summary>
    public static void SetStyle(XElement pPr, string styleId)
    {
        ArgumentNullException.ThrowIfNull(pPr);
        if (pPr.Element(W + "pStyle") is { } ps) ps.SetAttributeValue(W + "val", styleId);
        else pPr.AddFirst(new XElement(W + "pStyle", new XAttribute(W + "val", styleId)));
    }

    /// <summary>
    /// Adds a comment from <paramref name="startElement"/> to <paramref name="endElement"/> (runs,
    /// w:ins/w:del, or a whole paragraph), creating comments.xml if needed. Lines in
    /// <paramref name="text"/> become comment paragraphs. Returns the comment id.
    /// </summary>
    public AddedComment Comment(XElement startElement, XElement endElement, string text)
    {
        ArgumentNullException.ThrowIfNull(startElement);
        ArgumentNullException.ThrowIfNull(endElement);
        ArgumentException.ThrowIfNullOrEmpty(text);
        var id = _nextId++;
        // Comment markers inside a w:ins vanish with it on Reject All, orphaning the comment
        // (2026-10-15 review): anchor around the whole insertion instead.
        startElement = OutsideInsertion(startElement);
        endElement = OutsideInsertion(endElement);
        var startP = startElement.Name == W + "p" ? startElement : startElement.Ancestors(W + "p").First();
        var endP = endElement.Name == W + "p" ? endElement : endElement.Ancestors(W + "p").First();
        AddAtStart(startP, startElement, new XElement(W + "commentRangeStart", new XAttribute(W + "id", id)));
        AddAtEnd(endP, endElement,
            new XElement(W + "commentRangeEnd", new XAttribute(W + "id", id)),
            new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", id))));

        _comments ??= new XDocument(new XElement(W + "comments", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName)));
        var comment = new XElement(W + "comment",
            new XAttribute(W + "id", id),
            new XAttribute(W + "author", Author.Name),
            new XAttribute(W + "date", Author.WordDate),
            new XAttribute(W + "initials", Author.Initials));
        var first = true;
        foreach (var line in text.Split('\n'))
        {
            var p = new XElement(W + "p");
            if (first) p.Add(new XElement(W + "r", new XElement(W + "annotationRef")));
            first = false;
            p.Add(new XElement(W + "r", Preserved(W + "t", line.TrimEnd('\r'))));
            comment.Add(p);
        }
        _comments.Root!.Add(comment);
        var anchor = startP == endP ? Text(startP) : Text(startP) + " …";
        return new AddedComment(id, ParagraphText.Excerpt(anchor), text);
    }

    /// <summary>Writes document.xml and (if any comments exist) comments.xml back to the package.</summary>
    public void Save()
    {
        var main = Document.MainDocumentPart!;
        Docx.Save(main, _main);
        if (_comments is not null)
            Docx.Save(main.WordprocessingCommentsPart ?? main.AddNewPart<WordprocessingCommentsPart>(), _comments);
    }

    // ------------------------------------------------------------------ run model

    private sealed class Field
    {
        public List<XElement> Runs { get; } = [];
        public bool Closed { get; set; }
    }

    private sealed record Segment(XElement Run, int Start, int Length, XElement? Insertion, Field? Field, bool Foreign);

    private sealed record RunMap(List<Segment> Segments);

    private static bool Inside(XElement e, XElement paragraph, params string[] names) =>
        e.Ancestors().TakeWhile(a => a != paragraph).Any(a => a.Name.Namespace == W && names.Contains(a.Name.LocalName));

    private static int RunLength(XElement run) =>
        run.Elements().Sum(c => c.Name == W + "t" ? c.Value.Length : c.Name == W + "tab" ? 1 : 0);

    /// <summary>Live runs of the paragraph (not deleted, not in nested paragraphs) with offsets and context.</summary>
    private RunMap Map(XElement paragraph)
    {
        var runs = paragraph.Descendants(W + "r")
            .Where(r => r.Ancestors(W + "p").First() == paragraph && !Inside(r, paragraph, "del", "moveFrom"))
            .ToList();
        var foreignDepth = _fieldDepthAtStart.GetValueOrDefault(paragraph);
        var open = new Stack<Field>();
        var fieldOf = new Dictionary<XElement, Field>();
        var foreign = new HashSet<XElement>();
        foreach (var r in runs)
        {
            var kind = (string?)r.Element(W + "fldChar")?.Attribute(W + "fldCharType");
            if (kind == "begin") open.Push(new Field());
            if (open.Count > 0)
            {
                var outer = open.Last();
                outer.Runs.Add(r);
                fieldOf[r] = outer;
            }
            if (foreignDepth > 0 || Inside(r, paragraph, "fldSimple")) foreign.Add(r);
            if (kind == "end")
            {
                if (open.Count > 0)
                {
                    var f = open.Pop();
                    if (open.Count == 0) f.Closed = true;
                }
                else if (foreignDepth > 0) foreignDepth--;
            }
        }
        // A field opened here but closed in a later paragraph cannot be edited locally.
        foreach (var f in fieldOf.Values.Where(f => !f.Closed)) foreign.UnionWith(f.Runs);

        List<Segment> segments = [];
        var pos = 0;
        foreach (var r in runs)
        {
            var n = RunLength(r);
            var ins = r.Ancestors().TakeWhile(a => a != paragraph).FirstOrDefault(a => a.Name == W + "ins" || a.Name == W + "moveTo");
            var field = fieldOf.GetValueOrDefault(r);
            segments.Add(new Segment(r, pos, n, ins, field is { Closed: true } ? field : null, foreign.Contains(r) || field is { Closed: false }));
            pos += n;
        }
        return new RunMap(segments);
    }

    /// <summary>Split <paramref name="run"/> at text offset k: the run keeps [0, k), the returned run holds [k, end).</summary>
    private XElement SplitRun(XElement run, int k)
    {
        var right = new XElement(run.Name, run.Attributes());
        if (run.Element(W + "rPr") is { } rPr)
        {
            var copy = new XElement(rPr);
            // Both halves keep the tracked formatting change, each with its own w:id (ids must be
            // unique; a deep copy duplicated it, 2026-10-15 review).
            foreach (var change in copy.Elements(W + "rPrChange")) change.SetAttributeValue(W + "id", _nextId++);
            right.Add(copy);
        }
        var pos = 0;
        var moving = false;
        foreach (var c in run.Elements().Where(c => c.Name != W + "rPr").ToList())
        {
            if (!moving && pos >= k) moving = true;
            if (moving)
            {
                c.Remove();
                right.Add(c);
                continue;
            }
            var len = c.Name == W + "t" ? c.Value.Length : c.Name == W + "tab" ? 1 : 0;
            if (c.Name == W + "t" && pos + len > k)
            {
                var cut = k - pos;
                right.Add(Preserved(W + "t", c.Value[cut..]));
                c.Value = c.Value[..cut];
                c.SetAttributeValue(Ns.Xml + "space", "preserve");
                moving = true;
            }
            pos += len;
        }
        run.AddAfterSelf(right);
        return right;
    }

    private void SplitAt(XElement paragraph, int cut)
    {
        foreach (var s in Map(paragraph).Segments)
        {
            if (s.Start < cut && cut < s.Start + s.Length)
            {
                SplitRun(s.Run, cut - s.Start);
                return;
            }
        }
    }

    private List<XElement> Isolate(XElement paragraph, int start, int end, bool forEdit)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        var text = Text(paragraph);
        if (start < 0 || end <= start || end > text.Length)
            throw new ArgumentOutOfRangeException(nameof(end), $"range [{start}, {end}) outside paragraph of length {text.Length}");
        SplitAt(paragraph, end);
        SplitAt(paragraph, start);
        var map = Map(paragraph).Segments;
        var inRange = map.Where(s => s.Length > 0 && s.Start >= start && s.Start + s.Length <= end).ToList();
        var covered = string.Concat(inRange.Select(s => text.Substring(s.Start, s.Length)));
        if (covered != text[start..end])
            throw new InvalidOperationException($"could not isolate \"{text[start..end]}\" into whole runs");
        if (!forEdit) return inRange.Select(s => s.Run).ToList();

        var first = map.IndexOf(inRange[0]);
        var last = map.IndexOf(inRange[^1]);
        // Whole fields only: a range covering part of a field result would break the field.
        foreach (var f in inRange.Where(s => s.Field is not null).Select(s => s.Field!).Distinct())
        {
            var fieldSegs = map.Where(s => s.Field == f).ToList();
            if (fieldSegs.Any(s => s.Length > 0 && (s.Start < start || s.Start + s.Length > end)))
                throw new InvalidOperationException($"\"{text[start..end]}\" covers only part of a field result");
            first = Math.Min(first, map.IndexOf(fieldSegs[0]));
            last = Math.Max(last, map.IndexOf(fieldSegs[^1]));
        }
        var selected = map.Skip(first).Take(last - first + 1).ToList();
        foreach (var s in selected)
        {
            if (s.Foreign) throw new InvalidOperationException($"\"{text[start..end]}\" lies inside a field that spans paragraphs or a simple field");
            if (s.Insertion is not null) RefuseInsideRevision(s.Insertion, text[start..end]);
        }
        return selected.Select(s => s.Run).ToList();
    }

    private void RefuseInsideRevision(XElement revision, string what)
    {
        var author = (string?)revision.Attribute(W + "author") ?? "unknown";
        if (author != Author.Name)
            throw new InvalidOperationException($"\"{what}\" is inside a tracked insertion by {author}; resolve their revision first");
        throw new InvalidOperationException($"\"{what}\" is inside an insertion this redline already made; edit the original text instead");
    }

    private void RefuseOtherAuthor(XElement revision, string what)
    {
        var author = (string?)revision.Attribute(W + "author") ?? "unknown";
        if (author != Author.Name)
            throw new InvalidOperationException($"{what} already carry a tracked change by {author}; resolve their revision first");
    }

    /// <summary>Where a pure insertion at <paramref name="offset"/> goes, and the formatting it takes.</summary>
    private (XElement? RPr, Action<XElement> Place) InsertionPoint(XElement paragraph, int offset)
    {
        SplitAt(paragraph, offset);
        var map = Map(paragraph).Segments.Where(s => s.Length > 0).ToList();
        var before = map.LastOrDefault(s => s.Start + s.Length == offset);
        var after = map.FirstOrDefault(s => s.Start == offset);
        // Text in a field that spans paragraphs (TOC lines) is regenerated by Update Field, so an
        // insertion next to it would land in the result and be discarded silently (2026-10-15 review).
        if (StartsInsideField(paragraph) || before?.Foreign == true || after?.Foreign == true)
            throw new InvalidOperationException($"cannot insert at offset {offset}: the paragraph is part of a field result (e.g. a table of contents) that Word regenerates");
        if (before is not null && after is not null)
        {
            if (before.Field is not null && before.Field == after.Field)
                throw new InvalidOperationException($"cannot insert at offset {offset}: inside a field result");
            if (before.Insertion is not null && before.Insertion == after.Insertion)
                RefuseInsideRevision(before.Insertion, Text(paragraph)[before.Start..(after.Start + after.Length)]);
        }
        if (before is not null)
        {
            // After the whole field / insertion the preceding text belongs to, never inside it.
            var anchor = before.Field is { } f ? f.Runs[^1] : before.Run;
            var target = before.Insertion ?? anchor;
            return (CleanRPr(before.Run), e => target.AddAfterSelf(e));
        }
        if (after is not null)
        {
            var anchor = after.Field is { } f ? f.Runs[0] : after.Run;
            XElement target = after.Insertion ?? anchor;
            return (CleanRPr(after.Run), e => target.AddBeforeSelf(e));
        }
        return (null, e => AddAtStart(paragraph, paragraph, e));
    }

    private static IEnumerable<List<XElement>> AdjacentGroups(IReadOnlyList<XElement> runs)
    {
        List<XElement> group = [];
        foreach (var r in runs)
        {
            if (group.Count > 0 && group[^1].ElementsAfterSelf().FirstOrDefault() != r)
            {
                yield return group;
                group = [];
            }
            group.Add(r);
        }
        if (group.Count > 0) yield return group;
    }

    /// <summary>The outermost w:ins / w:moveTo containing <paramref name="node"/> within its paragraph, else the node.</summary>
    private static XElement OutsideInsertion(XElement node)
    {
        if (node.Name == W + "p") return node;
        return node.Ancestors().TakeWhile(a => a.Name != W + "p")
            .LastOrDefault(a => a.Name == W + "ins" || a.Name == W + "moveTo") ?? node;
    }

    private static void AddAtStart(XElement paragraph, XElement node, XElement added)
    {
        if (node != paragraph) node.AddBeforeSelf(added);
        else if (paragraph.Element(W + "pPr") is { } pPr) pPr.AddAfterSelf(added);
        else paragraph.AddFirst(added);
    }

    private static void AddAtEnd(XElement paragraph, XElement node, params XElement[] added)
    {
        if (node != paragraph) node.AddAfterSelf(added);
        else paragraph.Add(added);
    }

    private static XElement? CleanRPr(XElement run)
    {
        if (run.Element(W + "rPr") is not { } rPr) return null;
        var copy = new XElement(rPr);
        copy.Elements(W + "rPrChange").Remove();
        return copy;
    }

    private static XElement NewRun(XElement? rPr, XElement child) =>
        new(W + "r", rPr is null ? null : new XElement(rPr), child);

    private static IEnumerable<XElement> TextChildren(string text)
    {
        var parts = text.Split('\t');
        for (var i = 0; i < parts.Length; i++)
        {
            if (i > 0) yield return new XElement(W + "tab");
            if (parts[i].Length > 0) yield return Preserved(W + "t", parts[i]);
        }
    }

    private XElement Tracked(string localName) => new(W + localName,
        new XAttribute(W + "id", _nextId++),
        new XAttribute(W + "author", Author.Name),
        new XAttribute(W + "date", Author.WordDate));

    private static XElement FldChar(string kind) => new(W + "fldChar", new XAttribute(W + "fldCharType", kind));

    // Leading, trailing and whitespace-only text must carry xml:space="preserve"; without it a
    // deleted lone space vanishes (2026-10-15, a client redline: docx-tool's pack step dropped them).
    private static XElement Preserved(XName name, string text) => new(name, new XAttribute(Ns.Xml + "space", "preserve"), text);

    private static bool IsWord(char c) => char.IsLetterOrDigit(c);

    /// <summary>Lengths of the common prefix and suffix of a and b, snapped back to word boundaries.</summary>
    internal static (int Prefix, int Suffix) CommonEnds(string a, string b)
    {
        var p = 0;
        while (p < a.Length && p < b.Length && a[p] == b[p]) p++;
        while (p > 0 && IsWord(a[p - 1]) && ((p < a.Length && IsWord(a[p])) || (p < b.Length && IsWord(b[p])))) p--;
        var s = 0;
        while (s < a.Length - p && s < b.Length - p && a[a.Length - 1 - s] == b[b.Length - 1 - s]) s++;
        while (s > 0 && IsWord(a[a.Length - s])
               && ((s < a.Length - p && IsWord(a[a.Length - 1 - s])) || (s < b.Length - p && IsWord(b[b.Length - 1 - s]))))
            s--;
        return (p, s);
    }

    [GeneratedRegex(@"^_Ref(\d+)$")]
    private static partial Regex RefName();
}
