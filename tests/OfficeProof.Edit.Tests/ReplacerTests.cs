using System.IO.Compression;
using System.Xml.Linq;
using OfficeProof.Core;
using OfficeProof.TestKit;
using static OfficeProof.TestKit.DocxBuilder;

namespace OfficeProof.Edit.Tests;

public sealed class ReplacerTests
{
    private static readonly XNamespace W = Ns.W;
    private const string Body = "/word/document.xml";

    [Fact]
    public void Text_split_over_runs_is_matched_and_replaced()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("Minutes of the Au"), R("gust meet"), R("ing")))
            .Raw(P(R("P&L for Aug"), R("ust 2026")))
            .Para("Unrelated")
            .Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.Equal([new PartMatches(Body, 2)], result.Pairs.Single().Parts);
        Assert.Equal(["Minutes of the September meeting", "P&L for September 2026", "Unrelated"], Texts(dir.File("out.docx"), Body));
    }

    [Fact]
    public void Headers_footers_and_notes_are_counted_per_part_and_story_selects_them()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("August body"), NoteReference(1)))
            .Header("rIdH", P(R("Au"), R("gust header")))
            .Footer("rIdF", P(R("August footer, August")))
            .Footnote(1, P(R("August note")))
            .Save(dir.File("in.docx"));

        var all = Replace(input, dir.File("all.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(all.Succeeded, Messages(all));
        Assert.Equal(
            [new PartMatches(Body, 1), new PartMatches("/word/header1.xml", 1), new PartMatches("/word/footer1.xml", 2), new PartMatches("/word/footnotes.xml", 1)],
            all.Pairs.Single().Parts);
        Assert.Equal(["September header"], Texts(dir.File("all.docx"), "/word/header1.xml"));
        Assert.Equal(["September footer, September"], Texts(dir.File("all.docx"), "/word/footer1.xml"));
        Assert.Contains("September note", Texts(dir.File("all.docx"), "/word/footnotes.xml"));

        var headers = Replace(input, dir.File("headers.docx"), new ReplaceOptions { Stories = Stories.Headers }, ("August", "September"));

        Assert.True(headers.Succeeded, Messages(headers));
        Assert.Equal([new PartMatches("/word/header1.xml", 1)], headers.Pairs.Single().Parts);
        Assert.Equal(["/word/header1.xml"], headers.EditedParts);
        Assert.Equal(["August body"], Texts(dir.File("headers.docx"), Body));
    }

    [Fact]
    public void Regex_groups_substitute_and_literal_mode_escapes_both_sides()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("Next meeting: Monday, 19 Oct"), R("ober 2026 at 10am")))
            .Para("Cost $1 (a.b)")
            .Save(dir.File("in.docx"));

        var regex = Replace(input, dir.File("regex.docx"), new ReplaceOptions { Regex = true },
            (@"(\w+day), 19 October (\d{4})", "$1, 16 November $2"));

        Assert.True(regex.Succeeded, Messages(regex));
        Assert.Equal(["Next meeting: Monday, 16 November 2026 at 10am", "Cost $1 (a.b)"], Texts(dir.File("regex.docx"), Body));

        var literal = Replace(input, dir.File("literal.docx"), new ReplaceOptions(), ("$1 (a.b)", "$2 (c.d)"));

        Assert.True(literal.Succeeded, Messages(literal));
        Assert.Equal(1, literal.Total);
        Assert.Equal("Cost $2 (c.d)", Texts(dir.File("literal.docx"), Body)[1]);
    }

    [Fact]
    public void Pairs_apply_in_order_each_to_the_result_of_the_ones_before()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Para("Minutes of the August meeting").Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("August", "September"), ("September meeting", "September committee meeting"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.Equal([1, 1], result.Pairs.Select(p => p.Total));
        Assert.Equal(["Minutes of the September committee meeting"], Texts(dir.File("out.docx"), Body));
    }

    [Fact]
    public void A_read_only_input_gives_a_writable_output()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Para("Minutes of the August meeting").Save(dir.File("in.docx"));
        File.SetAttributes(input, FileAttributes.ReadOnly);
        var output = dir.File("out.docx");

        var result = Replace(input, output, new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.False(File.GetAttributes(output).HasFlag(FileAttributes.ReadOnly));
        Assert.Equal(["Minutes of the September meeting"], Texts(output, Body));
        File.SetAttributes(input, FileAttributes.Normal);
    }

    [Fact]
    public void A_pair_without_a_match_refuses_and_writes_nothing()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Para("Minutes of the August meeting").Save(dir.File("in.docx"));
        var output = dir.File("out.docx");

        var result = Replace(input, output, new ReplaceOptions(), ("August", "September"), ("Agust", "October"));

        var error = Assert.Single(result.Errors);
        Assert.Equal("replace-no-match", error.Rule);
        Assert.Contains("\"Agust\"", error.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Expect_mismatch_refuses_and_writes_nothing()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Para("August and August").Header("rIdH", P(R("August"))).Save(dir.File("in.docx"));
        var output = dir.File("out.docx");

        var wrong = Replace(input, output, new ReplaceOptions { Expect = 2 }, ("August", "September"));

        Assert.Equal("replace-expect", Assert.Single(wrong.Errors).Rule);
        Assert.Contains("3 match(es) in total, --expect 2", wrong.Errors[0].Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));

        var right = Replace(input, output, new ReplaceOptions { Expect = 3 }, ("August", "September"));

        Assert.True(right.Succeeded, Messages(right));
        Assert.True(File.Exists(output));
    }

    [Fact]
    public void Tracked_changes_anywhere_refuse_and_point_at_redline()
    {
        using var dir = new TempDir();
        var inBody = new DocxBuilder()
            .Para("August")
            .Raw(P(new XElement(W + "ins", new XAttribute(W + "id", 1), new XAttribute(W + "author", "a"), R("added"))))
            .Save(dir.File("body.docx"));
        // A formatting change in a header refuses even a main-story replace.
        var inHeader = new DocxBuilder()
            .Para("August")
            .Header("rIdH", P(new XElement(W + "pPr", new XElement(W + "pPrChange",
                new XAttribute(W + "id", 2), new XAttribute(W + "author", "a"), new XElement(W + "pPr"))), R("Header")))
            .Save(dir.File("header.docx"));

        foreach (var (input, part) in new[] { (inBody, Body), (inHeader, "/word/header1.xml") })
        {
            var output = dir.File("out.docx");
            var result = Replace(input, output, new ReplaceOptions { Stories = Stories.Main }, ("August", "September"));

            var error = Assert.Single(result.Errors);
            Assert.Equal("replace-revisions", error.Rule);
            Assert.Equal(part, error.Part);
            Assert.Contains("officeproof redline", error.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        }
    }

    [Fact]
    public void A_match_covering_a_field_code_refuses()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("Page "), Field("begin"), Instr(" PAGE "), Field("separate"), R("3"), Field("end"), R(" of 9")))
            .Save(dir.File("in.docx"));
        var output = dir.File("out.docx");

        var spanning = Replace(input, output, new ReplaceOptions { Regex = true }, ("Page .* of", "Seite"));

        var error = Assert.Single(spanning.Errors);
        Assert.Equal("replace-non-text", error.Rule);
        Assert.Contains("w:fldChar", error.Message, StringComparison.Ordinal);
        Assert.Contains("w:instrText", error.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));

        // The field result alone is text.
        var result = Replace(input, output, new ReplaceOptions(), (" of 9", " of 10"));

        Assert.True(result.Succeeded, Messages(result));
        using var doc = Docx.OpenRead(output);
        var p = Docx.Load(doc.MainDocumentPart!).Descendants(W + "p").Single();
        Assert.Equal(" PAGE ", p.Descendants(W + "instrText").Single().Value);
        Assert.Equal("Page 3 of 10", Wml.VisibleText(p));
    }

    [Fact]
    public void Replacement_takes_the_first_matched_characters_formatting()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("Minutes of the "), Run("Au", new XElement(W + "b")), Run("gust", new XElement(W + "i")), R(" meeting")))
            .Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        var runs = BodyParagraphs(dir.File("out.docx")).Single().Elements(W + "r").ToList();
        Assert.Equal(["Minutes of the ", "September", " meeting"], runs.Select(r => r.Element(W + "t")!.Value));
        var replaced = runs[1].Element(W + "rPr");
        Assert.NotNull(replaced?.Element(W + "b"));
        Assert.Null(replaced.Element(W + "i"));
    }

    [Fact]
    public void Leading_and_trailing_spaces_in_the_replacement_are_preserved()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Raw(P(Run("X", new XElement(W + "b")), R("tail"))).Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("X", "  a b "));

        Assert.True(result.Succeeded, Messages(result));
        var p = BodyParagraphs(dir.File("out.docx")).Single();
        var t = p.Elements(W + "r").First().Element(W + "t")!;
        Assert.Equal("  a b ", t.Value);
        Assert.Equal("preserve", (string?)t.Attribute(XNamespace.Xml + "space"));
        Assert.Equal("  a b tail", Wml.VisibleText(p));
    }

    [Fact]
    public void Text_in_a_hyperlink_is_replaced_inside_the_hyperlink()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Raw(P(R("See "), new XElement(W + "hyperlink", new XAttribute(Ns.R + "id", "rIdL"), R("Aug"), R("ust minutes"))))
            .Link("rIdL", "https://example.com/")
            .Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        var link = BodyParagraphs(dir.File("out.docx")).Single().Element(W + "hyperlink")!;
        Assert.Equal("rIdL", (string?)link.Attribute(Ns.R + "id"));
        Assert.Equal("September minutes", string.Concat(link.Descendants(W + "t").Select(t => t.Value)));
    }

    [Fact]
    public void A_text_box_is_replaced_in_both_copies_and_counted_once()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder().Raw(P(R("Box: "), ShapeTextBox(1, P(R("Aug"), R("ust box"))))).Save(dir.File("in.docx"));

        var result = Replace(input, dir.File("out.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.Equal(1, result.Total);
        var boxes = BodyParagraphs(dir.File("out.docx")).Where(p => p.Ancestors(W + "txbxContent").Any()).Select(Wml.VisibleText);
        Assert.Equal(["September box", "September box"], boxes);
    }

    [Fact]
    public void Every_part_but_the_edited_stories_keeps_its_bytes()
    {
        using var dir = new TempDir();
        var input = new DocxBuilder()
            .Para("Minutes of the August meeting")
            .Raw(P(Picture("rIdImg", 1)))
            .Raw(P(R("See "), new XElement(W + "hyperlink", new XAttribute(Ns.R + "id", "rIdL"), R("the site"))))
            .Link("rIdL", "https://example.com/")
            .Image("rIdImg", PngA)
            .Header("rIdH", P(R("Header without a match")))
            .Footer("rIdF", P(R("August footer")))
            .Footnote(1, P(R("A note")))
            .Comment(0, "August comment, not searched by default")
            .Save(dir.File("in.docx"));
        var output = dir.File("out.docx");

        var result = Replace(input, output, new ReplaceOptions(), ("August", "September"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.Equal([Body, "/word/footer1.xml"], result.EditedParts);
        var before = Entries(input);
        var after = Entries(output);
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));
        var untouched = before.Keys.Except(result.EditedParts.Select(p => p.TrimStart('/'))).ToList();
        Assert.Contains("word/header1.xml", untouched);
        Assert.Contains("word/comments.xml", untouched);
        Assert.Contains("word/styles.xml", untouched);
        Assert.Contains("[Content_Types].xml", untouched);
        Assert.Contains("word/_rels/document.xml.rels", untouched);
        Assert.Contains(untouched, n => n.EndsWith(".png", StringComparison.Ordinal));
        foreach (var name in untouched) Assert.True(before[name].SequenceEqual(after[name]), $"{name} changed");
        Assert.All(result.EditedParts, p => Assert.False(before[p.TrimStart('/')].SequenceEqual(after[p.TrimStart('/')])));
    }

    [Fact]
    public void Output_is_schema_clean_relative_to_the_input()
    {
        using var dir = new TempDir();
        var clean = new DocxBuilder()
            .Raw(P(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Head1"))), R("Au"), Run("gust", new XElement(W + "b")), R(" report")))
            .Footer("rIdF", P(R("August")))
            .Save(dir.File("clean.docx"));

        var result = Replace(clean, dir.File("clean-out.docx"), new ReplaceOptions(), ("August", " September\t"));

        Assert.True(result.Succeeded, Messages(result));
        Assert.Empty(result.NewSchemaErrors);
        Assert.Empty(Docx.Validate(dir.File("clean-out.docx")));

        // An error the input already has is not the replace's.
        var flawed = new DocxBuilder()
            .Para("August")
            .Raw(P(new XElement(W + "bogus"), R("text")))
            .Save(dir.File("flawed.docx"));
        Assert.NotEmpty(Docx.Validate(flawed));

        var again = Replace(flawed, dir.File("flawed-out.docx"), new ReplaceOptions(), ("August", "September"));

        Assert.True(again.Succeeded, Messages(again));
        Assert.Empty(again.NewSchemaErrors);
        Assert.Equal(Docx.Validate(flawed).Count, Docx.Validate(dir.File("flawed-out.docx")).Count);
    }

    // ---- helpers

    private static ReplaceResult Replace(string input, string output, ReplaceOptions options, params (string Find, string With)[] pairs) =>
        Replacer.Apply(input, output, [.. pairs.Select(p => new ReplacePair(p.Find, p.With))], options);

    private static string Messages(ReplaceResult r) => string.Join("; ", r.Errors.Select(e => e.Message));

    private static XElement Run(string text, params XElement[] properties) =>
        new(W + "r", new XElement(W + "rPr", properties), new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    private static XElement Field(string kind) => new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", kind)));

    private static XElement Instr(string code) =>
        new(W + "r", new XElement(W + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), code));

    private static List<XElement> BodyParagraphs(string path)
    {
        using var doc = Docx.OpenRead(path);
        return [.. Docx.Load(doc.MainDocumentPart!).Descendants(W + "p")];
    }

    /// <summary>Visible text of every paragraph in the part at <paramref name="uri"/>.</summary>
    private static List<string> Texts(string path, string uri)
    {
        using var doc = Docx.OpenRead(path);
        var part = Docx.StoryParts(doc.MainDocumentPart!).Single(p => p.Uri.ToString() == uri);
        return [.. Docx.Load(part).Descendants(W + "p").Select(Wml.VisibleText)];
    }

    private static Dictionary<string, byte[]> Entries(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var s = e.Open();
            using var m = new MemoryStream();
            s.CopyTo(m);
            return m.ToArray();
        });
    }
}
