using System.CommandLine;
using System.Globalization;
using OfficeProof.Core;
using OfficeProof.Redline;
using OfficeProof.Verify;

namespace OfficeProof.Cli;

/// <summary>
/// redline: tracked changes on someone else's document, then the same checks `verify` runs.
/// Every operation is optional; they run in a fixed order (edits, captions, headings, TOC font)
/// so content edits see the author's typed text before numbers turn into fields.
/// </summary>
internal static class RedlineCommand
{
    public static Command Create()
    {
        var input = new Argument<FileInfo>("input") { Description = "Document to redline (not modified)" };
        var output = new Argument<FileInfo>("output") { Description = "Where to write the tracked-change copy" };
        var author = new Option<string>("--author") { Description = "Revision author shown in Word", DefaultValueFactory = _ => "Claude" };
        var initials = new Option<string?>("--initials") { Description = "Comment initials (default: author's initials)" };
        var edits = new Option<FileInfo?>("--edits")
        {
            Description = "JSON edit script: [{\"find\": paragraph snippet, \"replace\": old, \"with\": new, \"comment\"?: text, \"nth\"?: n} | {\"find\", \"comment\", \"anchor\"?}]",
        };
        var captions = new Option<bool>("--captions") { Description = "Typed 'Figure N'/'Table N' captions become SEQ fields and in-text references become REF fields" };
        var fix = new Option<string[]>("--fix")
        {
            Description = "Retarget one typed reference: \"<paragraph snippet>|<Label>|<old>|<new>[|<nth>]\" (repeatable; implies --captions)",
        };
        var headings = new Option<string?>("--headings")
        {
            Description = "Number headings automatically: style=level pairs, e.g. Head1=1,Head2=2,Head3=3",
        };
        var level = new Option<string[]>("--level")
        {
            Description = "Put one heading at another level: \"<heading text prefix>=<level>\" (repeatable; needs --headings)",
        };
        var tocFont = new Option<string?>("--toc-font") { Description = "Set this font on TOC 1-9 styles so Update Field keeps it (untracked)" };
        var noVerify = new Option<bool>("--no-verify") { Description = "Skip the post-redline checks (not recommended)" };
        var json = Output.JsonOption();
        var force = Output.ForceOption();

        var cmd = new Command("redline", "Apply tracked changes and comments, then verify Reject All restores the original")
        {
            input, output, author, initials, edits, captions, fix, headings, level, tocFont, noVerify, force, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var inPath = parse.GetValue(input)!.FullName;
            var outPath = parse.GetValue(output)!.FullName;
            Output.RequireNewOutput(outPath, parse.GetValue(force));
            var name = parse.GetValue(author)!;
            var who = new RevisionAuthor(name, parse.GetValue(initials) ?? Initials(name), DateTimeOffset.UtcNow);
            var fixes = (parse.GetValue(fix) ?? []).Select(ParseFix).ToList();
            var headingLevels = parse.GetValue(headings) is { } h ? ParseLevels(h) : null;
            var overrides = (parse.GetValue(level) ?? []).Select(ParseOverride).ToDictionary(p => p.Key, p => p.Value);
            if (overrides.Count > 0 && headingLevels is null) throw new ArgumentException("--level needs --headings");
            var script = parse.GetValue(edits) is { } e ? EditScript.Load(e.FullName) : null;
            var font = parse.GetValue(tocFont);
            var doCaptions = parse.GetValue(captions) || fixes.Count > 0;
            if (script is null && !doCaptions && headingLevels is null && font is null)
                throw new ArgumentException("nothing to do: give --edits, --captions, --headings or --toc-font");

            var results = new List<(string Name, IRedlineResult Result)>();
            using (var doc = Docx.OpenCopy(inPath, outPath))
            {
                var editor = new TrackedEditor(doc, who);
                if (script is not null) results.Add(("edits", script.Apply(editor)));
                if (doCaptions) results.Add(("captions", CaptionFields.Apply(editor, fixes)));
                if (headingLevels is not null)
                    results.Add(("headings", HeadingNumbering.Apply(editor, headingLevels, new HeadingNumberingOptions { LevelOverrides = overrides })));
                if (font is not null) results.Add(("toc-font", TocStyles.SetFont(doc, font)));
                editor.Save();
            }

            foreach (var (opName, r) in results)
            {
                Output.Heading($"{opName}: {Summary(r)}");
                Output.List("comment", r.Comments.Select(c => $"#{c.Id} on \"{c.Anchor}\""));
                Output.List("note", r.Notes);
                // Word cannot track these; the author has to be told, and Reject All will not undo them.
                Output.List("UNTRACKED", r.UntrackedChanges);
            }
            Console.WriteLine($"output: {outPath}");
            if (parse.GetValue(noVerify)) return Output.Findings(Docx.Validate(outPath), parse.GetValue(json));
            var changedStyles = results.SelectMany(x => x.Result.ChangedStyles).ToHashSet(StringComparer.Ordinal);
            // The same check later: `officeproof verify <input> <output>` with these styles declared.
            Output.Heading(changedStyles.Count == 0 ? "verify" : $"verify --allow-style {string.Join(',', changedStyles.Order(StringComparer.Ordinal))}");
            return Output.Findings(RevisionCheck.Run(inPath, outPath, new VerifyOptions { IntendedStyles = changedStyles }), parse.GetValue(json));
        }, partialOutput: parse.GetValue(output)?.FullName));
        return cmd;
    }

    private static string Summary(IRedlineResult r) => r switch
    {
        EditScriptResult e => $"{e.Replacements} tracked replacement(s), {e.Comments.Count} comment(s)",
        CaptionFieldsResult c => $"{c.Captions} caption(s) -> SEQ, {c.References} reference(s) -> REF, {c.Retargeted} retargeted, {c.UnresolvedReferences.Count} unresolved",
        HeadingNumberingResult h => $"{h.NumberedHeadings} heading(s) numbered, {h.SectionReferences} section reference(s) linked, unnumbered: {string.Join(" | ", h.UnnumberedHeadings)}",
        TocFontResult => "TOC styles",
        _ => r.GetType().Name,
    };

    private static string Initials(string name) =>
        string.Concat(name.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(w => char.ToUpperInvariant(w[0])));

    private static ReferenceFix ParseFix(string spec)
    {
        var parts = spec.Split('|');
        if (parts.Length is < 4 or > 5
            || !int.TryParse(parts[2], CultureInfo.InvariantCulture, out var old)
            || !int.TryParse(parts[3], CultureInfo.InvariantCulture, out var @new)
            || (parts.Length == 5 && !int.TryParse(parts[4], CultureInfo.InvariantCulture, out _)))
            throw new ArgumentException($"--fix \"{spec}\": expected \"<paragraph snippet>|<Label>|<old>|<new>[|<nth>]\" with whole numbers");
        int? nth = parts.Length == 5 ? int.Parse(parts[4], CultureInfo.InvariantCulture) : null;
        return new ReferenceFix(parts[0], parts[1], old, @new, nth);
    }

    private static Dictionary<string, int> ParseLevels(string spec)
    {
        var levels = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var pair in spec.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var (k, v) = ParseOverride(pair);
            levels[k] = v;
        }
        return levels;
    }

    private static (string Key, int Value) ParseOverride(string spec)
    {
        var i = spec.LastIndexOf('=');
        if (i <= 0 || !int.TryParse(spec[(i + 1)..], CultureInfo.InvariantCulture, out var n) || n < 1)
            throw new ArgumentException($"\"{spec}\": expected <name>=<level>, level 1 or more");
        return (spec[..i].Trim(), n);
    }
}
