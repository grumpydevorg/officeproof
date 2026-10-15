using System.CommandLine;
using OfficeProof.Core;
using OfficeProof.Lint;
using OfficeProof.Verify;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Cli;

/// <summary>Read-only verbs: validate, lint, verify, outline.</summary>
internal static class CheckCommands
{
    public static Command Validate()
    {
        var input = new Argument<FileInfo>("input") { Description = ".docx, .dotx or .xlsx package" };
        var json = Output.JsonOption();
        var cmd = new Command("validate", "Schema-validate a package with the Open XML SDK validator (Office 2019)") { input, json };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var path = parse.GetValue(input)!.FullName;
            IReadOnlyList<Finding> findings;
            if (Path.GetExtension(path).ToLowerInvariant() is ".xlsx" or ".xlsm")
            {
                using var book = SpreadsheetDocument.Open(path, false);
                findings = Docx.Validate(book);
            }
            else findings = Docx.Validate(path);
            return Output.Findings(findings, parse.GetValue(json));
        }));
        return cmd;
    }

    public static Command Lint()
    {
        var input = new Argument<FileInfo?>("input") { Description = "Document to check", Arity = ArgumentArity.ZeroOrOne };
        var rules = new Option<string[]>("--rule") { Description = "Run only this rule (repeatable: --rule a --rule b); default all" };
        var list = new Option<bool>("--list-rules") { Description = "List the rules and exit" };
        var template = TemplateCommands.TemplateOption();
        template.Description += "; its template.yaml palette configures brand-colour, which is skipped (Info) without one";
        var json = Output.JsonOption();
        var cmd = new Command("lint", "Report what will break when the document is edited, renumbered or updated in Word") { input, rules, list, template, json };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            if (parse.GetValue(list))
            {
                foreach (var r in Linter.DefaultRules()) Console.WriteLine($"{r.Id,-26} {r.Severity,-8} {r.Description}");
                return 0;
            }
            var path = parse.GetValue(input)?.FullName ?? throw new ArgumentException("lint needs an input document");
            var palette = TemplateCommands.LintPalette(parse.GetValue(template));
            var selected = parse.GetValue(rules) is { Length: > 0 } ids ? Linter.Rules(ids, palette) : Linter.DefaultRules(palette);
            return Output.Findings(Linter.Run(path, selected), parse.GetValue(json));
        }));
        return cmd;
    }

    public static Command Verify()
    {
        var original = new Argument<FileInfo>("original") { Description = "The document before redlining" };
        var redline = new Argument<FileInfo>("redline") { Description = "The tracked-change copy" };
        var word = new Option<bool>("--word")
        {
            Description = "Also have Microsoft Word (macOS) reject all changes in both documents and compare every story and every printed page, " +
                          "compare its heading numbers in both views, update fields and report Error! results (about 2-3 minutes on a 40-page report)",
        };
        var keep = new Option<DirectoryInfo?>("--keep-pages") { Description = "With --word: directory for the page images of pages that differ" };
        var allowStyle = new Option<string[]>("--allow-style")
        {
            Description = "Style ids (w:styleId) whose existing definitions the redline changes on purpose, comma-separated or repeated, " +
                          "e.g. TOC1,TOC2,TOC3 after --toc-font; any other changed or removed existing style is an error",
        };
        var json = Output.JsonOption();
        var cmd = new Command("verify", "Check a redline: Reject All restores the original in every story, nothing hides in styles, references resolve")
        {
            original, redline, word, keep, allowStyle, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var o = parse.GetValue(original)!.FullName;
            var r = parse.GetValue(redline)!.FullName;
            var options = new VerifyOptions
            {
                IntendedStyles = (parse.GetValue(allowStyle) ?? [])
                    .SelectMany(s => s.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                    .ToHashSet(StringComparer.Ordinal),
            };
            var findings = RevisionCheck.Run(o, r, options).ToList();
            if (parse.GetValue(word))
            {
                findings.AddRange(WordFindings(r));
                findings.AddRange(WordCheck.RejectAll(o, r, parse.GetValue(keep)?.FullName, options: options));
            }
            return Output.Findings(findings, parse.GetValue(json));
        }));
        return cmd;
    }

    public static Command Outline()
    {
        var input = new Argument<FileInfo>("input") { Description = "Document to read" };
        var view = new Option<string>("--view")
        {
            Description = "as-is, accept (after Accept All) or reject (after Reject All)",
            DefaultValueFactory = _ => "as-is",
        };
        view.AcceptOnlyFromAmong("as-is", "accept", "reject");
        var word = new Option<bool>("--word") { Description = "Read the same outline from Microsoft Word and mark disagreements" };
        var cmd = new Command("outline", "Print headings and numbered paragraphs with the numbers Word would show") { input, view, word };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var path = parse.GetValue(input)!.FullName;
            var v = ParseView(parse.GetValue(view)!);
            var entries = OfficeProof.Verify.Outline.Read(path, v);
            Dictionary<int, OutlineEntry>? fromWord = null;
            if (parse.GetValue(word))
            {
                RequireWord();
                fromWord = WordOracle.Outline(path, v).ToDictionary(e => e.Index);
            }
            var disagreements = 0;
            foreach (var e in entries)
            {
                var mark = "";
                if (fromWord is not null)
                {
                    var same = fromWord.TryGetValue(e.Index, out var w) && w.ListNumber == e.ListNumber && w.Text == e.Text;
                    if (!same) disagreements++;
                    mark = same ? "  word: same" : $"  word: [{w?.ListNumber}] {w?.Text}";
                }
                var indent = new string(' ', 2 * (e.Level ?? 0));
                Console.WriteLine($"{e.Index,5} {indent}{(e.ListNumber.Length > 0 ? e.ListNumber + " " : "")}{e.Text}{mark}");
            }
            if (fromWord is not null) Console.WriteLine($"{disagreements} disagreement(s) with Word");
            return disagreements > 0 ? 1 : 0;
        }));
        return cmd;
    }

    private static OutlineView ParseView(string view) => view switch
    {
        "accept" => OutlineView.AcceptAll,
        "reject" => OutlineView.RejectAll,
        _ => OutlineView.AsIs,
    };

    private static void RequireWord()
    {
        if (!WordOracle.IsAvailable) throw new InvalidOperationException("--word needs Microsoft Word for Mac and osascript");
    }

    /// <summary>
    /// Word is the final oracle: Clippit models accept/reject and list numbering, but only Word
    /// shows what the reviewer will see (2026-10-15: style-definition revisions survive Reject All in Word).
    /// </summary>
    private static List<Finding> WordFindings(string redline)
    {
        RequireWord();
        var findings = new List<Finding>();
        foreach (var view in new[] { OutlineView.RejectAll, OutlineView.AcceptAll })
        {
            var ours = OfficeProof.Verify.Outline.Read(redline, view);
            var word = WordOracle.Outline(redline, view).ToDictionary(e => e.Index);
            var differ = ours.Where(e => !word.TryGetValue(e.Index, out var w) || w.ListNumber != e.ListNumber || w.Text != e.Text).ToList();
            foreach (var e in differ.Take(5))
            {
                word.TryGetValue(e.Index, out var w);
                findings.Add(new Finding("word-disagreement", Severity.Error,
                    $"{view}: officeproof reads [{e.ListNumber}] \"{e.Text}\", Word reads [{w?.ListNumber}] \"{w?.Text}\"",
                    "/word/document.xml", $"paragraph {e.Index}"));
            }
            findings.Add(new Finding("word-outline", differ.Count == 0 ? Severity.Info : Severity.Error,
                $"{view}: Word agrees on {ours.Count - differ.Count} of {ours.Count} headings and numbered paragraphs"));
        }
        var fields = WordOracle.UpdateFields(redline, OutlineView.AcceptAll);
        foreach (var error in fields.ErrorResults)
            findings.Add(new Finding("word-field-error", Severity.Error, $"after Accept All and Update Field, Word shows \"{error}\""));
        findings.Add(new Finding("word-toc", Severity.Info, $"TOC after update: {string.Join(" | ", fields.TocLines)}"));
        return findings;
    }
}
