using System.CommandLine;
using System.Text.Json;
using OfficeProof.Core;
using OfficeProof.Edit;

namespace OfficeProof.Cli;

/// <summary>replace: untracked find/replace in our own documents.</summary>
internal static class EditCommands
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public static Command Replace()
    {
        var input = new Argument<FileInfo>("input") { Description = "Document to edit (not modified)" };
        var output = new Argument<FileInfo>("output") { Description = "Where to write the edited copy" };
        var find = new Option<string[]>("--find")
        {
            Description = "Text to find (repeatable; the nth --find pairs with the nth --with). Literal unless --regex; matches text split over runs but never crosses paragraphs",
            Required = true,
        };
        var with = new Option<string[]>("--with")
        {
            Description = "Replacement text, in the formatting of the first matched character. With --regex, $1/${name} insert groups and $$ is a dollar sign",
            Required = true,
        };
        var regex = new Option<bool>("--regex") { Description = "Treat every --find as a .NET regular expression (case-sensitive)" };
        var story = new Option<string>("--story")
        {
            Description = "Stories to search: main, headers, footers, notes (footnotes and endnotes), comments, or all (every story but comments)",
            DefaultValueFactory = _ => "all",
        };
        story.AcceptOnlyFromAmong("main", "headers", "footers", "notes", "comments", "all");
        var expect = new Option<int?>("--expect") { Description = "Refuse unless the total number of matches over all pairs is exactly N" };
        var json = Output.JsonOption();
        var force = Output.ForceOption();
        var cmd = new Command("replace",
            "Untracked find/replace for our own documents, in every story; refuses documents with tracked changes (use redline), " +
            "a pair that matches nothing, a total other than --expect, and matches that cover fields or drawings")
        {
            input, output, find, with, regex, story, expect, force, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var finds = parse.GetValue(find)!;
            var withs = parse.GetValue(with)!;
            if (finds.Length != withs.Length)
                throw new ArgumentException($"{finds.Length} --find but {withs.Length} --with; give one --with per --find");
            var pairs = finds.Zip(withs, (f, w) => new ReplacePair(f, w)).ToList();
            var options = new ReplaceOptions
            {
                Regex = parse.GetValue(regex),
                Stories = ParseStory(parse.GetValue(story)!),
                Expect = parse.GetValue(expect),
            };
            var outPath = parse.GetValue(output)!.FullName;
            Output.RequireNewOutput(outPath, parse.GetValue(force));
            var asJson = parse.GetValue(json);

            var result = Replacer.Apply(parse.GetValue(input)!.FullName, outPath, pairs, options);
            if (!result.Succeeded)
            {
                // Refused replaces write nothing; exit 2 like every other refused operation.
                Output.Findings(result.Errors, asJson);
                return 2;
            }
            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(new
                {
                    pairs = result.Pairs.Select(p => new { find = p.Pair.Find, with = p.Pair.With, total = p.Total, parts = p.Parts }),
                    total = result.Total,
                    editedParts = result.EditedParts,
                    newSchemaErrors = result.NewSchemaErrors,
                    output = outPath,
                }, Json));
                return FindingReport.ExitCode(result.NewSchemaErrors);
            }
            foreach (var p in result.Pairs)
            {
                Console.WriteLine($"\"{p.Pair.Find}\" -> \"{p.Pair.With}\": {p.Total} match(es)");
                foreach (var part in p.Parts) Console.WriteLine($"  {part.Part}: {part.Matches}");
            }
            Console.WriteLine($"total {result.Total}; parts written: {string.Join(", ", result.EditedParts)}");
            if (result.NewSchemaErrors.Count > 0)
            {
                Output.Heading("new schema errors (the input does not have these)");
                FindingReport.WriteText(Console.Out, result.NewSchemaErrors);
            }
            else Console.WriteLine("schema: no new errors");
            Console.WriteLine($"output: {outPath}");
            return FindingReport.ExitCode(result.NewSchemaErrors);
        }, partialOutput: parse.GetValue(output)?.FullName));
        return cmd;
    }

    private static Stories ParseStory(string story) => story switch
    {
        "main" => Stories.Main,
        "headers" => Stories.Headers,
        "footers" => Stories.Footers,
        "notes" => Stories.Notes,
        "comments" => Stories.Comments,
        _ => Stories.All,
    };
}
