using System.CommandLine;
using OfficeProof.Core;
using OfficeProof.Lint;
using OfficeProof.Templating;

namespace OfficeProof.Cli;

/// <summary>generate and retemplate: a template directory applied as code.</summary>
internal static class TemplateCommands
{
    /// <summary>Environment variable naming the template directory when --template is not given.</summary>
    public const string TemplateVariable = "OFFICEPROOF_TEMPLATE";

    public static Option<DirectoryInfo?> TemplateOption() => new("--template")
    {
        Description = $"Template directory (template.yaml, styles.xml, theme1.xml, ...); default ${TemplateVariable}",
    };

    /// <summary>The template directory: <paramref name="explicitDirectory"/>, else <see cref="TemplateVariable"/>, else null.</summary>
    public static string? TemplateDirectory(DirectoryInfo? explicitDirectory) =>
        explicitDirectory?.FullName
        ?? (Environment.GetEnvironmentVariable(TemplateVariable) is { Length: > 0 } fromEnv ? Path.GetFullPath(fromEnv) : null);

    /// <summary>The lint palette a template declares; null when it declares none.</summary>
    public static BrandPalette? Palette(TemplateManifest manifest) =>
        manifest.Palette is { } p ? new BrandPalette(p.Colours, p.NearMisses) : null;

    /// <summary>
    /// The palette of the template named by <paramref name="explicitDirectory"/> or
    /// <see cref="TemplateVariable"/>; null without a template. Only template.yaml is read.
    /// </summary>
    public static BrandPalette? LintPalette(DirectoryInfo? explicitDirectory) =>
        TemplateDirectory(explicitDirectory) is { } dir ? Palette(TemplateManifest.Load(dir)) : null;

    private static Option<FileInfo?> MetadataOption(string what) => new("--metadata") { Description = what };

    private static Option<DirectoryInfo?> FontsOption() => new("--fonts")
    {
        Description = "Directory holding the font faces template.yaml names (default ~/Library/Fonts)",
    };

    public static Command Generate()
    {
        var template = TemplateOption();
        var metadata = MetadataOption("Report metadata YAML (a 'report:' map); default <template>/report.yaml");
        var fonts = FontsOption();
        var output = new Option<DirectoryInfo>("--output", "-o")
        {
            Description = "Directory for <name>-sample.docx and <name>.dotx (name from template.yaml)",
            DefaultValueFactory = _ => new DirectoryInfo(Environment.CurrentDirectory),
        };
        var cmd = new Command("generate", "Emit the template's Word template (.dotx) and a reference sample (.docx)")
        {
            template, metadata, fonts, output,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var source = Source(parse.GetValue(template));
            var meta = ReportMetadata.Load(parse.GetValue(metadata)?.FullName ?? source.MetadataPath);
            var result = Generator.Generate(source, meta, parse.GetValue(output)!.FullName, parse.GetValue(fonts)?.FullName);
            Docx.NormalizeEntryModes(result.SamplePath);
            Docx.NormalizeEntryModes(result.TemplatePath);
            Console.WriteLine($"template: {source.Directory}");
            Console.WriteLine($"metadata: {meta.Source}");
            Console.WriteLine($"generated: {result.SamplePath}");
            Console.WriteLine($"generated: {result.TemplatePath}");
            Console.WriteLine($"cover image: {result.Cover.Status}{(result.Cover.Image is { } i ? $" ({i})" : "")}");
            PrintFonts(result.Fonts);
            var findings = Docx.Validate(result.SamplePath).Concat(Docx.Validate(result.TemplatePath)).ToList();
            return Output.Findings(findings, json: false);
        }));
        return cmd;
    }

    public static Command Retemplate()
    {
        var input = new Argument<FileInfo>("input") { Description = "Report to restyle (not modified)" };
        var output = new Argument<FileInfo>("output") { Description = "Where to write the restyled copy" };
        var template = TemplateOption();
        var metadata = MetadataOption("Report metadata YAML; required unless the document carries the template's identityKeys itself");
        var fonts = FontsOption();
        var json = Output.JsonOption();
        var force = Output.ForceOption();
        var cmd = new Command("retemplate", "Apply a template to an existing report (untracked rewrite), then validate and lint it")
        {
            input, output, template, metadata, fonts, force, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var source = Source(parse.GetValue(template));
            var inPath = parse.GetValue(input)!.FullName;
            var outPath = parse.GetValue(output)!.FullName;
            Output.RequireNewOutput(outPath, parse.GetValue(force));
            var metaFile = parse.GetValue(metadata)?.FullName;

            // 2026-08-26: report.yaml is one hand-edited file, so whatever report was processed
            // last stamped its identity on the next. A document that does not name itself must be
            // given its metadata explicitly. 2026-10-15: an account-name dc:creator is not an
            // identity; the template's identityKeys are. A template with none always needs --metadata.
            var required = source.Manifest.IdentityKeys;
            if (metaFile is null)
            {
                if (required.Count == 0)
                    throw new ArgumentException($"{source.Manifest.Source} lists no identityKeys, so a document cannot identify itself; pass --metadata <report.yaml> for this report");
                DocumentIdentity identity;
                using (var doc = Docx.OpenRead(inPath)) identity = IdentityReader.Read(doc, source.Manifest);
                var missing = identity.Missing(required);
                if (missing.Count > 0)
                    throw new ArgumentException($"{Path.GetFileName(inPath)} does not carry its own {string.Join(" or ", missing)}; pass --metadata <report.yaml> for this report");
            }

            var seed = ReportMetadata.Load(metaFile ?? source.MetadataPath);
            var result = Retemplater.Retemplate(source, inPath, outPath, seed, parse.GetValue(fonts)?.FullName);
            Console.WriteLine($"identity: {result.IdentitySources} ({string.Join(", ", result.IdentityKeys)}); seed: {seed.Source}");
            var seeded = result.Metadata.Keys.Except(result.IdentityKeys).ToList();
            if (metaFile is null && seeded.Count > 0)
                Console.WriteLine($"  warning: {string.Join(", ", seeded)} taken from the template's report.yaml, which belongs to whichever report was last edited; check them");
            Console.WriteLine($"headings remapped {result.HeadingsRemapped}, typed numbers dropped {result.NumbersDropped}, direct formatting removed {result.DirectFormattingRemoved}");
            Console.WriteLine($"cover paragraphs replaced {result.CoverParagraphsReplaced} (block inserted: {result.CoverBlockInserted}), sections {result.Sections}, stale header/footer parts removed {result.StaleHeaderFooterParts}");
            Console.WriteLine($"cover image: {result.Cover.Status}");
            PrintFonts(result.Fonts);
            Console.WriteLine($"output: {result.OutputPath}");

            // Same post-checks the legacy CLI ran: schema, undefined styles, off-palette colours.
            var findings = Docx.Validate(outPath)
                .Concat(Linter.Run(outPath, Linter.Rules(["dangling-style", "brand-colour"], Palette(source.Manifest))))
                .ToList();
            return Output.Findings(findings, parse.GetValue(json));
        }, partialOutput: parse.GetValue(output)?.FullName));
        return cmd;
    }

    private static TemplateSource Source(DirectoryInfo? explicitDirectory) =>
        new(TemplateDirectory(explicitDirectory)
            ?? throw new ArgumentException($"no template directory: pass --template <dir> or set {TemplateVariable}"));

    private static void PrintFonts(FontEmbedResult? fonts)
    {
        if (fonts is null)
        {
            Console.WriteLine("fonts embedded: none (template.yaml names no fonts)");
            return;
        }
        Console.WriteLine($"fonts embedded: {fonts.Embedded.Count} ({fonts.Family})");
        Output.List("font missing", fonts.Missing);
    }
}
