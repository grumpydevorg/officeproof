using System.CommandLine;
using OfficeProof.Spreadsheet;

namespace OfficeProof.Cli;

/// <summary>xlsx inspect / xlsx set: workbook structure and part-preserving cell edits.</summary>
internal static class XlsxCommands
{
    public static Command Create() => new("xlsx", "Excel workbooks: inspect structure, edit cells without disturbing other parts")
    {
        Inspect(),
        Set(),
    };

    private static Command Inspect()
    {
        var input = new Argument<FileInfo>("input") { Description = "Workbook to inspect" };
        var findings = new Option<bool>("--findings") { Description = "Print risks (labels, external links, stale caches) as findings instead of the full report" };
        var json = Output.JsonOption();
        var cmd = new Command("inspect", "Sheets, formulas, merged cells, conditional formats, validations and package facts (labels, calcChain, customXml)")
        {
            input, findings, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var report = WorkbookInspector.Inspect(parse.GetValue(input)!.FullName);
            if (parse.GetValue(findings)) return Output.Findings(report.ToFindings(), parse.GetValue(json));
            if (parse.GetValue(json))
            {
                Console.WriteLine(report.ToJson());
                return 0;
            }
            foreach (var s in report.Sheets)
            {
                var f = s.Formulas;
                Console.WriteLine($"{s.Name} [{s.Kind}{(s.State is { } st && st != "visible" ? ", " + st : "")}] {s.Dimension}: {s.CellCount} cells, "
                    + $"formulas {f.Total} (shared masters {f.SharedMasters}, followers {f.SharedFollowers}, plain {f.Plain}, array {f.Array}), "
                    + $"merged {s.MergedRanges.Count}, conditional formats {s.ConditionalFormats.Count}, validations {s.DataValidations.Count}, "
                    + $"tables {s.Tables.Count}{(s.AutoFilter is { } af ? ", autoFilter " + af : "")}");
            }
            var p = report.Package;
            Console.WriteLine($"package: calcChain {p.CalcChain} ({p.CalcChainEntries}), fullCalcOnLoad {p.FullCalcOnLoad}, customXml {p.CustomXmlItems.Count}, "
                + $"sensitivity label {(p.SensitivityLabelPart ?? "none")}, VBA {p.VbaProject}, external links {p.ExternalLinks.Count}, "
                + $"printer settings {p.PrinterSettings}, pivot caches {p.PivotCaches}");
            Console.WriteLine($"shared strings {report.SharedStringCount}, cell formats {report.CellXfsCount}, defined names {report.DefinedNames.Count}");
            return 0;
        }));
        return cmd;
    }

    private static Command Set()
    {
        var input = new Argument<FileInfo>("input") { Description = "Workbook to edit (not modified)" };
        var output = new Argument<FileInfo>("output") { Description = "Where to write the edited copy" };
        var edits = new Argument<string[]>("edits")
        {
            Description = "Sheet!Cell=value (repeatable). A value starting with '=' is a formula; numbers and TRUE/FALSE are typed unless --kind says otherwise",
            Arity = ArgumentArity.OneOrMore,
        };
        var kind = new Option<CellValueKind>("--kind") { Description = "Force the value type of every edit", DefaultValueFactory = _ => CellValueKind.Auto };
        var json = Output.JsonOption();
        var force = Output.ForceOption();
        var cmd = new Command("set", "Write cell values; removes calcChain and sets fullCalcOnLoad so Excel recalculates on open")
        {
            input, output, edits, kind, force, json,
        };
        cmd.SetAction(parse => Output.Guard(() =>
        {
            var k = parse.GetValue(kind);
            var list = parse.GetValue(edits)!.Select(e => Parse(e, k)).ToList();
            var outPath = parse.GetValue(output)!.FullName;
            Output.RequireNewOutput(outPath, parse.GetValue(force));
            var result = CellEditor.Apply(parse.GetValue(input)!.FullName, outPath, list);
            if (!result.Succeeded)
            {
                // Refused edits write nothing; exit 2 like every other refused operation.
                Output.Findings(result.Errors, parse.GetValue(json));
                return 2;
            }
            foreach (var c in result.Changes)
                Console.WriteLine($"{c.Sheet}!{c.Cell}: {(c.Previous is null ? "(empty)" : $"\"{c.Previous}\"")} -> {c.Kind} \"{c.Value}\"");
            Console.WriteLine(result.CalcChainRemoved ? "calcChain removed; Excel recalculates on open (cached values are stale until then)" : "no calcChain");
            Console.WriteLine($"output: {outPath}");
            return 0;
        }, partialOutput: parse.GetValue(output)?.FullName));
        return cmd;
    }

    private static CellEdit Parse(string spec, CellValueKind kind)
    {
        var bang = spec.IndexOf('!', StringComparison.Ordinal);
        var eq = bang < 0 ? -1 : spec.IndexOf('=', bang);
        if (bang <= 0 || eq <= bang + 1)
            throw new ArgumentException($"\"{spec}\": expected Sheet!Cell=value");
        return new CellEdit(spec[..bang], spec[(bang + 1)..eq], spec[(eq + 1)..], kind);
    }
}
