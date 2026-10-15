using System.CommandLine;
using OfficeProof.Core;

namespace OfficeProof.Cli;

/// <summary>Printing and option helpers shared by every verb.</summary>
internal static class Output
{
    public static Option<bool> JsonOption() => new("--json") { Description = "Print findings as JSON instead of text" };

    public static Option<bool> ForceOption() => new("--force") { Description = "Overwrite the output file if it exists" };

    /// <summary>
    /// Writing verbs never replace an existing file unless asked: a failed run deletes its partial
    /// output, and that must never be someone's earlier result (2026-10-15 review: a failed
    /// redline deleted a pre-existing unrelated file at the output path).
    /// </summary>
    public static void RequireNewOutput(string path, bool force)
    {
        if (!force && File.Exists(path))
            throw new ArgumentException($"{path} exists; pass --force to overwrite it");
    }

    /// <summary>Prints findings in the chosen format and returns the exit code they imply.</summary>
    public static int Findings(IReadOnlyCollection<Finding> findings, bool json)
    {
        if (json) Console.WriteLine(FindingReport.ToJson(findings));
        else FindingReport.WriteText(Console.Out, findings);
        return FindingReport.ExitCode(findings);
    }

    public static void Heading(string text) => Console.WriteLine($"== {text}");

    public static void List(string label, IEnumerable<string> items)
    {
        foreach (var item in items) Console.WriteLine($"  {label}: {item}");
    }

    /// <summary>
    /// Runs <paramref name="action"/> and turns expected failures (bad input, refused edits)
    /// into a one-line error and exit code 2 instead of a stack trace.
    /// </summary>
    public static int Guard(Func<int> action, string? partialOutput = null)
    {
        // Only a file this run wrote may be removed: if the output path is the input (refused
        // by Docx.OpenCopy before any write) or an earlier result, its timestamp is unchanged.
        DateTime? before = partialOutput is not null && File.Exists(partialOutput) ? File.GetLastWriteTimeUtc(partialOutput) : null;
        try
        {
            var code = action();
            if (partialOutput is not null && File.Exists(partialOutput) && File.GetLastWriteTimeUtc(partialOutput) != before)
                Docx.NormalizeEntryModes(partialOutput);
            return code;
        }
        catch (Exception e) when (e is ArgumentException or InvalidOperationException or InvalidDataException
                                      or FormatException or IOException or System.Text.Json.JsonException
                                      or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
                                      or OfficeProof.Verify.WordOracleException)
        {
            // A failed edit can leave a half-written copy; never leave it looking like a result.
            if (partialOutput is not null && File.Exists(partialOutput) && File.GetLastWriteTimeUtc(partialOutput) != before)
                File.Delete(partialOutput);
            Console.Error.WriteLine($"officeproof: {e.Message}");
            return 2;
        }
    }
}
