using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeProof.Core;

public enum Severity { Info, Warning, Error }

/// <summary>
/// One result from any checking component (schema validation, lint, verify).
/// <paramref name="Rule"/> is a stable kebab-case id, e.g. "typed-heading-number".
/// <paramref name="Part"/> is the package part URI ("/word/document.xml"); <paramref name="Location"/>
/// is a human locator such as a paragraph index and text excerpt.
/// </summary>
public sealed record Finding(string Rule, Severity Severity, string Message, string? Part = null, string? Location = null);

/// <summary>Renders findings for people and machines; the CLI exit code follows <see cref="ExitCode"/>.</summary>
public static class FindingReport
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>1 when any finding is an error, else 0.</summary>
    public static int ExitCode(IEnumerable<Finding> findings) =>
        findings.Any(f => f.Severity == Severity.Error) ? 1 : 0;

    public static string ToJson(IEnumerable<Finding> findings) => JsonSerializer.Serialize(findings.ToList(), Json);

    public static void WriteText(TextWriter output, IReadOnlyCollection<Finding> findings)
    {
        if (findings.Count == 0)
        {
            output.WriteLine("no findings");
            return;
        }
        foreach (var f in findings)
        {
            var where = string.Join(" ", new[] { f.Part, f.Location }.Where(s => !string.IsNullOrEmpty(s)));
            output.WriteLine($"{Label(f.Severity)} {f.Rule}: {f.Message}{(where.Length > 0 ? $"  [{where}]" : "")}");
        }
        var e = findings.Count(f => f.Severity == Severity.Error);
        var w = findings.Count(f => f.Severity == Severity.Warning);
        output.WriteLine($"{e} error(s), {w} warning(s), {findings.Count - e - w} info");
    }

    private static string Label(Severity s) => s switch
    {
        Severity.Error => "ERROR",
        Severity.Warning => "WARN ",
        _ => "INFO ",
    };
}
