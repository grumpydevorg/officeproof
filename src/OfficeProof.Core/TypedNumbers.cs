using System.Text.RegularExpressions;

namespace OfficeProof.Core;

/// <summary>
/// The one definition of "a number someone typed" that Lint reports, Redline converts and
/// Templating strips. Kept in Core so the three never disagree (2026-10-15 review: Lint
/// capped heading numbers at three digits while Redline did not, so "2025 Field season"
/// linted clean yet lost "2025 " to --headings).
/// </summary>
public static partial class TypedNumbers
{
    /// <summary>
    /// A heading that starts with a typed number plus its trailing whitespace: "2.0 Aims",
    /// "3.1\tFrog Pond", "4. Results". The first component is capped at three digits so a year
    /// ("2025 Field season") is not a number. Group "n" is the number; the match length is
    /// what to delete.
    /// </summary>
    [GeneratedRegex(@"^\s*(?<n>\d{1,3}(?:\.\d+)*)\.?[ \t\u00A0]+")]
    public static partial Regex HeadingNumber();

    /// <summary>
    /// A caption label at paragraph start: "Figure 1.", "Table 2:", "Figure 3 – ". Chapter-style
    /// numbers ("Figure 2.1") are not matched: they are not a single SEQ counter.
    /// </summary>
    [GeneratedRegex(@"^\s*(?<label>Figure|Table)\s+(?<n>\d+)(?!\.\d)(?:[.:]|\s+[–—-]\s)")]
    public static partial Regex Caption();

    /// <summary>Any paragraph opening with "Figure N"/"Table N" (not "Figure N.M"), for collecting caption numbers.</summary>
    [GeneratedRegex(@"^\s*(?<label>Figure|Table)\s+(?<n>\d+)(?!\.\d)\b")]
    public static partial Regex CaptionStart();

    /// <summary>A single reference in running text: "Figure 1", "figure 2", "Table 3" (not "Figure 2.1").</summary>
    [GeneratedRegex(@"\b(?<label>Figure|Table)\s+(?<n>\d+)(?!\.\d)\b", RegexOptions.IgnoreCase)]
    public static partial Regex Reference();

    /// <summary>
    /// A plural reference list: "Figures 7 and 9", "Tables 1, 2 and 4". Each number is a capture
    /// of group "n" (<c>match.Groups["n"].Captures</c>).
    /// </summary>
    [GeneratedRegex(@"\b(?<label>Figure|Table)s\s+(?<n>\d+)(?!\.\d)(?:\s*(?:,|and|&|or)\s*(?<n>\d+)(?!\.\d))+\b", RegexOptions.IgnoreCase)]
    public static partial Regex PluralReference();

    /// <summary>A typed section reference: "section 3.1", "Section 2.4.1" (at least one dot).</summary>
    [GeneratedRegex(@"\b[Ss]ection\s+(?<n>\d+(?:\.\d+)+)\b")]
    public static partial Regex SectionReference();

    /// <summary>Word's text for a REF whose bookmark is gone.</summary>
    public const string BrokenReference = "Error! Reference source not found";

    /// <summary>"Figure"/"Table" in canonical case, so "figure 1" and "Figure 1" compare equal.</summary>
    public static string Label(string label) =>
        string.Equals(label, "figure", StringComparison.OrdinalIgnoreCase) ? "Figure" : "Table";
}
