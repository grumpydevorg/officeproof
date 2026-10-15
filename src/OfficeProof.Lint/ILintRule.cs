using OfficeProof.Core;

namespace OfficeProof.Lint;

/// <summary>
/// One read-only check over an open document. Rules never modify the package; they
/// inspect <see cref="LintContext"/> and return findings tagged with their <see cref="Id"/>.
/// </summary>
public interface ILintRule
{
    /// <summary>Stable kebab-case id, used for selection and as <see cref="Finding.Rule"/>.</summary>
    string Id { get; }

    /// <summary>The severity this rule's findings carry.</summary>
    Severity Severity { get; }

    /// <summary>One line saying what the rule catches, for rule listings.</summary>
    string Description { get; }

    /// <summary>Inspects the document and yields findings; an empty sequence means clean.</summary>
    IEnumerable<Finding> Check(LintContext ctx);
}
