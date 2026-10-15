using OfficeProof.Core;

namespace OfficeProof.Spreadsheet;

/// <summary>How <see cref="CellEdit.Value"/> is stored.</summary>
public enum CellValueKind
{
    /// <summary>Excel's typing rules: a leading '=' is a formula, an invariant-culture number is a number,
    /// TRUE/FALSE (any case) is a boolean, anything else is a string.</summary>
    Auto,
    /// <summary>A shared string, even if it looks like a number or starts with '='.</summary>
    Text,
    /// <summary>An invariant-culture number ("1234.5", "-2e3"); no thousands separators.</summary>
    Number,
    /// <summary>TRUE/FALSE (any case) or 1/0.</summary>
    Boolean,
    /// <summary>A formula; the leading '=' is optional. Written without a cached value.</summary>
    Formula,
}

/// <summary>Set one cell. <paramref name="Cell"/> is an A1 reference ("B2"); <paramref name="Sheet"/> matches the tab name case-insensitively.</summary>
public sealed record CellEdit(string Sheet, string Cell, string Value, CellValueKind Kind = CellValueKind.Auto);

/// <summary>One cell as written.</summary>
/// <param name="Sheet">The sheet's tab name as stored.</param>
/// <param name="Cell">Normalised A1 reference.</param>
/// <param name="Kind">The resolved kind (never <see cref="CellValueKind.Auto"/>).</param>
/// <param name="Value">The stored value: a formula with its leading '=', a number in invariant form, TRUE/FALSE, or the string.</param>
/// <param name="Previous">The cell's previous content in the same form, or null when it was absent or empty.</param>
public sealed record CellChange(string Sheet, string Cell, CellValueKind Kind, string Value, string? Previous);

/// <summary>
/// Outcome of <see cref="CellEditor.Apply"/>. When <see cref="Errors"/> is non-empty nothing was written:
/// every edit is checked against the input before the output is created.
/// </summary>
/// <param name="Changes">The cells written, in edit order.</param>
/// <param name="Errors">Why the edits were refused (rules xlsx-edit-sheet, xlsx-edit-cell, xlsx-edit-value, xlsx-shared-formula, xlsx-array-formula).</param>
/// <param name="CalcChainRemoved">xl/calcChain.xml existed and was deleted.</param>
public sealed record EditResult(IReadOnlyList<CellChange> Changes, IReadOnlyList<Finding> Errors, bool CalcChainRemoved)
{
    public bool Succeeded => Errors.Count == 0;
}
