using System.Globalization;

namespace OfficeProof.Spreadsheet;

/// <summary>A single-cell A1 reference, 1-based. Excel's grid is A1:XFD1048576.</summary>
internal readonly record struct CellAddress(int Column, int Row)
{
    public const int MaxColumn = 16384;
    public const int MaxRow = 1048576;

    /// <summary>Parses "B2" (case-insensitive, no '$', no sheet prefix).</summary>
    public static bool TryParse(string? text, out CellAddress address)
    {
        address = default;
        if (string.IsNullOrEmpty(text)) return false;
        var i = 0;
        var column = 0;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            if (column > MaxColumn) return false;
            i++;
        }
        if (i == 0 || i == text.Length || text[i] == '0') return false;
        if (!int.TryParse(text.AsSpan(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row > MaxRow)
            return false;
        address = new CellAddress(column, row);
        return true;
    }

    public static string ColumnName(int column)
    {
        Span<char> buffer = stackalloc char[3];
        var at = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--at] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(buffer[at..]);
    }

    public override string ToString() => ColumnName(Column) + Row.ToString(CultureInfo.InvariantCulture);
}

/// <summary>A rectangular range "A1:U262" (or a single cell "A1").</summary>
internal readonly record struct CellRange(CellAddress First, CellAddress Last)
{
    public static bool TryParse(string? text, out CellRange range)
    {
        range = default;
        if (string.IsNullOrEmpty(text)) return false;
        var colon = text.IndexOf(':', StringComparison.Ordinal);
        if (colon < 0)
        {
            if (!CellAddress.TryParse(text, out var single)) return false;
            range = new CellRange(single, single);
            return true;
        }
        if (!CellAddress.TryParse(text[..colon], out var a) || !CellAddress.TryParse(text[(colon + 1)..], out var b))
            return false;
        range = new CellRange(a, b);
        return true;
    }

    public CellRange Including(CellAddress cell) => new(
        new CellAddress(Math.Min(First.Column, cell.Column), Math.Min(First.Row, cell.Row)),
        new CellAddress(Math.Max(Last.Column, cell.Column), Math.Max(Last.Row, cell.Row)));

    public bool Contains(CellAddress cell) =>
        First.Column <= cell.Column && cell.Column <= Last.Column && First.Row <= cell.Row && cell.Row <= Last.Row;

    public override string ToString() => First == Last ? First.ToString() : $"{First}:{Last}";
}
