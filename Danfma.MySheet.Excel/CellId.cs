using System.Globalization;

namespace Danfma.MySheet.Excel;

/// <summary>Parses "A1"-style cell ids into 1-based (row, column) coordinates.</summary>
internal static class CellId
{
    public static (int Row, int Column) Parse(string id)
    {
        var index = 0;
        var column = 0;

        while (index < id.Length && char.IsLetter(id[index]))
        {
            column = column * 26 + (char.ToUpperInvariant(id[index]) - 'A' + 1);
            index++;
        }

        return (int.Parse(id.AsSpan(index), CultureInfo.InvariantCulture), column);
    }

    /// <summary>Formats 1-based (row, column) back into an "A1"-style id (inverse of Parse).</summary>
    public static string Format(int row, int column)
    {
        // Column letters build right-to-left (base-26, bijective); 7 chars cover int.MaxValue.
        Span<char> letters = stackalloc char[7];
        var start = letters.Length;

        while (column > 0)
        {
            column--;
            letters[--start] = (char)('A' + column % 26);
            column /= 26;
        }

        return string.Concat(letters[start..], row.ToString(CultureInfo.InvariantCulture));
    }
}
