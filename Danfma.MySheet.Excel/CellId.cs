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

        return (int.Parse(id[index..], CultureInfo.InvariantCulture), column);
    }

    /// <summary>Formats 1-based (row, column) back into an "A1"-style id (inverse of Parse).</summary>
    public static string Format(int row, int column)
    {
        var letters = string.Empty;

        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }

        return letters + row.ToString(CultureInfo.InvariantCulture);
    }
}
