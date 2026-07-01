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
}
