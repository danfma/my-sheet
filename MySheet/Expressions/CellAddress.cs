using System.Text;

namespace MySheet.Expressions;

/// <summary>
/// A parsed A1-style cell address (1-based column and row), used to enumerate ranges.
/// </summary>
internal readonly record struct CellAddress(int Column, int Row)
{
    public static CellAddress Parse(string id)
    {
        var letters = 0;
        while (letters < id.Length && char.IsLetter(id[letters]))
        {
            letters++;
        }

        if (letters == 0 || letters == id.Length || !int.TryParse(id[letters..], out var row))
        {
            throw new FormatException($"Invalid cell reference '{id}'.");
        }

        var column = 0;
        for (var i = 0; i < letters; i++)
        {
            column = column * 26 + (char.ToUpperInvariant(id[i]) - 'A' + 1);
        }

        return new CellAddress(column, row);
    }

    public string ToId()
    {
        var builder = new StringBuilder();
        var column = Column;

        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            column = (column - 1) / 26;
        }

        return builder.Append(Row).ToString();
    }
}
