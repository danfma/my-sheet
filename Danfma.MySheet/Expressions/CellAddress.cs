using System.Text;

namespace Danfma.MySheet.Expressions;

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

    /// <summary>
    /// Reads the (1-based) column and row out of a cell id WITHOUT allocating — no substring, no
    /// <see cref="int.Parse(string)"/>. This is the hot extractor used by the whole-column/row NaiveScan
    /// (see <c>OpenRangeReference</c>): the benchmark proved the substring-based <see cref="Parse"/> is
    /// 4.6× slower and allocates ~3.2 MB per scan of 100k cells. Returns <c>false</c> for a malformed id.
    /// </summary>
    public static bool TryGetColumnRow(string id, out int column, out int row)
    {
        column = 0;
        row = 0;

        var i = 0;
        while (i < id.Length && char.IsLetter(id[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(id[i]) - 'A' + 1);
            i++;
        }

        // Need at least one letter and at least one following digit.
        if (i == 0 || i == id.Length)
        {
            return false;
        }

        var parsedRow = 0;
        for (; i < id.Length; i++)
        {
            var c = id[i];
            if (c is < '0' or > '9')
            {
                return false;
            }

            parsedRow = parsedRow * 10 + (c - '0');
        }

        row = parsedRow;
        return true;
    }

    /// <summary>
    /// Parses an all-letters column label (e.g. <c>A</c>, <c>AB</c>) to its 1-based column number, stripping
    /// any absolute markers (<c>$</c>). Returns <c>false</c> when the text is empty or holds a non-letter.
    /// </summary>
    public static bool TryParseColumn(string label, out int column)
    {
        column = 0;

        var seen = false;
        foreach (var raw in label)
        {
            if (raw == '$')
            {
                continue;
            }

            if (!char.IsLetter(raw))
            {
                column = 0;
                return false;
            }

            column = column * 26 + (char.ToUpperInvariant(raw) - 'A' + 1);
            seen = true;
        }

        return seen;
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
