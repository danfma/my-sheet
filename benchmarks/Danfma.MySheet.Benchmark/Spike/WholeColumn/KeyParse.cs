using System.Globalization;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumn;

// Reimplements the parsing logic of the core `CellAddress.Parse` and the Excel `CellId.Parse`
// (both are `internal`, so the benchmark project cannot call them directly). Kept byte-faithful so
// the spike measures the same work the engine already does on the hot path.
internal static class KeyParse
{
    // Faithful mirror of CellAddress.Parse: the `id[letters..]` slice allocates a substring and
    // int.Parse(string) runs over it. This is what a naive "parse every key" scan pays today.
    public static (int Col, int Row) ParseAlloc(string id)
    {
        var letters = 0;
        while (letters < id.Length && char.IsLetter(id[letters]))
        {
            letters++;
        }

        var row = int.Parse(id[letters..], CultureInfo.InvariantCulture); // substring allocation
        var col = 0;
        for (var i = 0; i < letters; i++)
        {
            col = col * 26 + (char.ToUpperInvariant(id[i]) - 'A' + 1);
        }

        return (col, row);
    }

    // Same result, zero allocation: parses the row from a span slice instead of a substring.
    public static (int Col, int Row) ParseSpan(string id)
    {
        var span = id.AsSpan();
        var letters = 0;
        var col = 0;
        while (letters < span.Length && char.IsLetter(span[letters]))
        {
            col = col * 26 + (char.ToUpperInvariant(span[letters]) - 'A' + 1);
            letters++;
        }

        var row = int.Parse(span[letters..], provider: CultureInfo.InvariantCulture);
        return (col, row);
    }

    // Column only, no-alloc, for whole-column scan filtering (stops before the digits).
    public static int ColumnOf(string id)
    {
        var col = 0;
        var i = 0;
        while (i < id.Length && char.IsLetter(id[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(id[i]) - 'A' + 1);
            i++;
        }

        return col;
    }

    // Row only, no-alloc, for whole-row scan filtering.
    public static int RowOf(string id)
    {
        var i = 0;
        while (i < id.Length && char.IsLetter(id[i]))
        {
            i++;
        }

        return int.Parse(id.AsSpan(i), provider: CultureInfo.InvariantCulture);
    }

    // Mirror of CellAddress.ToId: builds an "A1" style id from 1-based (col, row).
    public static string ToId(int col, int row)
    {
        Span<char> letters = stackalloc char[8];
        var length = 0;
        var c = col;
        while (c > 0)
        {
            var remainder = (c - 1) % 26;
            letters[length++] = (char)('A' + remainder);
            c = (c - 1) / 26;
        }

        letters[..length].Reverse();
        return string.Concat(letters[..length], row.ToString(CultureInfo.InvariantCulture));
    }
}
