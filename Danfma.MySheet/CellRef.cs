using System.Globalization;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// A1-style id formatting for hosts that work with numeric addresses (see
/// <see cref="Workbook.GetValueReader"/> and <see cref="Sheet.CellAddresses"/>). The output is the
/// canonical id form the workbook itself uses: uppercase column letters, no <c>$</c>, no leading
/// zeros — <c>(2, 10)</c> → <c>"B10"</c>. (The instance half of this struct — (Sheet, Id) — is the
/// edit/recalculation currency; see RecalculationEngine.)
/// </summary>
public readonly partial record struct CellRef
{
    /// <summary>
    /// Formats the 1-based (column, row) address into <paramref name="destination"/> without
    /// allocating — the piece that makes an extraction pipeline fully allocation-free: format into a
    /// stack buffer and hand the span straight to a writer (e.g. <c>Utf8JsonWriter</c> accepts
    /// <c>ReadOnlySpan&lt;char&gt;</c>). Returns <c>false</c> when the address is out of range
    /// (column/row &lt; 1) or the buffer is too small; 18 chars always suffice.
    /// </summary>
    public static bool TryFormat(int column, int row, Span<char> destination, out int charsWritten)
    {
        charsWritten = 0;

        if (column < 1 || row < 1)
        {
            return false;
        }

        // Column letters build right-to-left (base-26, bijective); 7 chars cover int.MaxValue.
        Span<char> letters = stackalloc char[7];
        var start = letters.Length;

        while (column > 0)
        {
            column--;
            letters[--start] = (char)('A' + column % 26);
            column /= 26;
        }

        var letterCount = letters.Length - start;

        if (
            destination.Length < letterCount
            || !row.TryFormat(
                destination[letterCount..],
                out var digits,
                default,
                CultureInfo.InvariantCulture
            )
        )
        {
            return false;
        }

        letters[start..].CopyTo(destination);
        charsWritten = letterCount + digits;

        return true;
    }

    /// <summary>Formats the 1-based (column, row) address as a new string, e.g. <c>"B10"</c>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">column or row is less than 1.</exception>
    public static string Format(int column, int row)
    {
        if (column < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column), column, "Columns are 1-based.");
        }

        if (row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(row), row, "Rows are 1-based.");
        }

        return new CellAddress(column, row).ToId();
    }
}
