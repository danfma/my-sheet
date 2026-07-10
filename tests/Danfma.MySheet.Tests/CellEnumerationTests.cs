using NumberValue = Danfma.MySheet.Expressions.NumberValue;

namespace Danfma.MySheet.Tests;

/// <summary>
/// The public cell-enumeration surface: allocation-free numeric addresses (CellAddresses),
/// the (Id, Column, Row) convenience form (EnumerateCells), and the span-based id formatter
/// (CellRef.TryFormat) — together the alloc-free bulk-extraction pipeline.
/// </summary>
public class CellEnumerationTests
{
    private static Sheet BuildSheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");

        sheet["A1"] = new NumberValue(1);
        sheet["B10"] = new NumberValue(2);
        sheet["AB100"] = new NumberValue(3);

        return sheet;
    }

    [Test]
    public async Task CellAddresses_MatchesKeys_ThroughCellRefFormat()
    {
        var sheet = BuildSheet();
        var derived = new List<string>();

        foreach (var (column, row) in sheet.CellAddresses)
        {
            derived.Add(CellRef.Format(column, row));
        }

        await Assert.That(derived).IsEquivalentTo(sheet.Keys.ToList());
    }

    [Test]
    public async Task EnumerateCells_YieldsIdAndAddress_IncludingOverflow()
    {
        var sheet = BuildSheet();
        sheet["not an a1 id"] = new NumberValue(9); // overflow: host-chosen non-canonical key

        var seen = new Dictionary<string, (int Column, int Row)>();

        foreach (var (id, column, row) in sheet.EnumerateCells())
        {
            seen[id] = (column, row);
        }

        await Assert.That(seen.Count).IsEqualTo(4);
        await Assert.That(seen["B10"]).IsEqualTo((2, 10));
        await Assert.That(seen["AB100"]).IsEqualTo((28, 100));
        // Overflow ids have no numeric address, by contract.
        await Assert.That(seen["not an a1 id"]).IsEqualTo((0, 0));

        // The address collection, by contrast, lists only canonical cells.
        var addresses = 0;
        foreach (var _ in sheet.CellAddresses)
        {
            addresses++;
        }
        await Assert.That(addresses).IsEqualTo(3);
        await Assert.That(sheet.CellAddresses.Count).IsEqualTo(3);
    }

    [Test]
    public async Task CellRef_TryFormat_MatchesCanonicalIds()
    {
        Span<char> buffer = stackalloc char[18];

        var cases = new (int Column, int Row, string Expected)[]
        {
            (1, 1, "A1"),
            (2, 10, "B10"),
            (26, 3, "Z3"),
            (27, 3, "AA3"),
            (28, 100, "AB100"),
            (16_384, 1_048_576, "XFD1048576"),
        };

        // Collect results first (await and Span cannot mix in one frame).
        var results = new List<(string Expected, string Actual)>();

        foreach (var (column, row, expected) in cases)
        {
            var ok = CellRef.TryFormat(column, row, buffer, out var written);
            results.Add((expected, ok ? new string(buffer[..written]) : "<failed>"));
        }

        foreach (var (expected, actual) in results)
        {
            await Assert.That(actual).IsEqualTo(expected);
        }
    }

    [Test]
    public async Task CellRef_TryFormat_RejectsInvalidInput()
    {
        var tooSmall = CellRef.TryFormat(28, 100, new char[3], out var written);

        await Assert.That(tooSmall).IsFalse();
        await Assert.That(written).IsEqualTo(0);
        await Assert.That(CellRef.TryFormat(0, 1, new char[18], out _)).IsFalse();
        await Assert.That(CellRef.TryFormat(1, 0, new char[18], out _)).IsFalse();
        await Assert.That(() => CellRef.Format(0, 1)).Throws<ArgumentOutOfRangeException>();
    }
}
