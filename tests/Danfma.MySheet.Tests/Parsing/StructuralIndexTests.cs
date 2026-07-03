using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// The Layer-1 <see cref="SheetStructuralIndex"/> (whole-column scale) under the Phase-5 second-use admission:
/// a sheet's FIRST open-range read of an epoch NaiveScans (no index build); the SECOND builds the index once,
/// which every read after reuses. The index is dropped by <see cref="Workbook.InvalidateCache"/>, survives
/// <see cref="Workbook.Recalculate"/>, enumerates in column-then-row order down BOTH paths, and sorts each
/// column/row bucket lazily on first access.
/// </summary>
public class StructuralIndexTests
{
    private static (Workbook Workbook, Sheet Sheet) Sheet(params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return (workbook, sheet);
    }

    // Reads a whole-column/row open range's ids directly (bypasses the Layer-2 range cache), so each call is
    // exactly one Layer-1 read — the unit the admission policy counts.
    private static List<string> Read(Workbook workbook, OpenRangeReference range) =>
        range.PopulatedIds(new EvaluationContext(workbook)).ToList();

    private static OpenRangeReference WholeColumn(int column) =>
        OpenRangeReference.Create(column, column, null, null, "Sheet1");

    private static OpenRangeReference WholeRow(int row) =>
        OpenRangeReference.Create(null, null, row, row, "Sheet1");

    [Test]
    public async Task FirstRead_NaiveScans_DoesNotBuildIndex()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1));

        // No index object exists yet: the first read served itself from a NaiveScan.
        await Assert.That(workbook.PeekStructuralIndex("Sheet1")).IsNull();
    }

    [Test]
    public async Task SecondRead_BuildsIndexOnce()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds the index

        var index = workbook.PeekStructuralIndex("Sheet1");
        await Assert.That(index).IsNotNull();
        await Assert.That(index!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task Index_BuiltOnce_ReusedAcrossReads()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds
        Read(workbook, WholeColumn(1)); // reuses
        Read(workbook, WholeColumn(1)); // reuses

        await Assert.That(workbook.PeekStructuralIndex("Sheet1")!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task NaiveScan_And_Index_YieldIdenticalOrder()
    {
        // Insert deliberately out of spatial order; both the NaiveScan (first read) and the index (later
        // reads) must yield the SAME deterministic column-then-row order.
        var (workbook, _) = Sheet(
            ("B1", Number(10)),
            ("A3", Number(3)),
            ("A1", Number(1)),
            ("A2", Number(2)),
            ("B2", Number(20))
        );

        var openAToB = OpenRangeReference.Create(1, 2, null, null, "Sheet1"); // A:B

        var firstRead = Read(workbook, openAToB); // NaiveScan
        var secondRead = Read(workbook, openAToB); // index build
        var thirdRead = Read(workbook, openAToB); // index reuse

        await Assert.That(string.Join(",", firstRead)).IsEqualTo("A1,A2,A3,B1,B2");
        await Assert.That(string.Join(",", secondRead)).IsEqualTo("A1,A2,A3,B1,B2");
        await Assert.That(string.Join(",", thirdRead)).IsEqualTo("A1,A2,A3,B1,B2");
    }

    [Test]
    public async Task WholeRow_NaiveScan_And_Index_YieldRowThenColumnOrder()
    {
        var (workbook, _) = Sheet(
            ("C1", Number(30)),
            ("A1", Number(10)),
            ("B1", Number(20)),
            ("A2", Number(999))
        );

        var openRow1 = WholeRow(1);

        var firstRead = Read(workbook, openRow1); // NaiveScan
        var secondRead = Read(workbook, openRow1); // index build

        await Assert.That(string.Join(",", firstRead)).IsEqualTo("A1,B1,C1");
        await Assert.That(string.Join(",", secondRead)).IsEqualTo("A1,B1,C1");
    }

    [Test]
    public async Task Index_DroppedByInvalidateCache()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds
        await Assert.That(workbook.PeekStructuralIndex("Sheet1")).IsNotNull();

        workbook.InvalidateCache();

        // The index AND the seen markers are gone: the next read NaiveScans again (first read of a new epoch).
        await Assert.That(workbook.PeekStructuralIndex("Sheet1")).IsNull();
        Read(workbook, WholeColumn(1));
        await Assert.That(workbook.PeekStructuralIndex("Sheet1")).IsNull();
    }

    [Test]
    public async Task Index_SurvivesRecalculate()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds
        var before = workbook.PeekStructuralIndex("Sheet1");

        workbook.Recalculate();
        var after = workbook.PeekStructuralIndex("Sheet1");

        await Assert.That(ReferenceEquals(before, after)).IsTrue();

        // The surviving instance is not rebuilt on the next read.
        Read(workbook, WholeColumn(1));
        await Assert.That(after!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task WholeColumn_BuildsColumnMapOnly_NotRowMap()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(3)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds the column map

        var index = workbook.PeekStructuralIndex("Sheet1")!;
        await Assert.That(index.ColumnBuildCount).IsEqualTo(1);
        await Assert.That(index.RowBuildCount).IsEqualTo(0);
    }

    [Test]
    public async Task WholeRow_BuildsRowMapOnly_NotColumnMap()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(999)));

        Read(workbook, WholeRow(1)); // NaiveScan
        Read(workbook, WholeRow(1)); // builds the row map

        var index = workbook.PeekStructuralIndex("Sheet1")!;
        await Assert.That(index.RowBuildCount).IsEqualTo(1);
        await Assert.That(index.ColumnBuildCount).IsEqualTo(0);
    }

    [Test]
    public async Task ColumnSort_IsLazy_UnreadColumnStaysUnsorted()
    {
        // Two populated columns; only column A is ever read, so only its bucket is sorted.
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("B1", Number(3)), ("B2", Number(4)));

        Read(workbook, WholeColumn(1)); // NaiveScan
        Read(workbook, WholeColumn(1)); // builds + sorts column A only

        var index = workbook.PeekStructuralIndex("Sheet1")!;
        await Assert.That(index.IsColumnSorted(1)).IsTrue(); // A was read
        await Assert.That(index.IsColumnSorted(2)).IsFalse(); // B was bucketized but never read → unsorted
    }
}
