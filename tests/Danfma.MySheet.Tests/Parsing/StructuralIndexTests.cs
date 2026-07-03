using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// The 3.0 write-maintained, LIFETIME <see cref="SheetStructuralIndex"/> (whole-column scale): it is built
/// lazily on the sheet's FIRST open-range read (once per object life, proven by the build counter), then kept
/// current incrementally by the <see cref="Sheet.SetCell"/>/<see cref="Sheet.Remove"/> choke point. It SURVIVES
/// <see cref="Workbook.InvalidateCache"/> and <see cref="Workbook.Recalculate"/> (structure is orthogonal to
/// the value caches). An in-order write appends in O(1) with no re-sort; an out-of-order write dirties only its
/// bucket, which re-sorts on its next read; an overwrite leaves the index untouched; a delete removes the id in
/// place. A loaded workbook (MemoryPack populates the cells directly) rebuilds the index once on first read.
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
    // exactly one Layer-1 structural read.
    private static List<string> Read(Workbook workbook, OpenRangeReference range) =>
        range.PopulatedIds(new EvaluationContext(workbook)).ToList();

    private static OpenRangeReference WholeColumn(int column) =>
        OpenRangeReference.Create(column, column, null, null, "Sheet1");

    private static OpenRangeReference WholeRow(int row) =>
        OpenRangeReference.Create(null, null, row, row, "Sheet1");

    [Test]
    public async Task FirstRead_BuildsIndexOnce()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        // No index exists before the first open-range read.
        await Assert.That(sheet.PeekStructuralIndex()).IsNull();

        Read(workbook, WholeColumn(1));

        var index = sheet.PeekStructuralIndex();
        await Assert.That(index).IsNotNull();
        await Assert.That(index!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task Index_BuiltOnce_ReusedAcrossReads()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1));
        Read(workbook, WholeColumn(1));
        Read(workbook, WholeColumn(1));

        await Assert.That(sheet.PeekStructuralIndex()!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task Index_SurvivesInvalidateCache_BuiltOncePerLife()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // builds
        var before = sheet.PeekStructuralIndex();

        // N invalidations must NOT drop the (lifetime) index: same instance, still one build.
        for (var i = 0; i < 5; i++)
        {
            workbook.InvalidateCache();
            Read(workbook, WholeColumn(1));
        }

        var after = sheet.PeekStructuralIndex();
        await Assert.That(ReferenceEquals(before, after)).IsTrue();
        await Assert.That(after!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task Index_SurvivesRecalculate()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // builds
        var before = sheet.PeekStructuralIndex();

        workbook.Recalculate();
        Read(workbook, WholeColumn(1));

        var after = sheet.PeekStructuralIndex();
        await Assert.That(ReferenceEquals(before, after)).IsTrue();
        await Assert.That(after!.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task InOrderAppend_IsO1_StaysSorted_NoResort()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1)); // builds + sorts column A once
        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.ColumnSortCount).IsEqualTo(1);

        // Append below the current last row: an O(1) append that keeps the bucket sorted (no dirty).
        sheet["A4"] = Number(4);
        sheet["A5"] = Number(5);

        await Assert.That(index.ColumnAppendCount).IsEqualTo(2);
        await Assert.That(index.IsColumnSorted(1)).IsTrue();

        var read = Read(workbook, WholeColumn(1));
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A2,A3,A4,A5");
        // The bucket was already sorted, so the read did not re-sort it.
        await Assert.That(index.ColumnSortCount).IsEqualTo(1);
    }

    [Test]
    public async Task OutOfOrderInsert_DirtiesOnlyThatColumn_ResortsOnReadOfThatColumnOnly()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A3", Number(3)), ("B1", Number(10)));

        var openAToB = OpenRangeReference.Create(1, 2, null, null, "Sheet1"); // A:B
        Read(workbook, openAToB); // builds + sorts columns A and B
        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.ColumnSortCount).IsEqualTo(2);
        await Assert.That(index.IsColumnSorted(1)).IsTrue();
        await Assert.That(index.IsColumnSorted(2)).IsTrue();

        // Insert A2 out of order (row 2 < last row 3): column A goes dirty, column B is untouched.
        sheet["A2"] = Number(2);
        await Assert.That(index.ColumnAppendCount).IsEqualTo(0);
        await Assert.That(index.IsColumnSorted(1)).IsFalse();
        await Assert.That(index.IsColumnSorted(2)).IsTrue();

        var read = Read(workbook, openAToB);
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A2,A3,B1");
        // Only column A was re-sorted (2 → 3); column B was already sorted and was not touched.
        await Assert.That(index.ColumnSortCount).IsEqualTo(3);
    }

    [Test]
    public async Task NewCellInNewColumn_AfterBuild_IsIndexed()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        Read(workbook, WholeColumn(1)); // builds the column map

        sheet["C5"] = Number(5); // a brand-new column bucket
        var read = Read(workbook, WholeColumn(3));

        await Assert.That(string.Join(",", read)).IsEqualTo("C5");
        await Assert.That(sheet.PeekStructuralIndex()!.IsColumnSorted(3)).IsTrue(); // singleton is sorted
    }

    [Test]
    public async Task Overwrite_LeavesIndexUntouched()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // builds
        var index = sheet.PeekStructuralIndex()!;

        // Overwriting an EXISTING id is not a structural change: no append, no dirty, no duplicate.
        sheet["A1"] = Number(99);

        await Assert.That(index.ColumnAppendCount).IsEqualTo(0);
        await Assert.That(index.IsColumnSorted(1)).IsTrue();

        var read = Read(workbook, WholeColumn(1));
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A2");
    }

    [Test]
    public async Task Remove_UpdatesIndex_StaysSorted()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Read(workbook, WholeColumn(1)); // builds + sorts
        var index = sheet.PeekStructuralIndex()!;

        var removed = sheet.Remove("A2");
        await Assert.That(removed).IsTrue();
        await Assert.That(index.IsColumnSorted(1)).IsTrue(); // removal preserves order

        var read = Read(workbook, WholeColumn(1));
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A3");
        await Assert.That(index.ColumnSortCount).IsEqualTo(1); // no re-sort needed
    }

    [Test]
    public async Task Remove_MissingId_LeavesIndexUntouched()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Read(workbook, WholeColumn(1)); // builds

        var removed = sheet.Remove("A9"); // never existed
        await Assert.That(removed).IsFalse();

        var read = Read(workbook, WholeColumn(1));
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A2");
    }

    [Test]
    public async Task Remove_EmptyingColumn_DropsColumnKey()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(10)));

        Read(workbook, OpenRangeReference.Create(1, 2, null, null, "Sheet1")); // A:B builds both buckets
        var index = sheet.PeekStructuralIndex()!;

        sheet.Remove("A1"); // empties column A

        await Assert.That(index.ColumnKeys.Contains(1)).IsFalse();
        await Assert.That(index.ColumnKeys.Contains(2)).IsTrue();

        var read = Read(workbook, OpenRangeReference.Create(1, 2, null, null, "Sheet1"));
        await Assert.That(string.Join(",", read)).IsEqualTo("B1");
    }

    [Test]
    public async Task Maintenance_UpdatesBothMaps_WhenBothBuilt()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(2)));

        Read(workbook, WholeColumn(1)); // builds column map
        Read(workbook, WholeRow(1)); // builds row map
        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.ColumnBuildCount).IsEqualTo(1);
        await Assert.That(index.RowBuildCount).IsEqualTo(1);

        sheet["C1"] = Number(3); // new cell touches row 1 and column C

        var byColumn = Read(workbook, WholeColumn(3));
        var byRow = Read(workbook, WholeRow(1));
        await Assert.That(string.Join(",", byColumn)).IsEqualTo("C1");
        await Assert.That(string.Join(",", byRow)).IsEqualTo("A1,B1,C1");
    }

    [Test]
    public async Task WholeColumn_BuildsColumnMapOnly_NotRowMap()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(3)));

        Read(workbook, WholeColumn(1));

        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.ColumnBuildCount).IsEqualTo(1);
        await Assert.That(index.RowBuildCount).IsEqualTo(0);
    }

    [Test]
    public async Task WholeRow_BuildsRowMapOnly_NotColumnMap()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(999)));

        Read(workbook, WholeRow(1));

        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.RowBuildCount).IsEqualTo(1);
        await Assert.That(index.ColumnBuildCount).IsEqualTo(0);
    }

    [Test]
    public async Task WholeRow_YieldsRowThenColumnOrder()
    {
        var (workbook, _) = Sheet(
            ("C1", Number(30)),
            ("A1", Number(10)),
            ("B1", Number(20)),
            ("A2", Number(999))
        );

        var read = Read(workbook, WholeRow(1));
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,B1,C1");
    }

    [Test]
    public async Task Read_YieldsColumnThenRowOrder_IndependentOfInsertionOrder()
    {
        // Insert deliberately out of spatial order; the read must yield deterministic column-then-row order.
        var (workbook, _) = Sheet(
            ("B1", Number(10)),
            ("A3", Number(3)),
            ("A1", Number(1)),
            ("A2", Number(2)),
            ("B2", Number(20))
        );

        var read = Read(workbook, OpenRangeReference.Create(1, 2, null, null, "Sheet1")); // A:B
        await Assert.That(string.Join(",", read)).IsEqualTo("A1,A2,A3,B1,B2");
    }

    [Test]
    public async Task ColumnSort_IsLazy_UnreadColumnStaysUnsorted()
    {
        // Two populated columns; only column A is ever read, so only its bucket is sorted.
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("A2", Number(2)),
            ("B1", Number(3)),
            ("B2", Number(4))
        );

        Read(workbook, WholeColumn(1)); // builds the map + sorts column A only

        var index = sheet.PeekStructuralIndex()!;
        await Assert.That(index.IsColumnSorted(1)).IsTrue(); // A was read
        await Assert.That(index.IsColumnSorted(2)).IsFalse(); // B was bucketized but never read → unsorted
    }

    [Test]
    public async Task PostLoad_RebuildsLazyOncePerLife()
    {
        var (workbook, _) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        try
        {
            workbook.Save(path);
            var loaded = Workbook.Load(path);
            var loadedSheet = loaded["Sheet1"];

            // A loaded sheet has no index (MemoryPack bypasses field initializers and never serialized it).
            await Assert.That(loadedSheet.PeekStructuralIndex()).IsNull();

            var first = Read(loaded, WholeColumn(1));
            var second = Read(loaded, WholeColumn(1));

            await Assert.That(string.Join(",", first)).IsEqualTo("A1,A2,A3");
            await Assert.That(string.Join(",", second)).IsEqualTo("A1,A2,A3");
            await Assert.That(loadedSheet.PeekStructuralIndex()!.ColumnBuildCount).IsEqualTo(1);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
