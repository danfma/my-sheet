using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Directed tests for the Phase 1 <see cref="RangeSnapshot.Build"/> materialization: the pre-sized column-major
/// array, its block-copy fast path over fully-present pages (with the per-page seqlock re-check), and the
/// per-cell fallback when a page is only partially present. The invariant is that the materialized array is
/// bit-for-bit what a plain per-cell expansion produces — the block-copy is a pure optimization, never a
/// semantic change — including under concurrent writes (no torn read) and after a mutation + InvalidateCache.
/// </summary>
public class RangeSnapshotBuildTests
{
    // === Store-level block-copy mechanics ================================================================

    [Test]
    public async Task BlockCopyColumn_FullyPresent_FillsColumnMajor_AndReturnsTrue()
    {
        // Spans several 1024-row pages so the per-page copy loop and cross-page offsets are exercised.
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 2;
        const int minRow = 500;
        const int maxRow = 3000; // pages 0,1,2 partially/fully covered

        for (var row = 1; row <= 4000; row++)
        {
            store.SetDense(handle, col, row, ComputedValue.Number(row), tainted: false);
        }

        var rowCount = maxRow - minRow + 1;
        var dest = new ComputedValue[rowCount];

        var copied = store.TryBlockCopyColumn(handle, col, minRow, maxRow, dest, 0);

        await Assert.That(copied).IsTrue();
        var mismatches = 0;
        for (var i = 0; i < rowCount; i++)
        {
            if (!dest[i].TryGetNumber(out var number) || number != minRow + i)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    [Test]
    public async Task BlockCopyColumn_WithAHole_ReturnsFalse_AndLeavesDestUntouched()
    {
        // One absent slot inside the covered range must reject the whole column (it would need on-demand
        // evaluation), leaving dest untouched for the caller's per-cell fallback.
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 1;

        for (var row = 1; row <= 300; row++)
        {
            if (row == 150)
            {
                continue; // the hole
            }

            store.SetDense(handle, col, row, ComputedValue.Number(row), tainted: false);
        }

        var dest = new ComputedValue[300];
        for (var i = 0; i < dest.Length; i++)
        {
            dest[i] = ComputedValue.Text("sentinel");
        }

        var copied = store.TryBlockCopyColumn(handle, col, 1, 300, dest, 0);

        await Assert.That(copied).IsFalse();
        // Untouched: every slot still the sentinel (no partial fill from a rejected column).
        var touched = 0;
        foreach (var value in dest)
        {
            if (!value.TryGetText(out var text) || text != "sentinel")
            {
                touched++;
            }
        }

        await Assert.That(touched).IsEqualTo(0);
    }

    [Test]
    public async Task BlockCopyColumn_ConcurrentWritesDuringCopy_NeverTears()
    {
        // A writer flips one slot in the range between Text and Number while readers block-copy the column. The
        // seqlock re-check must make every copied value internally consistent (never kind=Text with a stale ref).
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 1;
        const int rows = 2000;

        for (var row = 1; row <= rows; row++)
        {
            store.SetDense(handle, col, row, ComputedValue.Number(row), tainted: false);
        }

        var done = 0;
        var torn = 0;
        const int flipRow = 1234;

        var writer = Task.Run(() =>
        {
            for (var i = 0; i < 1_500_000; i++)
            {
                store.SetDense(handle, col, flipRow, ComputedValue.Text("hello"), tainted: false);
                store.SetDense(handle, col, flipRow, ComputedValue.Number(flipRow), tainted: false);
            }

            Volatile.Write(ref done, 1);
        });

        var readers = new Task[4];
        for (var t = 0; t < readers.Length; t++)
        {
            readers[t] = Task.Run(() =>
            {
                var dest = new ComputedValue[rows];
                while (Volatile.Read(ref done) == 0)
                {
                    if (!store.TryBlockCopyColumn(handle, col, 1, rows, dest, 0))
                    {
                        continue;
                    }

                    var flipped = dest[flipRow - 1];
                    var consistent =
                        (flipped.TryGetText(out var text) && text == "hello")
                        || (flipped.TryGetNumber(out var number) && number == flipRow);

                    if (!consistent)
                    {
                        Interlocked.Exchange(ref torn, 1);
                        return;
                    }
                }
            });
        }

        await Task.WhenAll(readers.Append(writer));

        await Assert.That(torn).IsEqualTo(0);
    }

    // === RangeSnapshot.Build equivalence (block-copy AND fallback) =======================================

    [Test]
    public async Task Build_FullyWarmedRectangle_MatchesPerCellBypass_InColumnMajorOrder()
    {
        // A 4-column x 2500-row rectangle, fully warmed (every cell present) -> the block-copy path builds the
        // whole snapshot. It must equal a per-cell expansion, in column-major order.
        var (workbook, sheet, range) = BuildRectangleWorkbook(warmAll: true);
        var context = new EvaluationContext(workbook);

        var snapshot = RangeSnapshot.Build(range, context);
        var expected = ExpectedColumnMajor(workbook, sheet, range);

        await Assert.That(snapshot.Count).IsEqualTo(expected.Count);
        await Assert.That(DescribeAll(snapshot.Values)).IsEqualTo(DescribeAll(expected.ToArray()));
    }

    [Test]
    public async Task Build_PartiallyWarmedRectangle_UsesFallback_AndMatchesBypass()
    {
        // Nothing pre-warmed: every column has absent slots at Build time, so block-copy rejects each column and
        // the per-cell fallback evaluates the cells on demand. The result must still equal the bypass exactly.
        var (workbook, sheet, range) = BuildRectangleWorkbook(warmAll: false);
        var context = new EvaluationContext(workbook);

        var snapshot = RangeSnapshot.Build(range, context);
        var expected = ExpectedColumnMajor(workbook, sheet, range);

        await Assert.That(snapshot.Count).IsEqualTo(expected.Count);
        await Assert.That(DescribeAll(snapshot.Values)).IsEqualTo(DescribeAll(expected.ToArray()));
    }

    [Test]
    public async Task Build_AfterMutationAndInvalidate_RebuildsWithNewValues()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");
        for (var row = 1; row <= 500; row++)
        {
            sheet[$"A{row}"] = new NumberValue(row);
        }

        var range = new RangeReference("A1", "A500", "Data");
        var context = new EvaluationContext(workbook);

        var first = RangeSnapshot.Build(range, context);
        await Assert
            .That(first.Values[149].TryGetNumber(out var before) && before == 150d)
            .IsTrue();

        // Mutate a cell and drop the value cache: the next Build must see the new value (block-copy reads the
        // repopulated page, not a stale slot).
        sheet["A150"] = new NumberValue(99999);
        workbook.InvalidateCache();

        var second = RangeSnapshot.Build(range, context);
        await Assert
            .That(second.Values[149].TryGetNumber(out var after) && after == 99999d)
            .IsTrue();
    }

    [Test]
    public async Task Build_MixedColumns_SomeFullSomePartial_MatchesBypass()
    {
        // Column A fully warmed (block-copy), column B left cold (fallback) — proves the per-column decision and
        // the column-major offsets line up when the two paths interleave.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");
        for (var row = 1; row <= 1500; row++)
        {
            sheet[$"A{row}"] = new NumberValue(row);
            sheet[$"B{row}"] = new StringValue($"b{row}");
        }

        // Warm only column A.
        for (var row = 1; row <= 1500; row++)
        {
            workbook.GetCellValue("Data", $"A{row}");
        }

        var range = new RangeReference("A1", "B1500", "Data");
        var context = new EvaluationContext(workbook);

        var snapshot = RangeSnapshot.Build(range, context);
        var expected = ExpectedColumnMajor(workbook, sheet, range);

        await Assert.That(DescribeAll(snapshot.Values)).IsEqualTo(DescribeAll(expected.ToArray()));
    }

    // === Helpers =========================================================================================

    private static (Workbook Workbook, Sheet Sheet, RangeReference Range) BuildRectangleWorkbook(
        bool warmAll
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");

        const int columns = 4;
        const int rows = 2500;
        for (var col = 1; col <= columns; col++)
        {
            var letter = (char)('A' + col - 1);
            for (var row = 1; row <= rows; row++)
            {
                // A mix of every kind so Number/Text/Boolean/Error/Blank all round-trip through the block copy.
                // row % 5 == 3 is left genuinely empty (a Blank cell inside the closed rectangle).
                switch (row % 5)
                {
                    case 0:
                        sheet[$"{letter}{row}"] = new StringValue($"{letter}{row}");
                        break;
                    case 1:
                        sheet[$"{letter}{row}"] = new BooleanValue(row % 2 == 0);
                        break;
                    case 2:
                        sheet[$"{letter}{row}"] = ExpressionParser.Parse("=1/0", sheet); // Error kind
                        break;
                    case 3:
                        break; // empty -> Blank
                    default:
                        sheet[$"{letter}{row}"] = new NumberValue(col * 100000 + row);
                        break;
                }
            }
        }

        if (warmAll)
        {
            for (var col = 1; col <= columns; col++)
            {
                var letter = (char)('A' + col - 1);
                for (var row = 1; row <= rows; row++)
                {
                    workbook.GetCellValue("Data", $"{letter}{row}");
                }
            }
        }

        return (workbook, sheet, new RangeReference("A1", $"D{rows}", "Data"));
    }

    // The oracle: expand the closed range cell-by-cell via the public GetCellValue in column-major order.
    private static List<ComputedValue> ExpectedColumnMajor(
        Workbook workbook,
        Sheet sheet,
        RangeReference range
    )
    {
        var start = CellAddress.Parse(range.StartId);
        var end = CellAddress.Parse(range.EndId);
        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        var values = new List<ComputedValue>();
        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                values.Add(
                    workbook.GetCellValue(range.SheetName, new CellAddress(column, row).ToId())
                );
            }
        }

        return values;
    }

    private static string DescribeAll(ComputedValue[] values)
    {
        var parts = new string[values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            var value = values[i];
            parts[i] = value.Kind switch
            {
                ComputedValueKind.Number => value.TryGetNumber(out var n) ? $"n:{n:R}" : "n:?",
                ComputedValueKind.Boolean => value.TryGetBoolean(out var b) ? $"b:{b}" : "b:?",
                ComputedValueKind.Text => value.TryGetText(out var t) ? $"t:{t}" : "t:?",
                ComputedValueKind.Blank => "blank",
                ComputedValueKind.Error => "err",
                _ => "ref",
            };
        }

        return string.Join("|", parts);
    }
}
