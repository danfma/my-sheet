using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Dates;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

// Phase F1 (volatile functions) — the clock path: TODAY/NOW plus the epoch cache model. Time comes from an
// injected FixedTimeProvider so the values are reproducible and the LOCAL-time conversion is exercised.
// NOW() is the local wall-clock rendered as an Excel serial (OADate); TODAY() is its whole-day floor.
public class VolatileClockTests
{
    private const double Tolerance = 1e-9;

    // A zone three hours behind UTC, so a UTC instant and its local rendering differ — proving NOW()/TODAY()
    // use local time, not UTC.
    private static readonly TimeZoneInfo MinusThree = TimeZoneInfo.CreateCustomTimeZone(
        "UTC-03 (test)", TimeSpan.FromHours(-3), "UTC-03 (test)", "UTC-03 (test)");

    private static (Workbook Workbook, FixedTimeProvider Clock, Sheet Sheet) NewWorkbook(DateTimeOffset utcNow)
    {
        var clock = new FixedTimeProvider(utcNow, MinusThree);
        var workbook = new Workbook { TimeProvider = clock };
        var sheet = workbook.Sheets.Add("Sheet1");
        return (workbook, clock, sheet);
    }

    private static double Num(ComputedValue value) => value.AsObject() is double d ? d : double.NaN;

    private static double Cell(Workbook workbook, string id) => Num(workbook.GetCellValue("Sheet1", id));

    // --- Golden values against the injected clock ---

    [Test]
    public async Task Now_IsTheLocalClockAsAnExcelSerial()
    {
        // UTC 13:30 in a UTC-3 zone is local 10:30 on the same day.
        var (workbook, _, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);

        var expected = new DateTime(2026, 7, 2, 10, 30, 0).ToOADate();
        await Assert.That(Cell(workbook, "A1")).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Today_IsTheWholeDayFloorOfNow()
    {
        var (workbook, _, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=TODAY()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=NOW()", sheet);

        var expected = new DateTime(2026, 7, 2).ToOADate();
        await Assert.That(Cell(workbook, "A1")).IsEqualTo(expected).Within(Tolerance);
        // TODAY() == FLOOR(NOW()) within the same epoch.
        await Assert.That(Cell(workbook, "A1")).IsEqualTo(Math.Floor(Cell(workbook, "A2"))).Within(Tolerance);
    }

    // --- Epoch coherence ---

    [Test]
    public async Task WithoutRecalculate_TheValueIsStableEvenIfTheClockMoves()
    {
        var (workbook, clock, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);

        var first = Cell(workbook, "A1");
        clock.Advance(TimeSpan.FromDays(1)); // clock moves, but no epoch advance
        var second = Cell(workbook, "A1");

        await Assert.That(second).IsEqualTo(first).Within(Tolerance);
    }

    [Test]
    public async Task TwoVolatileCellsAgreeWithinAnEpoch()
    {
        var (workbook, _, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=NOW()", sheet);

        await Assert.That(Cell(workbook, "A1")).IsEqualTo(Cell(workbook, "A2")).Within(Tolerance);
    }

    // --- Recalculate advances the epoch ---

    [Test]
    public async Task Recalculate_ResamplesTheClock()
    {
        var (workbook, clock, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);

        var first = Cell(workbook, "A1");
        clock.Advance(TimeSpan.FromDays(1));
        workbook.Recalculate();
        var second = Cell(workbook, "A1");

        await Assert.That(second).IsEqualTo(first + 1d).Within(Tolerance);
    }

    [Test]
    public async Task Recalculate_RefreshesDependentsOfVolatiles_Contagion()
    {
        var (workbook, clock, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet); // volatile
        sheet["B1"] = ExpressionParser.Parse("=A1+1", sheet); // transitively volatile (references A1)

        var firstB = Cell(workbook, "B1");
        clock.Advance(TimeSpan.FromDays(1));
        workbook.Recalculate();
        var secondB = Cell(workbook, "B1");

        // B1 was cached-and-marked because it transitively touched a volatile, so Recalculate dropped it.
        await Assert.That(secondB).IsEqualTo(firstB + 1d).Within(Tolerance);
    }

    // --- Recalculate leaves the stable (non-volatile) cache untouched ---

    [Test]
    public async Task Recalculate_DoesNotRecomputeNonVolatileCells()
    {
        var (workbook, _, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));

        var evaluations = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                evaluations++;
                return 42d;
            });

        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet); // NOT volatile
        sheet["V1"] = ExpressionParser.Parse("=NOW()", sheet); // volatile, so the epoch has something to clear

        _ = Cell(workbook, "A1"); // evaluations == 1, now cached
        _ = Cell(workbook, "V1");
        workbook.Recalculate();
        _ = Cell(workbook, "A1"); // still cached: TICK must NOT run again

        await Assert.That(evaluations).IsEqualTo(1);
    }

    [Test]
    public async Task InvalidateCache_AlsoResamplesTheClock()
    {
        var (workbook, clock, sheet) = NewWorkbook(new DateTimeOffset(2026, 7, 2, 13, 30, 0, TimeSpan.Zero));
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);

        var first = Cell(workbook, "A1");
        clock.Advance(TimeSpan.FromDays(2));
        workbook.InvalidateCache();
        var second = Cell(workbook, "A1");

        await Assert.That(second).IsEqualTo(first + 2d).Within(Tolerance);
    }

    // --- Introspection: IsVolatile ---

    [Test]
    public async Task VolatileNodes_ReportIsVolatile()
    {
        await Assert.That(new Now([]).IsVolatile).IsTrue();
        await Assert.That(new Today([]).IsVolatile).IsTrue();
        await Assert.That(new NumberValue(1).IsVolatile).IsFalse();
        await Assert.That(new Year([new NumberValue(1)]).IsVolatile).IsFalse();
    }

    // --- Serialization: the new union tags survive a MemoryPack round-trip ---

    [Test]
    public async Task VolatileFormula_RoundTripsThroughMemoryPack()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=TODAY()", sheet);

        var path = Path.Combine(Path.GetTempPath(), $"volatile-{Guid.NewGuid():N}.msgpack.bin");
        try
        {
            workbook.Save(path);
            var reloaded = Workbook.Load(path);

            // A fixed clock on the reloaded workbook proves NOW/TODAY deserialized to the right nodes.
            reloaded.TimeProvider = new FixedTimeProvider(
                new DateTimeOffset(2026, 7, 2, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);

            var now = Num(reloaded.GetCellValue("Sheet1", "A1"));
            var today = Num(reloaded.GetCellValue("Sheet1", "A2"));

            await Assert.That(now).IsEqualTo(new DateTime(2026, 7, 2, 12, 0, 0).ToOADate()).Within(Tolerance);
            await Assert.That(today).IsEqualTo(new DateTime(2026, 7, 2).ToOADate()).Within(Tolerance);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
