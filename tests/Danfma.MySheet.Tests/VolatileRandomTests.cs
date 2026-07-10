using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

// Phase F1 (volatile functions) — the RNG path: RAND and RANDBETWEEN. The RNG is a persistent Random seeded
// from Workbook.RandomSeed (fixed seed => the whole run is reproducible), advanced across epochs and never
// re-seeded, so successive epochs differ while the per-epoch cache keeps a single cell stable within a pass.
// RAND() is [0,1); RANDBETWEEN(bottom, top) is an inclusive integer draw (bottom>top => #NUM!). Golden facts
// (range, no-args, volatility) are from the official support.microsoft.com RAND/RANDBETWEEN pages
// (fetched 2026-07-02); bottom>top => #NUM! and non-integer truncation are page-silent design decisions.
public class VolatileRandomTests
{
    private static Workbook Seeded(int seed)
    {
        var workbook = new Workbook { RandomSeed = seed };
        workbook.Sheets.Add("Sheet1");
        return workbook;
    }

    private static double Cell(Workbook workbook, string id) =>
        workbook.GetCellValue("Sheet1", id).AsObject() is double d ? d : double.NaN;

    // --- RAND is [0, 1) ---

    [Test]
    public async Task Rand_IsInTheHalfOpenUnitInterval()
    {
        var workbook = Seeded(1);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RAND()", workbook["Sheet1"]);

        for (var i = 0; i < 200; i++)
        {
            var value = Cell(workbook, "A1");
            await Assert.That(value).IsGreaterThanOrEqualTo(0d);
            await Assert.That(value).IsLessThan(1d);
            workbook.Recalculate();
        }
    }

    // --- Per-epoch cache: stable within an epoch, fresh after Recalculate ---

    [Test]
    public async Task Rand_IsStableWithinAnEpochAndFreshAfterRecalculate()
    {
        var workbook = Seeded(2);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RAND()", workbook["Sheet1"]);

        var first = Cell(workbook, "A1");
        var firstAgain = Cell(workbook, "A1"); // same epoch, cached
        await Assert.That(firstAgain).IsEqualTo(first);

        workbook.Recalculate();
        var second = Cell(workbook, "A1");
        await Assert.That(second).IsNotEqualTo(first);
    }

    // --- Distinct draws for distinct cells in the same epoch ---

    [Test]
    public async Task TwoRandCellsInTheSameEpochDrawDifferentValues()
    {
        var workbook = Seeded(3);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RAND()", workbook["Sheet1"]);
        workbook["Sheet1"]["A2"] = ExpressionParser.Parse("=RAND()", workbook["Sheet1"]);

        await Assert.That(Cell(workbook, "A1")).IsNotEqualTo(Cell(workbook, "A2"));
    }

    // --- Reproducibility: a fixed seed replays the same sequence ---

    [Test]
    public async Task FixedSeed_ReplaysTheSameSequenceAcrossEpochs()
    {
        static double[] Draw(int seed)
        {
            var workbook = Seeded(seed);
            workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RAND()", workbook["Sheet1"]);

            var draws = new double[5];
            for (var i = 0; i < draws.Length; i++)
            {
                draws[i] = workbook.GetCellValue("Sheet1", "A1").AsObject() is double d
                    ? d
                    : double.NaN;
                workbook.Recalculate();
            }

            return draws;
        }

        await Assert.That(Draw(12345)).IsEquivalentTo(Draw(12345));
        // A different seed diverges.
        await Assert.That(Draw(12345)).IsNotEquivalentTo(Draw(999));
    }

    // --- RANDBETWEEN range, inclusivity, and the reversed-bounds error ---

    [Test]
    public async Task RandBetween_StaysWithinTheInclusiveRange()
    {
        var workbook = Seeded(7);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RANDBETWEEN(1,6)", workbook["Sheet1"]);

        for (var i = 0; i < 500; i++)
        {
            var value = Cell(workbook, "A1");
            await Assert.That(value).IsGreaterThanOrEqualTo(1d);
            await Assert.That(value).IsLessThanOrEqualTo(6d);
            await Assert.That(value).IsEqualTo(Math.Truncate(value)); // an integer
            workbook.Recalculate();
        }
    }

    [Test]
    public async Task RandBetween_HitsBothInclusiveEndpoints()
    {
        // Over enough epochs a fair {0,1} draw hits both ends; the set proves inclusivity of both bounds.
        var workbook = Seeded(11);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RANDBETWEEN(0,1)", workbook["Sheet1"]);

        var seen = new HashSet<double>();
        for (var i = 0; i < 200; i++)
        {
            seen.Add(Cell(workbook, "A1"));
            workbook.Recalculate();
        }

        await Assert.That(seen).IsEquivalentTo(new HashSet<double> { 0d, 1d });
    }

    [Test]
    public async Task RandBetween_DegenerateRangeReturnsTheSingleValue()
    {
        var workbook = Seeded(13);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RANDBETWEEN(5,5)", workbook["Sheet1"]);
        await Assert.That(Cell(workbook, "A1")).IsEqualTo(5d);
    }

    [Test]
    public async Task RandBetween_ReversedBoundsIsNum()
    {
        var workbook = Seeded(1);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse("=RANDBETWEEN(6,1)", workbook["Sheet1"]);
        await Assert
            .That(workbook.GetCellValue("Sheet1", "A1").AsObject())
            .IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task RandBetween_TruncatesNonIntegerBounds()
    {
        // Page-silent design decision (the plan's default): truncate toward zero, then draw inclusively.
        // RANDBETWEEN(1.9, 3.2) truncates to [1, 3].
        var workbook = Seeded(17);
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse(
            "=RANDBETWEEN(1.9,3.2)",
            workbook["Sheet1"]
        );

        for (var i = 0; i < 200; i++)
        {
            var value = Cell(workbook, "A1");
            await Assert.That(value).IsGreaterThanOrEqualTo(1d);
            await Assert.That(value).IsLessThanOrEqualTo(3d);
            workbook.Recalculate();
        }
    }

    // --- Reproducibility for RANDBETWEEN too ---

    [Test]
    public async Task RandBetween_FixedSeedIsReproducible()
    {
        static double[] Draw(int seed)
        {
            var workbook = Seeded(seed);
            workbook["Sheet1"]["A1"] = ExpressionParser.Parse(
                "=RANDBETWEEN(1,1000000)",
                workbook["Sheet1"]
            );

            var draws = new double[6];
            for (var i = 0; i < draws.Length; i++)
            {
                draws[i] = workbook.GetCellValue("Sheet1", "A1").AsObject() is double d
                    ? d
                    : double.NaN;
                workbook.Recalculate();
            }

            return draws;
        }

        await Assert.That(Draw(2024)).IsEquivalentTo(Draw(2024));
    }

    // --- Introspection ---

    [Test]
    public async Task RandNodes_ReportIsVolatile()
    {
        await Assert.That(new Rand([]).IsVolatile).IsTrue();
        await Assert
            .That(new RandBetween([new NumberValue(1), new NumberValue(6)]).IsVolatile)
            .IsTrue();
    }

    // --- Serialization: the new union tags survive a MemoryPack round-trip ---

    [Test]
    public async Task RandFormulas_RoundTripThroughMemoryPack()
    {
        var workbook = new Workbook { RandomSeed = 5 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=RAND()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=RANDBETWEEN(1,6)", sheet);

        var path = Path.Combine(Path.GetTempPath(), $"volatile-rng-{Guid.NewGuid():N}.msgpack.bin");
        try
        {
            workbook.Save(path);
            var reloaded = Workbook.Load(path);
            reloaded.RandomSeed = 5;

            var rand = reloaded.GetCellValue("Sheet1", "A1").AsObject() as double?;
            var between = reloaded.GetCellValue("Sheet1", "A2").AsObject() as double?;

            await Assert.That(rand is >= 0d and < 1d).IsTrue();
            await Assert.That(between is >= 1d and <= 6d).IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
