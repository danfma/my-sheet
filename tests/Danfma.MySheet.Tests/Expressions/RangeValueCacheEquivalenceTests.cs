using System.Globalization;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Differential proof for the Layer-2 range value cache (whole-column scale, Phase 2; second-use admission,
/// Phase 4): the SAME battery of whole-column formulas must produce IDENTICAL results with the cache BYPASSED
/// (the pre-cache linear paths) and across the THREE reads of the admission policy — read 1 (linear fallback,
/// the range is only marked), read 2 (the snapshot is built) and read 3 (the built snapshot is reused) —
/// across data that stresses every derived accelerator and every documented fall-back (unsorted, duplicates,
/// mixed types, errors mid-range, blanks, zeros/empty text). The cache changes performance, never semantics.
/// </summary>
public class RangeValueCacheEquivalenceTests
{
    private const int Rows = 300; // comfortably above the 256-cell cache threshold

    private enum Scenario
    {
        AscendingNumbers,
        NumbersWithDuplicates,
        UnsortedNumbers,
        MixedTypes,
        WithErrors,
        WithZerosAndBlanks,
        TextKeys,
    }

    // The battery is deliberately broad: exact/approximate/wildcard/criteria/order-stat/aggregate, plus
    // misses, blank-equivalent lookups and separate sum ranges — every wired consumer and fall-back.
    private static readonly string[] Formulas =
    [
        "=MATCH(50,Data!A:A,0)",
        "=MATCH(\"k50\",Data!C:C,0)",
        "=MATCH(50,Data!A:A,1)",
        "=MATCH(50.5,Data!A:A,1)",
        "=MATCH(-999,Data!A:A,1)",
        "=MATCH(50,Data!A:A,-1)",
        "=MATCH(99999,Data!A:A,-1)",
        "=MATCH(0,Data!A:A,0)",
        "=MATCH(TRUE,Data!A:A,0)",
        "=XMATCH(50,Data!A:A)",
        "=XMATCH(\"k50\",Data!C:C)",
        "=XMATCH(99999,Data!A:A)",
        "=XLOOKUP(50,Data!A:A,Data!B:B)",
        "=XLOOKUP(99999,Data!A:A,Data!B:B,-1)",
        "=VLOOKUP(50,Data!A:B,2,FALSE)",
        "=VLOOKUP(50,Data!A:B,2,TRUE)",
        "=VLOOKUP(50.5,Data!A:B,2,TRUE)",
        "=VLOOKUP(99999,Data!A:B,2,FALSE)",
        "=VLOOKUP(0,Data!A:B,2,FALSE)",
        "=HLOOKUP(50,Data!A:B,2,FALSE)",
        "=LOOKUP(50,Data!A:A,Data!B:B)",
        "=SUMIF(Data!A:A,50)",
        "=SUMIF(Data!A:A,50,Data!B:B)",
        "=SUMIF(Data!A:A,\">100\")",
        "=SUMIF(Data!A:A,0)",
        "=COUNTIF(Data!A:A,50)",
        "=COUNTIF(Data!C:C,\"k50\")",
        "=COUNTIF(Data!C:C,\"k5*\")",
        "=COUNTIF(Data!A:A,\">100\")",
        "=COUNTIF(Data!A:A,0)",
        "=AVERAGEIF(Data!A:A,50)",
        "=AVERAGEIF(Data!A:A,\">100\")",
        "=SMALL(Data!A:A,5)",
        "=LARGE(Data!A:A,5)",
        "=MEDIAN(Data!A:A)",
        "=PERCENTILE(Data!A:A,0.9)",
        "=QUARTILE(Data!A:A,3)",
        "=SUM(Data!A:A)",
        "=COUNT(Data!A:A)",
        "=COUNTA(Data!A:A)",
        "=MAX(Data!A:A)",
        "=MIN(Data!A:A)",
        "=AVERAGE(Data!A:A)",
    ];

    [Test]
    public async Task Cache_MatchesBypass_AcrossTheThreeAdmissionReads_ScenariosAndFormulas()
    {
        var failures = new List<string>();
        var cases = 0;

        foreach (var scenario in Enum.GetValues<Scenario>())
        {
            var (workbook, sheet) = Build(scenario);

            foreach (var formula in Formulas)
            {
                cases++;

                var expected = Evaluate(workbook, sheet, formula, useCache: false);

                // With second-use admission the reads have distinct internal paths, all of which must agree
                // with the bypass: read 1 is the linear fallback (the range is only marked), read 2 builds the
                // snapshot, read 3 reuses it. One InvalidateCache, then three consecutive evaluations.
                workbook.RangeCacheDisabled = false;
                workbook.InvalidateCache();
                var read1 = Describe(ExpressionParser.Parse(formula, sheet).Evaluate(workbook));
                var read2 = Describe(ExpressionParser.Parse(formula, sheet).Evaluate(workbook));
                var read3 = Describe(ExpressionParser.Parse(formula, sheet).Evaluate(workbook));

                if (read1 != expected || read2 != expected || read3 != expected)
                {
                    failures.Add(
                        $"{scenario} :: {formula} :: expected={expected} read1={read1} read2={read2} read3={read3}"
                    );
                }
            }
        }

        await Assert.That(cases).IsGreaterThan(280);
        await Assert.That(failures).IsEmpty();
    }

    [Test]
    public async Task Cache_AdmitsOnSecondRead_AboveThreshold_AndNeverAdmits_Below()
    {
        var (workbook, _) = Build(Scenario.AscendingNumbers);
        var context = new EvaluationContext(workbook);

        var big = OpenRangeReference.Create(1, 1, null, null, "Data"); // A:A, 300 populated
        var small = new RangeReference("A1", "A10", "Data"); // 10 cells, below threshold

        // Second-use admission: the FIRST read of a big range only marks it (linear path → null); the SECOND
        // read builds and returns the snapshot; every read thereafter reuses that same instance.
        await Assert.That(workbook.TryGetRangeSnapshot(big, context) is null).IsTrue();
        var built = workbook.TryGetRangeSnapshot(big, context);
        await Assert.That(built is not null).IsTrue();
        await Assert
            .That(ReferenceEquals(workbook.TryGetRangeSnapshot(big, context), built))
            .IsTrue();

        // A sub-threshold range is never even marked, so it never admits no matter how often it is read.
        await Assert.That(workbook.TryGetRangeSnapshot(small, context) is null).IsTrue();
        await Assert.That(workbook.TryGetRangeSnapshot(small, context) is null).IsTrue();
    }

    [Test]
    public async Task ApproximateMatch_WithDuplicates_ReturnsLastPosition_LikeBypass()
    {
        // Ascending with duplicates: value 50 appears at several positions; MATCH(...,1) must return the
        // LAST of them (Excel's rule), identically cached and uncached.
        var (workbook, sheet) = Build(Scenario.NumbersWithDuplicates);
        const string formula = "=MATCH(50,Data!A:A,1)";

        var expected = Evaluate(workbook, sheet, formula, useCache: false);
        var cached = EvaluateBuilt(workbook, sheet, formula);

        await Assert.That(cached).IsEqualTo(expected);
    }

    [Test]
    public async Task ErrorInRange_PropagatesIdentically_ForOrderStatAndAggregate()
    {
        var (workbook, sheet) = Build(Scenario.WithErrors);

        foreach (
            var formula in new[] { "=SMALL(Data!A:A,5)", "=SUM(Data!A:A)", "=MEDIAN(Data!A:A)" }
        )
        {
            var expected = Evaluate(workbook, sheet, formula, useCache: false);
            var cached = EvaluateBuilt(workbook, sheet, formula);
            await Assert.That(cached).IsEqualTo(expected);
        }
    }

    // === Helpers =========================================================================================

    private static string Evaluate(Workbook workbook, Sheet sheet, string formula, bool useCache)
    {
        workbook.RangeCacheDisabled = !useCache;
        workbook.InvalidateCache();
        return Describe(ExpressionParser.Parse(formula, sheet).Evaluate(workbook));
    }

    // Drives the cache to its BUILT (second-read) path: the first read only marks the range and takes the
    // linear fallback, so the second read — evaluated here and returned — is the one served by the snapshot.
    private static string EvaluateBuilt(Workbook workbook, Sheet sheet, string formula)
    {
        workbook.RangeCacheDisabled = false;
        workbook.InvalidateCache();
        _ = ExpressionParser.Parse(formula, sheet).Evaluate(workbook); // read 1 → mark + linear
        return Describe(ExpressionParser.Parse(formula, sheet).Evaluate(workbook)); // read 2 → built snapshot
    }

    private static string Describe(ComputedValue value) =>
        value.Kind switch
        {
            ComputedValueKind.Number => "n:"
                + value.AsDouble()!.Value.ToString("R", CultureInfo.InvariantCulture),
            ComputedValueKind.Boolean => "b:" + value.AsBoolean(),
            ComputedValueKind.Text => "t:" + value.AsString(),
            ComputedValueKind.Blank => "blank",
            ComputedValueKind.Error => "e:" + value.AsObject(),
            _ => "ref",
        };

    private static (Workbook Workbook, Sheet Sheet) Build(Scenario scenario)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");

        for (var row = 1; row <= Rows; row++)
        {
            sheet[$"A{row}"] = KeyCell(scenario, row);
            sheet[$"B{row}"] = new NumberValue(row * 10);
            sheet[$"C{row}"] = new StringValue($"k{row}");
        }

        return (workbook, sheet);
    }

    private static Expression KeyCell(Scenario scenario, int row) =>
        scenario switch
        {
            Scenario.AscendingNumbers => new NumberValue(row),
            // 1,1,2,2,3,3,… — ascending with duplicates.
            Scenario.NumbersWithDuplicates => new NumberValue((row + 1) / 2),
            // A fixed pseudo-shuffle: deterministic, decidedly NOT sorted.
            Scenario.UnsortedNumbers => new NumberValue((row * 97) % Rows + 1),
            Scenario.MixedTypes => (row % 3) switch
            {
                0 => new NumberValue(row),
                1 => new StringValue($"k{row}"),
                _ => new BooleanValue(row % 2 == 0),
            },
            // Mostly ascending numbers with a #DIV/0! sprinkled in every 40th cell.
            Scenario.WithErrors => row % 40 == 0
                ? ExpressionParser.Parse("=1/0", new Sheet { Name = "Data" })
                : new NumberValue(row),
            // A mix of a zero, an empty-text, a genuinely blank cell (reference to an empty cell) and numbers.
            Scenario.WithZerosAndBlanks => (row % 4) switch
            {
                0 => new NumberValue(0),
                1 => new StringValue(string.Empty),
                2 => ExpressionParser.Parse("=ZZ9000", new Sheet { Name = "Data" }), // empty cell → Blank
                _ => new NumberValue(row),
            },
            Scenario.TextKeys => new StringValue($"k{row}"),
            _ => new NumberValue(row),
        };
}
