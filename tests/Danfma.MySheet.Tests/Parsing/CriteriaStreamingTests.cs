using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Guards the streaming/snapshot refactor of the *IFS family (SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS):
// the criteria/value ranges are now walked as parallel positional cursors instead of materialized lists.
// These tests drive BOTH cursor backings and assert equivalence against an independent in-test oracle:
//   * the STREAMING backing — a range read once per epoch (dense positional read, no list);
//   * the SNAPSHOT backing — a range read twice per epoch clears the 256-cell admission threshold, so the
//     second read serves the shared per-epoch snapshot array (zero-copy). N = 300 (> 256) reaches it.
public class CriteriaStreamingTests
{
    private const int N = 300; // > RangeCacheMinimumCells (256): large ranges are snapshot-eligible.

    // Deterministic data: A = group text, B = region text, C = amount (some blanks so both blank handling and
    // the numeric/text criteria mix are exercised).
    private static (Workbook Workbook, string[] Groups, string[] Regions, double?[] Amounts) BuildData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var groups = new string[N + 1];
        var regions = new string[N + 1];
        var amounts = new double?[N + 1];

        for (var r = 1; r <= N; r++)
        {
            var group = "G" + (r % 10);
            var region = "R" + (r % 5);
            groups[r] = group;
            regions[r] = region;
            sheet[$"A{r}"] = new Danfma.MySheet.Expressions.StringValue(group);
            sheet[$"B{r}"] = new Danfma.MySheet.Expressions.StringValue(region);

            // Every 7th amount is left blank to cover the blank branch on both value and criteria ranges.
            if (r % 7 == 0)
            {
                amounts[r] = null;
                continue;
            }

            var amount = (double)((r % 100) + 1);
            amounts[r] = amount;
            sheet[$"C{r}"] = new NumberValue(amount);
        }

        return (workbook, groups, regions, amounts);
    }

    private static object? Eval(Workbook workbook, string formula)
    {
        var sheet = workbook.Sheets["Sheet1"];
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task SumIfs_LargeRange_SnapshotAndStreaming_MatchOracle()
    {
        var (workbook, groups, _, amounts) = BuildData();

        // C is referenced TWICE (value range + third criteria) → its second read admits the snapshot; A is
        // referenced once → it streams. So this single formula exercises both backings at once.
        var formula = $"=SUMIFS(C1:C{N},A1:A{N},\"G3\",C1:C{N},\">=50\")";

        var expected = 0.0;
        for (var r = 1; r <= N; r++)
        {
            if (groups[r] == "G3" && amounts[r] is { } a && a >= 50)
            {
                expected += a;
            }
        }

        // First evaluation (mixed streaming/first-read) and a second (now snapshot-served) must both match.
        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task CountIfs_LargeRange_Streaming_MatchesOracle()
    {
        var (workbook, groups, regions, _) = BuildData();

        // A and B are each read once → pure streaming path over 300 cells, two criteria pairs.
        var formula = $"=COUNTIFS(A1:A{N},\"G3\",B1:B{N},\"R2\")";

        var expected = 0;
        for (var r = 1; r <= N; r++)
        {
            if (groups[r] == "G3" && regions[r] == "R2")
            {
                expected++;
            }
        }

        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
        // Second read of A and B now serves the snapshot; result is unchanged.
        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task AverageIfs_LargeRange_MatchesOracle()
    {
        var (workbook, groups, _, amounts) = BuildData();
        var formula = $"=AVERAGEIFS(C1:C{N},A1:A{N},\"G3\",C1:C{N},\">=50\")";

        var total = 0.0;
        var count = 0;
        for (var r = 1; r <= N; r++)
        {
            if (groups[r] == "G3" && amounts[r] is { } a && a >= 50)
            {
                total += a;
                count++;
            }
        }

        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(total / count);
    }

    [Test]
    public async Task MaxIfs_MinIfs_LargeRange_MatchOracle()
    {
        var (workbook, groups, regions, amounts) = BuildData();
        var maxFormula = $"=MAXIFS(C1:C{N},A1:A{N},\"G3\",B1:B{N},\"R2\")";
        var minFormula = $"=MINIFS(C1:C{N},A1:A{N},\"G3\",B1:B{N},\"R2\")";

        var max = double.NegativeInfinity;
        var min = double.PositiveInfinity;
        var any = false;
        for (var r = 1; r <= N; r++)
        {
            if (groups[r] == "G3" && regions[r] == "R2" && amounts[r] is { } a)
            {
                any = true;
                max = Math.Max(max, a);
                min = Math.Min(min, a);
            }
        }

        await Assert.That(Num(Eval(workbook, maxFormula))).IsEqualTo(any ? max : 0.0);
        await Assert.That(Num(Eval(workbook, minFormula))).IsEqualTo(any ? min : 0.0);
    }

    [Test]
    public async Task SumIfs_WildcardCriteria_LargeRange_MatchesOracle()
    {
        var (workbook, groups, _, amounts) = BuildData();

        // Text wildcard on a streamed criteria range (A read once), numeric comparison on the snapshot (C twice).
        var formula = $"=SUMIFS(C1:C{N},A1:A{N},\"G?\",C1:C{N},\"<>60\")";

        var expected = 0.0;
        for (var r = 1; r <= N; r++)
        {
            // "G?" matches G0..G9 (exactly one char after G); every group here is "G" + one digit → all match.
            var groupMatch = Criteria.WildcardMatch("G?", groups[r]);
            if (groupMatch && amounts[r] is { } a && a != 60)
            {
                expected += a;
            }
        }

        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task SumIfs_ShapeMismatch_LargeRange_IsValueError()
    {
        var (workbook, _, _, _) = BuildData();

        // Value range 1..N but criteria range 1..(N-1): length mismatch surfaces as #VALUE!, no materialization.
        var formula = $"=SUMIFS(C1:C{N},A1:A{N - 1},\"G3\")";

        await Assert.That(Eval(workbook, formula)).IsEqualTo(ErrorValue.NotValue);
    }
}
