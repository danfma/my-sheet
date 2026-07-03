using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The contract matrix for a reference to a NON-EXISTENT sheet. A missing sheet is a STRUCTURAL failure of
/// the reference (#REF!), fiel ao Excel: it must propagate through EVERY consuming function — INCLUDING the
/// error-ignoring COUNT family — instead of throwing <see cref="KeyNotFoundException"/> or being swallowed as
/// an empty range. The controls prove the per-cell value-error policy over an EXISTING sheet is unchanged.
/// </summary>
public class MissingSheetReferenceTests
{
    // Builds a workbook with a single sheet "Main". Seeds go into the given cells; the formula under test
    // lives in Z1 (a column outside the A:E ranges the matrix probes, so a whole-column control never sees
    // the formula cell itself).
    private static Workbook Build(string formula, params (string Id, string Formula)[] seed)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");

        foreach (var (id, cellFormula) in seed)
        {
            main[id] = ExpressionParser.Parse(cellFormula, main);
        }

        main["Z1"] = ExpressionParser.Parse(formula, main);

        return workbook;
    }

    private static ComputedValue Eval(string formula, params (string Id, string Formula)[] seed) =>
        Build(formula, seed).GetCellValue("Main", "Z1");

    private static async Task AssertRef(string formula, params (string Id, string Formula)[] seed)
    {
        var value = Eval(formula, seed);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);

        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // === Missing sheet → #REF! through every consuming function ==========================================

    [Test]
    public async Task Cell_MissingSheet_IsRef() => await AssertRef("=Ghost!A1");

    [Test]
    public async Task Text_MissingSheet_IsRef() => await AssertRef("=UPPER(Ghost!E9)");

    [Test]
    public async Task Operator_MissingSheet_IsRef() => await AssertRef("=Ghost!A1 + 1");

    [Test]
    public async Task Sum_MissingSheet_IsRef() => await AssertRef("=SUM(Ghost!A:A)");

    [Test]
    public async Task Count_MissingSheet_IsRef() => await AssertRef("=COUNT(Ghost!A:A)");

    [Test]
    public async Task CountA_MissingSheet_IsRef() => await AssertRef("=COUNTA(Ghost!A:A)");

    [Test]
    public async Task CountBlank_MissingSheet_IsRef() => await AssertRef("=COUNTBLANK(Ghost!A:A)");

    [Test]
    public async Task CountIf_MissingSheet_IsRef() => await AssertRef("=COUNTIF(Ghost!D:D,1)");

    [Test]
    public async Task SumIf_MissingSheet_IsRef() => await AssertRef("=SUMIF(Ghost!A:A,\">5\")");

    [Test]
    public async Task Average_MissingSheet_IsRef() => await AssertRef("=AVERAGE(Ghost!B2:B9)");

    [Test]
    public async Task VLookup_MissingSheet_IsRef() => await AssertRef("=VLOOKUP(1,Ghost!A:B,2)");

    [Test]
    public async Task Rows_MissingSheet_IsRef() => await AssertRef("=ROWS(Ghost!A:A)");

    [Test]
    public async Task Columns_MissingSheet_IsRef() => await AssertRef("=COLUMNS(Ghost!A:A)");

    [Test]
    public async Task Union_WithOneMissingSheet_IsRef() =>
        await AssertRef("=SUM(Main!A:A, Ghost!A:A)", ("A1", "=10"));

    // === Long tail: every reference-consuming function propagates the structural #REF! =====================

    [Test]
    public async Task SumProduct_FirstArgumentMissingSheet_IsRef() =>
        await AssertRef("=SUMPRODUCT(Ghost!A:A, Main!A:A)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task SumProduct_SecondArgumentMissingSheet_IsRef() =>
        await AssertRef("=SUMPRODUCT(Main!A1:A2, Ghost!B1:B2)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task SumX2MY2_MissingSheet_IsRef() =>
        await AssertRef("=SUMX2MY2(Ghost!A:A, Main!B:B)", ("B1", "=1"));

    [Test]
    public async Task Correl_MissingSheet_IsRef() =>
        await AssertRef("=CORREL(Ghost!A:A, Main!B:B)", ("B1", "=1"), ("B2", "=2"));

    [Test]
    public async Task Covar_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=COVAR(Main!A:A, Ghost!B:B)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task Forecast_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=FORECAST(1, Ghost!A:A, Main!B:B)", ("B1", "=1"), ("B2", "=2"));

    [Test]
    public async Task Prob_MissingSheet_IsRef() =>
        await AssertRef("=PROB(Ghost!A:A, Main!B:B, 1)", ("B1", "=1"));

    [Test]
    public async Task Match_MissingSheet_IsRef() => await AssertRef("=MATCH(1, Ghost!A:A)");

    [Test]
    public async Task XLookup_MissingSheet_IsRef() => await AssertRef("=XLOOKUP(1, Ghost!A:A, Main!B:B)");

    [Test]
    public async Task XMatch_MissingSheet_IsRef() => await AssertRef("=XMATCH(1, Ghost!A:A)");

    [Test]
    public async Task Lookup_MissingSheet_IsRef() => await AssertRef("=LOOKUP(1, Ghost!A:A)");

    [Test]
    public async Task Index_MissingSheet_IsRef() => await AssertRef("=INDEX(Ghost!A:A, 1)");

    // A BOUNDED ghost range (not just a whole-column open range): the lookup would otherwise scan its cells,
    // skip the per-cell #REF! keys, and degrade to #N/A. It must short-circuit to #REF! structurally.
    [Test]
    public async Task VLookup_BoundedMissingSheet_IsRef() => await AssertRef("=VLOOKUP(1, Ghost!A1:B5, 2)");

    [Test]
    public async Task HLookup_BoundedMissingSheet_IsRef() => await AssertRef("=HLOOKUP(1, Ghost!A1:B5, 2)");

    [Test]
    public async Task TextJoin_MissingSheet_IsRef() => await AssertRef("=TEXTJOIN(\",\", TRUE, Ghost!A:A)");

    [Test]
    public async Task Concat_MissingSheet_IsRef() => await AssertRef("=CONCAT(Ghost!A:A)");

    [Test]
    public async Task Concatenate_MissingSheet_IsRef() => await AssertRef("=CONCATENATE(Ghost!A:A)");

    [Test]
    public async Task Npv_ValuesMissingSheet_IsRef() => await AssertRef("=NPV(0.1, Ghost!A:A)");

    // The rate argument is itself a reference to a missing sheet — a structural #REF! before any math.
    [Test]
    public async Task Npv_RateMissingSheet_IsRef() => await AssertRef("=NPV(Ghost!A1, Main!A1)", ("A1", "=10"));

    [Test]
    public async Task Irr_MissingSheet_IsRef() => await AssertRef("=IRR(Ghost!A:A)");

    [Test]
    public async Task Mirr_MissingSheet_IsRef() => await AssertRef("=MIRR(Ghost!A:A, 0.1, 0.1)");

    [Test]
    public async Task XNpv_MissingSheet_IsRef() => await AssertRef("=XNPV(0.1, Ghost!A:A, Main!B:B)");

    [Test]
    public async Task XIrr_MissingSheet_IsRef() => await AssertRef("=XIRR(Ghost!A:A, Main!B:B)");

    [Test]
    public async Task FvSchedule_MissingSheet_IsRef() => await AssertRef("=FVSCHEDULE(100, Ghost!A:A)");

    [Test]
    public async Task NetworkDays_HolidaysMissingSheet_IsRef() =>
        await AssertRef("=NETWORKDAYS(Main!A1, Main!A2, Ghost!C:C)", ("A1", "=1"), ("A2", "=10"));

    [Test]
    public async Task Workday_HolidaysMissingSheet_IsRef() =>
        await AssertRef("=WORKDAY(Main!A1, 5, Ghost!C:C)", ("A1", "=1"));

    [Test]
    public async Task SeriesSum_MissingSheet_IsRef() => await AssertRef("=SERIESSUM(2, 1, 1, Ghost!A:A)");

    [Test]
    public async Task And_MissingSheet_IsRef() => await AssertRef("=AND(Ghost!A:A)");

    [Test]
    public async Task Or_MissingSheet_IsRef() => await AssertRef("=OR(Ghost!A:A)");

    [Test]
    public async Task Xor_MissingSheet_IsRef() => await AssertRef("=XOR(Ghost!A:A)");

    [Test]
    public async Task Row_MissingSheet_IsRef() => await AssertRef("=ROW(Ghost!A1)");

    [Test]
    public async Task Column_MissingSheet_IsRef() => await AssertRef("=COLUMN(Ghost!A1)");

    // Already reaches #REF! through the guarded Fold, but pinned here as part of the contract matrix.
    [Test]
    public async Task Median_MissingSheet_IsRef() => await AssertRef("=MEDIAN(Ghost!A:A)");

    [Test]
    public async Task Percentile_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=PERCENTILE(Ghost!A:A, 0.5)");

    // === Controls: the sheet EXISTS — the per-cell value-error policy is UNCHANGED =======================

    [Test]
    public async Task Match_ExistingSheet_NoMatch_IsNA()
    {
        // A whole-column MATCH over the SAME (existing) but empty column is #N/A, NOT #REF!: a missing match
        // over a real sheet is a value outcome, distinct from the structural missing-sheet failure.
        var value = Eval("=MATCH(1, A:A)");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task XLookup_ExistingSheet_NoMatch_IsNA()
    {
        var value = Eval("=XLOOKUP(1, A:A, B:B)");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task SumProduct_ExistingSheet_IsNumber()
    {
        var value = Eval("=SUMPRODUCT(A1:A2, B1:B2)", ("A1", "=1"), ("A2", "=2"), ("B1", "=3"), ("B2", "=4"));

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(11.0); // 1*3 + 2*4
    }

    [Test]
    public async Task Correl_ExistingSheet_IsNumber()
    {
        var value = Eval(
            "=CORREL(A1:A3, B1:B3)",
            ("A1", "=1"), ("A2", "=2"), ("A3", "=3"),
            ("B1", "=2"), ("B2", "=4"), ("B3", "=6")
        );

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(1.0); // perfectly correlated
    }

    // === Controls: the sheet EXISTS — the per-cell value-error policy is UNCHANGED =======================

    [Test]
    public async Task CountIf_ExistingSheet_NoMatch_IsZero()
    {
        // Whole-column COUNTIF over the SAME (existing) sheet with no match is 0, not #REF!.
        var value = Eval("=COUNTIF(D:D,1)", ("D1", "=2"), ("D2", "=3"));

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(0.0);
    }

    [Test]
    public async Task Count_ExistingSheet_IgnoresCellError()
    {
        // COUNT ignores a cell VALUE error (#DIV/0!) on an existing sheet: it still counts the numbers.
        var value = Eval("=COUNT(A:A)", ("A1", "=1/0"), ("A2", "=5"), ("A3", "=7"));

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(2.0);
    }

    [Test]
    public async Task Sum_ExistingSheet_PropagatesCellError()
    {
        // SUM propagates a cell VALUE error (#DIV/0!) on an existing sheet — unchanged behaviour.
        var value = Eval("=SUM(A:A)", ("A1", "=1/0"), ("A2", "=5"));

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);

        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.DivZero);
    }

    [Test]
    public async Task Cell_ExistingSheet_Empty_CoercesToZero()
    {
        // Control: a reference into an EXISTING sheet's empty cell is NOT #REF! (unlike a missing sheet).
        // Excel-parity update — a formula result is never blank at the CELL boundary, so =Main!A2 (A2 empty)
        // now displays 0 instead of blank. The control's real point (existing sheet ≠ #REF!) is preserved.
        var value = Eval("=Main!A2");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    // === No cell throws: a whole batch with dangling refs resolves, each to #REF! ========================

    [Test]
    public async Task Batch_WithDanglingRefs_NeverThrows()
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");

        main["A1"] = ExpressionParser.Parse("=10", main);
        main["A2"] = ExpressionParser.Parse("=UPPER(BOX11MNO_HIDE!E9)", main);
        main["A3"] = ExpressionParser.Parse("=SUM(Ghost!A:A)", main);
        main["A4"] = ExpressionParser.Parse("=COUNTIF(Ghost!D:D,1)", main);
        main["A5"] = ExpressionParser.Parse("=A1 + 1", main);

        var results = new Dictionary<string, ComputedValue>();

        foreach (var id in new[] { "A1", "A2", "A3", "A4", "A5" })
        {
            // The point of the fix: no cell throws mid-batch.
            results[id] = workbook.GetCellValue("Main", id);
        }

        await Assert.That(results["A1"].TryGetNumber(out var a1)).IsTrue();
        await Assert.That(a1).IsEqualTo(10.0);
        await Assert.That(results["A5"].TryGetNumber(out var a5)).IsTrue();
        await Assert.That(a5).IsEqualTo(11.0);

        foreach (var id in new[] { "A2", "A3", "A4" })
        {
            await Assert.That(results[id].Kind).IsEqualTo(ComputedValueKind.Error);
            results[id].TryGetError(out var error);
            await Assert.That(error).IsEqualTo(Error.Ref);
        }
    }

    // === TryGetSheet (host API) ==========================================================================

    [Test]
    public async Task TryGetSheet_KnownAndUnknown()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Main");

        await Assert.That(workbook.TryGetSheet("Main", out var main)).IsTrue();
        await Assert.That(main!.Name).IsEqualTo("Main");
        // Case-insensitive, like Excel.
        await Assert.That(workbook.TryGetSheet("MAIN", out _)).IsTrue();
        await Assert.That(workbook.TryGetSheet("Ghost", out _)).IsFalse();
    }
}
