using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Phase 7 (dynamic arrays) — the acceptance pins for the four mini-CSE producers <c>FILTER</c>,
/// <c>SORT</c>, <c>UNIQUE</c> and <c>SEQUENCE</c>. <b>Every test here is RED at the commit that adds
/// it</b>: none of the four is registered in <c>FunctionRegistry</c> yet, so each formula parses to a
/// generic <c>FunctionCall</c> and evaluates to <c>#NAME?</c> (or, under <c>INDEX</c>, to <c>#REF!</c>).
/// The observed-today value is recorded per test.
/// </summary>
/// <remarks>
/// <para>
/// <b>Golden values.</b> The documented behaviour comes from the four official Microsoft pages, fetched
/// <b>2026-09-10</b>: "FILTER function" (f4f7cb66-82eb-4767-8f7c-4877ad80c759), "SORT function"
/// (22f63bd0-ccc8-492f-953d-c20e8e44b86c), "UNIQUE function" (c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e) and
/// "SEQUENCE function" (57467a98-57e0-4817-9f14-2eb78519ca90), all on support.microsoft.com.
/// </para>
/// <para>
/// <b>Where a page is silent or disagrees, the oracle decides</b> (project rule P0, addendum 2026-09-09):
/// "Excel" means <b>Aspose.Cells 26.6.0 as measured</b>. Every number below was measured on that version
/// on <b>2026-09-10</b> against the fixtures in this file, in BOTH entry modes — <c>plain</c>
/// (<c>Cell.Formula</c>) and <c>CSE</c> (<c>Cell.SetArrayFormula(f, 1, 1)</c>). The two agree for every
/// pin here except the two named ones (the open range, and <c>COUNTA</c>/<c>COUNT</c> over
/// <c>SORT</c> — not pinned), and where they split <b>the CSE column is the target</b>, because the
/// mini-CSE implements the array-entered rule everywhere. Modes are never mixed inside one assertion.
/// </para>
/// <para>
/// The three places where the measured oracle beat a page or the phase design, all re-measured here:
/// <list type="bullet">
/// <item>a bad <c>SEQUENCE</c> argument is <c>#VALUE!</c>, never <c>#CALC!</c>;</item>
/// <item><c>FILTER</c>/<c>SORT</c>/<c>UNIQUE</c> <b>preserve</b> Blank instead of normalizing it to 0, and
/// <c>SORT</c> puts blanks and errors LAST ascending (errors first descending) instead of propagating;</item>
/// <item><c>UNIQUE</c> is case-SENSITIVE — and the Microsoft page as fetched today says nothing about case
/// either way, so there is no page to contradict.</item>
/// </list>
/// </para>
/// </remarks>
public class DynamicArrayTests
{
    // Harness copied from MathAggregateTests.cs:11-31, plus a `bool` arm for the mixed-type fixture
    // (SORT's number < text < FALSE < TRUE order needs logical cells).
    private static object? Calc(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value switch
            {
                string s when s.StartsWith('=') => ExpressionParser.Parse(s, sheet),
                string s => new Danfma.MySheet.Expressions.StringValue(s),
                bool b => b ? BooleanValue.True : BooleanValue.False,
                double d => new NumberValue(d),
                int i => new NumberValue(i),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // #CALC! has no Error code in the engine yet (Error.cs Displays has seven entries and none is
    // #CALC!, so Error.FromDisplay would fold it onto #VALUE!). The pin compares against the AST node,
    // whose equality is on the code string, so it is exact both before and after the code is added.
    private static readonly ErrorValue CalcError = new("#CALC!");

    // A1:A3 = 5, 0, 9 | B1:B3 = 1, 2, 3 | C1:C3 = "a", "A", "b". The row A1:C1 is 5, 1, "a".
    private static readonly (string, object)[] Grid =
    [
        ("A1", 5),
        ("A2", 0),
        ("A3", 9),
        ("B1", 1),
        ("B2", 2),
        ("B3", 3),
        ("C1", "a"),
        ("C2", "A"),
        ("C3", "b"),
    ];

    // A5:A8 = 7, <blank>, "t", 7 — A6 is deliberately never written.
    private static readonly (string, object)[] WithBlank = [("A5", 7), ("A7", "t"), ("A8", 7)];

    // E1:E3 = 5, #DIV/0!, 9.
    private static readonly (string, object)[] WithError = [("E1", 5), ("E2", "=1/0"), ("E3", 9)];

    // M1:M4 = TRUE, "x", 2, FALSE.
    private static readonly (string, object)[] MixedTypes =
    [
        ("M1", true),
        ("M2", "x"),
        ("M3", 2),
        ("M4", false),
    ];

    // N1:O4 = (1,"a"), (2,"b"), (1,"c"), (2,"d") — duplicate keys in column N, so the tie order in
    // column O is the stability proof.
    private static readonly (string, object)[] TwoColumns =
    [
        ("N1", 1),
        ("O1", "a"),
        ("N2", 2),
        ("O2", "b"),
        ("N3", 1),
        ("O3", "c"),
        ("N4", 2),
        ("O4", "d"),
    ];

    // Q1:Q4 = 9, 5, 9, 0 — 9 repeats, so first-appearance order and exactly_once both bite.
    private static readonly (string, object)[] Duplicates =
    [
        ("Q1", 9),
        ("Q2", 5),
        ("Q3", 9),
        ("Q4", 0),
    ];

    // ------------------------------------------------------------------ SEQUENCE

    [Test]
    public async Task Sequence_FillsRowMajor_WithTheGivenStartAndStep()
    {
        // support.microsoft.com SEQUENCE: "=SEQUENCE(rows,[columns],[start],[step])"; columns, start and
        // step default to 1. The page does not state the fill order; the oracle does — 1..6 across each
        // row in turn for SEQUENCE(2,3) (2 is (1,2) and 4 is (2,1), so it is row-major, not column-major).
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #NAME? (SUM) / #REF! (INDEX).
        await Assert.That(Num(Calc("=SUM(SEQUENCE(5))"))).IsEqualTo(15.0);
        await Assert.That(Num(Calc("=INDEX(SEQUENCE(2,3),1,3)"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=INDEX(SEQUENCE(2,3),2,1)"))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=INDEX(SEQUENCE(2,3),2,2)"))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=INDEX(SEQUENCE(5),4)"))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=SUM(SEQUENCE(3,1,10,5))"))).IsEqualTo(45.0);
    }

    [Test]
    public async Task Sequence_TruncatesItsSizes_AndDefaultsOmittedOptionalsToOne()
    {
        // Oracle 26.6.0, 2026-09-10, plain == CSE: SEQUENCE(2.7) is 2 rows (SUM 1+2), SEQUENCE(2,2.9) is
        // 2x2 (SUM 1+2+3+4), a fractional step is kept as given (1.5 alone), and the empty argument slots
        // of SEQUENCE(2,2,,) default to 1 exactly as an omitted argument does. Observed today: #NAME?.
        await Assert.That(Num(Calc("=SUM(SEQUENCE(2.7))"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=SUM(SEQUENCE(2,2.9))"))).IsEqualTo(10.0);
        await Assert.That(Num(Calc("=SUM(SEQUENCE(1,1,1.5,0.25))"))).IsEqualTo(1.5);
        await Assert.That(Num(Calc("=SUM(SEQUENCE(2,2,,))"))).IsEqualTo(10.0);
    }

    [Test]
    public async Task Sequence_ZeroNegativeOrNonNumericSize_IsValueError_NeverCalc()
    {
        // The Microsoft page is silent on zero, negative and non-numeric sizes. The oracle is not:
        // 26.6.0, 2026-09-10, plain == CSE — SUM(SEQUENCE(0)), SUM(SEQUENCE(-1)), SUM(SEQUENCE("x")) and
        // the bare SEQUENCE(0,0) are all #VALUE!, ERROR.TYPE(SEQUENCE(-1)) is 3 (= #VALUE!, not 14), and
        // an error in the size argument propagates unchanged. This is the phase design's correction B2:
        // the design said #CALC! for rows/columns < 1, and the oracle says #VALUE!.
        // Observed today: #NAME? for all four error rows, ERROR.TYPE = 5 (#NAME?), COUNT = 0.
        await Assert.That(Calc("=SUM(SEQUENCE(0))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(SEQUENCE(-1))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(SEQUENCE(\"x\"))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SEQUENCE(0,0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num(Calc("=ERROR.TYPE(SEQUENCE(-1))"))).IsEqualTo(3.0);
        await Assert.That(Calc("=SUM(SEQUENCE(1/0))")).IsEqualTo(ErrorValue.DivByZero);

        // Green today for the wrong reason (COUNT discards the #NAME? just as it will discard the
        // #VALUE!), kept because it is the oracle's answer and it must not move.
        await Assert.That(Num(Calc("=COUNT(SEQUENCE(-1))"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Sequence_BeyondTheGrid_IsNumError_TheOnePinnedDivergence()
    {
        // THE ONE DELIBERATE DIVERGENCE IN THIS FILE. MySheet bounds SEQUENCE by the grid
        // (1,048,576 rows x 16,384 columns) and answers #NUM! past it. The oracle has NO cap in a
        // consumed position: measured 26.6.0, 2026-09-10, plain == CSE, ROWS(SEQUENCE(1048577)) =
        // 1048577, COLUMNS(SEQUENCE(1,16385)) = 16385 and SUM(SEQUENCE(1048577)) = 549757386753. So the
        // cap is a MySheet choice, and the design's justification for it ("Excel bounds SEQUENCE by the
        // grid") is NOT what the oracle does. Pinned so the choice is visible and testable rather than
        // implicit; if the phase later drops the cap, both numbers are here. Observed today: #NAME?.
        await Assert.That(Calc("=SUM(SEQUENCE(1048577))")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=SUM(SEQUENCE(1,16385))")).IsEqualTo(ErrorValue.Number);
    }

    // ------------------------------------------------------------------ FILTER

    [Test]
    public async Task Filter_OverAColumn_KeepsTheMatchingRows()
    {
        // support.microsoft.com FILTER: "=FILTER(array,include,[if_empty])", include is "A Boolean array
        // whose height or width is the same as the array". A1:A3 = 5, 0, 9 keeps 5 and 9.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #NAME? (SUM) / #REF! (INDEX).
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,A1:A3>0))", Grid))).IsEqualTo(14.0);
        await Assert.That(Num(Calc("=INDEX(FILTER(A1:A3,A1:A3>0),2)", Grid))).IsEqualTo(9.0);

        // Both axes filter: A1:B3 keeps rows 1 and 3 (5+1+9+3), A1:C1 keeps all three columns of the row
        // (5 + 1, "a" is not numeric — and "a" > 0 is TRUE in the classic order, so it IS kept).
        await Assert.That(Num(Calc("=SUM(FILTER(A1:B3,A1:A3>0))", Grid))).IsEqualTo(18.0);
        await Assert.That(Num(Calc("=SUM(FILTER(A1:C1,A1:C1>0))", Grid))).IsEqualTo(6.0);

        // Past the end of the selection is #REF!, not the source's third row. Green today for the wrong
        // reason (INDEX over the #NAME? scalar is already out of range); kept as a must-not-move pin.
        await Assert
            .That(Calc("=INDEX(FILTER(A1:A3,A1:A3>0),3)", Grid))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Filter_WithNothingKept_IsCalcError_UnlessIfEmptyIsGiven()
    {
        // support.microsoft.com FILTER, verbatim: when nothing matches and if_empty is omitted the result
        // is a "#CALC! error" because "Excel does not currently support empty arrays". if_empty replaces
        // the whole result with a 1x1 array of that value — measured COUNTA = 1 and ROWS = 1.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #NAME? for the error and the value
        // rows; TRUE for ISERROR and 0 for COUNT (both green for the wrong reason, see below);
        // ERROR.TYPE = 5 (#NAME?) against the oracle's 14 (= #CALC!, correction M2).
        await Assert.That(Calc("=SUM(FILTER(A1:A3,A1:A3>100))", Grid)).IsEqualTo(CalcError);
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,A1:A3>100,0))", Grid))).IsEqualTo(0.0);
        await Assert.That(Calc("=FILTER(A1:A3,A1:A3>100,\"none\")", Grid)).IsEqualTo("none");
        await Assert
            .That(Num(Calc("=COUNTA(FILTER(A1:A3,A1:A3>100,\"none\"))", Grid)))
            .IsEqualTo(1.0);
        await Assert.That(Num(Calc("=ERROR.TYPE(FILTER(A1:A3,A1:A3>100))", Grid))).IsEqualTo(14.0);

        // Green today because #NAME? is also an error and COUNT also discards it. Both are the oracle's
        // answers, so they stay: the empty FILTER must remain an error the two agree on.
        await Assert.That(Calc("=ISERROR(FILTER(A1:A3,A1:A3>100))", Grid) as bool?).IsTrue();
        await Assert.That(Num(Calc("=COUNT(FILTER(A1:A3,A1:A3>100))", Grid))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Filter_AnIncludeOfTheWrongShape_IsValueError()
    {
        // The page requires an include "whose height or width is the same as the array". The oracle
        // rejects both a short vector and a 2-D include with #VALUE! (26.6.0, 2026-09-10, plain == CSE).
        // Observed today: #NAME?.
        await Assert.That(Calc("=SUM(FILTER(A1:A3,B1:B2>0))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(FILTER(A1:A3,A1:B3>0))", Grid)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Filter_AnErrorInInclude_Propagates_ButATextIncludeIsCalc()
    {
        // support.microsoft.com FILTER, verbatim: "If any value of the include argument is an error
        // (#N/A, #VALUE, etc.) or cannot be converted to a Boolean, the FILTER function will return an
        // error." The oracle splits those two clauses (26.6.0, 2026-09-10, plain == CSE): an ERROR
        // propagates as itself (#DIV/0!), while TEXT that cannot be converted keeps nothing and lands on
        // #CALC! — the empty-result error, not a coercion #VALUE!. A numeric include is truthy per
        // element (B1:B3 = 1, 2, 3 keeps everything; A1:A3 as its own include drops the 0).
        // Observed today: #NAME? for all four.
        await Assert
            .That(Calc("=SUM(FILTER(A1:A3,E1:E3>0))", [.. Grid, .. WithError]))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SUM(FILTER(A1:A3,C1:C3))", Grid)).IsEqualTo(CalcError);
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,B1:B3))", Grid))).IsEqualTo(14.0);
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,A1:A3))", Grid))).IsEqualTo(14.0);
    }

    [Test]
    public async Task Filter_AScalarInclude_Broadcasts_AndAFalsyOneIsValueError()
    {
        // Blocker B1's open question, closed by measurement (26.6.0, 2026-09-10, plain == CSE): a SCALAR
        // include broadcasts over the whole array rather than failing the height check — TRUE and 1 keep
        // every row — while a FALSY scalar keeps nothing and is #VALUE! (NOT the #CALC! of an empty
        // filter, and NOT the design's "1x1 #VALUE!" for the truthy case either). if_empty still answers
        // for the falsy scalar. Observed today: #NAME? for all four.
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,TRUE))", Grid))).IsEqualTo(14.0);
        await Assert.That(Num(Calc("=SUM(FILTER(A1:A3,1))", Grid))).IsEqualTo(14.0);
        await Assert.That(Calc("=SUM(FILTER(A1:A3,FALSE))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(FILTER(A1:A3,\"x\"))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=FILTER(A1:A3,FALSE,\"none\")", Grid)).IsEqualTo("none");
    }

    [Test]
    public async Task ASingleCellSource_IsAOneByOneProducer_ForAllThreeSelectors()
    {
        // Blocker B1: a 1x1 source must not be mistaken for a scalar operand. Oracle 26.6.0,
        // 2026-09-10, plain == CSE. Observed today: #NAME? for all three.
        await Assert.That(Num(Calc("=SUM(FILTER(A1,TRUE))", Grid))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=SUM(SORT(A1))", Grid))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=SUM(UNIQUE(A1))", Grid))).IsEqualTo(5.0);
    }

    [Test]
    public async Task Filter_OverAnOpenRange_IsRefused_AKnownDivergence()
    {
        // DIVERGENCE, pinned with both oracle columns because they disagree: measured 26.6.0,
        // 2026-09-10, SUM(FILTER(A:A,A:A>0)) is #VALUE! PLAIN and 28 ARRAY-ENTERED, while
        // ROWS(FILTER(A:A,A:A>0)) is 5 in BOTH modes (column A holds 5, 0, 9, 7, <blank>, "t", 7 and
        // ">0" keeps 5, 9, 7, "t", 7 — five rows summing to 28, since text sums as nothing but sorts and
        // compares above every number). MySheet implements the array-entered rule everywhere, so
        // the CSE column (28) is what P0 would demand; the mini-CSE nevertheless REFUSES an open range as
        // an array operand (ArrayEvaluation's OpenRangeReference arm) and the refusal reaches the
        // consumer as #VALUE!. That refusal is this phase's stated deviation, not a measurement, and it
        // is pinned here so it stays deliberate. Observed today: #NAME?.
        await Assert
            .That(Calc("=SUM(FILTER(A:A,A:A>0))", [.. Grid, .. WithBlank]))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // ------------------------------------------------------------------ SORT

    [Test]
    public async Task Sort_AscendingByDefault_AndDescendingOnMinusOne()
    {
        // support.microsoft.com SORT: "=SORT(array,[sort_index],[sort_order],[by_col])", sort_order
        // "1 for ascending order (default), -1 for descending order". A1:A3 = 5, 0, 9.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #REF! (INDEX) / #NAME? (SUM).
        await Assert.That(Num(Calc("=INDEX(SORT(A1:A3),1)", Grid))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:A3),2)", Grid))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:A3),3)", Grid))).IsEqualTo(9.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:A3,1,-1),1)", Grid))).IsEqualTo(9.0);
        await Assert.That(Num(Calc("=SUM(SORT(A1:A3))", Grid))).IsEqualTo(14.0);
    }

    [Test]
    public async Task Sort_AnOutOfRangeSortIndexOrAnUnknownSortOrder_IsValueError()
    {
        // The page states only the legal values. The oracle rejects everything else with #VALUE!
        // (26.6.0, 2026-09-10, plain == CSE): sort_index 0 or 2 on a one-column array, sort_order 0 or 2.
        // Observed today: #NAME? for all four.
        await Assert.That(Calc("=SUM(SORT(A1:A3,2))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(SORT(A1:A3,0))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(SORT(A1:A3,1,0))", Grid)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUM(SORT(A1:A3,1,2))", Grid)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Sort_ByCol_SortsColumnsAgainstTheChosenRow()
    {
        // support.microsoft.com SORT: by_col "TRUE to sort by column". A1:B3 is [[5,1],[0,2],[9,3]]; by
        // row 1 ascending the columns swap (1 before 5), and by row 2 descending they swap as well
        // (2 before 0), which is why both directions are pinned — the ascending-by-row-2 case is a no-op
        // on this fixture and would pin nothing. Oracle 26.6.0, 2026-09-10, plain == CSE.
        // Observed today: #REF!.
        await Assert.That(Num(Calc("=INDEX(SORT(A1:B3,1,1,TRUE),1,1)", Grid))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:B3,1,1,TRUE),1,2)", Grid))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:B3,2,-1,TRUE),1,1)", Grid))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A1:B3,2,-1,TRUE),2,1)", Grid))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Sort_IsStable_InBothDirections()
    {
        // N1:O4 = (1,"a"), (2,"b"), (1,"c"), (2,"d"): the keys tie in pairs, so the tie order in column O
        // is the stability proof. Ascending the keys are 1, 1, 2, 2 and the tags stay a, c, b, d;
        // descending the keys are 2, 2, 1, 1 and the tags stay b, d, a, c — a stable sort in BOTH
        // directions, not a reversal of the ascending result (that would be d, b, c, a).
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #REF!.
        await Assert.That(Num(Calc("=INDEX(SORT(N1:O4),1,1)", TwoColumns))).IsEqualTo(1.0);
        await Assert.That(Calc("=INDEX(SORT(N1:O4),1,2)", TwoColumns)).IsEqualTo("a");
        await Assert.That(Calc("=INDEX(SORT(N1:O4),2,2)", TwoColumns)).IsEqualTo("c");
        await Assert.That(Calc("=INDEX(SORT(N1:O4),3,2)", TwoColumns)).IsEqualTo("b");
        await Assert.That(Calc("=INDEX(SORT(N1:O4),4,2)", TwoColumns)).IsEqualTo("d");
        await Assert.That(Calc("=INDEX(SORT(N1:O4,1,-1),1,2)", TwoColumns)).IsEqualTo("b");
        await Assert.That(Calc("=INDEX(SORT(N1:O4,1,-1),2,2)", TwoColumns)).IsEqualTo("d");
        await Assert.That(Calc("=INDEX(SORT(N1:O4,1,-1),3,2)", TwoColumns)).IsEqualTo("a");
        await Assert.That(Calc("=INDEX(SORT(N1:O4,1,-1),4,2)", TwoColumns)).IsEqualTo("c");
    }

    [Test]
    public async Task Sort_MixedTypes_UseTheClassicNumberTextFalseTrueOrder()
    {
        // M1:M4 = TRUE, "x", 2, FALSE sorts to 2, "x", FALSE, TRUE — the classic Excel order
        // number < text < FALSE < TRUE, which is exactly what ValueCoercion.Compare already implements.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #REF!.
        await Assert.That(Num(Calc("=INDEX(SORT(M1:M4),1)", MixedTypes))).IsEqualTo(2.0);
        await Assert.That(Calc("=INDEX(SORT(M1:M4),2)", MixedTypes)).IsEqualTo("x");
        await Assert.That(Calc("=INDEX(SORT(M1:M4),3)", MixedTypes) as bool?).IsFalse();
        await Assert.That(Calc("=INDEX(SORT(M1:M4),4)", MixedTypes) as bool?).IsTrue();
    }

    [Test]
    public async Task Sort_PutsBlanksLast_InBothDirections()
    {
        // Correction M4, re-measured: A5:A8 = 7, <blank>, "t", 7 sorts to 7, 7, "t", <blank> ascending
        // and "t", 7, 7, <blank> descending — the blank goes LAST in BOTH directions. It is NOT sorted
        // as a 0 (a 0 would come FIRST ascending, which is what ValueCoercion.Compare's "blank counts as
        // 0" would do) and it is still a blank when it arrives: ISBLANK on the fourth row is TRUE, even
        // though INDEX of it reads as 0 through the number channel.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #REF! (INDEX) / FALSE (ISBLANK).
        await Assert.That(Num(Calc("=INDEX(SORT(A5:A8),1)", WithBlank))).IsEqualTo(7.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A5:A8),2)", WithBlank))).IsEqualTo(7.0);
        await Assert.That(Calc("=INDEX(SORT(A5:A8),3)", WithBlank)).IsEqualTo("t");
        await Assert.That(Calc("=ISBLANK(INDEX(SORT(A5:A8),4))", WithBlank) as bool?).IsTrue();
        await Assert.That(Calc("=INDEX(SORT(A5:A8,1,-1),1)", WithBlank)).IsEqualTo("t");
        await Assert.That(Num(Calc("=INDEX(SORT(A5:A8,1,-1),2)", WithBlank))).IsEqualTo(7.0);
        await Assert.That(Num(Calc("=INDEX(SORT(A5:A8,1,-1),3)", WithBlank))).IsEqualTo(7.0);
        await Assert.That(Calc("=ISBLANK(INDEX(SORT(A5:A8,1,-1),4))", WithBlank) as bool?).IsTrue();
    }

    [Test]
    public async Task Sort_SortsErrors_RatherThanPropagatingThem()
    {
        // Correction M4, re-measured: E1:E3 = 5, #DIV/0!, 9 sorts to 5, 9, #DIV/0! ascending and puts the
        // error FIRST descending. SORT itself does NOT propagate the error — it ranks it after every
        // value; SUM(SORT(E1:E3)) is #DIV/0! only because SUM propagates what it is handed.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #REF! (INDEX) / #NAME? (SUM).
        await Assert.That(Num(Calc("=INDEX(SORT(E1:E3),1)", WithError))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=INDEX(SORT(E1:E3),2)", WithError))).IsEqualTo(9.0);
        await Assert.That(Calc("=INDEX(SORT(E1:E3),3)", WithError)).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=INDEX(SORT(E1:E3,1,-1),1)", WithError))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SUM(SORT(E1:E3))", WithError)).IsEqualTo(ErrorValue.DivByZero);
    }

    // ------------------------------------------------------------------ UNIQUE

    [Test]
    public async Task Unique_KeepsFirstAppearanceOrder()
    {
        // support.microsoft.com UNIQUE: "=UNIQUE(array,[by_col],[exactly_once])"; with by_col omitted it
        // compares rows and returns "all distinct rows". Q1:Q4 = 9, 5, 9, 0 gives 9, 5, 0 — the order of
        // FIRST appearance, not sorted. Oracle 26.6.0, 2026-09-10, plain == CSE.
        // Observed today: #REF! (INDEX) / #NAME? (SUM, COUNTA).
        await Assert.That(Num(Calc("=INDEX(UNIQUE(Q1:Q4),1)", Duplicates))).IsEqualTo(9.0);
        await Assert.That(Num(Calc("=INDEX(UNIQUE(Q1:Q4),2)", Duplicates))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=INDEX(UNIQUE(Q1:Q4),3)", Duplicates))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=SUM(UNIQUE(Q1:Q4))", Duplicates))).IsEqualTo(14.0);
        await Assert.That(Num(Calc("=COUNTA(UNIQUE(Q1:Q4))", Duplicates))).IsEqualTo(3.0);

        // by_col compares columns instead: A1:B3 has two distinct columns, so nothing collapses.
        await Assert.That(Num(Calc("=COLUMNS(UNIQUE(A1:B3,TRUE))", Grid))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Unique_IsCaseSensitive_WhereTheCriteriaFamilyIsNot()
    {
        // Correction M4, re-measured and re-checked against the page. C1:C3 = "a", "A", "b" keeps all
        // THREE rows and the second one is "A": UNIQUE compares text case-SENSITIVELY on the oracle
        // (26.6.0, 2026-09-10, plain == CSE), so it is NOT ValueCoercion.AreEqual, which is
        // case-insensitive. The Microsoft UNIQUE page as fetched 2026-09-10 says nothing about case in
        // either direction, so there is no page to contradict — the measurement stands alone (P0).
        // The contrast rows are the control: COUNTIF and MATCH stay case-INSENSITIVE over the same data,
        // so the case rule belongs to UNIQUE's key comparison and nowhere else.
        // Observed today: #NAME? (ROWS, COUNTA), #REF! (INDEX), #N/A (MATCH); COUNTIF = 2 already green.
        await Assert.That(Num(Calc("=ROWS(UNIQUE(C1:C3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=COUNTA(UNIQUE(C1:C3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Calc("=INDEX(UNIQUE(C1:C3),2)", Grid)).IsEqualTo("A");
        await Assert.That(Num(Calc("=MATCH(\"A\",UNIQUE(C1:C3),0)", Grid))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=COUNTIF(C1:C3,\"a\")", Grid))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Unique_ExactlyOnce_ReturnsOnlyTheRowsThatOccurOnce()
    {
        // support.microsoft.com UNIQUE, verbatim: exactly_once "TRUE will return all distinct rows or
        // columns that occur exactly once from the range or array". Q1:Q4 = 9, 5, 9, 0 leaves 5 and 0.
        // Oracle 26.6.0, 2026-09-10, plain == CSE, for the two rows pinned here.
        // The SHAPE of that result is deliberately NOT pinned: the oracle answers
        // ROWS(UNIQUE(Q1:Q4,FALSE,TRUE)) = 3 with the third row a REPEAT of the last kept value
        // (5, 0, 0 — and 1, 3, 3 for 1, 2, 2, 3), i.e. it keeps the DISTINCT-count shape and pads. That
        // contradicts both the page and itself (a UNIQUE result containing a duplicate), so this file
        // pins only what the two readings agree on: the values in order at the front, and a SUM that is
        // the same either way here because the padded element is the 0. See the task report — the shape
        // is an open question for the producer contract, not a settled expectation.
        // Observed today: #REF! (INDEX) / #NAME? (SUM).
        await Assert
            .That(Num(Calc("=INDEX(UNIQUE(Q1:Q4,FALSE,TRUE),1)", Duplicates)))
            .IsEqualTo(5.0);
        await Assert
            .That(Num(Calc("=INDEX(UNIQUE(Q1:Q4,FALSE,TRUE),2)", Duplicates)))
            .IsEqualTo(0.0);
        await Assert.That(Num(Calc("=SUM(UNIQUE(Q1:Q4,FALSE,TRUE))", Duplicates))).IsEqualTo(5.0);
    }

    [Test]
    public async Task Unique_TreatsBlankAsItsOwnKey_AndKeepsAnErrorRow()
    {
        // Correction M4, re-measured. A5:A8 = 7, <blank>, "t", 7 has THREE distinct rows — 7, <blank>,
        // "t" — so the blank is its own key: it merges with neither the 0 nor the "" that
        // ValueCoercion.AreEqual equates it with (that equality is also non-transitive, since 0 and ""
        // are not equal to each other). It survives as a blank, in place: ISBLANK of the second row is
        // TRUE, COUNTA counts 2 (7 and "t") and COUNT counts 1 (only the 7 is a number, a blank is not a
        // 0). An error row is likewise kept as a row and as an error.
        // Oracle 26.6.0, 2026-09-10, plain == CSE. Observed today: #NAME? (ROWS), 1 (COUNTA), 0 (COUNT),
        // FALSE (ISBLANK), #REF! (INDEX).
        await Assert.That(Num(Calc("=ROWS(UNIQUE(A5:A8))", WithBlank))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=COUNTA(UNIQUE(A5:A8))", WithBlank))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=COUNT(UNIQUE(A5:A8))", WithBlank))).IsEqualTo(1.0);
        await Assert.That(Calc("=ISBLANK(INDEX(UNIQUE(A5:A8),2))", WithBlank) as bool?).IsTrue();
        await Assert.That(Num(Calc("=INDEX(UNIQUE(A5:A8),1)", WithBlank))).IsEqualTo(7.0);
        await Assert.That(Calc("=INDEX(UNIQUE(A5:A8),3)", WithBlank)).IsEqualTo("t");

        await Assert.That(Num(Calc("=ROWS(UNIQUE(E1:E3))", WithError))).IsEqualTo(3.0);
        await Assert
            .That(Calc("=INDEX(UNIQUE(E1:E3),2)", WithError))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task TheSelectors_PreserveBlank_TheyDoNotNormalizeItToZero()
    {
        // Correction M4's headline, re-measured and pointed the other way round from the correction
        // itself: there is NO normalization seam. FILTER hands the blank through untouched, so
        // COUNTA(FILTER(A5:A8,…)) equals COUNTA(A5:A8) = 3 rather than 4 — a normalized blank would be a
        // 0, which COUNTA counts. ISBLANK still sees a blank through the selection, while the same blank
        // still compares equal to 0 through the comparison channel (SUMPRODUCT counts one).
        // Oracle 26.6.0, 2026-09-10, plain == CSE for all four rows. NOTE the one shape whose two modes
        // DISAGREE and which is therefore not pinned: COUNTA(SORT(A5:A8)) is 4 plain but 3 CSE (COUNT:
        // 3 plain, 2 CSE) — the CSE column matches the raw range, so blanks survive SORT as well.
        // Observed today: 1 (COUNTA of FILTER), FALSE (ISBLANK), #NAME? (SUMPRODUCT); COUNTA(A5:A8) = 3
        // already green as the control.
        await Assert.That(Num(Calc("=COUNTA(A5:A8)", WithBlank))).IsEqualTo(3.0);
        await Assert
            .That(Num(Calc("=COUNTA(FILTER(A5:A8,A5:A8<>\"zzz\"))", WithBlank)))
            .IsEqualTo(3.0);
        await Assert
            .That(Calc("=ISBLANK(INDEX(FILTER(A5:A8,A5:A8<>\"zzz\"),2))", WithBlank) as bool?)
            .IsTrue();
        await Assert
            .That(Num(Calc("=SUMPRODUCT(--(FILTER(A5:A8,A5:A8<>\"zzz\")=0))", WithBlank)))
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task TheReferenceGuard_SeesAGhostSheet_ThroughTheThreeSelectors()
    {
        // Item 13. ReferenceGuard.MissingSheet is the SYNTACTIC pass the error-IGNORING family (COUNT, and
        // ROWS before its own resolution) runs over its argument nodes so a deleted sheet is a structural
        // #REF! rather than an empty range it silently counts as 0; without an arm for FILTER/SORT/UNIQUE the
        // guard's default arm ignores a producer and COUNT answers 0, exactly the hole the guard exists to
        // close. The three arms stand for the producer's SOURCE argument the way the unary-plus arm stands
        // for its operand; SEQUENCE has no reference argument and needs none.
        //
        // The project's policy here is set by MissingSheetReferenceTests (COUNT(Ghost!A:A) = #REF!), and the
        // oracle does NOT share it: Aspose.Cells 26.6.0, 2026-09-10, plain == CSE, COUNT(Ghost!A1:A3) = 0 and
        // COUNT(FILTER(Ghost!A1:A3,Ghost!B1:B3>0)) = COUNT(SORT(Ghost!A1:A3)) = COUNT(UNIQUE(Ghost!A1:A3)) = 0
        // — the same silent hole, one function up — while SUM(Ghost!A1:A3), SUM(FILTER(Ghost!…)) and
        // ROWS(FILTER(Ghost!A1:A3,Ghost!B1:B3>0)) are #REF! in both modes. The COUNT rows therefore pin the
        // existing policy, not the oracle; the ROWS row pins both.
        // Observed today (registered, no arm): 0 for the three COUNT rows and #REF! for ROWS.
        await Assert
            .That(Calc("=COUNT(FILTER(Ghost!A1:A3,Ghost!B1:B3>0))", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNT(SORT(Ghost!A1:A3))", Grid)).IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=COUNT(UNIQUE(Ghost!A1:A3))", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=ROWS(FILTER(Ghost!A1:A3,Ghost!B1:B3>0))", Grid))
            .IsEqualTo(ErrorValue.Reference);
    }

    // ------------------------------------------------------------------ consumers

    [Test]
    public async Task RowsAndColumns_OverAProducer_AnswerTheProducersShape()
    {
        // ROWS(FILTER(…)) is the formula that makes this feature testable at all, and it is the one
        // consumer that needs a change of its own (this phase's item 14: ROWS/COLUMNS must ask
        // ArrayEvaluation.TryStream for a shape instead of demanding a reference). Measured on this build
        // TODAY, for the reason: ROWS(A1:A3*2) and COLUMNS(A1:A3*2) are both #VALUE! — the gate is
        // missing for EVERY computed array, not just for producers — while SUM(A1:A3*2) is already 28.
        // Oracle 26.6.0, 2026-09-10, plain == CSE for every row here (ROWS(A1:A3*2) itself is 3 / 1).
        // Observed today: #NAME? for all of them.
        await Assert.That(Num(Calc("=ROWS(FILTER(A1:A3,A1:A3>0))", Grid))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=COLUMNS(FILTER(A1:A3,A1:A3>0))", Grid))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=ROWS(FILTER(A1:B3,A1:A3>0))", Grid))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=COLUMNS(FILTER(A1:C1,A1:C1>0))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(FILTER(A1:A3,1))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(FILTER(A1:A3,A1:A3>100,0))", Grid))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=ROWS(SEQUENCE(5))"))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=COLUMNS(SEQUENCE(2,3))"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(SORT(A1))", Grid))).IsEqualTo(1.0);

        // The error arm: an empty FILTER is #CALC! and a bad SEQUENCE size is #VALUE! through ROWS too,
        // so ROWS must hand the producer's own error out rather than reporting a missing reference.
        await Assert.That(Calc("=ROWS(FILTER(A1:A3,A1:A3>100))", Grid)).IsEqualTo(CalcError);
        await Assert.That(Calc("=ROWS(SEQUENCE(-1))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ROWS(SEQUENCE(0))")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task RowsAndColumns_OverAnyComputedArray_AnswerItsShape()
    {
        // Item 14's gate is ArrayEvaluation.TryStream, so it answers for EVERY computed array, not only for
        // the four producers: the binary, IF, lifted-function, unary and ROW forms all counted as a scalar
        // (or reported the argument's own #VALUE!) before it. Oracle 26.6.0, 2026-09-10, plain == CSE for
        // every row here. Observed today: #VALUE! for the first seven rows (ReferencePosition.TryResolve
        // reports the bare array expression's own error), 3 for the two producer rows (green already —
        // an error ELEMENT inside a 3x1 result is not the producer's own error).
        await Assert.That(Num(Calc("=ROWS(A1:A3*2)", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=COLUMNS(A1:B3*2)", Grid))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=ROWS(IF(A1:A3>0,A1:A3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(LEN(A1:A3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(-A1:A3)", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(ROW(A1:A3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(IF(A1:A3>0,1/0))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(FILTER(E1:E3,TRUE))", WithError))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ROWS(SORT(E1:E3))", WithError))).IsEqualTo(3.0);
    }

    [Test]
    public async Task RowsAndColumns_OverAOneByOneErrorArray_ReportThatError()
    {
        // The error arm of item 14, and the reason a 1x1 stream is treated as the scalar it stands for: a
        // producer's own failure is a 1x1 singleton carrying the error (ArrayShaping), and ROWS must hand it
        // out rather than count it as one row. The rule is "a 1x1 array whose only element is an error IS
        // that error" — the same treatment a scalar error already gets on the reference path (ROWS(1/0) is
        // #DIV/0! there) — because the engine cannot tell a producer's failure from a kept error element.
        // Oracle 26.6.0, 2026-09-10, plain == CSE for every pinned row.
        //
        // Not pinned, because the oracle's two modes split or the oracle distinguishes by PROVENANCE, which
        // no shape rule can follow: ROWS(FILTER(A1:A3,A1:A3>100,1/0)) is 1 CSE / #DIV/0! plain,
        // ROWS(IF(A1:A1>0,1/0)) is 1 CSE / #DIV/0! plain, and ROWS(SEQUENCE(1)/0) is 1 in BOTH modes while
        // ROWS(SEQUENCE(1)*E2) — the same 1x1 #DIV/0!, from a cell instead of a literal — is #DIV/0! in both.
        // MySheet answers #DIV/0! for all four under the rule above. Also outside this item: ROWS(E2) and
        // ROWS(E2:E2) are #DIV/0! on the oracle (both modes) and 1 here — the reference path, untouched.
        // Observed today: #DIV/0! for the FILTER/SORT/UNIQUE/SEQUENCE rows through TryResolve reporting the
        // producer's collapsed value, #VALUE! for the binary row — every row green or red for a reason the
        // gate does not own, which is why the mutation in the commit body is the proof, not the colour.
        (string, object)[] errors = [.. Grid, .. WithError];

        await Assert.That(Calc("=ROWS(SEQUENCE(1,1,1/0))")).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=ROWS(FILTER(E2:E2,TRUE))", errors))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=ROWS(FILTER(E1:E3,E1:E3=5))", errors))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=ROWS(SORT(E2))", errors)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=ROWS(UNIQUE(E2:E2))", errors)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=ROWS(SEQUENCE(1)*E2)", errors)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=COLUMNS(SEQUENCE(-1))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=COLUMNS(FILTER(A1:A3,A1:A3>100))", Grid)).IsEqualTo(CalcError);
    }

    [Test]
    public async Task TheFlatteningFamily_StreamsAComputedArray_ElementByElement()
    {
        // Item 15. ArgumentFlattening.FlattenComputedValues — COUNTA, CONCAT and TEXTJOIN's shared argument
        // walk — gains the mini-CSE arm (ArrayEvaluation.TryStream) at the top of its default branch, so a
        // computed array is walked element by element, row-major, instead of being evaluated once to a
        // scalar (a producer's top-left, or a bare array expression's #VALUE!). It lights up COUNTA over a
        // producer, Excel's distinct-count idiom, and fixes the pre-existing array forms with it.
        // Oracle 26.6.0, 2026-09-10; the CSE column is pinned. Plain and CSE agree for every producer row;
        // they split for the bare array forms (COUNTA(IF(…)) and COUNTA(A1:A3*2) are 1 plain, CONCAT and
        // TEXTJOIN over A1:A3*2 are #VALUE! plain).
        // Observed today: 1 (both COUNTA rows), #VALUE! (CONCAT/TEXTJOIN over A1:A3*2), 1 / "5" / "5" / …
        // for the producer rows (the top-left collapse).
        await Assert.That(Num(Calc("=COUNTA(IF(A1:A3>0,A1:A3))", Grid))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=COUNTA(A1:A3*2)", Grid))).IsEqualTo(3.0);
        await Assert.That(Calc("=CONCAT(A1:A3*2)", Grid) as string).IsEqualTo("10018");
        await Assert
            .That(Calc("=TEXTJOIN(\",\",TRUE,A1:A3*2)", Grid) as string)
            .IsEqualTo("10,0,18");

        await Assert.That(Num(Calc("=COUNTA(SEQUENCE(2,3))"))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=COUNTA(FILTER(A1:B3,A1:A3>0))", Grid))).IsEqualTo(4.0);
        await Assert.That(Calc("=CONCAT(SEQUENCE(3))") as string).IsEqualTo("123");
        await Assert.That(Calc("=CONCAT(SEQUENCE(2,3))") as string).IsEqualTo("123456");
        await Assert.That(Calc("=CONCAT(FILTER(A1:B3,A1:A3>0))", Grid) as string).IsEqualTo("5193");
        await Assert.That(Calc("=CONCAT(SORT(A1:A3,1,-1))", Grid) as string).IsEqualTo("950");
        await Assert
            .That(Calc("=TEXTJOIN(\",\",TRUE,FILTER(A1:A3,A1:A3>0))", Grid) as string)
            .IsEqualTo("5,9");
        await Assert
            .That(Calc("=TEXTJOIN(\",\",TRUE,SORT(A1:A3))", Grid) as string)
            .IsEqualTo("0,5,9");
        await Assert
            .That(Calc("=TEXTJOIN(\",\",TRUE,UNIQUE(A1:A3))", Grid) as string)
            .IsEqualTo("5,0,9");

        // The blank survives the walk as a blank: CONCAT skips it, TEXTJOIN with ignore_empty FALSE keeps
        // its slot.
        await Assert
            .That(Calc("=CONCAT(FILTER(A5:A8,A5:A8<>\"zzz\"))", WithBlank) as string)
            .IsEqualTo("7t7");
        await Assert
            .That(Calc("=TEXTJOIN(\",\",FALSE,FILTER(A5:A8,A5:A8<>\"zzz\"))", WithBlank) as string)
            .IsEqualTo("7,,t,7");

        // A producer's own failure is a 1x1 error element: COUNTA counts it (an error is not blank), the
        // text joiners propagate it.
        await Assert.That(Num(Calc("=COUNTA(SEQUENCE(-1))"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=COUNTA(FILTER(A1:A3,A1:A3>100))", Grid))).IsEqualTo(1.0);
        await Assert.That(Calc("=CONCAT(SEQUENCE(-1))")).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=TEXTJOIN(\",\",TRUE,FILTER(A1:A3,A1:A3>100))", Grid))
            .IsEqualTo(CalcError);
    }

    [Test]
    public async Task CountBlank_OverAComputedArray_IsRefError_LikeTheCriteriaFamily()
    {
        // COUNTBLANK is the one caller of FlattenComputedValues that must NOT stream: Excel defines it over
        // a range, and the oracle rejects every computed array in that slot the way the criteria family's
        // Rule B (Phase 11a, PositionalRange.RejectComputedArray) does — #REF! for a producer in BOTH
        // modes, #REF! array-entered / #VALUE! plain for a bare array expression. Oracle 26.6.0,
        // 2026-09-10; the CSE column is pinned, COUNTBLANK(A5:A8) = 1 is the control.
        // Not pinned: COUNTBLANK(IF(A5:A8<>"zzz",A5:A8)) is 1 CSE / #VALUE! plain, and COUNTIF/SUMIF over
        // the same IF(…) are 2 / 14 CSE — the oracle reads IF(range,…) as reference-returning, an exception
        // Rule B does not carve out either; MySheet answers #REF! for it here as it does in the criteria
        // family. Observed today: 0 for every computed row (the top-left collapse is never blank).
        await Assert.That(Num(Calc("=COUNTBLANK(A5:A8)", WithBlank))).IsEqualTo(1.0);
        await Assert
            .That(Calc("=COUNTBLANK(FILTER(A5:A8,TRUE))", WithBlank))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=COUNTBLANK(SORT(A5:A8))", WithBlank))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTBLANK(UNIQUE(A1:A3))", Grid)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTBLANK(SEQUENCE(1))")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTBLANK(A5:A8*1)", WithBlank)).IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=COUNTBLANK(FILTER(A5:A8,A5:A8<>\"zzz\",))", WithBlank))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Concatenate_TakesTheTopLeftOfAProducer_ItDoesNotExpandIt()
    {
        // CONCATENATE is the other caller that must not stream: it joins SCALARS, and over an array the
        // oracle answers the array's top-left (CONCATENATE(FILTER(A1:A3,A1:A3>0)) = "5",
        // CONCATENATE(SEQUENCE(3)) = "1", plain == CSE; CONCATENATE(LEN(A1:A3)) = "1" and
        // CONCATENATE(A1:A3) = "5" array-entered, #VALUE! plain) where CONCAT expands ("59", "123"). So its
        // walk keeps the pre-item-15 default branch, which evaluates a producer to its top-left through
        // ArrayEvaluation.FirstElement. Oracle 26.6.0, 2026-09-10. Observed today: "5" and "1" already —
        // this pin exists so the streaming arm cannot silently take CONCATENATE with it.
        // NOT covered, pre-existing: CONCATENATE(A1:A3) expands the RANGE to "509" here against the
        // oracle's #VALUE! / "5" — a classification question (CONCATENATE is scalar-only in Excel), not
        // this item's.
        await Assert
            .That(Calc("=CONCATENATE(FILTER(A1:A3,A1:A3>0))", Grid) as string)
            .IsEqualTo("5");
        await Assert.That(Calc("=CONCATENATE(SEQUENCE(3))") as string).IsEqualTo("1");
    }

    [Test]
    public async Task TheCriteriaFamily_OverAProducer_IsRefError()
    {
        // Phase 11a Rule B (PositionalRange.RejectComputedArray) already answers #REF! for a range slot
        // holding a non-reference node the mini-CSE would stream, so these rows should turn green with no
        // work here the moment the four functions are registered — they are this file's proof of that
        // handover, and CriteriaComputedArgumentTests' header asks for them by name.
        // Oracle 26.6.0, 2026-09-10: #REF! in BOTH entry modes for every row (a producer in a range slot
        // is one of the few shapes with no mode ambiguity at all).
        // Observed today, all silently wrong rather than loud: 0 for COUNTIF/SUMIF/COUNTIFS/COUNTIF over
        // SEQUENCE/SUMIF over SORT/COUNTIF over UNIQUE and for the sum_range form, #DIV/0! for AVERAGEIF,
        // #VALUE! for SUMIFS and MAXIFS (a length mismatch against the collapsed argument).
        await Assert
            .That(Calc("=COUNTIF(FILTER(A1:A3,A1:A3>0),\">5\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=SUMIF(FILTER(A1:A3,A1:A3>0),\">5\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=AVERAGEIF(FILTER(A1:A3,A1:A3>0),\">0\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTIF(SEQUENCE(5),\">3\")")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTIFS(SEQUENCE(5),\">3\")")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=SUMIF(SORT(A1:A3),\">0\")", Grid)).IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=COUNTIF(UNIQUE(A1:A3),\">0\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=SUMIFS(B1:B3,FILTER(A1:A3,A1:A3>0),\">0\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=MAXIFS(B1:B3,SEQUENCE(3),\">1\")", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=SUMIF(A1:A3,\">0\",FILTER(B1:B3,A1:A3>0))", Grid))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task ALetBoundProducer_InACriteriaSlot_IsAKnownLimitHandedOverByPhase11a()
    {
        // THE FIRST FORMULA A USER OF THIS PHASE WILL WRITE, and a standing limit rather than a fresh
        // bug: Rule B's gate is "not a bare reference node AND array-eligible", and a LET escapes it from
        // BOTH sides — LET(...) in the slot is an opaque scalar to the shape probe, and a LET-BOUND name
        // in the slot is a bare NameReference. Phase 11a pinned the measurable half at today's value
        // (CriteriaComputedArgumentTests:298-318: COUNTIF(LET(r,A1:A3,r*1),">0") = 0, re-measured 0 on
        // this build, against the oracle's #VALUE! plain / #REF! CSE) and handed the producer half to
        // this phase's LET routing correction (M1), which nobody else owns.
        // Oracle 26.6.0, 2026-09-10: #REF! in BOTH modes for the COUNTIF row, and 14 / 2 in both modes
        // for the SUM and ROWS rows — a LET-bound producer works everywhere EXCEPT the criteria slot.
        // Observed today: 0 for the COUNTIF row and #NAME? for the other two. The COUNTIF row will stay
        // RED after the producers land (Let.CaptureValue evaluates FILTER as a scalar before either gate
        // can see it, so it will answer the top-left collapse, not #REF!) until M1's arm through
        // Let/Choose/unary-plus lands. That is the intended shape of this pin, not an accident.
        await Assert
            .That(Calc("=LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,\">0\"))", Grid))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(Num(Calc("=LET(f,FILTER(A1:A3,A1:A3>0),SUM(f))", Grid))).IsEqualTo(14.0);
        await Assert.That(Num(Calc("=LET(f,FILTER(A1:A3,A1:A3>0),ROWS(f))", Grid))).IsEqualTo(2.0);
    }
}
