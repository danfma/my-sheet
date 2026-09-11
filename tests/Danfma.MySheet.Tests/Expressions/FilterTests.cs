using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using static Danfma.MySheet.Expressions.Expression;
using static Danfma.MySheet.Tests.Expressions.SelectionProducerFixture;
using Index = Danfma.MySheet.Expressions.Lookup.Index;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 7 item 8 — <see cref="Filter"/> as an <see cref="IArrayProducer"/>, driven through the AST (see
/// <see cref="SelectionProducerFixture"/>). Oracle: every named value was measured on <b>Aspose.Cells 26.6.0,
/// 2026-09-10</b> in BOTH entry modes (plain <c>Cell.Formula</c> and CSE <c>SetArrayFormula</c>), agreeing
/// unless a test names the column. An assertion with no formula named is an engine-mechanism pin.
/// </summary>
public class FilterTests
{
    private static Filter F(params Expression[] arguments) => new(arguments);

    [Test]
    public async Task KeepsTheRowsWhoseIncludeIsTrue_OnEitherAxis()
    {
        var (workbook, sheet) = Grid();

        // SUM(FILTER(A1:A3,A1:A3>0)) = 14, INDEX(…,2) = 9, INDEX(…,3) = #REF!; SUM(FILTER(A1:B3,A1:A3>0)) =
        // 18 (rows 1 and 3 of both columns); SUM(FILTER(A1:C1,A1:C1>0)) = 6 with COLUMNS 3 ("a" > 0 is
        // TRUE in the classic order, so the text column is kept).
        var byRows = F(Ref("A1:A3", sheet), Ref("A1:A3>0", sheet));
        await Assert.That(Num(Eval(new Sum([byRows]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Index([byRows, Number(2)]), workbook))).IsEqualTo(9.0);
        await Assert
            .That(Eval(new Index([byRows, Number(3)]), workbook))
            .IsEqualTo(ErrorValue.Reference);

        var twoColumns = Build(F(Ref("A1:B3", sheet), Ref("A1:A3>0", sheet)), workbook);
        await Assert.That((twoColumns.Rows, twoColumns.Columns)).IsEqualTo((2, 2));
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:B3", sheet), Ref("A1:A3>0", sheet))]), workbook)))
            .IsEqualTo(18.0);

        var byColumns = Build(F(Ref("A1:C1", sheet), Ref("A1:C1>0", sheet)), workbook);
        await Assert.That((byColumns.Rows, byColumns.Columns)).IsEqualTo((1, 3));
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:C1", sheet), Ref("A1:C1>0", sheet))]), workbook)))
            .IsEqualTo(6.0);
        await Assert.That(At(byColumns, 2).AsObject()).IsEqualTo("a");
    }

    [Test]
    public async Task NothingKept_IsCalc_UnlessIfEmptyAnswers_AndIfEmptyIsAValueEvenWhenBlank()
    {
        var (workbook, sheet) = Grid();
        var none = Ref("A1:A3>100", sheet);
        var some = Ref("A1:A3>0", sheet);

        // SUM(FILTER(A1:A3,A1:A3>100)) = #CALC!, ERROR.TYPE 14, ISERROR TRUE, COUNT 0 — a 1x1 singleton.
        var empty = F(Ref("A1:A3", sheet), none);
        await Assert.That(Eval(new Sum([empty]), workbook)).IsEqualTo(ErrorValue.Calculation);
        await Assert.That(Num(Eval(new ErrorType([empty]), workbook))).IsEqualTo(14.0);
        await Assert.That(Eval(new IsError([empty]), workbook) as bool?).IsTrue();
        await Assert.That(Num(Eval(new Count([empty]), workbook))).IsEqualTo(0.0);
        await Assert
            .That((Build(empty, workbook).Rows, Build(empty, workbook).Columns))
            .IsEqualTo((1, 1));

        // if_empty: SUM(FILTER(A1:A3,A1:A3>100,0)) = 0; FILTER(…,"none") = "none", 1 row (COUNTA 1); an ARRAY
        // if_empty is that array — SUM(FILTER(A1:A3,A1:A3>100,B1:B3)) = 6 with 3 rows and INDEX(…,3) = 3;
        // FILTER(A1:A3,A1:A3>100,1/0) = #DIV/0!, but SUM(FILTER(A1:A3,A1:A3>0,1/0)) = 14 — never read.
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:A3", sheet), none, Number(0))]), workbook)))
            .IsEqualTo(0.0);
        await Assert
            .That(Eval(F(Ref("A1:A3", sheet), none, new StringValue("none")), workbook))
            .IsEqualTo("none");
        await Assert
            .That(Build(F(Ref("A1:A3", sheet), none, new StringValue("none")), workbook).Rows)
            .IsEqualTo(1);
        var arrayIfEmpty = F(Ref("A1:A3", sheet), none, Ref("B1:B3", sheet));
        await Assert.That(Num(Eval(new Sum([arrayIfEmpty]), workbook))).IsEqualTo(6.0);
        await Assert.That(Build(arrayIfEmpty, workbook).Rows).IsEqualTo(3);
        await Assert.That(Num(Eval(new Index([arrayIfEmpty, Number(3)]), workbook))).IsEqualTo(3.0);
        await Assert
            .That(Eval(F(Ref("A1:A3", sheet), none, Ref("1/0", sheet)), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:A3", sheet), some, Ref("1/0", sheet))]), workbook)))
            .IsEqualTo(14.0);

        // An OPEN-RANGE if_empty is the phase's open-range deviation one slot over, and the FINAL REVIEW
        // found it claimed as "pinned" in Filter.cs's remark while no test asserted it. Pinned here, with the
        // number that justifies the refusal rather than merely recording it: the oracle answers
        // SUM(FILTER(A1:A3,A1:A3>100,A:A)) = 14 in both modes — and ROWS of the same formula is **1048576**,
        // the whole column. Matching that would mean materializing a million-row result from a slot nobody
        // reads, so the refusal is not a shortfall, it is the cost guard doing its job. if_empty is built
        // lazily and never probed, so this arrives as a 1x1 #VALUE! from the build, not as a refusal.
        // A CLOSED range in the same slot works and is the contrast that stops this passing vacuously.
        var openIfEmpty = F(Ref("A1:A3", sheet), none, Ref("A:A", sheet));
        await Assert.That(Eval(new Sum([openIfEmpty]), workbook)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Eval(new Rows([openIfEmpty]), workbook)).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:A3", sheet), none, Ref("A1:A3", sheet))]), workbook)))
            .IsEqualTo(14.0);

        // An EMPTY if_empty slot is a blank VALUE, not an omitted argument: FILTER(A1:A3,A1:A3>100,) is a
        // 1x1 blank (ISBLANK TRUE, COUNTA 0, ROWS 1, ERROR.TYPE #N/A), and so is a blank CELL there
        // (FILTER(…,A9): ISBLANK TRUE, COUNTA 0, SUM 0); the element is Blank-kind, which is what COUNTA
        // will see once item 15 lets it walk a producer. This is the one slot that does NOT follow
        // SEQUENCE's "BlankValue means omitted" rule — measured, not inferred.
        var emptySlot = F(Ref("A1:A3", sheet), none, BlankValue.Instance);
        await Assert.That(Eval(new IsBlank([emptySlot]), workbook) as bool?).IsTrue();
        await Assert
            .That(At(Build(emptySlot, workbook), 0).Kind)
            .IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(Build(emptySlot, workbook).Rows).IsEqualTo(1);
        await Assert
            .That(Eval(new ErrorType([emptySlot]), workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
        var blankCell = F(Ref("A1:A3", sheet), none, Ref("A9", sheet));
        await Assert.That(Eval(new IsBlank([blankCell]), workbook) as bool?).IsTrue();
        await Assert
            .That(At(Build(blankCell, workbook), 0).Kind)
            .IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(Num(Eval(new Sum([blankCell]), workbook))).IsEqualTo(0.0);
    }

    [Test]
    public async Task AnIncludeThatMatchesNeitherAxis_IsValueError_AOneByOneIncluded()
    {
        var (workbook, sheet) = Grid();

        // A short vector, a 2-D include, a row against a rectangle's height, and — the oracle's rule for
        // a 1x1 include, scalar or array, over a RECTANGLE — TRUE, SEQUENCE(1) and B1:B1>0 over A1:B3: all
        // #VALUE!. Correction B1's "a scalar include broadcasts" holds only over a vector source.
        foreach (
            var (source, include) in new (string, Expression)[]
            {
                ("A1:A3", Ref("B1:B2>0", sheet)),
                ("A1:A3", Ref("A1:B3>0", sheet)),
                ("A1:B3", Ref("A1:C1>0", sheet)),
                ("A1:B3", BooleanValue.True),
                ("A1:B3", new Sequence([Number(1)])),
                ("A1:B3", Ref("B1:B1>0", sheet)),
            }
        )
        {
            await Assert
                .That(Eval(new Sum([F(Ref(source, sheet), include)]), workbook))
                .IsEqualTo(ErrorValue.NotValue);
        }

        // The include's own shape decides, not the source's: A1:B1>0 (1x2) over A1:B3 keeps both columns
        // (SUM 20, COLUMNS 2).
        var byColumns = F(Ref("A1:B3", sheet), Ref("A1:B1>0", sheet));
        await Assert.That(Num(Eval(new Sum([byColumns]), workbook))).IsEqualTo(20.0);
        await Assert.That(Build(byColumns, workbook).Columns).IsEqualTo(2);
    }

    [Test]
    public async Task IncludeElements_AnErrorPropagates_UnconvertibleTextIsSkipped_TheTwoWordsCoerce()
    {
        var (workbook, sheet) = Grid();
        Filter Over(string include) => F(Ref("A1:A3", sheet), Ref(include, sheet));

        // SUM(FILTER(A1:A3,E1:E3>0)) = #DIV/0! — an error element is the whole answer.
        await Assert
            .That(Eval(new Sum([Over("E1:E3>0")]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);

        // Text that cannot convert is NOT kept: C1:C3 ("a","A","b") keeps nothing → #CALC!, and answers
        // if_empty ("none"); K1:K3 (" TRUE ","yes","1") the same #CALC!; F1:F3 (TRUE,"x",TRUE) and L1:L3
        // (TRUE," TRUE ",TRUE) keep rows 1 and 3 → SUM 14, COUNT 2 — the skipped element is not an error.
        await Assert
            .That(Eval(new Sum([Over("C1:C3")]), workbook))
            .IsEqualTo(ErrorValue.Calculation);
        await Assert
            .That(
                Eval(F(Ref("A1:A3", sheet), Ref("C1:C3", sheet), new StringValue("none")), workbook)
            )
            .IsEqualTo("none");
        await Assert
            .That(Eval(new Sum([Over("K1:K3")]), workbook))
            .IsEqualTo(ErrorValue.Calculation);
        await Assert.That(Num(Eval(new Sum([Over("F1:F3")]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([Over("F1:F3")]), workbook))).IsEqualTo(2.0);
        await Assert.That(Num(Eval(new Sum([Over("L1:L3")]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([Over("L1:L3")]), workbook))).IsEqualTo(2.0);

        // The words TRUE/FALSE as text coerce (Phase 11b's rule for this slot): G1:G3 ("TRUE","FALSE",
        // "true") keeps rows 1 and 3 → SUM 14, COUNT 2. Numbers are truthy per element: B1:B3 keeps all
        // (14); A1:A3 as its own include drops the 0 (14, COUNT 2).
        await Assert.That(Num(Eval(new Sum([Over("G1:G3")]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([Over("G1:G3")]), workbook))).IsEqualTo(2.0);
        await Assert.That(Num(Eval(new Sum([Over("B1:B3")]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Sum([Over("A1:A3")]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([Over("A1:A3")]), workbook))).IsEqualTo(2.0);
    }

    [Test]
    public async Task AOneElementInclude_KeepsAVectorWhole_OrIsValueError_TextWordsAndBoundaries()
    {
        var (workbook, sheet) = Grid();
        var column = Ref("A1:A3", sheet);
        Filter Over(Expression include) => F(column, include);
        Filter OverWithNone(Expression include) => F(column, include, new StringValue("none"));

        // Truthy: TRUE, 1, "TRUE", "true", the 1x1 array B1:B1>0 and the cell G1 ("TRUE") keep the whole
        // column (14); over A1:C1 the whole row (6); over A1 the cell (5).
        foreach (
            var include in new Expression[]
            {
                BooleanValue.True,
                Number(1),
                new StringValue("TRUE"),
                new StringValue("true"),
                Ref("B1:B1>0", sheet),
                Ref("G1:G1", sheet),
            }
        )
        {
            await Assert.That(Num(Eval(new Sum([Over(include)]), workbook))).IsEqualTo(14.0);
        }
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1:C1", sheet), BooleanValue.True)]), workbook)))
            .IsEqualTo(6.0);
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1", sheet), BooleanValue.True)]), workbook)))
            .IsEqualTo(5.0);
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1", sheet), Ref("B1:B1>0", sheet))]), workbook)))
            .IsEqualTo(5.0);

        // Falsy keeps nothing as #VALUE! (NOT #CALC!) — FALSE, 0, the blank cell A9, A9:A9, B1:B1>5 and
        // the cell G2 ("FALSE") — and if_empty answers for every one of them.
        foreach (
            var include in new Expression[]
            {
                BooleanValue.False,
                Number(0),
                Ref("A9", sheet),
                Ref("A9:A9", sheet),
                Ref("B1:B1>5", sheet),
                Ref("G2:G2", sheet),
            }
        )
        {
            await Assert
                .That(Eval(new Sum([Over(include)]), workbook))
                .IsEqualTo(ErrorValue.NotValue);
            await Assert.That(Eval(OverWithNone(include), workbook)).IsEqualTo("none");
        }

        // Unconvertible is #VALUE! here even WITH if_empty: "x", " TRUE ", "yes", "1" and the cells F2
        // ("x") and K1 (" TRUE ") — the boundaries of the two-word rule (no trimming, the two words only).
        foreach (
            var include in new Expression[]
            {
                new StringValue("x"),
                new StringValue(" TRUE "),
                new StringValue("yes"),
                new StringValue("1"),
                Ref("F2:F2", sheet),
                Ref("K1:K1", sheet),
            }
        )
        {
            await Assert
                .That(Eval(new Sum([Over(include)]), workbook))
                .IsEqualTo(ErrorValue.NotValue);
            await Assert.That(Eval(OverWithNone(include), workbook)).IsEqualTo(ErrorValue.NotValue);
        }

        // An error is itself, with or without if_empty: 1/0 and the cell E2.
        await Assert
            .That(Eval(new Sum([Over(Ref("1/0", sheet))]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Eval(OverWithNone(Ref("1/0", sheet)), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Eval(new Sum([Over(Ref("E2:E2", sheet))]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task PreservesBlank_ThroughTheSelection_AndKeepsAnErrorElement()
    {
        var (workbook, sheet) = Grid();

        // FILTER(A5:A8,A5:A8<>"zzz") keeps all four rows and the blank goes through untouched: the
        // operand's four elements are 7, Blank-kind, "t", 7 (a normalized 0 would be Number-kind, and
        // COUNTA — the oracle's 3 = COUNTA(A5:A8) — would count 4); ISBLANK(INDEX(…,2)) TRUE;
        // SUMPRODUCT(--(FILTER(…)=0)) = 1 — the same blank still compares equal to 0 through the
        // comparison channel. COUNTA itself is not asserted over the producer here: its argument walk
        // (ArgumentFlattening's default arm) still evaluates a computed array as one scalar — the phase's
        // item 15, pinned in DynamicArrayTests.TheSelectors_PreserveBlank — and counts 1 today.
        var kept = F(Ref("A5:A8", sheet), Ref("A5:A8<>\"zzz\"", sheet));
        var operand = Build(kept, workbook);
        await Assert.That((operand.Rows, operand.Columns)).IsEqualTo((4, 1));
        await Assert.That(Num(At(operand, 0).AsObject())).IsEqualTo(7.0);
        await Assert.That(At(operand, 1).Kind).IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(At(operand, 2).AsObject()).IsEqualTo("t");
        await Assert
            .That(Eval(new IsBlank([new Index([kept, Number(2)])]), workbook) as bool?)
            .IsTrue();
        var equalsZero = new BinaryOperation(BinaryOperator.Equal, kept, Number(0));
        var doubleNegated = new UnaryOperation(
            UnaryOperator.Negate,
            new UnaryOperation(UnaryOperator.Negate, equalsZero)
        );
        await Assert.That(Num(Eval(new SumProduct([doubleNegated]), workbook))).IsEqualTo(1.0);

        // A kept error element stays an element: INDEX(FILTER(E1:E3,B1:B3>0),2) = #DIV/0! with 3 rows.
        var withError = F(Ref("E1:E3", sheet), Ref("B1:B3>0", sheet));
        await Assert
            .That(Eval(new Index([withError, Number(2)]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Build(withError, workbook).Rows).IsEqualTo(3);
    }

    [Test]
    public async Task ASourceThatIsAScalar_OrAReferenceValue_IsAOneByOneOrTheRectangle()
    {
        var (workbook, sheet) = Grid();

        // Blocker B1: SUM(FILTER(A1,TRUE)) = 5; a bare =FILTER("x",TRUE) is "x"; =FILTER(1/0,TRUE) is
        // #DIV/0!. A REFERENCE-kind scalar is the rectangle it denotes, not a wrapped reference:
        // SUM(FILTER(OFFSET(A1,0,0,3,1),TRUE)) = 14 with 3 rows (OFFSET stands in for INDIRECT, which
        // the oracle answers identically but which cannot resolve without a current cell in this harness).
        await Assert
            .That(Num(Eval(new Sum([F(Ref("A1", sheet), BooleanValue.True)]), workbook)))
            .IsEqualTo(5.0);
        await Assert
            .That(Eval(F(new StringValue("x"), BooleanValue.True), workbook))
            .IsEqualTo("x");
        await Assert
            .That(Eval(F(Ref("1/0", sheet), BooleanValue.True), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        var offset = F(Ref("OFFSET(A1,0,0,3,1)", sheet), BooleanValue.True);
        await Assert.That(Num(Eval(new Sum([offset]), workbook))).IsEqualTo(14.0);
        await Assert.That(Build(offset, workbook).Rows).IsEqualTo(3);
    }

    [Test]
    public async Task ARefusedArrayOrInclude_RefusesTheBuild_InLockstepWithTheProbe()
    {
        var (workbook, sheet) = Grid();
        var context = new EvaluationContext(workbook);

        // An open range in the array or the include slot refuses the producer (the phase's pinned
        // deviation: SUM(FILTER(A:A,A:A>0)) is #VALUE! here against the oracle's 28 array-entered);
        // probe and build agree, and the consumer's scalar path reaches FirstElement's #VALUE!.
        foreach (
            var refused in new[]
            {
                F(Ref("A:A", sheet), Ref("A:A>0", sheet)),
                F(Ref("A1:A3", sheet), Ref("A:A>0", sheet)),
                F(Ref("A:A", sheet), BooleanValue.True),
            }
        )
        {
            await Assert.That(Eligible(refused, workbook)).IsFalse();
            await Assert.That(TryBuild(refused, workbook, out _)).IsFalse();
            await Assert.That(ArrayEvaluation.TryEvaluate(refused, context, out _)).IsFalse();
            await Assert.That(Eval(new Sum([refused]), workbook)).IsEqualTo(ErrorValue.NotValue);
        }

        // if_empty is NOT probed (it is built lazily), so an open range there never refuses: with rows
        // kept it is never read (14, still eligible); with nothing kept it is a loud 1x1 #VALUE! from the
        // build — the same open-range deviation one slot over (the oracle answers the column).
        var unread = F(Ref("A1:A3", sheet), Ref("A1:A3>0", sheet), Ref("A:A", sheet));
        await Assert.That(Eligible(unread, workbook)).IsTrue();
        await Assert.That(TryBuild(unread, workbook, out _)).IsTrue();
        await Assert.That(Num(Eval(new Sum([unread]), workbook))).IsEqualTo(14.0);
        var read = F(Ref("A1:A3", sheet), Ref("A1:A3>100", sheet), Ref("A:A", sheet));
        await Assert.That(Eligible(read, workbook)).IsTrue();
        await Assert.That(TryBuild(read, workbook, out _)).IsTrue();
        await Assert.That(Eval(new Sum([read]), workbook)).IsEqualTo(ErrorValue.NotValue);

        // Every producer error is a 1x1 array, never a refusal: eligible and built for the mismatch, the
        // include error and the empty result alike.
        foreach (
            var erroring in new[]
            {
                F(Ref("A1:A3", sheet), Ref("B1:B2>0", sheet)),
                F(Ref("A1:A3", sheet), Ref("E1:E3>0", sheet)),
                F(Ref("A1:A3", sheet), Ref("A1:A3>100", sheet)),
                F(Ref("A1:A3", sheet), BooleanValue.False),
            }
        )
        {
            await Assert.That(Eligible(erroring, workbook)).IsTrue();
            await Assert.That(TryBuild(erroring, workbook, out var operand)).IsTrue();
            await Assert.That((operand.Rows, operand.Columns)).IsEqualTo((1, 1));
        }
    }

    [Test]
    public async Task InACell_AnswersTheTopLeftElement()
    {
        var (workbook, sheet) = Grid();

        // =FILTER(A1:B3,A1:A3>0) is 5 (top-left of rows 1 and 3); =FILTER(A1:A3,A1:A3>100) is #CALC!.
        await Assert
            .That(Num(Eval(F(Ref("A1:B3", sheet), Ref("A1:A3>0", sheet)), workbook)))
            .IsEqualTo(5.0);
        await Assert
            .That(Eval(F(Ref("A1:A3", sheet), Ref("A1:A3>100", sheet)), workbook))
            .IsEqualTo(ErrorValue.Calculation);
    }
}
