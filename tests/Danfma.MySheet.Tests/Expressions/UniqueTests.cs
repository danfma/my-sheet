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
/// Phase 7 item 10 — <see cref="Unique"/> as an <see cref="IArrayProducer"/>, driven through the AST (see
/// <see cref="SelectionProducerFixture"/>). Oracle: every named value was measured on <b>Aspose.Cells 26.6.0,
/// 2026-09-10</b> in BOTH entry modes, agreeing — except the <c>exactly_once</c> SHAPE, which follows
/// Microsoft's page against the oracle and says so where it is pinned.
/// </summary>
public class UniqueTests
{
    private static Unique U(params Expression[] arguments) => new(arguments);

    private static Expression Nth(Expression producer, int row) =>
        new Index([producer, Number(row)]);

    [Test]
    public async Task KeepsFirstAppearanceOrder_ByRowOrByColumn()
    {
        var (workbook, sheet) = Grid();

        // Q1:Q4 = 9, 5, 9, 0 gives 9, 5, 0 (first appearance, not sorted): SUM 14, 3 rows (the oracle's
        // COUNTA 3 is pinned in DynamicArrayTests; COUNTA's own walk over a computed array is item 15).
        var distinct = U(Ref("Q1:Q4", sheet));
        await Assert.That(Num(Eval(Nth(distinct, 1), workbook))).IsEqualTo(9.0);
        await Assert.That(Num(Eval(Nth(distinct, 2), workbook))).IsEqualTo(5.0);
        await Assert.That(Num(Eval(Nth(distinct, 3), workbook))).IsEqualTo(0.0);
        await Assert.That(Num(Eval(new Sum([distinct]), workbook))).IsEqualTo(14.0);
        await Assert.That(Build(distinct, workbook).Rows).IsEqualTo(3);

        // Rows compare whole: N1:O4 has four distinct rows; T1:U3 = (1,"a"), (1,"a"), (1,"A") has two.
        await Assert.That(Build(U(Ref("N1:O4", sheet)), workbook).Rows).IsEqualTo(4);
        await Assert.That(Build(U(Ref("T1:U3", sheet)), workbook).Rows).IsEqualTo(2);

        // by_col compares columns: UNIQUE(A1:B3,TRUE) keeps 2 columns, UNIQUE(A1:C1,TRUE) 3,
        // UNIQUE(T1:U3,TRUE) 2, and UNIQUE(A1:A3,TRUE) is the one 3x1 column (SUM 14).
        await Assert
            .That(Build(U(Ref("A1:B3", sheet), BooleanValue.True), workbook).Columns)
            .IsEqualTo(2);
        await Assert
            .That(Build(U(Ref("A1:C1", sheet), BooleanValue.True), workbook).Columns)
            .IsEqualTo(3);
        await Assert
            .That(Build(U(Ref("T1:U3", sheet), BooleanValue.True), workbook).Columns)
            .IsEqualTo(2);
        var oneColumn = Build(U(Ref("A1:A3", sheet), BooleanValue.True), workbook);
        await Assert.That((oneColumn.Rows, oneColumn.Columns)).IsEqualTo((3, 1));
        await Assert
            .That(Num(Eval(new Sum([U(Ref("A1:A3", sheet), BooleanValue.True)]), workbook)))
            .IsEqualTo(14.0);
    }

    [Test]
    public async Task IsCaseSensitive_WhereMatchAndCountIfAreNot()
    {
        var (workbook, sheet) = Grid();

        // C1:C3 = "a", "A", "b" keeps all THREE rows with "A" second; MATCH("A",UNIQUE(…),0)
        // = 1 and COUNTIF(C1:C3,"a") = 2 stay case-INSENSITIVE beside it — the case rule belongs to
        // UNIQUE's key comparison and nowhere else. The Microsoft UNIQUE page as fetched 2026-09-10 says
        // nothing about case either way, so the measurement stands alone (P0).
        var distinct = U(Ref("C1:C3", sheet));
        await Assert.That(Build(distinct, workbook).Rows).IsEqualTo(3);
        await Assert.That(Eval(Nth(distinct, 2), workbook)).IsEqualTo("A");
        await Assert
            .That(Num(Eval(new Match([new StringValue("A"), distinct, Number(0)]), workbook)))
            .IsEqualTo(1.0);
        await Assert.That(Num(Eval(Ref("COUNTIF(C1:C3,\"a\")", sheet), workbook))).IsEqualTo(2.0);
    }

    [Test]
    public async Task ExactlyOnce_KeepsOnlyTheRowsThatOccurOnce_AsManyAsThereAre()
    {
        var (workbook, sheet) = Grid();

        // Q1:Q4 = 9, 5, 9, 0 leaves 5 and 0 (INDEX 1 = 5, INDEX 2 = 0, SUM 5 — oracle, both modes).
        // The SHAPE follows Microsoft's page ("all distinct rows or columns that occur exactly once"):
        // TWO rows. NOT the oracle, which keeps the distinct-count shape and pads it by repeating the
        // last kept value — ROWS(UNIQUE(Q1:Q4,FALSE,TRUE)) = 3 with INDEX(…,3) a second 0 — a UNIQUE
        // result holding a duplicate, contradicted by its own row count; that is why it is not matched.
        var once = U(Ref("Q1:Q4", sheet), BooleanValue.False, BooleanValue.True);
        await Assert.That(Num(Eval(Nth(once, 1), workbook))).IsEqualTo(5.0);
        await Assert.That(Num(Eval(Nth(once, 2), workbook))).IsEqualTo(0.0);
        await Assert.That(Num(Eval(new Sum([once]), workbook))).IsEqualTo(5.0);
        await Assert.That(Build(once, workbook).Rows).IsEqualTo(2);
        await Assert.That(Eval(Nth(once, 3), workbook)).IsEqualTo(ErrorValue.Reference);

        // Nothing occurring once is the EMPTY result — a 1x1 #CALC! (ERROR.TYPE 14, COUNT 0), the page's
        // reading again: the oracle's padding gives it a 1-row 4 for S1:S2 = 4, 4.
        var nothing = U(Ref("S1:S2", sheet), BooleanValue.False, BooleanValue.True);
        await Assert.That(Eval(new Sum([nothing]), workbook)).IsEqualTo(ErrorValue.Calculation);
        await Assert.That(Num(Eval(new ErrorType([nothing]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([nothing]), workbook))).IsEqualTo(0.0);
        await Assert
            .That((Build(nothing, workbook).Rows, Build(nothing, workbook).Columns))
            .IsEqualTo((1, 1));
    }

    [Test]
    public async Task ABlankIsItsOwnKey_KindsNeverCross_AndAnErrorRowIsKept()
    {
        var (workbook, sheet) = Grid();

        // A5:A8 = 7, <blank>, "t", 7: three rows — 7, <blank>, "t" — with the blank still a blank
        // (the element is Blank-kind, ISBLANK TRUE, COUNT 1; the oracle's COUNTA 2 is pinned in
        // DynamicArrayTests and waits on item 15's COUNTA walk); V1:V4 = 0, <blank>, "", FALSE: FOUR rows, the blank equal to
        // none of the three empties ValueCoercion.AreEqual would merge it with; R1:R4 = 1, "1", TRUE, 1:
        // three rows.
        var withBlank = U(Ref("A5:A8", sheet));
        await Assert.That(Build(withBlank, workbook).Rows).IsEqualTo(3);
        await Assert
            .That(At(Build(withBlank, workbook), 1).Kind)
            .IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(Num(Eval(new Count([withBlank]), workbook))).IsEqualTo(1.0);
        await Assert.That(Eval(new IsBlank([Nth(withBlank, 2)]), workbook) as bool?).IsTrue();
        await Assert.That(Num(Eval(Nth(withBlank, 1), workbook))).IsEqualTo(7.0);
        await Assert.That(Eval(Nth(withBlank, 3), workbook)).IsEqualTo("t");
        var empties = U(Ref("V1:V4", sheet));
        await Assert.That(Build(empties, workbook).Rows).IsEqualTo(4);
        await Assert.That(Eval(new IsBlank([Nth(empties, 2)]), workbook) as bool?).IsTrue();
        await Assert.That(Build(U(Ref("R1:R4", sheet)), workbook).Rows).IsEqualTo(3);

        // E1:E3 = 5, #DIV/0!, 9 keeps its error row (3 rows, INDEX 2 = #DIV/0!); W1:W5 = 2, #DIV/0!, 3,
        // #N/A, #DIV/0! keeps FOUR — the same error is the same key (ERROR.TYPE 2 at row 2, 7 at row 4).
        var withError = U(Ref("E1:E3", sheet));
        await Assert.That(Build(withError, workbook).Rows).IsEqualTo(3);
        await Assert.That(Eval(Nth(withError, 2), workbook)).IsEqualTo(ErrorValue.DivByZero);
        var many = U(Ref("W1:W5", sheet));
        await Assert.That(Build(many, workbook).Rows).IsEqualTo(4);
        await Assert.That(Num(Eval(new ErrorType([Nth(many, 2)]), workbook))).IsEqualTo(2.0);
        await Assert.That(Num(Eval(new ErrorType([Nth(many, 4)]), workbook))).IsEqualTo(7.0);
    }

    [Test]
    public async Task TheFlagsCoerceWithTheTextWords_AndTheirErrorsPropagateInOrder()
    {
        var (workbook, sheet) = Grid();
        var column = Ref("Q1:Q4", sheet);
        int Rows(params Expression[] arguments) => Build(U(arguments), workbook).Rows;
        object? Sum(params Expression[] arguments) => Eval(new Sum([U(arguments)]), workbook);

        // by_col over Q1:Q4 (9, 5, 9, 0): TRUE, "TRUE" and 2 compare the one column (4 rows); the blank
        // cell A9 and an empty slot are FALSE (3 rows); "x" is #VALUE!.
        await Assert.That(Rows(column, BooleanValue.True)).IsEqualTo(4);
        await Assert.That(Rows(column, new StringValue("TRUE"))).IsEqualTo(4);
        await Assert.That(Rows(column, Number(2))).IsEqualTo(4);
        await Assert.That(Rows(column, Ref("A9", sheet))).IsEqualTo(3);
        await Assert.That(Rows(column, BlankValue.Instance)).IsEqualTo(3);
        await Assert.That(Sum(column, new StringValue("x"))).IsEqualTo(ErrorValue.NotValue);

        // exactly_once: "TRUE" and 2 turn it on (the two once-rows, page shape); an empty slot is off.
        await Assert.That(Rows(column, BooleanValue.False, new StringValue("TRUE"))).IsEqualTo(2);
        await Assert.That(Rows(column, BooleanValue.False, Number(2))).IsEqualTo(2);
        await Assert.That(Rows(column, BooleanValue.False, BlankValue.Instance)).IsEqualTo(3);

        // Errors in argument order: (1/0,"x") is #DIV/0!, ("x",1/0) is #VALUE!, (FALSE,1/0) is #DIV/0!.
        await Assert
            .That(Sum(column, Ref("1/0", sheet), new StringValue("x")))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Sum(column, new StringValue("x"), Ref("1/0", sheet)))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Sum(column, BooleanValue.False, Ref("1/0", sheet)))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task ASourceThatIsAScalar_OrAReferenceValue_AndARefusedOne()
    {
        var (workbook, sheet) = Grid();

        // Blocker B1: SUM(UNIQUE(A1)) = 5; =UNIQUE("x") is "x"; SUM(UNIQUE(1/0)) is #DIV/0!;
        // SUM(UNIQUE(OFFSET(A1,0,0,3,1))) = 14 — the reference-kind scalar resolves to its rectangle.
        await Assert.That(Num(Eval(new Sum([U(Ref("A1", sheet))]), workbook))).IsEqualTo(5.0);
        await Assert.That(Eval(U(new StringValue("x")), workbook)).IsEqualTo("x");
        await Assert
            .That(Eval(new Sum([U(Ref("1/0", sheet))]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Num(Eval(new Sum([U(Ref("OFFSET(A1,0,0,3,1)", sheet))]), workbook)))
            .IsEqualTo(14.0);

        // An open range refuses the build, in lockstep with the probe; a bad flag never does.
        var refused = U(Ref("A:A", sheet));
        await Assert.That(Eligible(refused, workbook)).IsFalse();
        await Assert.That(TryBuild(refused, workbook, out _)).IsFalse();
        await Assert.That(Eval(new Sum([refused]), workbook)).IsEqualTo(ErrorValue.NotValue);
        var bad = U(Ref("A1:A3", sheet), new StringValue("x"));
        await Assert.That(Eligible(bad, workbook)).IsTrue();
        await Assert.That(TryBuild(bad, workbook, out var operand)).IsTrue();
        await Assert.That((operand.Rows, operand.Columns)).IsEqualTo((1, 1));
    }

    [Test]
    public async Task InACell_AnswersTheTopLeftElement()
    {
        var (workbook, sheet) = Grid();

        // =UNIQUE(A1:A3) is 5; =UNIQUE(Q1:Q4,FALSE,TRUE) is 5 (the first once-row).
        await Assert.That(Num(Eval(U(Ref("A1:A3", sheet)), workbook))).IsEqualTo(5.0);
        await Assert
            .That(
                Num(Eval(U(Ref("Q1:Q4", sheet), BooleanValue.False, BooleanValue.True), workbook))
            )
            .IsEqualTo(5.0);
    }
}
