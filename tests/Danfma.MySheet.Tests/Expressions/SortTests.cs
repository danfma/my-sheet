using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using static Danfma.MySheet.Expressions.Expression;
using static Danfma.MySheet.Tests.Expressions.SelectionProducerFixture;
using Index = Danfma.MySheet.Expressions.Lookup.Index;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 7 item 9 — <see cref="Sort"/> as an <see cref="IArrayProducer"/>, driven through the AST (see
/// <see cref="SelectionProducerFixture"/>). Oracle: every named value was measured on <b>Aspose.Cells 26.6.0,
/// 2026-09-10</b> in BOTH entry modes, agreeing unless a test names the column; three argument shapes
/// crash the oracle and are named as page-based where they appear.
/// </summary>
public class SortTests
{
    private static Sort S(params Expression[] arguments) => new(arguments);

    private static Expression Nth(Expression producer, int row, int? column = null) =>
        column is { } c
            ? new Index([producer, Number(row), Number(c)])
            : new Index([producer, Number(row)]);

    [Test]
    public async Task AscendingByDefault_DescendingOnMinusOne_ByTheChosenColumn()
    {
        var (workbook, sheet) = Grid();

        // INDEX(SORT(A1:A3),k) = 0, 5, 9; INDEX(SORT(A1:A3,1,-1),1) = 9; SUM(SORT(A1:A3)) = 14;
        // INDEX(SORT(N1:O4,2,-1),1,2) = "d" (by the second column, descending).
        var ascending = S(Ref("A1:A3", sheet));
        await Assert.That(Num(Eval(Nth(ascending, 1), workbook))).IsEqualTo(0.0);
        await Assert.That(Num(Eval(Nth(ascending, 2), workbook))).IsEqualTo(5.0);
        await Assert.That(Num(Eval(Nth(ascending, 3), workbook))).IsEqualTo(9.0);
        await Assert
            .That(Num(Eval(Nth(S(Ref("A1:A3", sheet), Number(1), Number(-1)), 1), workbook)))
            .IsEqualTo(9.0);
        await Assert.That(Num(Eval(new Sum([ascending]), workbook))).IsEqualTo(14.0);
        await Assert
            .That(Eval(Nth(S(Ref("N1:O4", sheet), Number(2), Number(-1)), 1, 2), workbook))
            .IsEqualTo("d");
    }

    [Test]
    public async Task ArgumentsTruncate_AndAreValidatedInOrder()
    {
        var (workbook, sheet) = Grid();
        var column = Ref("A1:A3", sheet);
        object? First(params Expression[] arguments) => Eval(Nth(S(arguments), 1), workbook);

        // sort_index truncates and coerces: 1.9, "1" and TRUE are 1 (first 0); an empty slot is omitted
        // (INDEX(SORT(A1:A3,,-1),1) = 9); out of range or unconvertible is #VALUE! — 0, 2, "x", the blank
        // CELL A9 (0), and SORT(A1:B3,3) against two columns.
        await Assert.That(Num(First(column, Number(1.9)))).IsEqualTo(0.0);
        await Assert.That(Num(First(column, new StringValue("1")))).IsEqualTo(0.0);
        await Assert.That(Num(First(column, BooleanValue.True))).IsEqualTo(0.0);
        await Assert.That(Num(First(column, BlankValue.Instance, Number(-1)))).IsEqualTo(9.0);
        await Assert.That(First(column, Number(0))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(First(column, Number(2))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(First(column, new StringValue("x"))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(First(column, Ref("A9", sheet))).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([S(Ref("A1:B3", sheet), Number(3))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);

        // sort_order truncates then must be 1 or -1: -1.5 → -1 (first 9, last 0), 1.5 → 1 (first 0),
        // "-1" → -1, TRUE → 1; 0, 2 and -0.5 are #VALUE!.
        await Assert.That(Num(First(column, Number(1), Number(-1.5)))).IsEqualTo(9.0);
        await Assert
            .That(Num(Eval(Nth(S(column, Number(1), Number(-1.5)), 3), workbook)))
            .IsEqualTo(0.0);
        await Assert.That(Num(First(column, Number(1), Number(1.5)))).IsEqualTo(0.0);
        await Assert.That(Num(First(column, Number(1), new StringValue("-1")))).IsEqualTo(9.0);
        await Assert.That(Num(First(column, Number(1), BooleanValue.True))).IsEqualTo(0.0);
        await Assert.That(First(column, Number(1), Number(0))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(First(column, Number(1), Number(2))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(First(column, Number(1), Number(-0.5))).IsEqualTo(ErrorValue.NotValue);

        // NOT measured — Aspose.Cells 26.6.0 throws NullReferenceException from CalculateFormula for an
        // empty, blank-cell or text sort_order — so these three follow the rule stated above and the
        // Microsoft page: an empty slot is omitted (1), a blank cell is 0 → #VALUE!, text is #VALUE!.
        await Assert.That(Num(First(column, Number(1), BlankValue.Instance))).IsEqualTo(0.0);
        await Assert
            .That(First(column, Number(1), Ref("A9", sheet)))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(First(column, Number(1), new StringValue("x")))
            .IsEqualTo(ErrorValue.NotValue);

        // Argument errors propagate in argument order: SORT(A1:A3,1,1/0) and SORT(A1:A3,1,1,1/0) are
        // #DIV/0!; SORT(A1:A3,"x",0) is #VALUE! (the index is read first).
        //
        // AND SO DOES AN ERROR IN sort_index, which is the ONE row in this phase where the engine follows the
        // oracle's PLAIN column rather than its array-entered one — both final reviewers flagged it, and the
        // sentence in Sort.cs's remark claimed it was pinned here when it was not. Oracle 26.6.0, 2026-09-10:
        // INDEX(SORT(A1:A3,1/0),1) is #DIV/0! PLAIN and #VALUE! ARRAY-ENTERED. This engine answers #DIV/0!,
        // deliberately, because an error in sort_index must not become a DIFFERENT error from one in
        // sort_order or by_col — the two assertions below it would otherwise disagree with this one on the
        // same kind of input. Pinned so that choosing the other column becomes a decision rather than a drift.
        await Assert.That(First(column, Ref("1/0", sheet))).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(First(column, Number(1), Ref("1/0", sheet)))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(First(column, Number(1), Number(1), Ref("1/0", sheet)))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(First(column, new StringValue("x"), Number(0)))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task ByCol_SortsColumnsAgainstTheChosenRow_AndTheFlagCoerces()
    {
        var (workbook, sheet) = Grid();
        var grid = Ref("A1:B3", sheet);
        object? TopLeft(Expression flag, int index = 1, int order = 1) =>
            Eval(Nth(S(grid, Number(index), Number(order), flag), 1, 1), workbook);

        // A1:B3 = [[5,1],[0,2],[9,3]]: by row 1 ascending the columns swap (INDEX(…,1,1) = 1, (1,2) = 5);
        // by row 2 descending they swap too ((1,1) = 1, (2,1) = 2). sort_index runs over ROWS then:
        // SORT(A1:B3,3,1,TRUE) sorts (SUM 20), SORT(A1:B3,4,1,TRUE) is #VALUE!.
        await Assert.That(Num(TopLeft(BooleanValue.True))).IsEqualTo(1.0);
        await Assert
            .That(Num(Eval(Nth(S(grid, Number(1), Number(1), BooleanValue.True), 1, 2), workbook)))
            .IsEqualTo(5.0);
        await Assert.That(Num(TopLeft(BooleanValue.True, 2, -1))).IsEqualTo(1.0);
        await Assert
            .That(Num(Eval(Nth(S(grid, Number(2), Number(-1), BooleanValue.True), 2, 1), workbook)))
            .IsEqualTo(2.0);
        await Assert
            .That(Num(Eval(new Sum([S(grid, Number(3), Number(1), BooleanValue.True)]), workbook)))
            .IsEqualTo(20.0);
        await Assert
            .That(Eval(new Sum([S(grid, Number(4), Number(1), BooleanValue.True)]), workbook))
            .IsEqualTo(ErrorValue.NotValue);

        // The flag coerces with the text words: "TRUE" and 2 sort by column (1); the blank cell A9 and an
        // empty slot are FALSE (by row: 0); "x" is #VALUE!; 1/0 is #DIV/0!.
        await Assert.That(Num(TopLeft(new StringValue("TRUE")))).IsEqualTo(1.0);
        await Assert.That(Num(TopLeft(Number(2)))).IsEqualTo(1.0);
        await Assert.That(Num(TopLeft(Ref("A9", sheet)))).IsEqualTo(0.0);
        await Assert.That(Num(TopLeft(BlankValue.Instance))).IsEqualTo(0.0);
        await Assert.That(TopLeft(new StringValue("x"))).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(TopLeft(Ref("1/0", sheet))).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task IsStable_InBothDirections_AndTextComparesCaseInsensitively()
    {
        var (workbook, sheet) = Grid();

        // N1:O4 = (1,"a"), (2,"b"), (1,"c"), (2,"d"): ascending tags a, c, b, d; descending b, d, a, c —
        // a stable sort in BOTH directions, not a reversal (that would be d, b, c, a).
        var ascending = S(Ref("N1:O4", sheet));
        var descending = S(Ref("N1:O4", sheet), Number(1), Number(-1));
        await Assert.That(Num(Eval(Nth(ascending, 1, 1), workbook))).IsEqualTo(1.0);
        await Assert.That(Eval(Nth(ascending, 1, 2), workbook)).IsEqualTo("a");
        await Assert.That(Eval(Nth(ascending, 2, 2), workbook)).IsEqualTo("c");
        await Assert.That(Eval(Nth(ascending, 3, 2), workbook)).IsEqualTo("b");
        await Assert.That(Eval(Nth(ascending, 4, 2), workbook)).IsEqualTo("d");
        await Assert.That(Eval(Nth(descending, 1, 2), workbook)).IsEqualTo("b");
        await Assert.That(Eval(Nth(descending, 2, 2), workbook)).IsEqualTo("d");
        await Assert.That(Eval(Nth(descending, 3, 2), workbook)).IsEqualTo("a");
        await Assert.That(Eval(Nth(descending, 4, 2), workbook)).IsEqualTo("c");

        // C1:C3 = "a", "A", "b": ascending a, A, b and descending b, a, A — "a" and "A" tie
        // (case-insensitive keys) and keep their source order either way.
        var text = S(Ref("C1:C3", sheet));
        var textDescending = S(Ref("C1:C3", sheet), Number(1), Number(-1));
        await Assert.That(Eval(Nth(text, 1), workbook)).IsEqualTo("a");
        await Assert.That(Eval(Nth(text, 2), workbook)).IsEqualTo("A");
        await Assert.That(Eval(Nth(text, 3), workbook)).IsEqualTo("b");
        await Assert.That(Eval(Nth(textDescending, 1), workbook)).IsEqualTo("b");
        await Assert.That(Eval(Nth(textDescending, 2), workbook)).IsEqualTo("a");
        await Assert.That(Eval(Nth(textDescending, 3), workbook)).IsEqualTo("A");
    }

    [Test]
    public async Task MixedTypes_UseTheClassicNumberTextFalseTrueOrder()
    {
        var (workbook, sheet) = Grid();

        // M1:M4 = TRUE, "x", 2, FALSE sorts to 2, "x", FALSE, TRUE.
        var sorted = S(Ref("M1:M4", sheet));
        await Assert.That(Num(Eval(Nth(sorted, 1), workbook))).IsEqualTo(2.0);
        await Assert.That(Eval(Nth(sorted, 2), workbook)).IsEqualTo("x");
        await Assert.That(Eval(Nth(sorted, 3), workbook) as bool?).IsFalse();
        await Assert.That(Eval(Nth(sorted, 4), workbook) as bool?).IsTrue();
    }

    [Test]
    public async Task BlanksGoLast_InBothDirections_AndArriveAsBlanks()
    {
        var (workbook, sheet) = Grid();

        // A5:A8 = 7, <blank>, "t", 7: ascending 7, 7, "t", <blank>; descending "t", 7, 7, <blank>. The
        // blank is not a 0 (a 0 would lead ascending) and is still a blank when read (ISBLANK TRUE, and
        // the operand's own element is Blank-kind).
        var ascending = S(Ref("A5:A8", sheet));
        var descending = S(Ref("A5:A8", sheet), Number(1), Number(-1));
        await Assert.That(Num(Eval(Nth(ascending, 1), workbook))).IsEqualTo(7.0);
        await Assert.That(Num(Eval(Nth(ascending, 2), workbook))).IsEqualTo(7.0);
        await Assert.That(Eval(Nth(ascending, 3), workbook)).IsEqualTo("t");
        await Assert.That(Eval(new IsBlank([Nth(ascending, 4)]), workbook) as bool?).IsTrue();
        await Assert
            .That(At(Build(ascending, workbook), 3).Kind)
            .IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(Eval(Nth(descending, 1), workbook)).IsEqualTo("t");
        await Assert.That(Num(Eval(Nth(descending, 2), workbook))).IsEqualTo(7.0);
        await Assert.That(Num(Eval(Nth(descending, 3), workbook))).IsEqualTo(7.0);
        await Assert.That(Eval(new IsBlank([Nth(descending, 4)]), workbook) as bool?).IsTrue();
    }

    [Test]
    public async Task ErrorsAreSorted_LastAscending_FirstDescending_InSourceOrderAmongThemselves()
    {
        var (workbook, sheet) = Grid();

        // E1:E3 = 5, #DIV/0!, 9: ascending 5, 9, #DIV/0!; descending #DIV/0! first; SUM(SORT(E1:E3)) is
        // #DIV/0! only because SUM propagates what it is handed.
        var ascending = S(Ref("E1:E3", sheet));
        await Assert.That(Num(Eval(Nth(ascending, 1), workbook))).IsEqualTo(5.0);
        await Assert.That(Num(Eval(Nth(ascending, 2), workbook))).IsEqualTo(9.0);
        await Assert.That(Eval(Nth(ascending, 3), workbook)).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Eval(Nth(S(Ref("E1:E3", sheet), Number(1), Number(-1)), 1), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Eval(new Sum([ascending]), workbook)).IsEqualTo(ErrorValue.DivByZero);

        // W1:W5 = 2, #DIV/0!, 3, #N/A, #DIV/0!: ascending 2, 3, #DIV/0!, #N/A, #DIV/0! (ERROR.TYPE 2, 7,
        // 2 at rows 3-5) and descending #DIV/0!, #N/A, #DIV/0!, 3, 2 — the errors keep their SOURCE order
        // in both directions rather than reversing.
        var many = S(Ref("W1:W5", sheet));
        var manyDescending = S(Ref("W1:W5", sheet), Number(1), Number(-1));
        await Assert.That(Num(Eval(new ErrorType([Nth(many, 3)]), workbook))).IsEqualTo(2.0);
        await Assert.That(Num(Eval(new ErrorType([Nth(many, 4)]), workbook))).IsEqualTo(7.0);
        await Assert.That(Num(Eval(new ErrorType([Nth(many, 5)]), workbook))).IsEqualTo(2.0);
        await Assert
            .That(Num(Eval(new ErrorType([Nth(manyDescending, 1)]), workbook)))
            .IsEqualTo(2.0);
        await Assert
            .That(Num(Eval(new ErrorType([Nth(manyDescending, 2)]), workbook)))
            .IsEqualTo(7.0);
        await Assert
            .That(Num(Eval(new ErrorType([Nth(manyDescending, 3)]), workbook)))
            .IsEqualTo(2.0);
        await Assert.That(Num(Eval(Nth(manyDescending, 4), workbook))).IsEqualTo(3.0);
    }

    [Test]
    public async Task ASourceThatIsAScalar_OrAReferenceValue_AndARefusedOne()
    {
        var (workbook, sheet) = Grid();

        // Blocker B1: SUM(SORT(A1)) = 5 with ROWS 1; =SORT("x") is "x"; SUM(SORT(1/0)) is #DIV/0!;
        // SUM(SORT(OFFSET(A1,0,0,3,1))) = 14 — the reference-kind scalar resolves to its rectangle (the
        // oracle answers INDIRECT("A1:A3") the same, but INDIRECT cannot resolve without a current cell
        // in this harness).
        var cell = S(Ref("A1", sheet));
        await Assert.That(Num(Eval(new Sum([cell]), workbook))).IsEqualTo(5.0);
        await Assert
            .That((Build(cell, workbook).Rows, Build(cell, workbook).Columns))
            .IsEqualTo((1, 1));
        await Assert.That(Eval(S(new StringValue("x")), workbook)).IsEqualTo("x");
        await Assert
            .That(Eval(new Sum([S(Ref("1/0", sheet))]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Num(Eval(new Sum([S(Ref("OFFSET(A1,0,0,3,1)", sheet))]), workbook)))
            .IsEqualTo(14.0);

        // An open range refuses the build, in lockstep with the probe; a bad argument never does.
        var refused = S(Ref("A:A", sheet));
        await Assert.That(Eligible(refused, workbook)).IsFalse();
        await Assert.That(TryBuild(refused, workbook, out _)).IsFalse();
        await Assert.That(Eval(new Sum([refused]), workbook)).IsEqualTo(ErrorValue.NotValue);
        var bad = S(Ref("A1:A3", sheet), Number(2));
        await Assert.That(Eligible(bad, workbook)).IsTrue();
        await Assert.That(TryBuild(bad, workbook, out var operand)).IsTrue();
        await Assert.That((operand.Rows, operand.Columns)).IsEqualTo((1, 1));
    }

    [Test]
    public async Task InACell_AnswersTheTopLeftElement()
    {
        var (workbook, sheet) = Grid();

        // =SORT(A1:B3,1,-1) is 9 (rows ordered 3, 1, 2 — the top-left is A3); =SORT(A1:A3) is 0.
        await Assert
            .That(Num(Eval(S(Ref("A1:B3", sheet), Number(1), Number(-1)), workbook)))
            .IsEqualTo(9.0);
        await Assert.That(Num(Eval(S(Ref("A1:A3", sheet)), workbook))).IsEqualTo(0.0);
    }
}
