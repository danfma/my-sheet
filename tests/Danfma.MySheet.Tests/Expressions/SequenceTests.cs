using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;
using Index = Danfma.MySheet.Expressions.Lookup.Index;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 7 item 7 — <see cref="Sequence"/> as a real <see cref="IArrayProducer"/>, driven through the AST
/// directly because <c>SEQUENCE</c> is not registered yet (registration, the union tag and the
/// classification counts are one later commit): the formula-level pins in <c>DynamicArrayTests</c> stay
/// red until then, and these are the same numbers reached without the parser.
/// </summary>
/// <remarks>
/// Oracle: every named value was measured on <b>Aspose.Cells 26.6.0, 2026-09-10</b>, in BOTH entry modes
/// (plain <c>Cell.Formula</c> and CSE <c>SetArrayFormula</c>), over <c>B1:B3</c> = 1, 2, 3 with <c>A9</c>
/// never written; the two modes agree unless a test names the column. An assertion with no formula named
/// is an engine-mechanism pin (operand shape, projection, the invariant guard), not an oracle claim.
/// </remarks>
public class SequenceTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["B1"] = Number(1);
        sheet["B2"] = Number(2);
        sheet["B3"] = Number(3);

        return (workbook, sheet);
    }

    private static Sequence Seq(params Expression[] arguments) => new(arguments);

    private static object? Eval(Expression node, Workbook workbook) =>
        node.Evaluate(workbook).AsObject();

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static Expression Ref(string formula, Sheet sheet) =>
        ExpressionParser.Parse("=" + formula, sheet);

    private static BinaryOperation DivideByZero() =>
        new(BinaryOperator.Divide, Number(1), Number(0));

    private static ArrayOperand Build(Sequence sequence, Workbook workbook)
    {
        ((IArrayProducer)sequence).TryBuildArrayOperand(
            new EvaluationContext(workbook),
            out var operand
        );
        return operand;
    }

    [Test]
    public async Task FillsRowMajor_WithTheGivenStartAndStep()
    {
        var (workbook, _) = Grid();

        // SUM(SEQUENCE(5)) = 15; INDEX(SEQUENCE(2,3),r,c): (1,3) = 3, (2,1) = 4, (2,2) = 5 — row-major;
        // INDEX(SEQUENCE(5),4) = 4; SUM(SEQUENCE(3,1,10,5)) = 45; SUM(SEQUENCE(2,3,7,1)) = 57.
        await Assert.That(Num(Eval(new Sum([Seq(Number(5))]), workbook))).IsEqualTo(15.0);
        await Assert
            .That(Num(Eval(new Index([Seq(Number(2), Number(3)), Number(1), Number(3)]), workbook)))
            .IsEqualTo(3.0);
        await Assert
            .That(Num(Eval(new Index([Seq(Number(2), Number(3)), Number(2), Number(1)]), workbook)))
            .IsEqualTo(4.0);
        await Assert
            .That(Num(Eval(new Index([Seq(Number(2), Number(3)), Number(2), Number(2)]), workbook)))
            .IsEqualTo(5.0);
        await Assert
            .That(Num(Eval(new Index([Seq(Number(5)), Number(4)]), workbook)))
            .IsEqualTo(4.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(3), Number(1), Number(10), Number(5))]), workbook)))
            .IsEqualTo(45.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(2), Number(3), Number(7), Number(1))]), workbook)))
            .IsEqualTo(57.0);

        // The operand itself: 2x3, and the values at its own extent are 1..6 in row order.
        var operand = Build(Seq(Number(2), Number(3)), workbook);
        await Assert.That(operand.IsArray).IsTrue();
        await Assert.That(operand.Rows).IsEqualTo(2);
        await Assert.That(operand.Columns).IsEqualTo(3);
        await Assert.That(Num(operand.At(4, 2, 3).AsObject())).IsEqualTo(5.0);
    }

    [Test]
    public async Task InACell_AnswersTheTopLeftElement()
    {
        var (workbook, _) = Grid();

        // =SEQUENCE(2,3,7,1) = 7 and =SEQUENCE(2,,,) = 1 — Excel's @ on an array (FirstElement).
        await Assert
            .That(Num(Eval(Seq(Number(2), Number(3), Number(7), Number(1)), workbook)))
            .IsEqualTo(7.0);
        await Assert
            .That(
                Num(
                    Eval(
                        Seq(
                            Number(2),
                            BlankValue.Instance,
                            BlankValue.Instance,
                            BlankValue.Instance
                        ),
                        workbook
                    )
                )
            )
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task TruncatesRowsAndColumns_AndKeepsStartAndStepAsGiven()
    {
        var (workbook, _) = Grid();

        // SUM(SEQUENCE(2.7)) = 3, SUM(SEQUENCE(2,2.9)) = 10, SUM(SEQUENCE(1,1,1.5,0.25)) = 1.5,
        // SUM(SEQUENCE(2,2,0,0)) = 0 (a zero step is a value), SUM(SEQUENCE(2,2,-3,-1)) = -18.
        await Assert.That(Num(Eval(new Sum([Seq(Number(2.7))]), workbook))).IsEqualTo(3.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(2), Number(2.9))]), workbook)))
            .IsEqualTo(10.0);
        await Assert
            .That(
                Num(Eval(new Sum([Seq(Number(1), Number(1), Number(1.5), Number(0.25))]), workbook))
            )
            .IsEqualTo(1.5);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(2), Number(2), Number(0), Number(0))]), workbook)))
            .IsEqualTo(0.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(2), Number(2), Number(-3), Number(-1))]), workbook)))
            .IsEqualTo(-18.0);
    }

    [Test]
    public async Task AnOmittedOptional_DefaultsToOne_ButABlankCellIsAValueAndCoercesToZero()
    {
        var (workbook, sheet) = Grid();

        // SUM(SEQUENCE(2,2,,)) = 10: the empty slots the parser leaves as BlankValue default to 1, exactly
        // as an absent argument does. A blank CELL is not omitted — it coerces to 0 like any blank:
        // SUM(SEQUENCE(2,2,A9)) = 6 (start 0), SUM(SEQUENCE(2,2,1,A9)) = 4 (step 0), and
        // SUM(SEQUENCE(A9)) = #VALUE! (0 rows). TRUE and "3" coerce as numbers: SUM(SEQUENCE(TRUE)) = 1,
        // SUM(SEQUENCE("3")) = 6.
        await Assert
            .That(
                Num(
                    Eval(
                        new Sum([
                            Seq(Number(2), Number(2), BlankValue.Instance, BlankValue.Instance),
                        ]),
                        workbook
                    )
                )
            )
            .IsEqualTo(10.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(2), Number(2), Ref("A9", sheet))]), workbook)))
            .IsEqualTo(6.0);
        await Assert
            .That(
                Num(
                    Eval(
                        new Sum([Seq(Number(2), Number(2), Number(1), Ref("A9", sheet))]),
                        workbook
                    )
                )
            )
            .IsEqualTo(4.0);
        await Assert
            .That(Eval(new Sum([Seq(Ref("A9", sheet))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num(Eval(new Sum([Seq(BooleanValue.True)]), workbook))).IsEqualTo(1.0);
        await Assert.That(Num(Eval(new Sum([Seq(new StringValue("3"))]), workbook))).IsEqualTo(6.0);
    }

    [Test]
    public async Task AZeroNegativeOrNonNumericSize_IsValueError_NeverCalc_AsAOneByOneSingleton()
    {
        var (workbook, _) = Grid();

        // Correction B2: the phase design said #CALC! for rows/columns < 1; the oracle says #VALUE! —
        // SUM(SEQUENCE(0)), SUM(SEQUENCE(-1)), SUM(SEQUENCE("x")), =SEQUENCE(0,0), SUM(SEQUENCE(1,0)),
        // SUM(SEQUENCE(0.5)) (truncates to 0), SUM(SEQUENCE(-0.5)) all #VALUE!; ERROR.TYPE(SEQUENCE(-1))
        // = 3, not 14. #CALC! is a real code now (Error.Calc), so this assertion cannot pass through the
        // unknown-display fold; it is exact.
        await Assert.That(Eval(new Sum([Seq(Number(0))]), workbook)).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([Seq(Number(-1))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([Seq(new StringValue("x"))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Eval(Seq(Number(0), Number(0)), workbook)).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([Seq(Number(1), Number(0))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([Seq(Number(0.5))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Eval(new Sum([Seq(Number(-0.5))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num(Eval(new ErrorType([Seq(Number(-1))]), workbook))).IsEqualTo(3.0);
        await Assert
            .That(Eval(new Sum([Seq(Number(-1))]), workbook))
            .IsNotEqualTo(ErrorValue.Calculation);

        // The error is the producer's OWN answer — a 1x1 array carrying it (ArrayShaping's invariant),
        // not a refused build: the probe and the build stay (true, true) / true in lockstep.
        var operand = Build(Seq(Number(-1)), workbook);
        await Assert.That(operand.IsArray).IsTrue();
        await Assert.That(operand.Rows).IsEqualTo(1);
        await Assert.That(operand.Columns).IsEqualTo(1);
        await Assert.That(operand.At(0, 1, 1).AsObject()).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(((IArrayProducer)Seq(Number(-1))).ProbeArray(new EvaluationContext(workbook)))
            .IsEqualTo((true, true));
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Seq(Number(-1)), new EvaluationContext(workbook)))
            .IsTrue();

        // COUNT discards the error channel by design, so it tallies nothing: COUNT(SEQUENCE(-1)) = 0.
        await Assert.That(Num(Eval(new Count([Seq(Number(-1))]), workbook))).IsEqualTo(0.0);
    }

    [Test]
    public async Task AnErrorInAnArgument_PropagatesUnchanged()
    {
        var (workbook, _) = Grid();

        // SUM(SEQUENCE(1/0)) = #DIV/0!, SUM(SEQUENCE(2,2,1/0)) = #DIV/0!, SUM(SEQUENCE(2,2,"x")) and
        // SUM(SEQUENCE(2,2,1,"x")) = #VALUE! — the coercion error of the offending argument, whichever
        // slot it is in.
        await Assert
            .That(Eval(new Sum([Seq(DivideByZero())]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Eval(new Sum([Seq(Number(2), Number(2), DivideByZero())]), workbook))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Eval(new Sum([Seq(Number(2), Number(2), new StringValue("x"))]), workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(
                Eval(
                    new Sum([Seq(Number(2), Number(2), Number(1), new StringValue("x"))]),
                    workbook
                )
            )
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task BeyondTheGrid_IsNumError_TheOnePinnedDivergence()
    {
        var (workbook, _) = Grid();

        // MySheet's OWN cost guard, not the oracle's rule: the oracle has no cap in a consumed position
        // (ROWS(SEQUENCE(1048577)) = 1048577, COLUMNS(SEQUENCE(1,16385)) = 16385,
        // SUM(SEQUENCE(1048577)) = 549757386753 — Aspose.Cells 26.6.0, 2026-09-10, both modes). The cap
        // is rows > 1,048,576, columns > 16,384, or rows * columns > 1,048,576 -> a 1x1 #NUM!; exactly
        // at the cap is allowed (SUM(SEQUENCE(1048576)) = 549756338176 here AND on the oracle by the
        // same arithmetic, SUM(SEQUENCE(1,16384)) = 134225920).
        await Assert
            .That(Eval(new Sum([Seq(Number(1048577))]), workbook))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Eval(new Sum([Seq(Number(1), Number(16385))]), workbook))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Eval(new Sum([Seq(Number(1024), Number(1025))]), workbook))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(1048576))]), workbook)))
            .IsEqualTo(549756338176.0);
        await Assert
            .That(Num(Eval(new Sum([Seq(Number(1), Number(16384))]), workbook)))
            .IsEqualTo(134225920.0);
    }

    [Test]
    public async Task ComposesUnderAnOperator_ByProjection()
    {
        var (workbook, sheet) = Grid();

        // SUM(SEQUENCE(3)*B1:B3) = 14 (CSE column; plain entry is #VALUE!): 1*1 + 2*2 + 3*3.
        await Assert
            .That(
                Num(
                    Eval(
                        new Sum([
                            new BinaryOperation(
                                BinaryOperator.Multiply,
                                Seq(Number(3)),
                                Ref("B1:B3", sheet)
                            ),
                        ]),
                        workbook
                    )
                )
            )
            .IsEqualTo(14.0);

        // Projection mechanics: a 2x3 operand read at a 3x3 extent covers rows 0-1 and answers #N/A for
        // row 2; a 1-extent axis repeats.
        var operand = new SequenceOperand(2, 3, 7, 1);
        await Assert.That(Num(operand.At(5, 2, 3).AsObject())).IsEqualTo(12.0);
        await Assert.That(operand.At(6, 3, 3).AsObject()).IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Num(new SequenceOperand(1, 3, 0, 1).At(7, 3, 3).AsObject()))
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task TheOperand_RefusesAZeroExtent_AtConstruction()
    {
        // ArrayShaping's invariant, enforced by the operand's own constructor: nothing else guards an
        // operand class that lives outside ArrayShaping.cs, and a 0-extent operand streams nothing —
        // SUM 0, COUNT 0 in silence (ArrayProducerContractTests shows the breakage). Sequence itself
        // never constructs one (a size below 1 is the #VALUE! singleton), so the guard is reached here.
        await Assert
            .That(() => new SequenceOperand(0, 1, 1, 1))
            .Throws<InvalidOperationException>()
            .WithMessageContaining("Rows >= 1 && Columns >= 1");
        await Assert
            .That(() => new SequenceOperand(1, 0, 1, 1))
            .Throws<InvalidOperationException>();
        await Assert
            .That(() => new SequenceOperand(-1, 3, 1, 1))
            .Throws<InvalidOperationException>();
    }
}
