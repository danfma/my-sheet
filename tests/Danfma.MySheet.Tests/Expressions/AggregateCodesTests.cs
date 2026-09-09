using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Direct contract of the shared aggregate core extracted from SUBTOTAL — the accumulator, the 1-13 fold
/// and the 14-19 positional map. SUBTOTAL only ever reaches codes 1-11 with <c>ignoreErrors: false</c>, so
/// the widened halves (12, 13, the unreachable <c>default</c> guard, 14-19 and every
/// <c>ignoreErrors: true</c> rule) have no end-to-end caller until AGGREGATE lands: they are pinned here,
/// on plain populations and hand-built values, exactly as StatisticsMathFoldTests pins the folds.
/// </summary>
public class AggregateCodesTests
{
    private static double Number(ComputedValue value) =>
        value.TryGetNumber(out var number) ? number : double.NaN;

    // Turns "the call was expected to fail" into a readable failure instead of a silent NaN.
    private static Error Failure(ComputedValue value) =>
        value.TryGetError(out var error)
            ? error
            : throw new InvalidOperationException($"expected an error, but got {value.Kind}");

    private static Error Failure(Error? error) =>
        error ?? throw new InvalidOperationException("expected an error, but the call succeeded");

    // --- Fold: the 1-13 code map ---

    [Test]
    public async Task Fold_Code9_IsTheSum()
    {
        await Assert.That(Number(AggregateCodes.Fold(9, [1.0, 2.0, 3.0]))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Fold_Code11_IsPopulationVariance()
    {
        // O `default:` do Subtotal.Aggregate era 11 (VAR.P) e engolia qualquer código maior; virou um
        // `case 11:` explícito, então o comportamento de 11 continua pinado aqui.
        await Assert.That(Number(AggregateCodes.Fold(11, [1.0, 2.0, 3.0, 4.0]))).IsEqualTo(1.25);
    }

    [Test]
    public async Task Fold_Code12_IsTheMedianOfTheSortedPopulation()
    {
        // População DESORDENADA de propósito: a mediana só sai 2.5 se o Fold ordenar antes de dobrar.
        // (VAR.P dessa mesma população — o que o `default:` antigo devolveria — é 1747.25.)
        await Assert.That(Number(AggregateCodes.Fold(12, [100.0, 1.0, 3.0, 2.0]))).IsEqualTo(2.5);
    }

    [Test]
    public async Task Fold_Code13_IsTheModeInScanOrder()
    {
        // MODE.SNGL desempata pelo PRIMEIRO valor que atinge a contagem vencedora, em ordem de varredura:
        // {3,3,2,2} → 3. Se o Fold ordenasse (como faz para 12), a resposta viraria 2.
        await Assert.That(Number(AggregateCodes.Fold(13, [3.0, 3.0, 2.0, 2.0]))).IsEqualTo(3.0);
    }

    [Test]
    public async Task Fold_Code13_NoRepeatedValue_IsNAError()
    {
        await Assert.That(Failure(AggregateCodes.Fold(13, [1.0, 2.0, 3.0]))).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task Fold_CodeOutsideTheMap_IsValueError()
    {
        // Guarda inalcançável (Subtotal filtra 1-11, AGGREGATE filtra 1-13 antes de chamar): o antigo
        // `default:` calcularia VAR.P silenciosamente para 14 ou 20.
        await Assert.That(Failure(AggregateCodes.Fold(14, [1.0, 2.0, 3.0]))).IsEqualTo(Error.Value);
        await Assert.That(Failure(AggregateCodes.Fold(20, [1.0, 2.0, 3.0]))).IsEqualTo(Error.Value);
    }

    // --- Positional: the 14-19 code map (k is coerced AFTER the population is gathered) ---

    private static ComputedValue Positional(int code, double[] sorted, Expression kArgument)
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Sheet1");

        return AggregateCodes.Positional(code, sorted, kArgument, new EvaluationContext(workbook));
    }

    [Test]
    public async Task Positional_Code14_IsTheKthLargest()
    {
        // LARGE({1,2,3,10}, 2) = 3 (≠ SMALL, que daria 2).
        await Assert
            .That(Number(Positional(14, [1.0, 2.0, 3.0, 10.0], new NumberValue(2))))
            .IsEqualTo(3.0);
    }

    [Test]
    public async Task Positional_Code15_IsTheKthSmallest()
    {
        await Assert
            .That(Number(Positional(15, [1.0, 2.0, 3.0, 10.0], new NumberValue(2))))
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Positional_Code16_IsPercentileInclusive()
    {
        // PERCENTILE.INC({1,2,3,4,10}, 0.25): posição 0.25·4 = 1 → sorted[1] = 2 (a EXC dá 1.5).
        await Assert
            .That(Number(Positional(16, [1.0, 2.0, 3.0, 4.0, 10.0], new NumberValue(0.25))))
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Positional_Code17_IsQuartileInclusive()
    {
        // QUARTILE.INC({1,2,3,4}, 1) = PERCENTILE.INC em 0.25 → 1 + 0.75·(2−1) = 1.75 (a EXC dá 1.25).
        await Assert
            .That(Number(Positional(17, [1.0, 2.0, 3.0, 4.0], new NumberValue(1))))
            .IsEqualTo(1.75);
    }

    [Test]
    public async Task Positional_Code18_IsPercentileExclusive()
    {
        // PERCENTILE.EXC({1,2,3,4,10}, 0.25): rank 0.25·6 = 1.5 → 1 + 0.5·(2−1) = 1.5.
        await Assert
            .That(Number(Positional(18, [1.0, 2.0, 3.0, 4.0, 10.0], new NumberValue(0.25))))
            .IsEqualTo(1.5);
    }

    [Test]
    public async Task Positional_Code19_IsQuartileExclusive()
    {
        // QUARTILE.EXC({1,2,3,4}, 1) = PERCENTILE.EXC em 0.25 → rank 1.25 → 1 + 0.25·(2−1) = 1.25.
        await Assert
            .That(Number(Positional(19, [1.0, 2.0, 3.0, 4.0], new NumberValue(1))))
            .IsEqualTo(1.25);
    }

    [Test]
    public async Task Positional_CodeOutsideTheMap_IsValueError()
    {
        // Guarda inalcançável, espelhando a do Fold: AGGREGATE só encaminha 14-19 para cá. Um `default:`
        // que fosse o próprio 19 (QUARTILE.EXC) devolveria 1.25 silenciosamente para o código 20.
        await Assert
            .That(Failure(Positional(20, [1.0, 2.0, 3.0, 4.0], new NumberValue(1))))
            .IsEqualTo(Error.Value);
    }

    [Test]
    public async Task Positional_KOutOfRange_IsNumError()
    {
        await Assert
            .That(Failure(Positional(14, [1.0, 2.0, 3.0, 10.0], new NumberValue(5))))
            .IsEqualTo(Error.Num);
    }

    [Test]
    public async Task Positional_KThatCannotBeCoerced_IsTheCoercionError()
    {
        await Assert
            .That(Failure(Positional(15, [1.0, 2.0, 3.0], new StringValue("abc"))))
            .IsEqualTo(Error.Value);
    }

    // --- Accumulator: the two ignoreErrors rules (AGGREGATE's option bit 1) ---

    private static ComputedValue Accumulate(
        int code,
        bool ignoreErrors,
        params ComputedValue[] values
    )
    {
        var accumulator = new AggregateCodes.Accumulator(code, ignoreErrors);

        foreach (var value in values)
        {
            if (accumulator.Add(value) is { } error)
            {
                return ComputedValue.Error(error);
            }
        }

        return accumulator.Finish();
    }

    private static readonly ComputedValue[] OneErrorAndTwoNumbers =
    [
        ComputedValue.Number(1),
        ComputedValue.Error(Error.DivZero),
        ComputedValue.Number(4),
    ];

    [Test]
    public async Task Accumulator_CountA_CountsErrorCells_WhenErrorsAreNotIgnored()
    {
        // Medido no SUBTOTAL de hoje: =SUBTOTAL(3,A1:A3) com uma célula de erro → 3. Continua assim.
        await Assert
            .That(Number(Accumulate(3, ignoreErrors: false, OneErrorAndTwoNumbers)))
            .IsEqualTo(3.0);
    }

    [Test]
    public async Task Accumulator_CountA_ExcludesErrorCells_WhenErrorsAreIgnored()
    {
        // A única diferença comportamental que a opção "ignorar erros" introduz no COUNTA.
        await Assert
            .That(Number(Accumulate(3, ignoreErrors: true, OneErrorAndTwoNumbers)))
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Accumulator_Count_IgnoresErrorCellsEitherWay()
    {
        // COUNT só conta Number, então já ignorava erros: os dois modos dão 2.
        await Assert
            .That(Number(Accumulate(2, ignoreErrors: false, OneErrorAndTwoNumbers)))
            .IsEqualTo(2.0);
        await Assert
            .That(Number(Accumulate(2, ignoreErrors: true, OneErrorAndTwoNumbers)))
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Accumulator_Numeric_PropagatesTheFirstError_WhenErrorsAreNotIgnored()
    {
        await Assert
            .That(Failure(Accumulate(9, ignoreErrors: false, OneErrorAndTwoNumbers)))
            .IsEqualTo(Error.DivZero);
    }

    [Test]
    public async Task Accumulator_Numeric_SwallowsTheError_WhenErrorsAreIgnored()
    {
        // O acumulador já excluía a célula de erro da população numérica; "ignorar erros" é só a
        // supressão do canal de propagação → SUM = 1 + 4.
        await Assert
            .That(Number(Accumulate(9, ignoreErrors: true, OneErrorAndTwoNumbers)))
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Accumulator_Numbers_ExposesTheNumericPopulation()
    {
        // A forma-array de AGGREGATE (14-19) precisa da população crua, não do Finish().
        var accumulator = new AggregateCodes.Accumulator(9, ignoreErrors: true);

        foreach (var value in OneErrorAndTwoNumbers)
        {
            accumulator.Add(value);
        }

        await Assert.That(accumulator.Numbers).IsEquivalentTo(new List<double> { 1.0, 4.0 });
    }

    // --- Gather: the nested-skip predicate ---

    // Runs the per-cell scan over A1:A3, where A3 holds <paramref name="nested"/> — a nested aggregate cell
    // worth 3, so the population is 1 + 2 (+ 3 when the cell is NOT skipped). A `ref Accumulator` cannot be
    // captured by a lambda, so the whole call lives in this method.
    private static double GatherOver(
        AggregateCodes.NestedSkip skip,
        string nested = "=SUBTOTAL(9,A1:A2)"
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = ExpressionParser.Parse(nested, sheet);

        var accumulator = new AggregateCodes.Accumulator(9, ignoreErrors: false);

        AggregateCodes.Gather(
            ExpressionParser.Parse("=A1:A3", sheet),
            new EvaluationContext(workbook),
            ref accumulator,
            skip
        );

        return Number(accumulator.Finish());
    }

    [Test]
    public async Task Gather_None_KeepsEveryCell()
    {
        // Sem exclusão: 1 + 2 + SUBTOTAL(9,A1:A2) = 6.
        await Assert.That(GatherOver(AggregateCodes.NestedSkip.None)).IsEqualTo(6.0);
    }

    [Test]
    public async Task Gather_Subtotal_DropsTheNestedSubtotalCell()
    {
        await Assert.That(GatherOver(AggregateCodes.NestedSkip.Subtotal)).IsEqualTo(3.0);
    }

    [Test]
    public async Task Gather_SubtotalAndAggregate_DropsBothNestedNodeKinds()
    {
        // O arm largo (options 0-3 do AGGREGATE) pula as DUAS espécies de célula aninhada.
        await Assert
            .That(GatherOver(AggregateCodes.NestedSkip.SubtotalAndAggregate))
            .IsEqualTo(3.0);
        await Assert
            .That(
                GatherOver(AggregateCodes.NestedSkip.SubtotalAndAggregate, "=AGGREGATE(9,0,A1:A2)")
            )
            .IsEqualTo(3.0);

        // …e os outros dois arms continuam ESTREITOS sobre a mesma célula AGGREGATE — é isso que prova que
        // o predicado largo não é só o estreito com outro nome (senão o par acima passaria de graça).
        await Assert
            .That(GatherOver(AggregateCodes.NestedSkip.Subtotal, "=AGGREGATE(9,0,A1:A2)"))
            .IsEqualTo(6.0);
        await Assert
            .That(GatherOver(AggregateCodes.NestedSkip.None, "=AGGREGATE(9,0,A1:A2)"))
            .IsEqualTo(6.0);
    }

    // --- CollectStream: the mini-CSE feed into the same accumulator ---

    private static ArrayEvaluation.ArrayStream StreamOf(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);

        var expression = ExpressionParser.Parse(formula, sheet);

        if (
            !ArrayEvaluation.TryEvaluateStream(
                expression,
                new EvaluationContext(workbook),
                out var stream
            )
        )
        {
            throw new InvalidOperationException($"{formula} is not array-eligible");
        }

        return stream;
    }

    [Test]
    public async Task CollectStream_FeedsEveryElementIntoTheAccumulator()
    {
        // (A1:A3<>0)*1 = {1;0;1} → SUM = 2, COUNT = 3.
        var sum = new AggregateCodes.Accumulator(9, ignoreErrors: false);
        var count = new AggregateCodes.Accumulator(2, ignoreErrors: false);

        await Assert
            .That(AggregateCodes.CollectStream(StreamOf("=(A1:A3<>0)*1"), ref sum))
            .IsNull();
        await Assert.That(Number(sum.Finish())).IsEqualTo(2.0);

        await Assert
            .That(AggregateCodes.CollectStream(StreamOf("=(A1:A3<>0)*1"), ref count))
            .IsNull();
        await Assert.That(Number(count.Finish())).IsEqualTo(3.0);
    }

    [Test]
    public async Task CollectStream_HonoursTheAccumulatorsErrorRules()
    {
        // 1/A1:A3 = {0.2; #DIV/0!; 1/9} — o elemento do meio é um erro de verdade.
        var propagating = new AggregateCodes.Accumulator(9, ignoreErrors: false);
        var ignoring = new AggregateCodes.Accumulator(9, ignoreErrors: true);

        await Assert
            .That(Failure(AggregateCodes.CollectStream(StreamOf("=1/A1:A3"), ref propagating)))
            .IsEqualTo(Error.DivZero);

        await Assert
            .That(AggregateCodes.CollectStream(StreamOf("=1/A1:A3"), ref ignoring))
            .IsNull();
        await Assert.That(Number(ignoring.Finish())).IsEqualTo(0.2 + (1.0 / 9.0));
    }
}
