using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// Aritmética escalar e agregações numéricas simples: MOD (sinal do divisor, como o Excel),
// QUOTIENT, SIGN, PI, PRODUCT, SUMSQ, MULTINOMIAL e SERIESSUM.

[MemoryPackable]
public sealed partial record Mod(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var divisor) is { } divisorError)
        {
            return ComputedValue.Error(divisorError);
        }

        if (divisor == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        // MOD(n,d) = n - d*INT(n/d): o resultado tem o sinal do divisor (≠ do operador % do C#).
        return ComputedValue.Number(number - divisor * Math.Floor(number / divisor));
    }
}

[MemoryPackable]
public sealed partial record Quotient(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var numerator) is { } numeratorError)
        {
            return ComputedValue.Error(numeratorError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var denominator) is { } denominatorError)
        {
            return ComputedValue.Error(denominatorError);
        }

        if (denominator == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(Math.Truncate(ExcelMath.Snap(numerator / denominator)));
    }
}

[MemoryPackable]
public sealed partial record Sign(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return double.IsNaN(number)
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Sign(number));
    }
}

[MemoryPackable]
public sealed partial record Pi(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(Math.PI);
}

[MemoryPackable]
public sealed partial record Product(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new ProductFold { Product = 1 };

        if (NumericAggregation.Fold(Arguments, context, ref fold) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Sem nenhum valor numérico o produto é 0 (Excel), não o produto vazio 1.
        return ComputedValue.Number(fold.Any ? fold.Product : 0);
    }

    private struct ProductFold : INumericFold
    {
        public double Product;
        public bool Any;

        public void Accept(double value)
        {
            Product *= value;
            Any = true;
        }
    }
}

[MemoryPackable]
public sealed partial record SumSq(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new SumSqFold();

        return NumericAggregation.Fold(Arguments, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.Total);
    }

    private struct SumSqFold : INumericFold
    {
        public double Total;

        public void Accept(double value) => Total += value * value;
    }
}

[MemoryPackable]
public sealed partial record Multinomial(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new NumericListFold { Values = [] };

        if (NumericAggregation.Fold(Arguments, context, ref fold) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var sum = 0d;
        var denominator = 1d;

        foreach (var value in fold.Values)
        {
            if (value < 0)
            {
                return ComputedValue.Error(Error.Num);
            }

            var truncated = Math.Truncate(value);

            sum += truncated;
            denominator *= ExcelMath.Factorial(truncated);
        }

        var result = ExcelMath.Factorial(sum) / denominator;

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record SeriesSum(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } xError)
        {
            return ComputedValue.Error(xError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var initialPower) is { } powerError)
        {
            return ComputedValue.Error(powerError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var step) is { } stepError)
        {
            return ComputedValue.Error(stepError);
        }

        var sum = 0d;
        var index = 0;

        foreach (var value in ArgumentFlattening.ExpandComputedValues(Arguments[3], context))
        {
            if (value.CoerceToNumber(out var coefficient) is { } coefficientError)
            {
                return ComputedValue.Error(coefficientError);
            }

            sum += coefficient * Math.Pow(x, initialPower + index * step);
            index++;
        }

        return double.IsFinite(sum)
            ? ComputedValue.Number(sum)
            : ComputedValue.Error(Error.Num);
    }
}
