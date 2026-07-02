using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// Combinatória: FACT, FACTDOUBLE, COMBIN, COMBINA, GCD e LCM. Argumentos não-inteiros são truncados
// (regra documentada do Excel); domínio inválido/overflow → #NUM!.

[MemoryPackable]
public sealed partial record Fact(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (number < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = ExcelMath.Factorial(Math.Truncate(number));

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record FactDouble(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (number < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = 1d;

        for (var i = Math.Truncate(number); i >= 2; i -= 2)
        {
            result *= i;
        }

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Combin(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var numberChosen) is { } chosenError)
        {
            return ComputedValue.Error(chosenError);
        }

        var total = Math.Truncate(number);
        var chosen = Math.Truncate(numberChosen);

        if (total < 0 || chosen < 0 || total < chosen)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = ExcelMath.Binomial(total, chosen);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record CombinA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var numberChosen) is { } chosenError)
        {
            return ComputedValue.Error(chosenError);
        }

        var total = Math.Truncate(number);
        var chosen = Math.Truncate(numberChosen);

        if (total < 0 || chosen < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        // COMBINA(n,k) = COMBIN(n+k-1,k); k = 0 vale 1 mesmo com n = 0.
        if (chosen == 0)
        {
            return ComputedValue.Number(1);
        }

        if (total == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = ExcelMath.Binomial(total + chosen - 1, chosen);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Gcd(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new NumericListFold { Values = [] };

        if (NumericAggregation.Fold(Arguments, context, ref fold) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var result = 0L;

        foreach (var value in fold.Values)
        {
            // Excel: negativo → #NUM!; parâmetro >= 2^53 → #NUM!.
            if (value < 0 || value >= ExcelMath.MaxSafeInteger)
            {
                return ComputedValue.Error(Error.Num);
            }

            result = ExcelMath.Gcd(result, (long)Math.Truncate(value));
        }

        return ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record Lcm(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new NumericListFold { Values = [] };

        if (NumericAggregation.Fold(Arguments, context, ref fold) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var hasZero = false;
        var result = 1L;

        foreach (var value in fold.Values)
        {
            if (value < 0 || value >= ExcelMath.MaxSafeInteger)
            {
                return ComputedValue.Error(Error.Num);
            }

            var truncated = (long)Math.Truncate(value);

            if (truncated == 0)
            {
                hasZero = true;
                continue;
            }

            var divisor = ExcelMath.Gcd(result, truncated);
            var candidate = (double)(result / divisor) * truncated;

            // Excel: LCM(a,b) >= 2^53 → #NUM!.
            if (candidate >= ExcelMath.MaxSafeInteger)
            {
                return ComputedValue.Error(Error.Num);
            }

            result = result / divisor * truncated;
        }

        return ComputedValue.Number(hasZero ? 0 : result);
    }
}
