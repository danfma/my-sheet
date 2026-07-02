using MemoryPack;

namespace Danfma.MySheet.Expressions;

// Potência, raiz e logaritmos: SQRT, POWER, EXP, LN, LOG, LOG10, SQRTPI. Domínio inválido → #NUM!
// (Excel); overflow numérico (> ~1E+308) → #NUM!.

[MemoryPackable]
public sealed partial record Sqrt(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return number < 0
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Sqrt(number));
    }
}

[MemoryPackable]
public sealed partial record Power(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var exponent) is { } exponentError)
        {
            return ComputedValue.Error(exponentError);
        }

        // Excel: 0^0 → #NUM!; 0^negativo → #DIV/0!; base negativa com expoente fracionário → #NUM!.
        if (number == 0 && exponent == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (number == 0 && exponent < 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var result = Math.Pow(number, exponent);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Exp(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var result = Math.Exp(number);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Ln(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return number <= 0
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Log(number));
    }
}

[MemoryPackable]
public sealed partial record Log(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var newBase = 10.0;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out newBase) is { } baseError
        )
        {
            return ComputedValue.Error(baseError);
        }

        // Excel: número ou base <= 0 → #NUM!; base 1 → #DIV/0! (ln(1) = 0 no denominador).
        if (number <= 0 || newBase <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (newBase == 1)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(Math.Log(number) / Math.Log(newBase));
    }
}

[MemoryPackable]
public sealed partial record Log10(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return number <= 0
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Log10(number));
    }
}

[MemoryPackable]
public sealed partial record SqrtPi(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return number < 0
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Sqrt(number * Math.PI));
    }
}
