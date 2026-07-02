using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// Trigonometria completa: SIN/COS/TAN, os recíprocos COT/SEC/CSC (com o limite documentado de 2^27 e
// #DIV/0! no polo em zero), os inversos ASIN/ACOS/ATAN/ATAN2/ACOT, as hiperbólicas e DEGREES/RADIANS.
// ATAN2 usa a ordem de argumentos do Excel — ATAN2(x, y) — que é a INVERSA do Math.Atan2(y, x) do .NET.

[MemoryPackable]
public sealed partial record Sin(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Sin(number));
    }
}

[MemoryPackable]
public sealed partial record Cos(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Cos(number));
    }
}

[MemoryPackable]
public sealed partial record Tan(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Tan(number));
    }
}

[MemoryPackable]
public sealed partial record Cot(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (Math.Abs(number) >= Trig.MaxArgument)
        {
            return ComputedValue.Error(Error.Num);
        }

        // Docs: COT(0) → #DIV/0!.
        if (number == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(Math.Cos(number) / Math.Sin(number));
    }
}

[MemoryPackable]
public sealed partial record Sec(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (Math.Abs(number) >= Trig.MaxArgument)
        {
            return ComputedValue.Error(Error.Num);
        }

        return ComputedValue.Number(1 / Math.Cos(number));
    }
}

[MemoryPackable]
public sealed partial record Csc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (Math.Abs(number) >= Trig.MaxArgument)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (number == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(1 / Math.Sin(number));
    }
}

[MemoryPackable]
public sealed partial record Asin(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Math.Abs(number) > 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Asin(number));
    }
}

[MemoryPackable]
public sealed partial record Acos(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Math.Abs(number) > 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Acos(number));
    }
}

[MemoryPackable]
public sealed partial record Atan(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Atan(number));
    }
}

[MemoryPackable]
public sealed partial record Atan2(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } xError)
        {
            return ComputedValue.Error(xError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var y) is { } yError)
        {
            return ComputedValue.Error(yError);
        }

        // Docs: ATAN2(0,0) → #DIV/0!.
        if (x == 0 && y == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(Math.Atan2(y, x));
    }
}

[MemoryPackable]
public sealed partial record Acot(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Resultado no intervalo (0, π), como documentado.
        return ComputedValue.Number(Math.PI / 2 - Math.Atan(number));
    }
}

[MemoryPackable]
public sealed partial record Sinh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var result = Math.Sinh(number);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Cosh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var result = Math.Cosh(number);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Tanh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Tanh(number));
    }
}

[MemoryPackable]
public sealed partial record Coth(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (number == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(1 / Math.Tanh(number));
    }
}

[MemoryPackable]
public sealed partial record Sech(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(1 / Math.Cosh(number));
    }
}

[MemoryPackable]
public sealed partial record Csch(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (number == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return ComputedValue.Number(1 / Math.Sinh(number));
    }
}

[MemoryPackable]
public sealed partial record Asinh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Asinh(number));
    }
}

[MemoryPackable]
public sealed partial record Acosh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return number < 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Acosh(number));
    }
}

[MemoryPackable]
public sealed partial record Atanh(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Math.Abs(number) >= 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Atanh(number));
    }
}

[MemoryPackable]
public sealed partial record Acoth(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Docs: |número| deve ser maior que 1.
        return Math.Abs(number) <= 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(Math.Atanh(1 / number));
    }
}

[MemoryPackable]
public sealed partial record Degrees(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(number * (180 / Math.PI));
    }
}

[MemoryPackable]
public sealed partial record Radians(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(number * (Math.PI / 180));
    }
}

file static class Trig
{
    /// <summary>2^27 — limite documentado do argumento das funções trig de 2013 (COT/SEC/CSC):
    /// fora do intervalo → #NUM!.</summary>
    public const double MaxArgument = 134217728d;
}
