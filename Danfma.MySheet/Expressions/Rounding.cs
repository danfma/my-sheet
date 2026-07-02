using MemoryPack;

namespace Danfma.MySheet.Expressions;

// Arredondamentos por significância e paridade: ROUNDDOWN, TRUNC, MROUND, CEILING (legada, com as
// regras de sinal do Excel), CEILING.MATH, CEILING.PRECISE, ISO.CEILING, FLOOR (legada), FLOOR.MATH,
// FLOOR.PRECISE, EVEN e ODD. Quocientes intermediários passam por ExcelMath.Snap para reproduzir o
// arredondamento cosmético do Excel (ver golden values documentados nos testes).

[MemoryPackable]
public sealed partial record RoundDown(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var digits) is { } digitsError)
        {
            return ComputedValue.Error(digitsError);
        }

        return ComputedValue.Number(ExcelMath.TruncateToDigits(number, digits));
    }
}

[MemoryPackable]
public sealed partial record Trunc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var digits = 0.0;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out digits) is { } digitsError
        )
        {
            return ComputedValue.Error(digitsError);
        }

        return ComputedValue.Number(ExcelMath.TruncateToDigits(number, digits));
    }
}

[MemoryPackable]
public sealed partial record MRound(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var multiple) is { } multipleError)
        {
            return ComputedValue.Error(multipleError);
        }

        // Excel: número e múltiplo com sinais opostos → #NUM!.
        if ((number > 0 && multiple < 0) || (number < 0 && multiple > 0))
        {
            return ComputedValue.Error(Error.Num);
        }

        if (multiple == 0)
        {
            return ComputedValue.Number(0);
        }

        var quotient = Math.Round(
            ExcelMath.Snap(number / multiple),
            MidpointRounding.AwayFromZero
        );

        return ComputedValue.Number(quotient * multiple);
    }
}

[MemoryPackable]
public sealed partial record Ceiling(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var significance) is { } significanceError)
        {
            return ComputedValue.Error(significanceError);
        }

        // Excel (legada): número positivo com significância negativa → #NUM!; significância 0 → 0.
        if (number > 0 && significance < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (significance == 0)
        {
            return ComputedValue.Number(0);
        }

        return ComputedValue.Number(
            Math.Ceiling(ExcelMath.Snap(number / significance)) * significance
        );
    }
}

[MemoryPackable]
public sealed partial record CeilingMath(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var significance = 1.0;

        if (
            Arguments.Length >= 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out significance) is { } significanceError
        )
        {
            return ComputedValue.Error(significanceError);
        }

        var mode = 0.0;

        if (
            Arguments.Length >= 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out mode) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        significance = Math.Abs(significance);

        if (significance == 0)
        {
            return ComputedValue.Number(0);
        }

        var quotient = ExcelMath.Snap(number / significance);

        // O mode só afeta números negativos: mode 0 arredonda em direção a +∞ (rumo ao zero);
        // mode ≠ 0 arredonda para longe do zero.
        var units = number < 0 && mode != 0 ? Math.Floor(quotient) : Math.Ceiling(quotient);

        return ComputedValue.Number(units * significance);
    }
}

[MemoryPackable]
public sealed partial record CeilingPrecise(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PreciseRounding.Evaluate(Arguments, context, roundUp: true);
}

[MemoryPackable]
public sealed partial record IsoCeiling(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PreciseRounding.Evaluate(Arguments, context, roundUp: true);
}

[MemoryPackable]
public sealed partial record Floor(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var significance) is { } significanceError)
        {
            return ComputedValue.Error(significanceError);
        }

        // Excel (legada): significância 0 → #DIV/0!; número positivo com significância negativa → #NUM!.
        if (significance == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        if (number > 0 && significance < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return ComputedValue.Number(
            Math.Floor(ExcelMath.Snap(number / significance)) * significance
        );
    }
}

[MemoryPackable]
public sealed partial record FloorMath(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var significance = 1.0;

        if (
            Arguments.Length >= 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out significance) is { } significanceError
        )
        {
            return ComputedValue.Error(significanceError);
        }

        var mode = 0.0;

        if (
            Arguments.Length >= 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out mode) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        significance = Math.Abs(significance);

        if (significance == 0)
        {
            return ComputedValue.Number(0);
        }

        var quotient = ExcelMath.Snap(number / significance);

        // O mode só afeta números negativos: mode 0 arredonda em direção a -∞ (longe do zero);
        // mode ≠ 0 arredonda rumo ao zero.
        var units = number < 0 && mode != 0 ? Math.Ceiling(quotient) : Math.Floor(quotient);

        return ComputedValue.Number(units * significance);
    }
}

[MemoryPackable]
public sealed partial record FloorPrecise(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PreciseRounding.Evaluate(Arguments, context, roundUp: false);
}

[MemoryPackable]
public sealed partial record Even(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Para longe do zero até o próximo par; par exato não muda.
        var magnitude = Math.Ceiling(Math.Abs(number) / 2) * 2;

        return ComputedValue.Number(number < 0 ? -magnitude : magnitude);
    }
}

[MemoryPackable]
public sealed partial record Odd(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Para longe do zero até o próximo ímpar; ímpar exato não muda (ODD(0)=1).
        var magnitude = Math.Ceiling((Math.Abs(number) - 1) / 2) * 2 + 1;

        return ComputedValue.Number(number < 0 ? -magnitude : magnitude);
    }
}

/// <summary>
/// CEILING.PRECISE / ISO.CEILING / FLOOR.PRECISE: o sinal da significância é ignorado e o
/// arredondamento é sempre em direção a +∞ (ceiling) ou -∞ (floor); significância 0 → 0.
/// </summary>
file static class PreciseRounding
{
    public static ComputedValue Evaluate(Expression[] arguments, EvaluationContext context, bool roundUp)
    {
        if (arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var significance = 1.0;

        if (
            arguments.Length == 2
            && arguments[1].Evaluate(context).CoerceToNumber(out significance) is { } significanceError
        )
        {
            return ComputedValue.Error(significanceError);
        }

        significance = Math.Abs(significance);

        if (significance == 0)
        {
            return ComputedValue.Number(0);
        }

        var quotient = ExcelMath.Snap(number / significance);
        var units = roundUp ? Math.Ceiling(quotient) : Math.Floor(quotient);

        return ComputedValue.Number(units * significance);
    }
}
