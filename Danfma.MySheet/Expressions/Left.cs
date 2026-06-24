using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Left(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(context), out var text) is { } textError)
        {
            return textError;
        }

        var count = 1.0;

        if (
            Arguments.Length == 2
            && ValueCoercion.TryToNumber(Arguments[1].Compute(context), out count) is { } countError
        )
        {
            return countError;
        }

        if (count < 0)
        {
            return ErrorValue.NotValue;
        }

        return text[..Math.Min((int)count, text.Length)];
    }
}
