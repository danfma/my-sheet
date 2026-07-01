using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Left(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        var count = 1.0;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out count) is { } countError
        )
        {
            return ComputedValue.Error(countError);
        }

        if (count < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Text(text[..Math.Min((int)count, text.Length)]);
    }
}
