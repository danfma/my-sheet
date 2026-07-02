using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record Mid(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var start) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var count) is { } countError)
        {
            return ComputedValue.Error(countError);
        }

        if (start < 1 || count < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (start > text.Length)
        {
            return ComputedValue.Text(string.Empty);
        }

        var available = text.Length - ((int)start - 1);

        return ComputedValue.Text(text.Substring((int)start - 1, Math.Min((int)count, available)));
    }
}
