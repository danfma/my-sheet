using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Match(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var lookup = Arguments[0].Compute(context);
        var array = ArgumentFlattening.Expand(Arguments[1], context);

        var matchType = 1.0;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out matchType) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        if (matchType == 0)
        {
            for (var i = 0; i < array.Count; i++)
            {
                if (ValueCoercion.AreEqual(array[i], lookup))
                {
                    return ComputedValue.Number(i + 1);
                }
            }

            return ComputedValue.Error(Error.NA);
        }

        // Approximate: matchType > 0 assumes ascending (largest value <= lookup); < 0 assumes
        // descending (smallest value >= lookup). Cross-type ordering (ValueCoercion.Compare) lets text
        // keys sort lexicographically, exactly like the <= operator — not only numeric keys.
        if (lookup is ErrorValue lookupError)
        {
            return ComputedValue.From(lookupError);
        }

        var position = -1;

        for (var i = 0; i < array.Count; i++)
        {
            var value = array[i];
            if (value is null or ErrorValue)
            {
                continue;
            }

            var comparison = ValueCoercion.Compare(value, lookup);

            if (matchType > 0 && comparison <= 0)
            {
                position = i + 1;
            }
            else if (matchType < 0 && comparison >= 0)
            {
                position = i + 1;
            }
        }

        return position >= 1 ? ComputedValue.Number(position) : ComputedValue.Error(Error.NA);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
