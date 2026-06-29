using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Match(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var lookup = Arguments[0].Compute(context);
        var array = ArgumentFlattening.Expand(Arguments[1], context);

        var matchType = 1.0;

        if (
            Arguments.Length == 3
            && ValueCoercion.TryToNumber(Arguments[2].Compute(context), out matchType)
                is { } typeError
        )
        {
            return typeError;
        }

        if (matchType == 0)
        {
            for (var i = 0; i < array.Count; i++)
            {
                if (ValueCoercion.AreEqual(array[i], lookup))
                {
                    return (double)(i + 1);
                }
            }

            return ErrorValue.NotAvailable;
        }

        // Approximate: matchType > 0 assumes ascending (largest value <= lookup); < 0 assumes
        // descending (smallest value >= lookup). Cross-type ordering (ValueCoercion.Compare) lets text
        // keys sort lexicographically, exactly like the <= operator — not only numeric keys.
        if (lookup is ErrorValue lookupError)
        {
            return lookupError;
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

        return position >= 1 ? (double)position : ErrorValue.NotAvailable;
    }
}
