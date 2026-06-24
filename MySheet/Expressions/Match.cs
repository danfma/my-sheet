using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Match(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var lookup = Arguments[0].Compute(workbook);
        var array = ArgumentFlattening.Expand(Arguments[1], workbook);

        var matchType = 1.0;

        if (Arguments.Length == 3 &&
            ValueCoercion.TryToNumber(Arguments[2].Compute(workbook), out matchType) is { } typeError)
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
        // descending (smallest value >= lookup).
        if (ValueCoercion.TryToNumber(lookup, out var target) is { } lookupError)
        {
            return lookupError;
        }

        var position = -1;

        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not double value)
            {
                continue;
            }

            if (matchType > 0 && value <= target)
            {
                position = i + 1;
            }
            else if (matchType < 0 && value >= target)
            {
                position = i + 1;
            }
        }

        return position >= 1 ? (double)position : ErrorValue.NotAvailable;
    }
}
