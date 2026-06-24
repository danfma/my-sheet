using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found]) — exact match (other modes pending).
    public override object? Compute(EvaluationContext context)
    {
        var lookup = Arguments[0].Compute(context);
        var lookupArray = ArgumentFlattening.Expand(Arguments[1], context);
        var returnArray = ArgumentFlattening.Expand(Arguments[2], context);

        var count = Math.Min(lookupArray.Count, returnArray.Count);

        for (var i = 0; i < count; i++)
        {
            if (ValueCoercion.AreEqual(lookupArray[i], lookup))
            {
                return returnArray[i];
            }
        }

        return Arguments.Length >= 4 ? Arguments[3].Compute(context) : ErrorValue.NotAvailable;
    }
}
