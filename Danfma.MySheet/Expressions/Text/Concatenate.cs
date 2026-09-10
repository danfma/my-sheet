using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record Concatenate(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet range is a structural #REF! — an open-range ghost would otherwise be read as empty
        // and concatenated into "".
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var builder = new StringBuilder();

        // streamArrays:false — CONCATENATE joins SCALARS, and over a computed array the oracle answers the
        // array's top-left where CONCAT expands it: CONCATENATE(FILTER(A1:A3,A1:A3>0)) = "5" and
        // CONCATENATE(SEQUENCE(3)) = "1" against CONCAT's "59" and "123" (Aspose.Cells 26.6.0, 2026-09-10,
        // both entry modes; DynamicArrayTests). The default branch's scalar Evaluate is that top-left for a
        // producer (ArrayEvaluation.FirstElement). The RANGE expansion this walk still performs here is
        // pre-existing and unmeasured against that rule.
        foreach (
            var value in ArgumentFlattening.FlattenComputedValues(
                Arguments,
                context,
                streamArrays: false
            )
        )
        {
            if (value.CoerceToText(out var text) is { } error)
            {
                return ComputedValue.Error(error);
            }

            builder.Append(text);
        }

        return ComputedValue.Text(builder.ToString());
    }
}
