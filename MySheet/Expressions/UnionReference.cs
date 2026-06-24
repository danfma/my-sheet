using MemoryPack;

namespace MySheet.Expressions;

/// <summary>
/// The union of several reference areas, written <c>(A1:A3, C1:C3)</c>. Used inside a function that
/// accepts ranges (e.g. <c>SUM((A1:A3, C1:C3))</c>); in a scalar context it is a <c>#VALUE!</c> error.
/// </summary>
[MemoryPackable]
public sealed partial record UnionReference(Expression[] Areas) : Reference
{
    public override object? Compute(EvaluationContext context) => ErrorValue.NotValue;

    public IEnumerable<object?> ExpandValues(EvaluationContext context)
    {
        foreach (var area in Areas)
        {
            switch (area)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandValues(context))
                    {
                        yield return value;
                    }

                    break;

                default:
                    yield return area.Compute(context);
                    break;
            }
        }
    }
}
