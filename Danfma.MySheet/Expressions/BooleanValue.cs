using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record BooleanValue(bool Value) : ValueExpression
{
    // Only two possible instances ever exist in practice; sharing them avoids an allocation per
    // boolean literal (formula parsing and, especially, .xlsx load — every "b"-typed cell hits this).
    // Static fields on a [MemoryPackable] record aren't part of the wire format (MemoryPack only
    // serializes the declared members, i.e. Value), so this is purely an allocation optimization with
    // no effect on serialized bytes.
    public static readonly BooleanValue True = new(true);
    public static readonly BooleanValue False = new(false);

    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Value);
}
