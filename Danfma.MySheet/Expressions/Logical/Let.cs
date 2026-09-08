using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record Let(Expression[] Arguments) : Function
{
    // LET(name1, value1, …, calculation) — binds each name to its value (later values may use earlier
    // names), then evaluates the final calculation in that scope.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments.Length < 3 || Arguments.Length % 2 == 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        var scope = context;

        for (var i = 0; i < Arguments.Length - 1; i += 2)
        {
            if (Arguments[i] is not NameReference name)
            {
                return ComputedValue.Error(Error.Value);
            }

            // A range value stays a range (reference value) so `LET(r, A1:C1, MATCH(x, r, 0))` works; see
            // NamedReferences.CaptureValue.
            scope = scope.WithName(
                name.Name,
                NamedReferences.CaptureValue(Arguments[i + 1], scope)
            );
        }

        return Arguments[^1].Evaluate(scope);
    }
}
