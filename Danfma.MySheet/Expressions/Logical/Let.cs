using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record Let(Expression[] Arguments) : Function
{
    // LET(name1, value1, …, calculation) — binds each name to its value (later values may use earlier
    // names), then evaluates the final calculation in that scope. A binding is what ArrayBindings.Capture
    // makes of it: a range stays a reference value (so `LET(r, A1:C1, MATCH(x, r, 0))` works — see
    // NamedReferences.CaptureValue), a scalar is its scalar, and an array-eligible expression is built ONCE
    // and bound as the operand, so every read of the name — SUM(f), ROWS(f), f*2, a bare f (its top-left) —
    // sees the same array (Phase 11c).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryBind(context, ArrayBindings.Capture, out var scope)
            ? Arguments[^1].Evaluate(scope)
            : ComputedValue.Error(Error.Value);

    /// <summary>
    /// Walks the bindings in order, each captured by <paramref name="capture"/> in the scope the earlier ones
    /// made (so <c>LET(a,FILTER(…),b,SORT(a),…)</c> finds <c>a</c>), and hands back the scope the body
    /// evaluates in; <c>false</c> for a malformed <c>LET</c> — fewer than three arguments, an even count, a
    /// non-name in a name slot — which is the scalar path's <c>#VALUE!</c>. The walk is shared by
    /// <see cref="Evaluate"/> (<see cref="ArrayBindings.Capture"/>, the single evaluation) and by the
    /// mini-CSE's <c>Let</c> arm (<see cref="ArrayBindings.Shape"/> for the probe, <see cref="ArrayBindings.Capture"/>
    /// for the build), so the three cannot bind differently.
    /// </summary>
    internal bool TryBind(
        EvaluationContext context,
        Func<Expression, EvaluationContext, ArrayBindings.Binding> capture,
        out EvaluationContext scope
    )
    {
        scope = context;

        if (Arguments.Length < 3 || Arguments.Length % 2 == 0)
        {
            return false;
        }

        for (var i = 0; i < Arguments.Length - 1; i += 2)
        {
            if (Arguments[i] is not NameReference name)
            {
                return false;
            }

            scope = scope.WithName(name.Name, capture(Arguments[i + 1], scope));
        }

        return true;
    }
}
