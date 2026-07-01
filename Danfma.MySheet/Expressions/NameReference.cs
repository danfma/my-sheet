using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A bare name in a formula (not a cell reference or function call). Resolved at evaluation time against
/// the context's LET bindings; an unbound name is a <c>#NAME?</c> error.
/// </summary>
[MemoryPackable]
public sealed partial record NameReference(string Name) : Expression
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        context.TryGetName(Name, out var value)
            ? ComputedValue.From(value)
            : ComputedValue.Error(Error.Name);

}
