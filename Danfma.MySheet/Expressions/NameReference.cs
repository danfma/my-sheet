using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A bare name in a formula (not a cell reference or function call). Resolved at evaluation time in this
/// order: (1) the context's LET bindings (so a LET name shadows a workbook name), (2) the workbook's
/// <see cref="Workbook.DefinedNames"/> (evaluated with a name→name cycle guard), (3) otherwise a
/// <c>#NAME?</c> error.
/// </summary>
[MemoryPackable]
public sealed partial record NameReference(string Name) : Expression
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (context.TryGetName(Name, out var value))
        {
            return value;
        }

        if (context.Workbook.DefinedNames.TryGetValue(Name, out var definition))
        {
            return NamedReferences.EvaluateDefinition(definition, context, Name);
        }

        return ComputedValue.Error(Error.Name);
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference) =>
        NamedReferences.TryResolveReference(this, context, out reference);
}
