using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A bare name in a formula (not a cell reference or function call). Resolved at evaluation time in this
/// order: (1) the context's LET bindings (so a LET name shadows a workbook name) — an ARRAY binding reads as
/// its top-left element, the same <c>@</c> rule a bare producer follows (Phase 11c, <see cref="ArrayBindings"/>),
/// a scalar or range binding as the value it holds; (2) the workbook's <see cref="Workbook.DefinedNames"/>
/// (evaluated with a name→name cycle guard); (3) the workbook's <see cref="Workbook.Tables"/> — a bare table
/// name (<c>=Tabela1</c>) answers its data body's reference value (Phase 5 ruling R3); (4) otherwise a
/// <c>#NAME?</c> error. Tables and defined names share one namespace, so (2) and (3) cannot both match.
/// </summary>
[MemoryPackable]
public sealed partial record NameReference(string Name) : Expression
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (context.TryGetArrayBinding(Name, out var operand))
        {
            return ArrayEvaluation.FirstElement(operand);
        }

        if (context.TryGetName(Name, out var value))
        {
            return value;
        }

        if (context.Workbook.DefinedNames.TryGetValue(Name, out var definition))
        {
            return NamedReferences.EvaluateDefinition(definition, context, Name);
        }

        // Phase 5 ruling R3: a bare TABLE name evaluates to its resolved range's reference VALUE — the same
        // answer TableReference.Evaluate gives, arrived at by delegating to it so the #NAME?/#REF! mapping
        // stays in that one primitive. The reference kind is what lets the cell boundary intersect it
        // (=Tabela1[Valor] in the table's own row answers the row's value; a bare 2-D =Tabela1 is #VALUE!)
        // and range-aware consumers expand it (SUM 66, ROWS 3, COUNTA 9 over the oracle fixture).
        if (context.Workbook.Tables.ContainsKey(Name))
        {
            return new TableReference(Name, null, TableArea.Data).Evaluate(context);
        }

        return ComputedValue.Error(Error.Name);
    }

    // An array binding is not a reference: NamedReferences answers false for it BEFORE the defined-name
    // fallback, so a workbook name spelled like the LET name stays shadowed. The consumers that ask (INDEX,
    // ROWS, ISREF, the criteria gate) reach an array binding through ArrayEvaluation's NameReference arms
    // and the context-aware IsBareReferenceNode instead.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference) =>
        NamedReferences.TryResolveReference(this, context, out reference);
}
