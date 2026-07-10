using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

// Onda 2 — lógicas: TRUE, FALSE (formas função dos literais), XOR (paridade de TRUEs), IFS e SWITCH
// (ambas lazy: só o ramo que casa é avaliado, como o IF).

[MemoryPackable]
public sealed partial record TrueFunction(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(true);
}

[MemoryPackable]
public sealed partial record FalseFunction(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(false);
}

[MemoryPackable]
public sealed partial record Xor(Expression[] Arguments) : Function
{
    // Excel: TRUE when the number of TRUE inputs is odd. Text/blank cells reached through a reference
    // (a single cell as well as a range) are ignored; a call whose arguments contribute no logical value
    // at all -> #VALUE! (per the docs). Shares the reduction with AND/OR via LogicalReduction.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference argument to a missing sheet is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (
            LogicalReduction.Reduce(Arguments, context, out var trueCount, out var total) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        return total == 0
            ? ComputedValue.Error(Error.Value)
            : ComputedValue.Boolean(trueCount % 2 == 1);
    }
}

[MemoryPackable]
public sealed partial record Ifs(Expression[] Arguments) : Function
{
    // IFS(test1, value1, [test2, value2], …) — lazy like IF: conditions run in order and only the
    // value of the pair that matched is computed. No TRUE condition -> #N/A (per the Excel docs).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments.Length % 2 != 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        for (var i = 0; i + 1 < Arguments.Length; i += 2)
        {
            if (Arguments[i].Evaluate(context).CoerceToBool(out var condition) is { } error)
            {
                return ComputedValue.Error(error);
            }

            if (condition)
            {
                return Arguments[i + 1].Evaluate(context);
            }
        }

        return ComputedValue.Error(Error.NA);
    }
}

[MemoryPackable]
public sealed partial record Switch(Expression[] Arguments) : Function
{
    // SWITCH(expression, value1, result1, …, [default]) — the default is the odd trailing argument.
    // Lazy: candidates are compared in order (via the '=' equality semantics) and only the matched
    // result — or the default — is computed. No match and no default -> #N/A (per the Excel docs).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var expression = Arguments[0].Evaluate(context);

        if (expression.TryGetError(out var expressionError))
        {
            return ComputedValue.Error(expressionError);
        }

        var i = 1;

        for (; i + 1 < Arguments.Length; i += 2)
        {
            var candidate = Arguments[i].Evaluate(context);

            if (candidate.TryGetError(out var candidateError))
            {
                return ComputedValue.Error(candidateError);
            }

            if (ValueCoercion.AreEqual(expression, candidate))
            {
                return Arguments[i + 1].Evaluate(context);
            }
        }

        return i < Arguments.Length
            ? Arguments[i].Evaluate(context)
            : ComputedValue.Error(Error.NA);
    }
}
