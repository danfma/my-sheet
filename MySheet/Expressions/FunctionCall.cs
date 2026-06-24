using MemoryPack;

namespace MySheet.Expressions;

/// <summary>
/// A call to a function that is not a built-in record — resolved at evaluation time against the
/// workbook's custom-function registry. Unknown at runtime → <c>#NAME?</c>. This is the extension
/// point for user/host-defined functions; only the name and arguments are serialized.
/// </summary>
[MemoryPackable]
public sealed partial record FunctionCall(string Name, Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context) =>
        context.Workbook.TryGetFunction(Name, out var function)
            ? function(Arguments, context.Workbook)
            : ErrorValue.Name;
}
