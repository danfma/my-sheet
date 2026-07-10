using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A call to a function that is not a built-in record — resolved at evaluation time against the
/// workbook's custom-function registry. Unknown at runtime → <c>#NAME?</c>. This is the extension
/// point for user/host-defined functions; only the name and arguments are serialized.
/// </summary>
[MemoryPackable]
public sealed partial record FunctionCall(string Name, Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (!context.Workbook.TryGetFunctionEntry(Name, out var entry))
        {
            return ComputedValue.Error(Error.Name);
        }

        // Arity is validated (when declared -- see Workbook.RegisterFunction) BEFORE invoking the delegate,
        // the same way a built-in's ParseException guards a wrong argument count at parse time. A custom
        // function's arity is only known once registered, so this guard lives here instead: an out-of-range
        // call becomes #VALUE! instead of an unhandled IndexOutOfRangeException from inside host code that
        // assumed its declared argument count.
        if (Arguments.Length < entry.MinArgs || Arguments.Length > entry.MaxArgs)
        {
            return ComputedValue.Error(Error.Value);
        }

        return entry.Function(Arguments, context.Workbook);
    }
}
