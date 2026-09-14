using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record ErrorValue(string ErrorCode) : ValueExpression
{
    // Item 43 round 2 (M-2): the one classic error that had no singleton until now — see Error.cs's
    // ToErrorValue, which used to fall through to the `new ErrorValue(Display)` default for code 0.
    public static readonly ErrorValue Null = new("#NULL!");

    public static readonly ErrorValue NotValue = new("#VALUE!");
    public static readonly ErrorValue Name = new("#NAME?");
    public static readonly ErrorValue Reference = new("#REF!");
    public static readonly ErrorValue DivByZero = new("#DIV/0!");
    public static readonly ErrorValue NotAvailable = new("#N/A");
    public static readonly ErrorValue Calculation = new("#CALC!");

    // `new`: intentionally hides the inherited Expression.Number(double) factory on this type; the
    // error singleton is the natural name for #NUM! and the factory stays reachable via Expression.
    public static new readonly ErrorValue Number = new("#NUM!");

    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Error(AsError());

    /// <summary>Identidade do erro como <see cref="Error"/> (código struct alloc-free).</summary>
    internal Error AsError() => Error.FromDisplay(ErrorCode);
}
