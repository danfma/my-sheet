using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record ErrorValue(string ErrorCode) : ValueExpression
{
    public static readonly ErrorValue NotValue = new("#VALUE!");
    public static readonly ErrorValue Name = new("#NAME?");
    public static readonly ErrorValue Reference = new("#REF!");
    public static readonly ErrorValue DivByZero = new("#DIV/0!");
    public static readonly ErrorValue NotAvailable = new("#N/A");

    // `new`: intentionally hides the inherited Expression.Number(double) factory on this type; the
    // error singleton is the natural name for #NUM! and the factory stays reachable via Expression.
    public static new readonly ErrorValue Number = new("#NUM!");

    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(AsError());

    // Compute keeps returning `this` (the exact node): all ErrorValues in the engine are the well-known
    // singletons, so the ComputedValue bridge is lossless, but preserving identity here is zero-risk.
    public override object? Compute(EvaluationContext context) => this;

    /// <summary>Identidade do erro como <see cref="Error"/> (código struct alloc-free).</summary>
    internal Error AsError() => Error.FromDisplay(ErrorCode);
}
