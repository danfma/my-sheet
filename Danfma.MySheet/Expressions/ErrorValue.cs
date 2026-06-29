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

    public override object? Compute(EvaluationContext context) => this;
}
