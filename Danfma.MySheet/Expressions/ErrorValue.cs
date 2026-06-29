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
    public static readonly ErrorValue Number = new("#NUM!");

    public override object? Compute(EvaluationContext context) => this;
}
