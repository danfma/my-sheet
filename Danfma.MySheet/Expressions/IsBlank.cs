using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record IsBlank(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context) =>
        Arguments[0].Compute(context) is null;
}
