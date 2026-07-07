namespace Danfma.MySheet.Expressions;

public abstract record Reference : Expression
{
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = this;
        return true;
    }
}
