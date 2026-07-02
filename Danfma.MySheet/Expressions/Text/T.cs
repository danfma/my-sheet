using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

// Nasceu na onda 2 junto com as funções de informação, mas a categoria publicada (function-reference,
// espelhando o catálogo do Excel) é Text — por isso vive aqui.
[MemoryPackable]
public sealed partial record T(Expression[] Arguments) : Function
{
    // Text -> itself; every other value -> "" (empty text). Errors propagate.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = Arguments[0].Evaluate(context);

        return value.Kind switch
        {
            ComputedValueKind.Text or ComputedValueKind.Error => value,
            _ => ComputedValue.Text(string.Empty),
        };
    }
}
