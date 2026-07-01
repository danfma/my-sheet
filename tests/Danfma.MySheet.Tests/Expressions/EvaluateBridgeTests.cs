using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class EvaluateBridgeTests
{
    private static Expression[] ValueNodes() =>
    [
        Number(3),
        String("hi"),
        new BooleanValue(true),
        BlankValue.Instance,
        ErrorValue.NotValue,
    ];

    [Test]
    public async Task Evaluate_AsObject_MatchesLegacyCompute()
    {
        var workbook = new Workbook();

        foreach (var node in ValueNodes())
        {
            await Assert.That(node.Evaluate(workbook).AsObject()).IsEqualTo(node.Compute(workbook));
        }
    }

    [Test]
    public async Task Evaluate_ProducesTypedKind()
    {
        var workbook = new Workbook();

        await Assert.That(Number(3).Evaluate(workbook).Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(String("x").Evaluate(workbook).Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(new BooleanValue(true).Evaluate(workbook).Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BlankValue.Instance.Evaluate(workbook).Kind).IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(ErrorValue.NotValue.Evaluate(workbook).Kind).IsEqualTo(ComputedValueKind.Error);
    }

    [Test]
    public async Task Evaluate_LogicNodes_Native()
    {
        var workbook = new Workbook();

        // IF(TRUE, 1, <#DIV/0!>) → 1 (o ramo falso NÃO é avaliado — curto-circuito nativo).
        var ifNode = new If([new BooleanValue(true), Number(1), Divide(Number(1), Number(0))]);
        var taken = ifNode.Evaluate(workbook);
        await Assert.That(taken.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(taken.ToDouble()).IsEqualTo(1.0);

        // IFERROR(#VALUE!, 42) → 42.
        await Assert.That(new IfError([ErrorValue.NotValue, Number(42)]).Evaluate(workbook).ToDouble()).IsEqualTo(42.0);

        // IFNA só pega #N/A; outros erros passam.
        await Assert.That(new IfNa([ErrorValue.NotAvailable, Number(7)]).Evaluate(workbook).ToDouble()).IsEqualTo(7.0);
        await Assert.That(new IfNa([ErrorValue.NotValue, Number(7)]).Evaluate(workbook).TryGetError(out var passed)).IsTrue();
        await Assert.That(passed).IsEqualTo(Error.Value);

        // AND(TRUE, FALSE) → false; NOT(FALSE) → true.
        await Assert.That(new And([new BooleanValue(true), new BooleanValue(false)]).Evaluate(workbook).TryGetBoolean(out var and)).IsTrue();
        await Assert.That(and).IsFalse();
        await Assert.That(new Not([new BooleanValue(false)]).Evaluate(workbook).TryGetBoolean(out var not)).IsTrue();
        await Assert.That(not).IsTrue();
    }

    [Test]
    public async Task Evaluate_Operators_Native()
    {
        var workbook = new Workbook();

        // Aritmética.
        await Assert.That(Add(Number(2), Number(3)).Evaluate(workbook).ToDouble()).IsEqualTo(5.0);
        await Assert.That(Divide(Number(1), Number(0)).Evaluate(workbook).TryGetError(out var div)).IsTrue();
        await Assert.That(div).IsEqualTo(Error.DivZero);

        // Comparação cross-type (número < texto) e igualdade case-insensitive.
        await Assert.That(GreaterThan(Number(3), Number(2)).Evaluate(workbook).TryGetBoolean(out var gt)).IsTrue();
        await Assert.That(gt).IsTrue();

        // Concatenação (&) e negação unária.
        await Assert
            .That(new BinaryOperation(BinaryOperator.Concat, String("a"), Number(1)).Evaluate(workbook).ToText())
            .IsEqualTo("a1");
        await Assert.That(Negate(Number(4)).Evaluate(workbook).ToDouble()).IsEqualTo(-4.0);

        // Propagação de erro através do operador.
        await Assert.That(Add(Divide(Number(1), Number(0)), Number(1)).Evaluate(workbook).TryGetError(out _)).IsTrue();
    }

    [Test]
    public async Task Evaluate_UnmigratedNode_BridgesFromCompute()
    {
        var workbook = new Workbook();

        // Sum ainda não foi migrado → usa o Evaluate default da base (From(Compute())).
        var result = Sum(Number(1), Number(2)).Evaluate(workbook);

        await Assert.That(result.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(result.ToDouble()).IsEqualTo(3.0);
    }

    [Test]
    public async Task NativeCoercion_CoerceToNumber()
    {
        // Forma fluente (extension method): value.CoerceToNumber(out ...).
        await Assert.That(ComputedValue.Number(2).CoerceToNumber(out var number)).IsNull();
        await Assert.That(number).IsEqualTo(2.0);

        await Assert.That(ComputedValue.Boolean(true).CoerceToNumber(out var boolean)).IsNull();
        await Assert.That(boolean).IsEqualTo(1.0);

        await Assert.That(ComputedValue.Text("3.5").CoerceToNumber(out var text)).IsNull();
        await Assert.That(text).IsEqualTo(3.5);

        await Assert.That(ComputedValue.Blank.CoerceToNumber(out var blank)).IsNull();
        await Assert.That(blank).IsEqualTo(0.0);

        // Texto não-numérico → #VALUE!; erro propaga.
        await Assert.That(ComputedValue.Text("abc").CoerceToNumber(out _)!.Value).IsEqualTo(Error.Value);
        await Assert
            .That(ComputedValue.Error(Error.DivZero).CoerceToNumber(out _)!.Value)
            .IsEqualTo(Error.DivZero);
    }

    [Test]
    public async Task NativeCoercion_CoerceToTextAndBool()
    {
        await Assert.That(ComputedValue.Number(12.5).CoerceToText(out var text)).IsNull();
        await Assert.That(text).IsEqualTo("12.5");

        await Assert.That(ComputedValue.Boolean(true).CoerceToText(out var boolText)).IsNull();
        await Assert.That(boolText).IsEqualTo("TRUE");

        await Assert.That(ComputedValue.Number(1).CoerceToBool(out var truthy)).IsNull();
        await Assert.That(truthy).IsTrue();

        await Assert.That(ComputedValue.Blank.CoerceToBool(out var falsy)).IsNull();
        await Assert.That(falsy).IsFalse();

        // Texto num contexto booleano → #VALUE!.
        await Assert.That(ComputedValue.Text("x").CoerceToBool(out _)!.Value).IsEqualTo(Error.Value);
    }
}
