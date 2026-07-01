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
