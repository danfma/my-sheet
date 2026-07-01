using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class ComputedValueTests
{
    [Test]
    public async Task Factories_ProduceExpectedKind()
    {
        await Assert.That(ComputedValue.Blank.Kind).IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(ComputedValue.Number(1).Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(ComputedValue.Boolean(true).Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(ComputedValue.Text("x").Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(ComputedValue.Error(Error.Value).Kind).IsEqualTo(ComputedValueKind.Error);
    }

    [Test]
    public async Task Text_NullBecomesBlank()
    {
        await Assert.That(ComputedValue.Text(null).Kind).IsEqualTo(ComputedValueKind.Blank);
    }

    [Test]
    public async Task TryGetNumber_IsStrict_NotBoolean()
    {
        await Assert.That(ComputedValue.Number(3).TryGetNumber(out var n)).IsTrue();
        await Assert.That(n).IsEqualTo(3.0);

        // Boolean NÃO é número (tipo exato).
        await Assert.That(ComputedValue.Boolean(true).TryGetNumber(out _)).IsFalse();
        await Assert.That(ComputedValue.Text("3").TryGetNumber(out _)).IsFalse();
    }

    [Test]
    public async Task TryGetBoolean_IsStrict()
    {
        await Assert.That(ComputedValue.Boolean(true).TryGetBoolean(out var b)).IsTrue();
        await Assert.That(b).IsTrue();
        await Assert.That(ComputedValue.Number(1).TryGetBoolean(out _)).IsFalse();
    }

    [Test]
    public async Task TryGetText_And_TryGetError_And_TryGetReference()
    {
        await Assert.That(ComputedValue.Text("hi").TryGetText(out var s)).IsTrue();
        await Assert.That(s).IsEqualTo("hi");

        await Assert.That(ComputedValue.Error(Error.DivZero).TryGetError(out var e)).IsTrue();
        await Assert.That(e).IsEqualTo(Error.DivZero);

        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        var range = Range("A1", "A2", sheet);
        await Assert.That(ComputedValue.Reference(range).TryGetReference(out var reference)).IsTrue();
        await Assert.That(reference).IsEqualTo(range);
        await Assert.That(ComputedValue.Number(1).TryGetReference(out _)).IsFalse();
    }

    [Test]
    public async Task As_ReturnsNullOnMismatch()
    {
        await Assert.That(ComputedValue.Number(2).AsDouble()).IsEqualTo(2.0);
        await Assert.That(ComputedValue.Boolean(true).AsDouble()).IsNull();
        await Assert.That(ComputedValue.Text("x").AsString()).IsEqualTo("x");
        await Assert.That(ComputedValue.Number(2).AsString()).IsNull();
    }

    [Test]
    public async Task To_ReturnsOnMatch_ThrowsOnMismatch()
    {
        await Assert.That(ComputedValue.Number(5).ToDouble()).IsEqualTo(5.0);

        var threw = false;
        try
        {
            _ = ComputedValue.Text("x").ToDouble();
        }
        catch (InvalidOperationException)
        {
            threw = true;
        }

        await Assert.That(threw).IsTrue();
    }

    [Test]
    public async Task Implicit_Conversions_AreInputOnly()
    {
        ComputedValue fromDouble = 3.0;
        ComputedValue fromBool = true;
        ComputedValue fromText = "abc";
        ComputedValue fromNull = (string?)null;

        await Assert.That(fromDouble.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(fromBool.Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(fromText.Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(fromNull.Kind).IsEqualTo(ComputedValueKind.Blank);
    }

    [Test]
    public async Task AsObject_BridgesToLegacyWorld()
    {
        await Assert.That(ComputedValue.Number(4).AsObject()).IsEqualTo(4.0);
        await Assert.That((bool)ComputedValue.Boolean(false).AsObject()!).IsFalse();
        await Assert.That(ComputedValue.Text("t").AsObject()).IsEqualTo("t");
        await Assert.That(ComputedValue.Blank.AsObject()).IsNull();
        await Assert.That(ComputedValue.Error(Error.Value).AsObject()).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task EnumerateValues_OverRange_YieldsCellValues()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["B1"] = Number(3);
        sheet["B2"] = Number(4);

        var values = ComputedValue.Reference(Range("A1", "B2", sheet)).EnumerateValues(workbook).ToList();

        await Assert.That(values.Count).IsEqualTo(4);
        await Assert.That(values.Sum(v => v.ToDouble())).IsEqualTo(10.0);
    }

    [Test]
    public async Task EnumerateValues_OnNonReference_IsEmpty()
    {
        var workbook = new Workbook();
        await Assert.That(ComputedValue.Number(1).EnumerateValues(workbook).Any()).IsFalse();
    }
}
