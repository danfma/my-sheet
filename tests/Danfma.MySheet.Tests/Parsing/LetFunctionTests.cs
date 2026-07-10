using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class LetFunctionTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task Let_SingleBinding()
    {
        await Assert.That(Calc("=LET(x,2,x*x)") as double?).IsEqualTo(4.0);
    }

    [Test]
    public async Task Let_MultipleBindings()
    {
        await Assert.That(Calc("=LET(x,2,y,3,x+y)") as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task Let_LaterBindingSeesEarlierName()
    {
        // x = 2 ; y = x + 1 = 3 ; x + y = 5
        await Assert.That(Calc("=LET(x,2,y,x+1,x+y)") as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task Let_DeepNestedShadowing()
    {
        // x = 1 ; inner x = 1 + 1 = 2 ; innermost x = 2 + 1 = 3
        await Assert.That(Calc("=LET(x,1,LET(x,x+1,LET(x,x+1,x)))") as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task BareUnknownName_IsStillNameError()
    {
        await Assert.That(Calc("=NOPE")).IsEqualTo(ErrorValue.Name);
    }
}
