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

    // --- A range bound by LET stays a RANGE (reference value), so range consumers see the cells — not the
    // #VALUE! a bare range evaluates to. Regression for issue #8: `LET(hdr, Data!$1:$1, MATCH(x, hdr, 0))`
    // returned #N/A because `hdr` was bound to #VALUE!.

    private static object? CalcOnGrid(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new Danfma.MySheet.Expressions.StringValue("a");
        sheet["B1"] = new Danfma.MySheet.Expressions.StringValue("b");
        sheet["C1"] = new Danfma.MySheet.Expressions.StringValue("c");
        sheet["A2"] = new NumberValue(10);
        sheet["B2"] = new NumberValue(20);
        sheet["C2"] = new NumberValue(30);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    [Arguments("=LET(r,A2:C2,SUM(r))", 60.0)]
    [Arguments("=LET(r,A1:C1,MATCH(\"b\",r,0))", 2.0)]
    [Arguments("=LET(r,A1:C2,INDEX(r,2,3))", 30.0)]
    [Arguments("=LET(r,A1:C2,ROWS(r))", 2.0)]
    [Arguments("=LET(r,1:1,COUNTA(r))", 3.0)] // whole row (open range)
    [Arguments("=LET(r,$1:$1,COUNTA(r))", 3.0)] // absolute whole row (issue #8 shape)
    [Arguments("=LET(r,A:A,SUM(r))", 10.0)] // whole column
    [Arguments("=LET(r,(A2:A2,C2:C2),SUM(r))", 40.0)] // union
    [Arguments("=LET(r,A1:C1,LET(n,MATCH(\"c\",r,0),n*10))", 30.0)] // nested LET over a bound range
    public async Task Let_RangeBinding_IsUsableAsARange(string formula, double expected)
    {
        await Assert.That(CalcOnGrid(formula) as double?).IsEqualTo(expected);
    }

    [Test]
    public async Task Let_SingleCellBinding_IsItsValue()
    {
        // A single cell is captured by VALUE (as a defined name would be): x = A2 = 10.
        await Assert.That(CalcOnGrid("=LET(x,A2,x+1)") as double?).IsEqualTo(11.0);
    }

    [Test]
    public async Task Let_RangeBinding_UsedAsScalar_IsValueError()
    {
        // Using the bound range where a scalar is required is #VALUE!, exactly like `=A1:C1+1`.
        await Assert.That(CalcOnGrid("=LET(r,A2:C2,r+1)")).IsEqualTo(ErrorValue.NotValue);
    }
}
