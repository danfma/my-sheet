using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class ScalarReferenceValueEvaluationTests
{
    [Test]
    [Arguments("=ISERR(OFFSET(A1,TICK(),0))", true, false)]
    [Arguments("=IFNA(OFFSET(A1,TICK(),0),99)", null, true)]
    public async Task FailedReferenceProducer_IsEvaluatedOnce(
        string formula,
        bool? expectedBoolean,
        bool expectedReferenceError
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(7);
        var invocations = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++invocations == 1 ? -1 : 0);

        var result = ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        // The first draw moves OFFSET above row 1 and yields #REF!; a second draw would select A1=7.
        if (expectedReferenceError)
        {
            await Assert.That(result).IsEqualTo(ErrorValue.Reference);
        }
        else
        {
            await Assert.That(result as bool?).IsEqualTo(expectedBoolean);
        }
        await Assert.That(invocations).IsEqualTo(1);
    }
}
