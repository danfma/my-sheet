using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests.Expressions;

public class ErrorTests
{
    [Test]
    public async Task Display_MatchesExcelCodes()
    {
        await Assert.That(Error.Null.Display).IsEqualTo("#NULL!");
        await Assert.That(Error.DivZero.Display).IsEqualTo("#DIV/0!");
        await Assert.That(Error.Value.Display).IsEqualTo("#VALUE!");
        await Assert.That(Error.Ref.Display).IsEqualTo("#REF!");
        await Assert.That(Error.Name.Display).IsEqualTo("#NAME?");
        await Assert.That(Error.Num.Display).IsEqualTo("#NUM!");
        await Assert.That(Error.NA.Display).IsEqualTo("#N/A");
    }

    [Test]
    public async Task ToString_PrintsDisplay()
    {
        await Assert.That(Error.DivZero.ToString()).IsEqualTo("#DIV/0!");
        await Assert.That($"{Error.Value}").IsEqualTo("#VALUE!");
    }

    [Test]
    public async Task Equality_ByCode()
    {
        await Assert.That(Error.FromDisplay("#VALUE!") == Error.Value).IsTrue();
        await Assert.That(Error.Value != Error.Num).IsTrue();
        await Assert.That(Error.Value.Equals(Error.Value)).IsTrue();
        await Assert.That(Error.Value.Equals(Error.Ref)).IsFalse();
    }

    [Test]
    public async Task RoundTrips_ThroughErrorValueNode()
    {
        // Error -> ErrorValue (nó de AST) -> Error, preservando a identidade.
        await Assert
            .That(Error.FromDisplay(Error.DivZero.ToErrorValue().ErrorCode))
            .IsEqualTo(Error.DivZero);
        await Assert.That(Error.FromDisplay(Error.NA.ToErrorValue().ErrorCode)).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task Calc_IsTheEighthError_AndRoundTripsExactly()
    {
        // #CALC! is Excel's "empty array" error (the FILTER page: "Otherwise, a #CALC! error will result,
        // as Excel does not currently support empty arrays"). It must be a REAL code, because FromDisplay
        // folds any unknown display onto #VALUE! — so before this code existed, an assertion that a
        // formula answers #CALC! could pass for the wrong reason through that fold. Every seam the value
        // layer crosses is exercised here: display, code, the AST node singleton, AsObject, and the
        // warm-start surrogate (CachedCellValue carries the int code).
        await Assert.That(Error.Calc.Display).IsEqualTo("#CALC!");
        await Assert.That(Error.Calc.Code).IsEqualTo(7);
        await Assert.That(Error.FromCode(7)).IsEqualTo(Error.Calc);
        await Assert.That(Error.FromCode(7).Display).IsEqualTo("#CALC!");
        await Assert.That(Error.FromDisplay("#CALC!")).IsEqualTo(Error.Calc);
        await Assert.That(Error.Calc != Error.Value).IsTrue();

        await Assert.That(ErrorValue.Calculation.ErrorCode).IsEqualTo("#CALC!");
        await Assert.That(Error.Calc.ToErrorValue()).IsSameReferenceAs(ErrorValue.Calculation);
        await Assert.That(ErrorValue.Calculation.AsError()).IsEqualTo(Error.Calc);
        await Assert
            .That(ComputedValue.Error(Error.Calc).AsObject())
            .IsEqualTo(ErrorValue.Calculation);

        var surrogate = CachedCellValue.TryFrom("Sheet1", "A1", ComputedValue.Error(Error.Calc));
        await Assert.That(surrogate!.ErrorCode).IsEqualTo(7);
        await Assert.That(surrogate.ToComputedValue().TryGetError(out var restored)).IsTrue();
        await Assert.That(restored).IsEqualTo(Error.Calc);
    }

    [Test]
    public async Task AnUnknownDisplay_StillFoldsOntoValue_AndCalcIsNoLongerUnknown()
    {
        // The fold is the pre-existing rule for a code the engine does not know (a #SPILL! read from a
        // file, a malformed literal); adding #CALC! must move #CALC! out of it and nothing else.
        await Assert.That(Error.FromDisplay("#SPILL!")).IsEqualTo(Error.Value);
        await Assert.That(Error.FromDisplay("#ERR?")).IsEqualTo(Error.Value);
        await Assert.That(new ErrorValue("#SPILL!").AsError()).IsEqualTo(Error.Value);
        await Assert.That(new ErrorValue("#CALC!").AsError()).IsEqualTo(Error.Calc);
        await Assert.That(new ErrorValue("#CALC!").AsError() != Error.Value).IsTrue();

        // A code past the table still displays as the unknown marker, so a snapshot written by a NEWER
        // build with a ninth code degrades rather than crashes; 7 is no longer past the table.
        await Assert.That(Error.FromCode(8).Display).IsEqualTo("#ERR?");
    }
}
