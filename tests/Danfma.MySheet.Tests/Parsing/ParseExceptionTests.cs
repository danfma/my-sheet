using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// <see cref="ParseException"/> carries structured detail — <see cref="ParseException.Kind"/>,
/// <see cref="ParseException.Token"/>, <see cref="ParseException.Position"/> — so an integrator can tell a
/// syntax failure apart from a calculation error and report WHAT broke WHERE without parsing the message
/// (issue #8, diagnosability note). Positions are 0-based offsets into the formula body (after the '=').
/// </summary>
public class ParseExceptionTests
{
    private static ParseException Throws(string formula)
    {
        var sheet = new Sheet { Name = "Sheet1" };

        try
        {
            ExpressionParser.Parse(formula, sheet);
        }
        catch (ParseException exception)
        {
            return exception;
        }

        throw new InvalidOperationException($"'{formula}' parsed without error.");
    }

    [Test]
    public async Task UnexpectedCharacter()
    {
        var error = Throws("=1 # 2");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnexpectedCharacter);
        await Assert.That(error.Token).IsEqualTo("#");
        await Assert.That(error.Position).IsEqualTo(2);
    }

    [Test]
    public async Task UnterminatedString()
    {
        var error = Throws("=1&\"abc");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnterminatedString);
        await Assert.That(error.Token).IsEqualTo("\"abc");
        await Assert.That(error.Position).IsEqualTo(2);
    }

    [Test]
    public async Task UnterminatedQuotedName()
    {
        var error = Throws("='Other Sheet");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnterminatedQuotedName);
        await Assert.That(error.Token).IsEqualTo("'Other Sheet");
        await Assert.That(error.Position).IsEqualTo(0);
    }

    [Test]
    public async Task UnexpectedToken_TrailingInput()
    {
        var error = Throws("=1 2");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnexpectedToken);
        await Assert.That(error.Token).IsEqualTo("2");
        await Assert.That(error.Position).IsEqualTo(2); // body-relative: index 3 in the full "=1 2"
    }

    [Test]
    public async Task UnexpectedToken_DanglingOperator()
    {
        var error = Throws("=*2");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnexpectedToken);
        await Assert.That(error.Token).IsEqualTo("*");
        await Assert.That(error.Position).IsEqualTo(0);
    }

    [Test]
    public async Task UnexpectedToken_EmptyFormula_HasEmptyToken()
    {
        var error = Throws("=");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnexpectedToken);
        await Assert.That(error.Token).IsEqualTo(string.Empty);
        await Assert.That(error.Position).IsEqualTo(0);
    }

    [Test]
    public async Task ExpectedToken_UnclosedParenthesis_HasEmptyToken()
    {
        var error = Throws("=(1+2");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.ExpectedToken);
        await Assert.That(error.Token).IsEqualTo(string.Empty); // ran out of input
        await Assert.That(error.Position).IsEqualTo(4);
    }

    [Test]
    public async Task ExpectedToken_MissingComma()
    {
        var error = Throws("=SUM(A1 A2)");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.ExpectedToken);
        await Assert.That(error.Token).IsEqualTo("A2");
        await Assert.That(error.Position).IsEqualTo(7);
    }

    [Test]
    public async Task ExpectedCellReference_AfterSheetQualifier()
    {
        var error = Throws("=COUNTA(Other!*)");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.ExpectedCellReference);
        await Assert.That(error.Token).IsEqualTo("*");
        await Assert.That(error.Position).IsEqualTo(13);
    }

    [Test]
    public async Task ExpectedCellReference_QualifiedRangeWithBadEndpoint()
    {
        var error = Throws("=SUM(Other!A1:*)");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.ExpectedCellReference);
        await Assert.That(error.Position).IsEqualTo(10); // reported at the LEFT endpoint (the range start)
    }

    [Test]
    public async Task InvalidArgumentCount_NamesTheFunction()
    {
        var error = Throws("=1+ROUND(1)");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.InvalidArgumentCount);
        await Assert.That(error.Token).IsEqualTo("ROUND");
        await Assert.That(error.Position).IsEqualTo(2);
    }

    [Test]
    public async Task NestingTooDeep()
    {
        var error = Throws("=" + new string('(', 300) + "1" + new string(')', 300));

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.NestingTooDeep);
    }

    [Test]
    public async Task Message_KeepsThePositionSuffix()
    {
        // The human-readable message is unchanged by the structured properties.
        var error = Throws("=1 2");

        await Assert.That(error.Message).IsEqualTo("Unexpected token '2' (at position 2).");
    }

    [Test]
    public async Task LegacyConstructor_IsUnspecifiedKind_WithEmptyToken()
    {
        // The pre-3.16 (message, position) overload still compiles and behaves; only the parser sets a Kind.
        var error = new ParseException("Something is off", 7);

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.Unspecified);
        await Assert.That(error.Token).IsEqualTo(string.Empty);
        await Assert.That(error.Position).IsEqualTo(7);
        await Assert.That(error.Message).IsEqualTo("Something is off (at position 7).");
    }
}
