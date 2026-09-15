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

    // === The three structured-reference kinds (Phase 4) ==================================================

    // The tokenizer's own reject: the balanced-bracket scan runs off the end, so the token is the whole
    // remainder from the '[' and the position is the '[' itself. This is the ONE structured kind raised
    // before the parser ever runs.
    [Test]
    public async Task UnterminatedBracketedReference()
    {
        var error = Throws("=Tabela1[Valor");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnterminatedBracketedReference);
        await Assert.That(error.Token).IsEqualTo("[Valor");
        await Assert.That(error.Position).IsEqualTo(7);
    }

    // Balanced brackets whose content the oracle REJECTS (Aspose.Cells 26.6.0, PLAIN entry): an unknown
    // specifier is `Invalid table rows type`, two columns is `Unknown token with bracket`, and an illegal
    // specifier pair is "Specified rows to make up the contiguous range can only be one of following items:
    // Headers, Data, Totals, Data and Headers, Data and Totals, CurrentRow". So the kind is Invalid, not
    // Unsupported — the design's "classify Unsupported when unsure whether Excel accepts a shape" tie-break
    // no longer applies to any of them. Two rows the design listed here are ACCEPTED and moved out:
    // `Tabela1[]` (the whole data body) and `Tabela1[[Valor],[#Data]]` (accepted and REORDERED).
    [Test]
    [Arguments("=Tabela1[#Bogus]", "[#Bogus]", 8)]
    [Arguments("=Tabela1[[A],[B]]", "[[A],[B]]", 13)] // reported at the SECOND column's own offset
    [Arguments("=Tabela1[[#All],[#Data]]", "[[#All],[#Data]]", 8)]
    public async Task InvalidStructuredReference(string formula, string token, int position)
    {
        var error = Throws(formula);

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.InvalidStructuredReference);
        await Assert.That(error.Token).IsEqualTo(token);
        await Assert.That(error.Position).IsEqualTo(position);
    }

    // Two different reasons live under this test, and the rows do not say which. The TABLE-QUALIFIED rows
    // are valid Excel that MySheet does not model — an S1 scope decision, not parity: measured,
    // `Tabela1[@Valor]` outside the table evaluates to #VALUE! (it is not rejected) and is STORED as the
    // `[[#This Row],[Valor]]` item list, `Tabela1[#This Row]` is accepted and rewritten to `Tabela1[@]`, a
    // even `Data!Tabela1[Valor]` resolves — the oracle answers 60 and stores the formula with the qualifier
    // STRIPPED. The BARE-PREFIX rows
    // (`=SUM([Valor])`, `=[@Valor]`, `=[@]`) are rejected by BOTH sides: the oracle set-throws all three
    // with the same message ("Invalid table reference, formula should be in table when specifing no table
    // name" — measured, Aspose.Cells 26.6.0, PLAIN, 2026-09-11), and MySheet throws the named kind below.
    [Test]
    [Arguments("=Tabela1[@Valor]", "[@Valor]", 8)]
    [Arguments("=Tabela1[#This Row]", "[#This Row]", 8)]
    [Arguments("=Tabela1[[#This Row],[Valor]]", "[[#This Row],[Valor]]", 9)]
    [Arguments("=SUM([Valor])", "[Valor]", 4)]
    [Arguments("=[@Valor]", "[@Valor]", 0)] // the current-row form with no table name: a PREFIX-position arm
    [Arguments("=[@]", "[@]", 0)]
    [Arguments("=Data!Tabela1[Valor]", "Tabela1", 5)] // the table NAME's position, not the bracket's
    public async Task UnsupportedStructuredReference(string formula, string token, int position)
    {
        var error = Throws(formula);

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnsupportedStructuredReference);
        await Assert.That(error.Token).IsEqualTo(token);
        await Assert.That(error.Position).IsEqualTo(position);
    }

    // A `[...]` in PREFIX position — nothing before it to be a table name — is one of three different
    // mistakes, and the whole point of that arm is a message that says WHICH. The first two are pinned by
    // their payload; the third cannot be: `[Valor]` (an implicit-table column) and `[Book1.xlsx]` (an external
    // reference BY NAME) are indistinguishable from the token, so the message names both rather than
    // asserting one. Only the INDEX spelling `[1]` is decidable, and it is the only external spelling a saved
    // file carries, because the package keeps the workbook name in an externalLink part.
    [Test]
    [Arguments("=[@Valor]", "This-row structured references")]
    [Arguments("=[1]Sheet1!A1", "External-workbook references")]
    [Arguments("=SUM([Valor])", "A structured reference with no table name")]
    [Arguments("=[Book1.xlsx]Sheet1!A1", "A structured reference with no table name")]
    public async Task UnsupportedStructuredReference_NamesTheShape(string formula, string message)
    {
        var error = Throws(formula);

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnsupportedStructuredReference);
        await Assert.That(error.Message).Contains(message);
    }

    // The external-workbook form, pinned here BECAUSE this file is the only place UnexpectedCharacter is
    // asserted (the first test above): the kind for `=[1]Sheet1!A1` deliberately changed from
    // UnexpectedCharacter — the tokenizer had no '[' at all — to UnsupportedStructuredReference, which names
    // the shape instead of the character.
    [Test]
    public async Task UnsupportedStructuredReference_ExternalWorkbook()
    {
        var error = Throws("=[1]Sheet1!A1");

        await Assert.That(error.Kind).IsEqualTo(ParseErrorKind.UnsupportedStructuredReference);
        await Assert.That(error.Token).IsEqualTo("[1]");
        await Assert.That(error.Position).IsEqualTo(0);
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
