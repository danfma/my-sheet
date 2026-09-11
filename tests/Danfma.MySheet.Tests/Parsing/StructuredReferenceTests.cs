using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Phase 4 T5 (items 8-10): the PARSER's structured-reference arms — the node a formula text actually
/// produces, which is the half <c>StructuredReferenceSyntaxTests</c> cannot see (it drives
/// <c>StructuredReferenceSyntax</c> directly, with a hand-built token). Every accepted row here is a shape
/// the grammar already pins there; what this file adds is that the whole path — tokenizer, the
/// <c>ParseIdentifier</c> arm, the grammar — reaches it from a formula string, in a function argument and on
/// both shared-formula entry points.
/// <para>
/// Oracle for every "measured" note: <b>Aspose.Cells 26.6.0, PLAIN entry, 2026-09-11</b>, over a
/// <c>Data!Tabela1</c> at A1:B4 (header <c>Item</c>/<c>Valor</c>, three data rows 10/20/30, so a column SUM
/// is 60) plus an empty second sheet <c>Report</c>.
/// </para>
/// </summary>
public class StructuredReferenceTests
{
    private static Expression Parse(string formula) =>
        ExpressionParser.Parse("=" + formula, new Sheet { Name = "Sheet1" });

    private static ParseException Throws(string formula)
    {
        var exception = Assert.Throws<ParseException>(() => Parse(formula));

        return exception!;
    }

    // === The accepted S1 forms, as NODES =================================================================

    // Item 14's rows, with the three the controller's re-measurement moved: `'My Table'[Valor]` is gone (its
    // own test below, as a rejection), and `[ Col ]` / `[[ Col ]]` both keep their spaces because the oracle
    // does NOT trim — `SUM(Tabela1[ Valor ])` over a header "Valor" is rejected (`Invalid table column:
    // Valor`) while the same spelling over a header `" Padded "` resolves. `Tabela1[]` (the data body) and
    // `Tabela1[[]]` (#All) are task 2's measurements, which no item mentions.
    [Test]
    [Arguments("Tabela1[Valor]", "Valor", TableArea.Data)]
    [Arguments("Tabela1[#All]", null, TableArea.All)]
    [Arguments("Tabela1[#Data]", null, TableArea.Data)]
    [Arguments("Tabela1[#Headers]", null, TableArea.Headers)]
    [Arguments("Tabela1[#Totals]", null, TableArea.Totals)]
    [Arguments("Tabela1[#totals]", null, TableArea.Totals)] // case-insensitive
    [Arguments("Tabela1[]", null, TableArea.Data)] // the whole data body
    [Arguments("Tabela1[[]]", null, TableArea.All)] // ...but the empty ITEM means #All
    [Arguments("Tabela1[[#Data],[Valor]]", "Valor", TableArea.Data)]
    [Arguments("Tabela1[[#All],[Valor]]", "Valor", TableArea.All)]
    [Arguments("Tabela1[[#Headers],[#Data]]", null, TableArea.HeadersAndData)]
    [Arguments("Tabela1[[#Data],[#Totals]]", null, TableArea.DataAndTotals)]
    [Arguments("Tabela1[[#Headers],[#Data],[% Comissao]]", "% Comissao", TableArea.HeadersAndData)]
    [Arguments("Tabela1[[Valor]]", "Valor", TableArea.Data)]
    [Arguments("Tabela1[Sales Amount]", "Sales Amount", TableArea.Data)]
    [Arguments("Tabela1[[Total $ Amount]]", "Total $ Amount", TableArea.Data)]
    [Arguments("Tabela1[Total (USD)]", "Total (USD)", TableArea.Data)]
    [Arguments("Tabela1[a,b]", "a,b", TableArea.Data)]
    [Arguments("Tabela1[a:b]", "a:b", TableArea.Data)]
    [Arguments("Tabela1[Unit_Price]", "Unit_Price", TableArea.Data)]
    [Arguments("Tabela1['#OfItems]", "#OfItems", TableArea.Data)]
    [Arguments("Tabela1['[bracket']]", "[bracket]", TableArea.Data)]
    [Arguments("Tabela1[a'']", "a'", TableArea.Data)]
    [Arguments("Tabela1[ Col ]", " Col ", TableArea.Data)] // NOT trimmed: the spaces are the name's
    [Arguments("Tabela1[[ Col ]]", " Col ", TableArea.Data)]
    [Arguments("Tabela1[[a@b]]", "a@b", TableArea.Data)]
    public async Task Parses_ToATableReference(string formula, string? column, TableArea area)
    {
        await Assert.That(Parse(formula)).IsEqualTo(new TableReference("Tabela1", column, area));
    }

    // `Tabela1[ Valor ]` is a deliberate KIND divergence, recorded rather than fixed: the oracle rejects it
    // at PARSE time (`Invalid table column:  Valor`), MySheet parses the column `" Valor "` and misses at
    // RESOLUTION with #REF!. Closing it would need the parser to know the table's headers, which it cannot
    // (it holds a sheet name and nothing else), so the divergence is in the error's kind and timing only.
    [Test]
    public async Task PaddedColumnName_ParsesAndMissesAtResolution_ADivergenceInKind()
    {
        await Assert
            .That(Parse("Tabela1[ Valor ]"))
            .IsEqualTo(new TableReference("Tabela1", " Valor ", TableArea.Data));

        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Valor");
        data["A2"] = new NumberValue(10);
        workbook.DefineTable("Tabela1", "Data", "A1:A2", ["Valor"]);
        data["C1"] = ExpressionParser.Parse("=SUM(Tabela1[ Valor ])", data);

        await Assert.That(workbook.GetCellValue("Data", "C1").TryGetError(out var error)).IsTrue();
        await Assert.That(error.Display).IsEqualTo("#REF!");
    }

    // The commonest real shape, and the one the writer renders back verbatim. `SUM(Tabela1[Valor])` = 60 on
    // the oracle, stored identically.
    [Test]
    public async Task Parses_AsAFunctionArgument()
    {
        var parsed = Parse("SUM(Tabela1[Valor])");

        await Assert.That(parsed is Sum { Expressions: [TableReference] }).IsTrue();
        await Assert
            .That(((Sum)parsed).Expressions[0])
            .IsEqualTo(new TableReference("Tabela1", "Valor", TableArea.Data));
        await Assert.That(parsed.ToFormula("Sheet1")).IsEqualTo("SUM(Tabela1[Valor])");
    }

    // === Where the arm sits, and why =====================================================================

    // The measured reason the arm goes BEFORE the IsCellReference check: that check is unbounded (any
    // letters-then-digits string) and answers TRUE for every one of these names, Excel's own defaults
    // included — so an arm placed after it would build a CellReference and leave the bracket token dangling.
    // TRUE/FALSE cover the IsBoolean check two lines above it, which costs nothing to jump.
    [Test]
    [Arguments("Tabela1")]
    [Arguments("Table1")]
    [Arguments("Sales2024")]
    [Arguments("ABC123")]
    [Arguments("A1")] // a name Excel could never give a table: still a TableReference, see below
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    public async Task TheArmWins_OverTheCellReferenceAndBooleanArms(string table)
    {
        await Assert
            .That(Parse(table + "[Valor]"))
            .IsEqualTo(new TableReference(table, "Valor", TableArea.Data));
    }

    // The context-free consequence, accepted deliberately: the parser holds only a sheet name, so it cannot
    // know that `A1` is not a table. The failure is therefore a resolution-time #NAME?, not a syntax error.
    [Test]
    public async Task AGridShapedTableName_FailsAtResolution_NotAtParse()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");
        sheet["C1"] = ExpressionParser.Parse("=SUM(A1[Valor])", sheet);

        await Assert.That(workbook.GetCellValue("Data", "C1").TryGetError(out var error)).IsTrue();
        await Assert.That(error.Display).IsEqualTo("#NAME?");
    }

    // === The bracket must be ADJACENT to the table name ==================================================

    // Both shapes a gap admits are REJECTED by the oracle, so the arm requires adjacency and they keep the
    // dangling-token error they have today: `SUM('Tabela1'[Valor])` is `Invalid "'"` and
    // `SUM(Tabela1 [Valor])` is "Invalid table reference, formula should be in table when specifing no table
    // name" (the space breaks the association, leaving an implicit-table reference). Quoting matters twice
    // over: ReadQuotedName hands this arm a DECODED identifier, so `'My Table'[Valor]` would produce a node
    // the writer renders as the unparsable `My Table[Valor]`, and a table name with a space cannot be
    // registered anyway (`Table.ValidateName`, and Aspose refuses the DisplayName: "Invalid text for the
    // defined name").
    [Test]
    [Arguments("'My Table'[Valor]", ParseErrorKind.UnexpectedToken)]
    [Arguments("'Tabela1'[Valor]", ParseErrorKind.UnexpectedToken)]
    [Arguments("Tabela1 [Valor]", ParseErrorKind.UnexpectedToken)]
    // In an argument slot the same token is what a ')' was expected instead of: no arm is taken, so the kind
    // is whatever the surrounding grammar says about a token it cannot use, never a structured kind.
    [Arguments("SUM('Tabela1'[Valor])", ParseErrorKind.ExpectedToken)]
    public async Task ANonAdjacentBracket_IsNotATableReference(string formula, ParseErrorKind kind)
    {
        var error = Throws(formula);

        await Assert.That(error.Kind).IsEqualTo(kind);
        await Assert.That(error.Token).IsEqualTo("[Valor]");
    }

    // === Both shared-formula entry points build the same node ============================================

    // A structured reference has no position component, which is why AnchoredFormulaSupport keeps it on the
    // shared-master fast path (StructuredReferenceSharedFormulaTests owns that verdict). The parser half of
    // that claim is here: the ANCHORED master parse and the per-slave token-delta parse — the two modes the
    // loader chooses between — produce the IDENTICAL node, whatever the delta.
    [Test]
    [Arguments(0, 0)]
    [Arguments(7, 3)]
    [Arguments(-2, 0)]
    public async Task AnchoredMasterAndShiftedSlave_BuildTheIdenticalNode(
        int deltaRow,
        int deltaColumn
    )
    {
        var sheet = new Sheet { Name = "Sheet1" };
        var tokens = ExpressionParser.TokenizeFormulaBody("Tabela1[Valor]");
        var expected = new TableReference("Tabela1", "Valor", TableArea.Data);

        var anchored = ExpressionParser.ParseAnchoredMasterBody(tokens, sheet);
        var shifted = ExpressionParser.ParseSharedFormulaBody(tokens, sheet, deltaRow, deltaColumn);

        await Assert.That(anchored).IsEqualTo(expected);
        await Assert.That(shifted).IsEqualTo(expected);
        await Assert.That(anchored).IsEqualTo(shifted);
    }
}
