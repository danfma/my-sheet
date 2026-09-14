using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The contract matrix for a reference to a NON-EXISTENT sheet. A missing sheet is a STRUCTURAL failure of
/// the reference (#REF!), fiel ao Excel: it must propagate through EVERY consuming function — INCLUDING the
/// error-ignoring COUNT family — instead of throwing <see cref="KeyNotFoundException"/> or being swallowed as
/// an empty range. The controls prove the per-cell value-error policy over an EXISTING sheet is unchanged.
/// </summary>
public class MissingSheetReferenceTests
{
    // Builds a workbook with a single sheet "Main". Seeds go into the given cells; the formula under test
    // lives in Z1 (a column outside the A:E ranges the matrix probes, so a whole-column control never sees
    // the formula cell itself).
    private static Workbook Build(string formula, params (string Id, string Formula)[] seed)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");

        foreach (var (id, cellFormula) in seed)
        {
            main[id] = ExpressionParser.Parse(cellFormula, main);
        }

        main["Z1"] = ExpressionParser.Parse(formula, main);

        return workbook;
    }

    private static ComputedValue Eval(string formula, params (string Id, string Formula)[] seed) =>
        Build(formula, seed).GetCellValue("Main", "Z1");

    private static async Task AssertRef(string formula, params (string Id, string Formula)[] seed)
    {
        var value = Eval(formula, seed);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);

        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // === Missing sheet → #REF! through every consuming function ==========================================

    [Test]
    public async Task Cell_MissingSheet_IsRef() => await AssertRef("=Ghost!A1");

    [Test]
    public async Task Text_MissingSheet_IsRef() => await AssertRef("=UPPER(Ghost!E9)");

    [Test]
    public async Task Operator_MissingSheet_IsRef() => await AssertRef("=Ghost!A1 + 1");

    [Test]
    public async Task Sum_MissingSheet_IsRef() => await AssertRef("=SUM(Ghost!A:A)");

    [Test]
    public async Task Count_MissingSheet_IsRef() => await AssertRef("=COUNT(Ghost!A:A)");

    [Test]
    public async Task CountA_MissingSheet_IsRef() => await AssertRef("=COUNTA(Ghost!A:A)");

    [Test]
    public async Task CountBlank_MissingSheet_IsRef() => await AssertRef("=COUNTBLANK(Ghost!A:A)");

    [Test]
    public async Task CountIf_MissingSheet_IsRef() => await AssertRef("=COUNTIF(Ghost!D:D,1)");

    [Test]
    public async Task SumIf_MissingSheet_IsRef() => await AssertRef("=SUMIF(Ghost!A:A,\">5\")");

    [Test]
    public async Task Average_MissingSheet_IsRef() => await AssertRef("=AVERAGE(Ghost!B2:B9)");

    [Test]
    public async Task VLookup_MissingSheet_IsRef() => await AssertRef("=VLOOKUP(1,Ghost!A:B,2)");

    [Test]
    public async Task Rows_MissingSheet_IsRef() => await AssertRef("=ROWS(Ghost!A:A)");

    [Test]
    public async Task Columns_MissingSheet_IsRef() => await AssertRef("=COLUMNS(Ghost!A:A)");

    [Test]
    public async Task Union_WithOneMissingSheet_IsRef() =>
        await AssertRef("=SUM(Main!A:A, Ghost!A:A)", ("A1", "=10"));

    // === Long tail: every reference-consuming function propagates the structural #REF! =====================

    [Test]
    public async Task SumProduct_FirstArgumentMissingSheet_IsRef() =>
        await AssertRef("=SUMPRODUCT(Ghost!A:A, Main!A:A)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task SumProduct_SecondArgumentMissingSheet_IsRef() =>
        await AssertRef("=SUMPRODUCT(Main!A1:A2, Ghost!B1:B2)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task SumX2MY2_MissingSheet_IsRef() =>
        await AssertRef("=SUMX2MY2(Ghost!A:A, Main!B:B)", ("B1", "=1"));

    [Test]
    public async Task Correl_MissingSheet_IsRef() =>
        await AssertRef("=CORREL(Ghost!A:A, Main!B:B)", ("B1", "=1"), ("B2", "=2"));

    [Test]
    public async Task Covar_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=COVAR(Main!A:A, Ghost!B:B)", ("A1", "=1"), ("A2", "=2"));

    [Test]
    public async Task Forecast_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=FORECAST(1, Ghost!A:A, Main!B:B)", ("B1", "=1"), ("B2", "=2"));

    [Test]
    public async Task Prob_MissingSheet_IsRef() =>
        await AssertRef("=PROB(Ghost!A:A, Main!B:B, 1)", ("B1", "=1"));

    [Test]
    public async Task Match_MissingSheet_IsRef() => await AssertRef("=MATCH(1, Ghost!A:A)");

    [Test]
    public async Task XLookup_MissingSheet_IsRef() =>
        await AssertRef("=XLOOKUP(1, Ghost!A:A, Main!B:B)");

    [Test]
    public async Task XMatch_MissingSheet_IsRef() => await AssertRef("=XMATCH(1, Ghost!A:A)");

    [Test]
    public async Task Lookup_MissingSheet_IsRef() => await AssertRef("=LOOKUP(1, Ghost!A:A)");

    [Test]
    public async Task Index_MissingSheet_IsRef() => await AssertRef("=INDEX(Ghost!A:A, 1)");

    // Oracle (Aspose.Cells 26.7.0, Main!AZ5000; PLAIN/CSE): #REF! / #REF!, TRUE / TRUE, 999 / 999.
    [Test]
    public async Task Index_ZeroAxisMissingSheet_IsRefBeforeTheReferenceEscapes()
    {
        await AssertRef("=INDEX(Ghost!A1:A3,0,1)");
        await Assert.That(Eval("=ISERROR(INDEX(Ghost!A1:A3,0,1))").AsObject() as bool?).IsTrue();
        await Assert.That(Eval("=IFERROR(INDEX(Ghost!A1:A3,0,1),999)").AsObject()).IsEqualTo(999.0);
        await AssertRef("=SUM(INDEX(Ghost!A1:A3,0,1))");
    }

    // A BOUNDED ghost range (not just a whole-column open range): the lookup would otherwise scan its cells,
    // skip the per-cell #REF! keys, and degrade to #N/A. It must short-circuit to #REF! structurally.
    [Test]
    public async Task VLookup_BoundedMissingSheet_IsRef() =>
        await AssertRef("=VLOOKUP(1, Ghost!A1:B5, 2)");

    [Test]
    public async Task HLookup_BoundedMissingSheet_IsRef() =>
        await AssertRef("=HLOOKUP(1, Ghost!A1:B5, 2)");

    [Test]
    public async Task TextJoin_MissingSheet_IsRef() =>
        await AssertRef("=TEXTJOIN(\",\", TRUE, Ghost!A:A)");

    [Test]
    public async Task Concat_MissingSheet_IsRef() => await AssertRef("=CONCAT(Ghost!A:A)");

    [Test]
    public async Task Concatenate_MissingSheet_IsRef() =>
        await AssertRef("=CONCATENATE(Ghost!A:A)");

    [Test]
    public async Task Npv_ValuesMissingSheet_IsRef() => await AssertRef("=NPV(0.1, Ghost!A:A)");

    // The rate argument is itself a reference to a missing sheet — a structural #REF! before any math.
    [Test]
    public async Task Npv_RateMissingSheet_IsRef() =>
        await AssertRef("=NPV(Ghost!A1, Main!A1)", ("A1", "=10"));

    [Test]
    public async Task Irr_MissingSheet_IsRef() => await AssertRef("=IRR(Ghost!A:A)");

    [Test]
    public async Task Mirr_MissingSheet_IsRef() => await AssertRef("=MIRR(Ghost!A:A, 0.1, 0.1)");

    [Test]
    public async Task XNpv_MissingSheet_IsRef() =>
        await AssertRef("=XNPV(0.1, Ghost!A:A, Main!B:B)");

    [Test]
    public async Task XIrr_MissingSheet_IsRef() => await AssertRef("=XIRR(Ghost!A:A, Main!B:B)");

    [Test]
    public async Task FvSchedule_MissingSheet_IsRef() =>
        await AssertRef("=FVSCHEDULE(100, Ghost!A:A)");

    [Test]
    public async Task NetworkDays_HolidaysMissingSheet_IsRef() =>
        await AssertRef("=NETWORKDAYS(Main!A1, Main!A2, Ghost!C:C)", ("A1", "=1"), ("A2", "=10"));

    [Test]
    public async Task Workday_HolidaysMissingSheet_IsRef() =>
        await AssertRef("=WORKDAY(Main!A1, 5, Ghost!C:C)", ("A1", "=1"));

    [Test]
    public async Task SeriesSum_MissingSheet_IsRef() =>
        await AssertRef("=SERIESSUM(2, 1, 1, Ghost!A:A)");

    [Test]
    public async Task And_MissingSheet_IsRef() => await AssertRef("=AND(Ghost!A:A)");

    [Test]
    public async Task Or_MissingSheet_IsRef() => await AssertRef("=OR(Ghost!A:A)");

    [Test]
    public async Task Xor_MissingSheet_IsRef() => await AssertRef("=XOR(Ghost!A:A)");

    [Test]
    public async Task Row_MissingSheet_IsRef() => await AssertRef("=ROW(Ghost!A1)");

    [Test]
    public async Task Column_MissingSheet_IsRef() => await AssertRef("=COLUMN(Ghost!A1)");

    // Already reaches #REF! through the guarded Fold, but pinned here as part of the contract matrix.
    [Test]
    public async Task Median_MissingSheet_IsRef() => await AssertRef("=MEDIAN(Ghost!A:A)");

    [Test]
    public async Task Percentile_LegacyAlias_MissingSheet_IsRef() =>
        await AssertRef("=PERCENTILE(Ghost!A:A, 0.5)");

    // === Phase 5: a STRUCTURED reference (ruling R2) =====================================================
    //
    // Two failures wear the same word "unresolvable" and behave nothing alike, which is the whole content of
    // ruling R2:
    //
    //   * the table RESOLVES but its sheet is gone -> a STRUCTURAL failure of the reference, the same class as
    //     Ghost!A:A above, so the guard short-circuits every consumer to #REF!. This matrix is MySheet's OWN
    //     convention and cannot be measured on Aspose.Cells: deleting the worksheet deletes the ListObject
    //     with it, so the oracle has no such shape to answer for. It is pinned here as the convention, and the
    //     reason it is not a free choice is that SUBTOTAL/AGGREGATE index workbook.Sheets[range.SheetName]
    //     with the THROWING indexer (AggregateCodes.Gather), so without the guard arm SUBTOTAL(9,T[Valor]) is
    //     an unhandled KeyNotFoundException rather than a wrong number;
    //
    //   * the table does NOT resolve (unknown table, unknown column, [#Totals] with no totals row) -> a plain
    //     error VALUE, so the guard returns null and each consumer does with it what it does with any
    //     error-valued argument. COUNT answers 0 and COUNTA answers 1, NOT #REF! - measured on the oracle,
    //     both entry modes (see StructuredReference_ThatDoesNotResolve_IsAnErrorValue_NotAShortCircuit).
    //
    // Every tree here is built by hand, as TableReferenceTests does — a test-isolation choice now that the
    // parser emits the node (Phase 4's parse arms), not a need.

    // The oracle's own fixture on a sheet of its own: Data!Tabela1 = A1:C4, header Item/Valor/Qtd, data rows
    // a,10,1 / b,20,2 / c,30,3, no totals row. The formula under test still lives in Main!Z1.
    private static Workbook TableFixture(bool removeTheSheet)
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Main");
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["C1"] = new StringValue("Qtd");
        data["A2"] = new StringValue("a");
        data["B2"] = new NumberValue(10);
        data["C2"] = new NumberValue(1);
        data["A3"] = new StringValue("b");
        data["B3"] = new NumberValue(20);
        data["C3"] = new NumberValue(2);
        data["A4"] = new StringValue("c");
        data["B4"] = new NumberValue(30);
        data["C4"] = new NumberValue(3);
        workbook.DefineTable("Tabela1", "Data", "A1:C4", ["Item", "Valor", "Qtd"]);

        if (removeTheSheet)
        {
            // Removing the SHEET does not unregister the table: the registry keeps the geometry, so the node
            // still resolves to Data!B2:B4 and the guard is the only thing standing between that resolved
            // range and the throwing Sheets[...] indexer.
            workbook.Sheets.TryRemove("Data", out _);
        }

        return workbook;
    }

    private static Expression TableColumn(
        TableArea area = TableArea.Data,
        string? column = "Valor"
    ) => new TableReference("Tabela1", column, area);

    // The consuming functions of the matrix, by name, because [Arguments] takes constants only.
    private static Expression Consumer(string consumer, Expression node) =>
        consumer switch
        {
            "SUM" => new Sum([node]),
            "COUNT" => new Count([node]),
            "COUNTA" => new CountA([node]),
            "SUBTOTAL(9)" => new Subtotal([new NumberValue(9), node]),
            "SUBTOTAL(3)" => new Subtotal([new NumberValue(3), node]),
            "COUNTIF" => new CountIf([node, new StringValue(">0")]),
            "ROWS" => new Rows([node]),
            "COLUMNS" => new Columns([node]),
            "VLOOKUP(1)" => new VLookup([new NumberValue(1), node, new NumberValue(1)]),
            "VLOOKUP(20)" => new VLookup([new NumberValue(20), node, new NumberValue(1)]),
            "AVERAGE" => new Average([node]),
            "MAX" => new Max([node]),
            "SUMIF" => new SumIf([node, new StringValue(">0")]),
            "COUNTBLANK" => new CountBlank([node]),
            "ISREF" => new IsRef([node]),
            "MATCH" => new Match([new NumberValue(1), node, new NumberValue(0)]),
            "INDEX" => new Danfma.MySheet.Expressions.Lookup.Index([
                node,
                new NumberValue(1),
                new NumberValue(1),
            ]),
            "HLOOKUP(1)" => new HLookup([new NumberValue(1), node, new NumberValue(1)]),
            "XMATCH" => new XMatch([new NumberValue(1), node]),
            "LOOKUP(1)" => new Lookup([new NumberValue(1), node]),
            "OFFSET" => new Offset([node, new NumberValue(1), new NumberValue(1)]),
            "FORMULATEXT" => new FormulaText([node]),
            "XLOOKUP" => new XLookup([
                new NumberValue(1),
                node,
                new RangeReference("B2", "B4", "Data"),
            ]),
            _ => throw new ArgumentOutOfRangeException(
                nameof(consumer),
                consumer,
                "unknown consumer"
            ),
        };

    // One fresh workbook per row (a reused one would serve Z1 from the memo) and the consumer's name travels
    // with the value, so a failing row names itself.
    private static async Task AssertConsumers(
        bool removeTheSheet,
        Expression node,
        params (string Consumer, object? Expected)[] rows
    )
    {
        foreach (var (consumer, expected) in rows)
        {
            var workbook = TableFixture(removeTheSheet);
            workbook["Main"]["Z1"] = Consumer(consumer, node);

            var actual = workbook.GetCellValue("Main", "Z1").AsObject();

            await Assert.That((consumer, actual)).IsEqualTo((consumer, expected));
        }
    }

    // MySheet's convention, unmeasurable on the oracle (deleting the sheet deletes the table): a table whose
    // RESOLVED range names a sheet that no longer exists is the same structural #REF! as Ghost!A:A, through
    // every consumer INCLUDING the error-ignoring COUNT family. Measured on this branch with the guard arm
    // removed, which is what each row below is worth:
    //
    //   SUBTOTAL(9) and SUBTOTAL(3) THROW KeyNotFoundException | COUNTA 3 | COUNT 0 | COUNTIF 0 | SUMIF 0
    //   COUNTBLANK 0 | VLOOKUP(1,...,1) #N/A | (SUM, ROWS, COLUMNS, AVERAGE, MAX already #REF!)
    //
    // So SUBTOTAL is the row that does not merely answer wrong without the guard, and the COUNT family is the
    // row that answers a number where a range on a ghost sheet answers #REF!.
    [Test]
    [Arguments("SUM")]
    [Arguments("COUNT")]
    [Arguments("COUNTA")]
    [Arguments("SUBTOTAL(9)")]
    [Arguments("SUBTOTAL(3)")]
    [Arguments("COUNTIF")]
    [Arguments("ROWS")]
    [Arguments("COLUMNS")]
    [Arguments("VLOOKUP(1)")]
    [Arguments("AVERAGE")]
    [Arguments("MAX")]
    [Arguments("SUMIF")]
    [Arguments("COUNTBLANK")]
    public async Task StructuredReference_OnARemovedSheet_IsRefThroughEveryConsumer(
        string consumer
    ) =>
        await AssertConsumers(
            removeTheSheet: true,
            TableColumn(),
            (consumer, ErrorValue.Reference)
        );

    // The one consumer the guard does not reach, and it is NOT a table quirk: ISREF is a syntactic question
    // (NamedReferences.TryResolveReference), so a table on a removed sheet is TRUE — byte for byte what the
    // literal baseline answers, re-measured on this branch: ISREF(Ghost!A1:A3) is TRUE too, while
    // COUNTA(Ghost!A1:A3) is #REF!. Contrast with the UNRESOLVABLE table below, where ISREF is FALSE because
    // no reference is formed at all.
    [Test]
    public async Task IsRef_OverAStructuredReferenceOnARemovedSheet_IsTrue_LikeTheLiteralBaseline()
    {
        await AssertConsumers(removeTheSheet: true, TableColumn(), ("ISREF", true));

        var workbook = TableFixture(removeTheSheet: true);
        workbook["Main"]["Z1"] = new IsRef([new RangeReference("B2", "B4", "Data")]);

        await Assert.That(workbook.GetCellValue("Main", "Z1").AsObject() as bool?).IsTrue();
    }

    // The control that makes the matrix above a proof instead of a tautology: the SAME consumers over the same
    // table with its sheet intact answer the oracle's numbers (Aspose.Cells 26.6.0, 2026-09-11, identical in
    // the PLAIN and the array-entered columns). Without it, a guard that returned #REF! for every structured
    // reference would pass the matrix.
    [Test]
    public async Task StructuredReference_OnALiveSheet_AnswersTheOraclesNumbers() =>
        await AssertConsumers(
            removeTheSheet: false,
            TableColumn(),
            ("SUM", 60.0),
            ("COUNT", 3.0),
            ("COUNTA", 3.0),
            ("SUBTOTAL(9)", 60.0),
            ("SUBTOTAL(3)", 3.0),
            ("COUNTIF", 3.0),
            ("ROWS", 3.0),
            ("COLUMNS", 1.0),
            ("VLOOKUP(20)", 20.0),
            ("AVERAGE", 20.0),
            ("MAX", 30.0),
            ("SUMIF", 60.0),
            ("COUNTBLANK", 0.0),
            ("ISREF", true)
        );

    // Phase 5 T5 (R3): the BARE table name reaches this file's guard through the NameReference arm, which
    // resolves it (NamedReferences.TryResolveRaw -> the Tables arm -> the concrete data-body rectangle) and
    // re-checks the RESOLVED range's sheet. So a name whose table lives on a REMOVED sheet is the same
    // structural #REF! as the spelled node — the same MySheet convention as the matrix above (unmeasurable
    // on the oracle: deleting the sheet deletes the table), and this is the row that verifies the arm "sees
    // the node for free": no new guard code, only the resolution underneath it.
    [Test]
    [Arguments("SUM")]
    [Arguments("COUNTA")]
    [Arguments("ROWS")]
    [Arguments("COLUMNS")]
    [Arguments("SUBTOTAL(9)")]
    public async Task ABareTableName_OnARemovedSheet_IsRefThroughTheNameArm(string consumer) =>
        await AssertConsumers(
            removeTheSheet: true,
            new NameReference("Tabela1"),
            (consumer, ErrorValue.Reference)
        );

    // The live-sheet control for the rows above, with the oracle's own bare-name numbers (Aspose.Cells
    // 26.6.0, 2026-09-11, PLAIN and array-entered agreeing on every row here): the name answers 66 through
    // the guard's resolution, not the 0 of the blank cell the token used to be mistaken for.
    [Test]
    public async Task ABareTableName_OnALiveSheet_AnswersTheOraclesNumbers() =>
        await AssertConsumers(
            removeTheSheet: false,
            new NameReference("Tabela1"),
            ("SUM", 66.0),
            ("COUNTA", 9.0),
            ("ROWS", 3.0),
            ("COLUMNS", 3.0),
            ("ISREF", true)
        );

    // Ruling R2, and the reason the guard arm must return null rather than the table's error: an unresolvable
    // structured reference is an error VALUE. Measured on Aspose.Cells 26.6.0 (2026-09-11) over
    // Tabela1[#Totals] on a table with NO totals row - the one unresolvable structured reference Aspose
    // accepts at formula-set time - with the PLAIN and the array-entered columns agreeing on every row:
    //
    //   SUM #REF! | COUNT 0 | COUNTA 1 | SUBTOTAL(9) #REF! | SUBTOTAL(3) #REF! | ROWS #REF! | COLUMNS #REF!
    //   VLOOKUP(1,...,1) #REF! | AVERAGE #REF! | MAX #REF! | ISREF FALSE
    //   COUNTIF(...,">0") / SUMIF / COUNTBLANK #REF!  <- the criteria family diverges; see the next test
    //
    // COUNT 0 and COUNTA 1 are the rows that forbid the short-circuit: a guard returning the table's #REF!
    // would make both answer #REF!, two silent divergences inside the very family the guard exists for.
    [Test]
    public async Task StructuredReference_ThatDoesNotResolve_IsAnErrorValue_NotAShortCircuit() =>
        await AssertConsumers(
            removeTheSheet: false,
            TableColumn(area: TableArea.Totals, column: null),
            ("SUM", ErrorValue.Reference),
            ("COUNT", 0.0),
            ("COUNTA", 1.0),
            ("SUBTOTAL(9)", ErrorValue.Reference),
            ("SUBTOTAL(3)", ErrorValue.Reference),
            ("ROWS", ErrorValue.Reference),
            ("COLUMNS", ErrorValue.Reference),
            ("VLOOKUP(1)", ErrorValue.Reference),
            ("AVERAGE", ErrorValue.Reference),
            ("MAX", ErrorValue.Reference),
            ("ISREF", false)
        );

    // Sweep item 34(a), CLOSED 2026-09-11: the criteria family's range slot propagates an error-valued
    // argument. What it was: COUNTIF/SUMIF/COUNTBLANK all answered 0 here — the node's error streamed as
    // the one element the criteria discards — where the oracle (Aspose.Cells 26.6.0, both entry modes)
    // answers #REF! over Tabela1[#Totals]. The fix is the one general arm in PositionalRange.Open's
    // fallback (RangeValueCursor's mirror and COUNTBLANK's slot included); the sweep reference is
    // phase-11-excel-compatibility-sweep.md item 34, and the pre-fix zeros lived in
    // TheCriteriaFamily_OverAnUnresolvableStructuredReference_IsZero_ADivergence.
    [Test]
    public async Task TheCriteriaFamily_OverAnUnresolvableStructuredReference_PropagatesTheNodesError() =>
        await AssertConsumers(
            removeTheSheet: false,
            TableColumn(area: TableArea.Totals, column: null),
            // Oracle (both entry modes): #REF! on all three. Was: 0 on all three.
            ("COUNTIF", ErrorValue.Reference),
            ("SUMIF", ErrorValue.Reference),
            ("COUNTBLANK", ErrorValue.Reference)
        );

    // Sweep item 34(b), CLOSED 2026-09-11: a consumer that resolves its own reference argument reports the
    // node's OWN error when the node does not resolve — the #NAME? an unknown table's structured reference
    // carries (the design's mapping, a table name living in the same namespace as a defined name) —
    // instead of its fallback code. What it was: VLOOKUP/INDEX answered their own #REF! and MATCH its
    // not-found #N/A here, where the oracle's unknown-NAME column (the directly measurable twin, both
    // entry modes) answers #NAME? for all three. The rule is one for the whole NAME CLASS — see
    // TheResolvingFamily_OverAnUnresolvedName_ReportsTheNodesError below — landed through
    // ReferencePosition's shared unresolved arm; the sweep reference is phase-11 item 34 and the pre-fix
    // codes lived in AResolvingConsumer_OverAnUnknownTable_ReportsItsOwnCode_ADivergence.
    [Test]
    public async Task AResolvingConsumer_OverAnUnknownTable_ReportsTheNodesError() =>
        await AssertConsumers(
            removeTheSheet: false,
            new TableReference("NoSuch", "Valor", TableArea.Data),
            // Oracle (the unknown-name column, both entry modes): #NAME? on all three.
            // Was: #REF! / #REF! / #N/A.
            ("VLOOKUP(1)", ErrorValue.Name),
            ("INDEX", ErrorValue.Name),
            ("MATCH", ErrorValue.Name)
        );

    // The one rule of the closing tests above, over the whole resolving family the sweep had me measure
    // first (VLOOKUP, HLOOKUP, INDEX, MATCH, XMATCH, OFFSET, LOOKUP, FORMULATEXT — every consumer that
    // resolves its own reference argument), over an unknown DEFINED name (Aspose.Cells 26.6.0, 2026-09-11,
    // PLAIN and array-entered agreeing on every row: #NAME? on all eight). Before the rule the resolvers
    // answered their own #REF! (VLOOKUP/HLOOKUP/INDEX/OFFSET) and the scans their not-found #N/A (MATCH/
    // XMATCH/LOOKUP/FORMULATEXT's uniform #VALUE! for a non-reference).
    [Test]
    public async Task TheResolvingFamily_OverAnUnresolvedName_ReportsTheNodesError() =>
        await AssertConsumers(
            removeTheSheet: false,
            new NameReference("NoSuch"),
            ("VLOOKUP(1)", ErrorValue.Name),
            ("HLOOKUP(1)", ErrorValue.Name),
            ("INDEX", ErrorValue.Name),
            ("MATCH", ErrorValue.Name),
            ("XMATCH", ErrorValue.Name),
            ("OFFSET", ErrorValue.Name),
            ("LOOKUP(1)", ErrorValue.Name),
            ("FORMULATEXT", ErrorValue.Name)
        );

    // The same family over the same shape's unknown-TABLE twin (not directly measurable on the oracle:
    // Aspose rejects an unknown-table specifier at formula-set time, so the design's #NAME? — the node's
    // own error, the same value an unknown defined name carries — governs, exactly as it does for SUM and
    // ROWS in StructuredReference_WithAnUnknownTable_IsTheNameErrorValue above).
    [Test]
    public async Task TheResolvingFamily_OverAnUnknownTable_ReportsTheNodesError() =>
        await AssertConsumers(
            removeTheSheet: false,
            new TableReference("NoSuch", "Valor", TableArea.Data),
            ("VLOOKUP(1)", ErrorValue.Name),
            ("HLOOKUP(1)", ErrorValue.Name),
            ("INDEX", ErrorValue.Name),
            ("MATCH", ErrorValue.Name),
            ("XMATCH", ErrorValue.Name),
            ("OFFSET", ErrorValue.Name),
            ("LOOKUP(1)", ErrorValue.Name),
            ("FORMULATEXT", ErrorValue.Name)
        );

    // And over the one unresolvable structured reference the oracle itself accepts at formula-set time:
    // every member reports the node's own #REF! (Aspose.Cells 26.6.0, 2026-09-11, both entry modes).
    // MATCH/XMATCH/LOOKUP answered #N/A here before the rule; VLOOKUP/HLOOKUP/INDEX/OFFSET answered
    // #REF! — their own code, which only coincided with the node's error on this shape.
    [Test]
    public async Task TheResolvingFamily_OverAnUnresolvableStructuredReference_ReportsTheNodesError() =>
        await AssertConsumers(
            removeTheSheet: false,
            TableColumn(area: TableArea.Totals, column: null),
            ("VLOOKUP(1)", ErrorValue.Reference),
            ("HLOOKUP(1)", ErrorValue.Reference),
            ("INDEX", ErrorValue.Reference),
            ("MATCH", ErrorValue.Reference),
            ("XMATCH", ErrorValue.Reference),
            ("OFFSET", ErrorValue.Reference),
            ("LOOKUP(1)", ErrorValue.Reference),
            ("FORMULATEXT", ErrorValue.Reference)
        );

    // The measured exception that keeps the rule honest (Aspose.Cells 26.6.0, 2026-09-11, both entry
    // modes): XLOOKUP does NOT report an unresolved name's #NAME? in its array slots — the oracle answers
    // its own #N/A for the lookup array (and #VALUE! for the return array) where every sibling reports the
    // node's error. Reproducing both XLOOKUP rows would need exactly the per-shape arms the sweep forbids,
    // so MySheet keeps the agreeing #N/A for the name shape. The [#Totals] shape stays a RECORDED
    // divergence for the same reason: the oracle answers the node's #REF! there (but not the #NAME? in the
    // name shape — its own inconsistency), MySheet answers #N/A, and no one rule produces both.
    [Test]
    public async Task XLookup_OverAnUnresolvedName_KeepsItsOwnCode_WhereTheOracleDoesToo() =>
        await AssertConsumers(
            removeTheSheet: false,
            new NameReference("NoSuch"),
            ("XLOOKUP", ErrorValue.NotAvailable)
        );

    // Unknown TABLE. Aspose rejects Tabela1x[Valor] at formula-set time ("Invalid table reference"), so this
    // column is NOT measurable directly; it is the design's #NAME? (a table name lives in the same namespace
    // as a defined name) over the measured shape of the oracle's unknown-NAME column, which is the same value
    // error with a different code: SUM(NoSuch) #NAME?, COUNT 0, COUNTA 1, ROWS #NAME?, SUBTOTAL(9) #NAME?,
    // ISREF FALSE (measured 2026-09-11, both modes).
    [Test]
    public async Task StructuredReference_WithAnUnknownTable_IsTheNameErrorValue() =>
        await AssertConsumers(
            removeTheSheet: false,
            new TableReference("NoSuch", "Valor", TableArea.Data),
            ("SUM", ErrorValue.Name),
            ("COUNT", 0.0),
            ("COUNTA", 1.0),
            ("SUBTOTAL(9)", ErrorValue.Name),
            ("ROWS", ErrorValue.Name),
            ("COLUMNS", ErrorValue.Name),
            ("AVERAGE", ErrorValue.Name),
            ("MAX", ErrorValue.Name),
            ("ISREF", false)
        );

    // Unknown COLUMN: the same shape again, with #REF! - Excel's own repair marker (a deleted column's
    // specifier becomes Table1[#REF!]). Also not measurable directly (Aspose: "Invalid table column").
    [Test]
    public async Task StructuredReference_WithAnUnknownColumn_IsTheRefErrorValue() =>
        await AssertConsumers(
            removeTheSheet: false,
            TableColumn(column: "NoSuch"),
            ("SUM", ErrorValue.Reference),
            ("COUNT", 0.0),
            ("COUNTA", 1.0),
            ("SUBTOTAL(9)", ErrorValue.Reference),
            ("ROWS", ErrorValue.Reference),
            ("ISREF", false)
        );

    // The error-valued table beside a GOOD argument: the error lands on the channel each function already has
    // for an error element, so COUNT still counts the three numbers and COUNTA counts the error as one more
    // element, in either argument order, while SUM propagates. Measured (both modes): COUNT 3, COUNTA 4,
    // SUM(Data!B2:B4,T[#Totals]) #REF! - and COUNT(Tabela1[Valor],NoSuch) 3 on the unknown-name twin.
    [Test]
    public async Task AnUnresolvableStructuredReference_BesideAGoodArgument_OnlyErrorsWhereItShould()
    {
        var totals = TableColumn(area: TableArea.Totals, column: null);
        var valor = TableColumn();
        var range = new RangeReference("B2", "B4", "Data");

        foreach (
            var (node, expected) in new (Expression Node, object? Expected)[]
            {
                (new Count([valor, totals]), 3.0),
                (new Count([totals, valor]), 3.0),
                (new CountA([valor, totals]), 4.0),
                (new Sum([range, totals]), ErrorValue.Reference),
                (new Sum([totals, range]), ErrorValue.Reference),
            }
        )
        {
            var workbook = TableFixture(removeTheSheet: false);
            workbook["Main"]["Z1"] = node;

            await Assert.That(workbook.GetCellValue("Main", "Z1").AsObject()).IsEqualTo(expected);
        }
    }

    // === Controls: the sheet EXISTS — the per-cell value-error policy is UNCHANGED =======================

    [Test]
    public async Task Match_ExistingSheet_NoMatch_IsNA()
    {
        // A whole-column MATCH over the SAME (existing) but empty column is #N/A, NOT #REF!: a missing match
        // over a real sheet is a value outcome, distinct from the structural missing-sheet failure.
        var value = Eval("=MATCH(1, A:A)");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task XLookup_ExistingSheet_NoMatch_IsNA()
    {
        var value = Eval("=XLOOKUP(1, A:A, B:B)");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task SumProduct_ExistingSheet_IsNumber()
    {
        var value = Eval(
            "=SUMPRODUCT(A1:A2, B1:B2)",
            ("A1", "=1"),
            ("A2", "=2"),
            ("B1", "=3"),
            ("B2", "=4")
        );

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(11.0); // 1*3 + 2*4
    }

    [Test]
    public async Task Correl_ExistingSheet_IsNumber()
    {
        var value = Eval(
            "=CORREL(A1:A3, B1:B3)",
            ("A1", "=1"),
            ("A2", "=2"),
            ("A3", "=3"),
            ("B1", "=2"),
            ("B2", "=4"),
            ("B3", "=6")
        );

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(1.0); // perfectly correlated
    }

    // === Controls: the sheet EXISTS — the per-cell value-error policy is UNCHANGED =======================

    [Test]
    public async Task CountIf_ExistingSheet_NoMatch_IsZero()
    {
        // Whole-column COUNTIF over the SAME (existing) sheet with no match is 0, not #REF!.
        var value = Eval("=COUNTIF(D:D,1)", ("D1", "=2"), ("D2", "=3"));

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(0.0);
    }

    [Test]
    public async Task Count_ExistingSheet_IgnoresCellError()
    {
        // COUNT ignores a cell VALUE error (#DIV/0!) on an existing sheet: it still counts the numbers.
        var value = Eval("=COUNT(A:A)", ("A1", "=1/0"), ("A2", "=5"), ("A3", "=7"));

        await Assert.That(value.TryGetNumber(out var number)).IsTrue();
        await Assert.That(number).IsEqualTo(2.0);
    }

    [Test]
    public async Task Sum_ExistingSheet_PropagatesCellError()
    {
        // SUM propagates a cell VALUE error (#DIV/0!) on an existing sheet — unchanged behaviour.
        var value = Eval("=SUM(A:A)", ("A1", "=1/0"), ("A2", "=5"));

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);

        value.TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.DivZero);
    }

    [Test]
    public async Task Cell_ExistingSheet_Empty_CoercesToZero()
    {
        // Control: a reference into an EXISTING sheet's empty cell is NOT #REF! (unlike a missing sheet).
        // Excel-parity update — a formula result is never blank at the CELL boundary, so =Main!A2 (A2 empty)
        // now displays 0 instead of blank. The control's real point (existing sheet ≠ #REF!) is preserved.
        var value = Eval("=Main!A2");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    // === No cell throws: a whole batch with dangling refs resolves, each to #REF! ========================

    [Test]
    public async Task Batch_WithDanglingRefs_NeverThrows()
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");

        main["A1"] = ExpressionParser.Parse("=10", main);
        main["A2"] = ExpressionParser.Parse("=UPPER(BOX11MNO_HIDE!E9)", main);
        main["A3"] = ExpressionParser.Parse("=SUM(Ghost!A:A)", main);
        main["A4"] = ExpressionParser.Parse("=COUNTIF(Ghost!D:D,1)", main);
        main["A5"] = ExpressionParser.Parse("=A1 + 1", main);

        var results = new Dictionary<string, ComputedValue>();

        foreach (var id in new[] { "A1", "A2", "A3", "A4", "A5" })
        {
            // The point of the fix: no cell throws mid-batch.
            results[id] = workbook.GetCellValue("Main", id);
        }

        await Assert.That(results["A1"].TryGetNumber(out var a1)).IsTrue();
        await Assert.That(a1).IsEqualTo(10.0);
        await Assert.That(results["A5"].TryGetNumber(out var a5)).IsTrue();
        await Assert.That(a5).IsEqualTo(11.0);

        foreach (var id in new[] { "A2", "A3", "A4" })
        {
            await Assert.That(results[id].Kind).IsEqualTo(ComputedValueKind.Error);
            results[id].TryGetError(out var error);
            await Assert.That(error).IsEqualTo(Error.Ref);
        }
    }

    // === TryGetSheet (host API) ==========================================================================

    [Test]
    public async Task TryGetSheet_KnownAndUnknown()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Main");

        await Assert.That(workbook.TryGetSheet("Main", out var main)).IsTrue();
        await Assert.That(main!.Name).IsEqualTo("Main");
        // Case-insensitive, like Excel.
        await Assert.That(workbook.TryGetSheet("MAIN", out _)).IsTrue();
        await Assert.That(workbook.TryGetSheet("Ghost", out _)).IsFalse();
    }
}
