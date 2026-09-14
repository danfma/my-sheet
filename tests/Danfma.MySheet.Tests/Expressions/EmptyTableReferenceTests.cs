using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Sweep item 33: a structured reference over a band with ZERO rows — a header-only table's
/// <c>[#Data]</c> — is an EMPTY reference, not <c>#REF!</c>. Every row is read back through a real cell
/// (<see cref="Workbook.GetCellValue(string,string)"/>) on the oracle's own fixture: Main!H20 holds the
/// formula, Main!A1:A3 = 5, 0, 9 and B1:B3 = 1, 2, 3, and Data!Tabela1 = A1:C1 (Item / Valor / Qtd, a
/// header row and nothing else).
///
/// <para>THE ORACLE IS ORDER-DEPENDENT HERE, and which of its answers is binding had to be ruled. Measured
/// on Aspose.Cells 26.6.0 (2026-09-14, one formula per workbook, both entry modes): over a header-only
/// table its value readers depend on whether an EARLIER formula already resolved the same structured
/// reference. "Cold" they read the header and the row below the table (<c>COUNTA(Tabela1[Valor])</c> 1,
/// <c>CONCAT</c> "Valor"); "primed" by a sibling formula they read nothing (<c>COUNTA</c> 0, <c>CONCAT</c>
/// ""), and even primed a value placed directly UNDER the header leaks in (<c>SUM</c> 7). The Aspose-authored
/// <c>f7-header-only.xlsx</c> flips between the two when its own formulas are cleared. The controller's
/// ruling: MySheet follows the order-independent zero-row reading — the PRIMED answers over a fixture with
/// nothing below the table, which are also the answers Phase 5's re-verification recorded. An answer that
/// can be reached from a zero-row reference anchored at (header row + 1, the column) with the function's
/// own semantics is matched (geometry, an explicit resize); an answer that needs the zero-row reference to
/// yield a CELL is an oracle defect, registered below with both numbers and never matched.</para>
/// </summary>
public class EmptyTableReferenceTests
{
    private enum Shape
    {
        HeaderOnly,

        // Data!A2:C2 = "z", 7, 5 directly under the header — OUTSIDE the table, which guards the anchor:
        // geometry must see row 2, value readers must not.
        Sentinel,

        // Data!Tabela1 = A1:C2 with a header row and a totals row (A2 "Total", C2 0) and no data between.
        HeaderAndTotals,
    }

    private static Workbook Fixture(Shape shape)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(5);
        main["A2"] = new NumberValue(0);
        main["A3"] = new NumberValue(9);
        main["B1"] = new NumberValue(1);
        main["B2"] = new NumberValue(2);
        main["B3"] = new NumberValue(3);

        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["C1"] = new StringValue("Qtd");

        if (shape == Shape.Sentinel)
        {
            data["A2"] = new StringValue("z");
            data["B2"] = new NumberValue(7);
            data["C2"] = new NumberValue(5);
        }

        if (shape == Shape.HeaderAndTotals)
        {
            data["A2"] = new StringValue("Total");
            data["C2"] = new NumberValue(0);
        }

        workbook.DefineTable(
            "Tabela1",
            "Data",
            shape == Shape.HeaderAndTotals ? "A1:C2" : "A1:C1",
            ["Item", "Valor", "Qtd"],
            hasHeaderRow: true,
            hasTotalsRow: shape == Shape.HeaderAndTotals
        );

        return workbook;
    }

    // Stores the formula in a real cell of a pristine fixture and reads it back, so the value is the one a
    // host sees (the cell boundary's implicit intersection included).
    private static string InCell(
        string formula,
        Shape shape = Shape.HeaderOnly,
        string sheetName = "Main",
        string id = "H20"
    )
    {
        var workbook = Fixture(shape);
        var sheet = workbook[sheetName];
        sheet[id] = ExpressionParser.Parse(formula, sheet);

        return Display(workbook.GetCellValue(sheetName, id));
    }

    private static string Display(ComputedValue value)
    {
        if (value.TryGetError(out var error))
        {
            return error.ToString();
        }

        if (value.TryGetNumber(out var number))
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        if (value.TryGetBoolean(out var boolean))
        {
            return boolean ? "TRUE" : "FALSE";
        }

        return value.TryGetText(out var text) ? $"\"{text}\"" : value.Kind.ToString();
    }

    // === The matched rows: PLAIN and array-entered agree on the primed oracle =============================
    //
    // Every row answers the same in both entry modes on the oracle (primed, nothing below the table). Before
    // this item every row that reads the data band answered #REF! (ruling R1's recorded divergence): SUM,
    // ROWS, COLUMNS, AREAS, AVERAGE, the lookups, the criteria family, SUBTOTAL, SUMPRODUCT, FILTER, ROW and
    // COLUMN #REF!, ISREF FALSE, COUNTA 1 (the error counted as one element), ERROR.TYPE(SUM(…)) 4.
    // The criteria family's 0 falls out of the empty stream, not a special case. Rows whose answer depends on
    // a PRE-EXISTING general gap (one that already diverges over an ordinary range) are pinned separately in
    // TheRowsThatDependOnAnOrdinaryRangeGap, with both numbers.
    [Test]
    [Arguments("=SUM(Tabela1[Valor])", "0")]
    [Arguments("=COUNT(Tabela1[Valor])", "0")]
    [Arguments("=COUNTA(Tabela1[Valor])", "0")]
    [Arguments("=ROWS(Tabela1[Valor])", "0")]
    [Arguments("=COLUMNS(Tabela1[Valor])", "1")]
    [Arguments("=AREAS(Tabela1[Valor])", "1")]
    [Arguments("=ISREF(Tabela1[Valor])", "TRUE")]
    [Arguments("=ERROR.TYPE(SUM(Tabela1[Valor]))", "#N/A")]
    [Arguments("=AVERAGE(Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=MIN(Tabela1[Valor])", "0")]
    [Arguments("=MAX(Tabela1[Valor])", "0")]
    [Arguments("=SMALL(Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=LARGE(Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=MEDIAN(Tabela1[Valor])", "#NUM!")]
    [Arguments("=INDEX(Tabela1[Valor],1,1)", "#REF!")]
    [Arguments("=INDEX(Tabela1[Valor],1)", "#REF!")]
    [Arguments("=INDEX(Tabela1[#Data],1,2)", "#REF!")]
    [Arguments("=MATCH(1,Tabela1[Valor],0)", "#N/A")]
    [Arguments("=MATCH(1,Tabela1[Valor])", "#N/A")]
    [Arguments("=MATCH(1,Tabela1[Valor],-1)", "#N/A")]
    [Arguments("=MATCH(\"x\",Tabela1[Valor],0)", "#N/A")]
    [Arguments("=XMATCH(1,Tabela1[Valor])", "#N/A")]
    [Arguments("=VLOOKUP(1,Tabela1[Valor],1,FALSE)", "#N/A")]
    [Arguments("=VLOOKUP(1,Tabela1[Valor],1)", "#N/A")]
    [Arguments("=VLOOKUP(1,Tabela1[#Data],2,FALSE)", "#N/A")]
    // Array-entered #N/A; typed PLAIN the oracle throws a CellsException out of CalculateFormula.
    [Arguments("=HLOOKUP(1,Tabela1[Valor],1,FALSE)", "#N/A")]
    [Arguments("=LOOKUP(1,Tabela1[Valor])", "#N/A")]
    [Arguments("=LOOKUP(1,Tabela1[Valor],Tabela1[Qtd])", "#N/A")]
    [Arguments("=XLOOKUP(1,Tabela1[Valor],Tabela1[Qtd])", "#N/A")]
    [Arguments("=XLOOKUP(1,Tabela1[Valor],Tabela1[Qtd],\"nf\")", "\"nf\"")]
    [Arguments("=COUNTIF(Tabela1[Valor],\">0\")", "0")]
    [Arguments("=COUNTIF(Tabela1[Valor],\"\")", "0")]
    // Final-review fix wave, minor M2: the claim that used to sit here — that a stream of one error element
    // (what the reference's own #VALUE! would be if a consumer ever read it) "matches '<>' and '#VALUE!', so
    // these would answer 1" — was FALSE, measured: over a fixture holding a genuine error cell, an error
    // element matches NEITHER criterion (COUNTIF over it with "<>" counts it as 2 of 3, unchanged whether the
    // third slot is empty or an error; with "#DIV/0!" it counts 0 either way). The row above,
    // COUNTIF(Tabela1[Valor],"") = 0, is the ONE live guard: an error element DOES match the empty criterion
    // (a genuinely empty stream never would), which is the row the M2 mutation (an EMPTY-stream arm forced to
    // yield the reference's own #VALUE! as one element) actually turns red. The two rows below do not guard
    // anything — they would answer 0 whichever way the mutation went — and only pin the ordinary answer (the
    // oracle answers 0 / 0 for both, primed; the sentinel row's "<>" is 1, an unrelated, unfixed defect).
    [Arguments("=COUNTIF(Tabela1[Valor],\"<>\")", "0")]
    [Arguments("=COUNTIF(Tabela1[Valor],\"#VALUE!\")", "0")]
    [Arguments("=COUNTIFS(Tabela1[Valor],\"<>\")", "0")]
    [Arguments("=SUMIF(Tabela1[Valor],\">0\")", "0")]
    [Arguments("=SUMIF(Tabela1[Valor],\">0\",Tabela1[Qtd])", "0")]
    [Arguments("=SUMIF(A1:A3,\">0\",Tabela1[Valor])", "0")]
    [Arguments("=AVERAGEIF(Tabela1[Valor],\">0\")", "#DIV/0!")]
    [Arguments("=COUNTIFS(Tabela1[Valor],\">0\")", "0")]
    [Arguments("=COUNTIFS(Tabela1[Valor],\">0\",Tabela1[Qtd],\">0\")", "0")]
    [Arguments("=COUNTIFS(Tabela1[Valor],\">0\",A1:A3,\">0\")", "#VALUE!")]
    [Arguments("=SUMIFS(Tabela1[Qtd],Tabela1[Valor],\">0\")", "0")]
    [Arguments("=AVERAGEIFS(Tabela1[Qtd],Tabela1[Valor],\">0\")", "#DIV/0!")]
    [Arguments("=MAXIFS(Tabela1[Qtd],Tabela1[Valor],\">0\")", "0")]
    [Arguments("=MINIFS(Tabela1[Qtd],Tabela1[Valor],\">0\")", "0")]
    [Arguments("=COUNTBLANK(Tabela1[Valor])", "0")]
    [Arguments("=COUNTIF(Tabela1[#Data],\">0\")", "0")]
    [Arguments("=COUNTBLANK(Tabela1[#Data])", "0")]
    [Arguments("=SUBTOTAL(9,Tabela1[Valor])", "0")]
    [Arguments("=SUBTOTAL(3,Tabela1[Valor])", "0")]
    [Arguments("=SUBTOTAL(1,Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=SUBTOTAL(109,Tabela1[Valor])", "0")]
    [Arguments("=AGGREGATE(9,4,Tabela1[Valor])", "0")]
    [Arguments("=AGGREGATE(14,6,Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=SUMPRODUCT(Tabela1[Valor])", "0")]
    // Phase 5 recorded "#VALUE! plain / 0 CSE" for this row; that split is EVALUATION ORDER in its
    // one-workbook batch (swapping the two cells swaps the answers), and primed both modes are 0.
    [Arguments("=SUMPRODUCT(Tabela1[Valor],Tabela1[Qtd])", "0")]
    [Arguments("=SUMPRODUCT(Tabela1[Valor]*Tabela1[Qtd])", "0")]
    [Arguments("=ROWS(Tabela1[Valor]*2)", "0")]
    [Arguments("=COUNT((Tabela1[Valor]<>\"\")*1)", "0")]
    [Arguments("=ROWS(Tabela1[Valor]+1)", "0")]
    [Arguments("=SUM(OFFSET(Tabela1[Valor],0,0,1,1))", "0")]
    [Arguments("=COLUMNS(OFFSET(Tabela1[Valor],0,0))", "1")]
    [Arguments("=ROWS(Tabela1[Valor]:Tabela1[Qtd])", "0")]
    [Arguments("=COLUMNS(Tabela1[Valor]:Tabela1[Qtd])", "2")]
    [Arguments("=SUM(Tabela1[Valor]:Tabela1[Qtd])", "0")]
    [Arguments("=SUM(Tabela1)", "0")]
    [Arguments("=ROWS(Tabela1)", "0")]
    [Arguments("=ISREF(Tabela1)", "TRUE")]
    [Arguments("=SUM(Tabela1[#Data])", "0")]
    [Arguments("=ROWS(Tabela1[#Data])", "0")]
    [Arguments("=COLUMNS(Tabela1[#Data])", "3")]
    [Arguments("=ROWS(Tabela1[[#Headers],[#Data]])", "1")]
    [Arguments("=COUNTA(Tabela1[[#Headers],[#Data]])", "3")]
    [Arguments("=ROWS(Tabela1[[#Data],[#Totals]])", "0")]
    [Arguments("=SUM(Tabela1[[#Data],[#Totals]])", "0")]
    [Arguments("=ROWS(Tabela1[#All])", "1")]
    [Arguments("=COUNTA(Tabela1[#All])", "3")]
    [Arguments("=SUM(LET(x,Tabela1[Valor],x))", "0")]
    [Arguments("=ROWS(LET(x,Tabela1[Valor],x*1))", "0")]
    [Arguments("=SUM(IF(TRUE,Tabela1[Valor],0))", "0")]
    [Arguments("=ROWS(IF(TRUE,Tabela1[Valor],0))", "0")]
    [Arguments("=SUM(CHOOSE(1,Tabela1[Valor]))", "0")]
    [Arguments("=ROWS(CHOOSE(1,Tabela1[Valor]))", "0")]
    [Arguments("=SUM(INDIRECT(\"Tabela1[Valor]\"))", "0")]
    [Arguments("=ISNUMBER(Tabela1[Valor])", "FALSE")]
    [Arguments("=SUM(IFERROR(Tabela1[Valor],0))", "0")]
    [Arguments("=COLUMN(Tabela1[Valor])", "2")]
    [Arguments("=COLUMN(Tabela1[#Data])", "1")]
    [Arguments("=FILTER(Tabela1[Valor],Tabela1[Valor]>0)", "#CALC!")]
    [Arguments("=ROWS(FILTER(Tabela1[Valor],Tabela1[Valor]>0))", "#CALC!")]
    [Arguments("=FILTER(Tabela1[Valor],Tabela1[Valor]>0,\"none\")", "\"none\"")]
    [Arguments("=PRODUCT(Tabela1[Valor])", "0")]
    [Arguments("=STDEV(Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=VAR.P(Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=MODE(Tabela1[Valor])", "#N/A")]
    [Arguments("=RANK(1,Tabela1[Valor])", "#N/A")]
    [Arguments("=PERCENTILE(Tabela1[Valor],0.5)", "#NUM!")]
    [Arguments("=QUARTILE(Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=AVERAGEA(Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=MAXA(Tabela1[Valor])", "0")]
    [Arguments("=SUMSQ(Tabela1[Valor])", "0")]
    [Arguments("=CONCAT(Tabela1[Valor])", "\"\"")]
    [Arguments("=TEXTJOIN(\",\",TRUE,Tabela1[Valor])", "\"\"")]
    [Arguments("=AND(Tabela1[Valor])", "#VALUE!")]
    [Arguments("=OR(Tabela1[Valor])", "#VALUE!")]
    [Arguments("=CORREL(Tabela1[Valor],Tabela1[Qtd])", "#DIV/0!")]
    [Arguments("=ISFORMULA(Tabela1[Valor])", "FALSE")]
    [Arguments("=FORMULATEXT(Tabela1[Valor])", "#N/A")]
    [Arguments("=SUM(Tabela1[Valor],5)", "5")]
    [Arguments("=COUNTA(Tabela1[Valor],1)", "1")]
    [Arguments("=MAX(Tabela1[Valor],-3)", "-3")]
    [Arguments("=AVERAGE(Tabela1[Valor],4)", "4")]
    // Sweep item 37 (was pinned #REF! in TheRowsThatDependOnAnOrdinaryRangeGap_KeepTheEnginesAnswer, an
    // inherited gap; INDEX with a zero row/column now returns the empty band ITSELF as a reference, so the
    // usual empty-reference readings apply — a narrower zero-row band for a multi-column [#Data]).
    [Arguments("=SUM(INDEX(Tabela1[Valor],0,1))", "0")]
    [Arguments("=ROWS(INDEX(Tabela1[#Data],0,1))", "0")]
    public async Task OverAHeaderOnlyTable_TheDataBandIsAnEmptyReference(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === The rows the oracle SPLITS by entry mode =========================================================
    //
    // MySheet has one entry mode, so each row takes the reading its existing convention gives a NON-empty
    // reference of the same shape (the controller's Q3 ruling: no empty-only special case), and the other
    // oracle column is recorded here. Measured on the oracle, primed, PLAIN / CSE:
    //   SUM(T*2), SUM((T<>"")*1), SUM(T+T[Qtd]), SUM(-T)   #VALUE! / 0  — the mini-CSE lift is the CSE reading
    //   =Tabela1[Valor], =Tabela1                          #VALUE! / 0  — the cell boundary intersects (plain)
    //   ROW(T[Valor]), ROW(T[#Data]), ROW(Tabela1)          2 / 1        — the top row of the reference; the CSE 1
    //                                                                      is the header row, which only a
    //                                                                      normalized inverted rectangle reads
    //   ISERROR(T[Valor])                                   TRUE / FALSE — MySheet evaluates a table reference to
    //                                                                      its reference VALUE (non-empty table:
    //                                                                      oracle TRUE / FALSE, MySheet FALSE)
    //   ISBLANK(T[Valor])                                   FALSE / TRUE — non-empty table: oracle FALSE / FALSE,
    //                                                                      MySheet FALSE
    //   N(T[Valor])                                         #VALUE! / 0  — non-empty table: oracle #VALUE! / 10,
    //                                                                      MySheet 0 (neither column, pre-existing)
    //   IFERROR(T[Valor],"e")                               "e" / 0      — non-empty table: oracle "e" / 10,
    //                                                                      MySheet #VALUE! (neither, pre-existing)
    [Test]
    [Arguments("=SUM(Tabela1[Valor]*2)", "0")]
    [Arguments("=SUM((Tabela1[Valor]<>\"\")*1)", "0")]
    [Arguments("=SUM(Tabela1[Valor]+Tabela1[Qtd])", "0")]
    [Arguments("=SUM(-Tabela1[Valor])", "0")]
    [Arguments("=Tabela1[Valor]", "#VALUE!")]
    [Arguments("=Tabela1", "#VALUE!")]
    [Arguments("=ROW(Tabela1[Valor])", "2")]
    [Arguments("=ROW(Tabela1[#Data])", "2")]
    [Arguments("=ROW(Tabela1)", "2")]
    [Arguments("=ISERROR(Tabela1[Valor])", "FALSE")]
    [Arguments("=ISBLANK(Tabela1[Valor])", "FALSE")]
    [Arguments("=N(Tabela1[Valor])", "0")]
    [Arguments("=IFERROR(Tabela1[Valor],\"e\")", "#VALUE!")]
    public async Task TheEntryModeSplitRows_FollowTheEnginesConventionForTheShape(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === The anchor, guarded by a value under the header ==================================================
    //
    // The zero-row reference sits at (header row + 1, the column): geometry reports that anchor and an
    // explicit resize from it reads the real cell there, while nothing that READS the reference ever sees it.
    // Oracle, primed, sentinel fixture: ROWS 0, COLUMNS 1, AREAS 1, ISREF TRUE, ROW 2 (plain; CSE 1),
    // COLUMN 2, SUM(OFFSET(T,0,0,1,1)) 7, ROWS(T[Valor]:T[Qtd]) 0, COLUMNS 2, ROWS(T[#Data]) 0,
    // COLUMNS(T[#Data]) 3, INDEX(T,1,1) #REF! — both modes except ROW.
    [Test]
    [Arguments("=ROWS(Tabela1[Valor])", "0")]
    [Arguments("=COLUMNS(Tabela1[Valor])", "1")]
    [Arguments("=AREAS(Tabela1[Valor])", "1")]
    [Arguments("=ISREF(Tabela1[Valor])", "TRUE")]
    [Arguments("=ROW(Tabela1[Valor])", "2")]
    [Arguments("=COLUMN(Tabela1[Valor])", "2")]
    [Arguments("=SUM(OFFSET(Tabela1[Valor],0,0,1,1))", "7")]
    [Arguments("=ROWS(Tabela1[Valor]:Tabela1[Qtd])", "0")]
    [Arguments("=COLUMNS(Tabela1[Valor]:Tabela1[Qtd])", "2")]
    [Arguments("=ROWS(Tabela1[#Data])", "0")]
    [Arguments("=COLUMNS(Tabela1[#Data])", "3")]
    [Arguments("=INDEX(Tabela1[Valor],1,1)", "#REF!")]
    // SUMIF resizes the zero-row sum range from its B2 anchor to A1:A3's 3x1 shape. This was 0 before
    // item 35; Aspose.Cells 26.7.0 gives 7 in both modes on the sentinel fixture.
    [Arguments("=SUMIF(A1:A3,\">0\",Tabela1[Valor])", "7")]
    public async Task TheEmptyReferenceIsAnchoredUnderTheHeader(string formula, string expected)
    {
        await Assert.That(InCell(formula, Shape.Sentinel)).IsEqualTo(expected);
    }

    // === Oracle defects: the zero-row reference never yields a cell =======================================
    //
    // Registered, never matched (controller ruling): each oracle answer below needs the zero-row reference
    // to yield a cell, and it changes with evaluation order. Oracle numbers, sentinel fixture (Data!B2 = 7,
    // C2 = 5, A2 "z"), PRIMED / COLD, both entry modes unless noted:
    //   SUM 7 / 7, COUNT 1 / 1, COUNTA 1 / 2, AVERAGE 7 / 7, MIN 7 / 7, SMALL(…,1) 7 / 7, SUBTOTAL(9) 7 / 7,
    //   SUMPRODUCT(T[Valor],T[Qtd]) 35 / 35, SUM(T[Valor]:T[Qtd]) 12 / 12, SUM(Tabela1) 12 / 12,
    //   COUNTIF(T[#Data],">0") 2 / 2, COUNTBLANK(T[#Data]) -3 / 0, MATCH(1,T,-1) 1 / 2,
    //   FILTER(T,T>0) 7 / #VALUE!, SUM(IF(TRUE,T,0)) 7 / 7, SUM(LET(x,T,x)) 7 / 7,
    //   SUM(INDIRECT("Tabela1[Valor]")) 7 / 7, CONCAT "7" / "Valor7", and two top-left readings:
    //   =+Tabela1[Valor] (primed: #VALUE! plain / 7 CSE) — a '+' over a bare reference keeps the reference,
    //   so the cell boundary intersects it and a zero-row reference has no row to intersect: #VALUE!, the
    //   plain column (this pin was written red-first as #N/A, a wrong prediction of the path, and corrected
    //   to the engine's reference convention before the commit) — and =LET(x,Tabela1[Valor]*2,x) (primed:
    //   #VALUE! plain / 14 CSE, core fixture #VALUE! / 0): an ARRAY binding reads its top-left, and an
    //   empty array has none, so the position is uncovered (#N/A, the broadcasting rule) — never the anchor
    //   cell's 7 doubled.
    [Test]
    [Arguments("=SUM(Tabela1[Valor])", "0")]
    [Arguments("=COUNT(Tabela1[Valor])", "0")]
    [Arguments("=COUNTA(Tabela1[Valor])", "0")]
    [Arguments("=AVERAGE(Tabela1[Valor])", "#DIV/0!")]
    [Arguments("=MIN(Tabela1[Valor])", "0")]
    [Arguments("=SMALL(Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=SUBTOTAL(9,Tabela1[Valor])", "0")]
    [Arguments("=SUMPRODUCT(Tabela1[Valor],Tabela1[Qtd])", "0")]
    [Arguments("=SUM(Tabela1[Valor]:Tabela1[Qtd])", "0")]
    [Arguments("=SUM(Tabela1)", "0")]
    [Arguments("=COUNTIF(Tabela1[#Data],\">0\")", "0")]
    [Arguments("=COUNTBLANK(Tabela1[#Data])", "0")]
    [Arguments("=MATCH(1,Tabela1[Valor],-1)", "#N/A")]
    [Arguments("=FILTER(Tabela1[Valor],Tabela1[Valor]>0)", "#CALC!")]
    [Arguments("=SUM(IF(TRUE,Tabela1[Valor],0))", "0")]
    [Arguments("=SUM(LET(x,Tabela1[Valor],x))", "0")]
    [Arguments("=SUM(INDIRECT(\"Tabela1[Valor]\"))", "0")]
    [Arguments("=CONCAT(Tabela1[Valor])", "\"\"")]
    [Arguments("=+Tabela1[Valor]", "#VALUE!")]
    [Arguments("=LET(x,Tabela1[Valor]*2,x)", "#N/A")]
    public async Task TheZeroRowReference_NeverReadsTheRowUnderTheHeader(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula, Shape.Sentinel)).IsEqualTo(expected);
    }

    // The one defect row that needs no sentinel: the oracle answers ROWS(INDIRECT("Tabela1[Valor]")) 2 in
    // every context measured (cold, primed, the loaded f7 file), while ROWS(Tabela1[Valor]) is 0. MySheet's
    // INDIRECT parses the text into the same structured reference, so it resolves to the same empty
    // reference.
    [Test]
    public async Task Indirect_SharesTheStructuredReferencesResolution()
    {
        await Assert.That(InCell("=ROWS(INDIRECT(\"Tabela1[Valor]\"))")).IsEqualTo("0");
    }

    // === A producer over the empty reference ==============================================================
    //
    // MySheet's producer contract (ArrayShaping.RequireProducerShape): a producer's result is never a
    // zero-extent array — an empty result is the 1x1 #CALC! singleton, Excel's empty-array error, which is
    // what FILTER and UNIQUE already answer when nothing survives. SORT over a zero-row source is that empty
    // result too (it used to THROW the contract's InvalidOperationException: SORT had never met an empty
    // source before). Oracle, primed, core fixture PLAIN / CSE (sentinel in brackets):
    //   SORT(T[Valor])          #VALUE! / 0   [#VALUE! / 7]  — the CSE reads the cell under the header
    //   SUM(SORT(T[Valor]))     0 / 0         [7 / 7]
    //   ROWS(SORT(T[Valor]))    0 / 0         — a zero-row RESULT, which the producer contract rules out
    //   UNIQUE(T[Valor])        0 / 0         [7 / 7]
    //   ROWS(UNIQUE(T[Valor]))  1 / 1         — a one-row result read from the cell under the header
    [Test]
    [Arguments("=SORT(Tabela1[Valor])", "#CALC!")]
    [Arguments("=SUM(SORT(Tabela1[Valor]))", "#CALC!")]
    [Arguments("=ROWS(SORT(Tabela1[Valor]))", "#CALC!")]
    [Arguments("=UNIQUE(Tabela1[Valor])", "#CALC!")]
    [Arguments("=ROWS(UNIQUE(Tabela1[Valor]))", "#CALC!")]
    public async Task AProducerOverTheEmptyReference_IsTheEmptyResultError(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === Rows that depend on an ordinary-range gap ========================================================
    //
    // Out of item 33's scope by the controller's ruling: each row below follows a MySheet behaviour that
    // already diverges from the oracle over an ORDINARY range, so the empty reference inherits the engine's
    // answer and the gap is registered as its own sweep item. Oracle PLAIN / CSE, primed (the ordinary-range
    // twin that shows the gap, oracle vs MySheet, in brackets):
    //   OFFSET's omitted height/width are 1, not the base's size
    //                                    [ROWS(OFFSET(A1:A3,0,0)) 3 / 3 vs 1; SUM 14 / 14 vs 5]
    //     ROWS(OFFSET(T[Valor],0,0)) 0 / 0; ROWS(OFFSET(T[Valor],1,0)) 0 / 0; sentinel SUM(OFFSET(T,0,0)) 0 / 0
    //   XLOOKUP does not check the arrays' sizes
    //                                    [XLOOKUP(5,A1:A3,B1:B2) #VALUE! / #VALUE! vs 1]
    //     XLOOKUP(5,A1:A3,T[Valor]) #VALUE! / #VALUE!; with "nf" #VALUE! / #VALUE!
    //   ROWS of a LET whose body is a bound bare reference falls back to 1
    //                                    [ROWS(LET(x,A1:A3,x)) 3 / 3 vs 1]
    //     ROWS(LET(x,T[Valor],x)) 0 / 0
    //   (sweep item 37 closed "INDEX with row 0 is #REF!" — the two rows this bullet used to list moved to
    //   OverAHeaderOnlyTable_TheDataBandIsAnEmptyReference, now matched: SUM(INDEX(T[Valor],0,1)) 0,
    //   ROWS(INDEX(T[#Data],0,1)) 0)
    [Test]
    [Arguments("=ROWS(OFFSET(Tabela1[Valor],0,0))", false, "1")]
    [Arguments("=ROWS(OFFSET(Tabela1[Valor],1,0))", false, "1")]
    [Arguments("=SUM(OFFSET(Tabela1[Valor],0,0))", true, "7")]
    [Arguments("=XLOOKUP(5,A1:A3,Tabela1[Valor])", false, "#N/A")]
    [Arguments("=XLOOKUP(5,A1:A3,Tabela1[Valor],\"nf\")", false, "\"nf\"")]
    [Arguments("=ROWS(LET(x,Tabela1[Valor],x))", false, "1")]
    public async Task TheRowsThatDependOnAnOrdinaryRangeGap_KeepTheEnginesAnswer(
        string formula,
        bool sentinel,
        string expected
    )
    {
        await Assert
            .That(InCell(formula, sentinel ? Shape.Sentinel : Shape.HeaderOnly))
            .IsEqualTo(expected);
    }

    // === A selection over a zero-row source (fix wave: BUG 1 and BUG 2) ===================================
    //
    // A selection along the COLUMNS axis of a zero-row source — FILTER with a row include, UNIQUE/SORT with
    // by_col — keeps its columns and the source's zero rows. Before the fix every such row THREW
    // InvalidOperationException out of Evaluate (the producer-shape guard rejected the 0-row result); at
    // c6fde1f they were #REF!. The oracle keeps the zero-row result (primed, both modes; sentinel in brackets):
    //   FILTER(T[Valor],TRUE)                               0 [7]         ROWS 0 [1]    SUM 0 [7]
    //   FILTER(T[#Data],T[#Headers]<>"Item")                0 [7]         ROWS 0 [1]    COLUMNS 2 [2]
    //     COUNT 0 [2], INDEX(…,1,1) #REF! [7], SUM(…*1) 0 [12]
    //   FILTER(FILTER(T[#Data],…<>"Item"),TRUE)             ROWS #VALUE! (a 1x1 include against a 0x2 source)
    //   FILTER(T[#Data],T[#Headers]="none")                 ROWS #CALC!; with "e" → "e"
    //   UNIQUE(T[Valor],TRUE)                               0 [7]         ROWS 0 [1]
    //   UNIQUE(T[#Data],TRUE)                               THREW inside Aspose ["z"]; COLUMNS 1 [3], ROWS 0 [1]
    //   UNIQUE(T[#Data],TRUE,TRUE)                          THREW inside Aspose
    //   COLUMNS(UNIQUE(FILTER(T[#Data],…<>"Item"),TRUE))    2 — the oracle contradicts its own direct row (1):
    //                                                       zero-row columns have equal (empty) keys, so 1
    //   SORT(FILTER(T[#Data],…<>"Item"))                    0, ROWS 0 — MySheet: an empty SELECTION is the
    //                                                       #CALC! singleton, as for SORT(T[Valor]) above
    //   SORT(T[#Data],1,1,TRUE)                             #VALUE! (sort_index past a zero-row cross extent)
    //   SEQUENCE(ROWS(T[Valor]))                            #VALUE! (a zero size)
    // A cell showing a zero-row result has no top-left: #N/A, the uncovered-position rule (the oracle's 0 / 7
    // there reads the cell under the header, a registered defect).
    [Test]
    [Arguments("=FILTER(Tabela1[Valor],TRUE)", "#N/A")]
    [Arguments("=ROWS(FILTER(Tabela1[Valor],TRUE))", "0")]
    [Arguments("=SUM(FILTER(Tabela1[Valor],TRUE))", "0")]
    [Arguments("=FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\")", "#N/A")]
    [Arguments("=ROWS(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"))", "0")]
    [Arguments("=COLUMNS(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"))", "2")]
    [Arguments("=COUNT(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"))", "0")]
    [Arguments("=INDEX(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"),1,1)", "#REF!")]
    [Arguments("=SUM(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\")*1)", "0")]
    [Arguments("=ROWS(FILTER(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"),TRUE))", "#VALUE!")]
    [Arguments("=ROWS(FILTER(Tabela1[#Data],Tabela1[#Headers]=\"none\"))", "#CALC!")]
    [Arguments("=FILTER(Tabela1[#Data],Tabela1[#Headers]=\"none\",\"e\")", "\"e\"")]
    [Arguments("=UNIQUE(Tabela1[Valor],TRUE)", "#N/A")]
    [Arguments("=ROWS(UNIQUE(Tabela1[Valor],TRUE))", "0")]
    [Arguments("=UNIQUE(Tabela1[#Data],TRUE)", "#N/A")]
    [Arguments("=COLUMNS(UNIQUE(Tabela1[#Data],TRUE))", "1")]
    [Arguments("=ROWS(UNIQUE(Tabela1[#Data],TRUE))", "0")]
    [Arguments("=UNIQUE(Tabela1[#Data],TRUE,TRUE)", "#CALC!")]
    [Arguments("=COLUMNS(UNIQUE(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"),TRUE))", "1")]
    [Arguments("=SORT(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"))", "#CALC!")]
    [Arguments("=SORT(Tabela1[#Data],1,1,TRUE)", "#VALUE!")]
    [Arguments("=SEQUENCE(ROWS(Tabela1[Valor]))", "#VALUE!")]
    public async Task ASelectionOverAZeroRowSource_NeverThrows_AndKeepsTheZeroRows(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // The same selections with a value under the header: the zero-row result never reaches it (the oracle's
    // 7 / 1 / 12 there are the registered sentinel defect).
    [Test]
    [Arguments("=SUM(FILTER(Tabela1[Valor],TRUE))", "0")]
    [Arguments("=ROWS(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\"))", "0")]
    [Arguments("=SUM(FILTER(Tabela1[#Data],Tabela1[#Headers]<>\"Item\")*1)", "0")]
    [Arguments("=FILTER(Tabela1[Valor],TRUE)", "#N/A")]
    public async Task ASelectionOverAZeroRowSource_NeverReadsTheRowUnderTheHeader(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula, Shape.Sentinel)).IsEqualTo(expected);
    }

    // === SUMPRODUCT compares a zero-row shape (fix wave: RISK 1) ==========================================
    //
    // A zero-row rectangle HAS a shape (0 x its columns); it is not the "no shape known" of a name or a
    // scalar. Oracle, primed, both modes: SUMPRODUCT(T[Valor]*1,T[#Data]*1) and SUMPRODUCT(T[Valor],T[#Data])
    // #VALUE! (0x1 against 0x3, like SUMPRODUCT(A1:A3*1,A1:B3*1) #VALUE!), SUMPRODUCT(T[Valor]*1,T[Qtd]*1)
    // and SUMPRODUCT(T[#Data]*1,T[#Data]*1) 0. Before the fix MySheet answered 0 for all four.
    [Test]
    [Arguments("=SUMPRODUCT(Tabela1[Valor]*1,Tabela1[#Data]*1)", "#VALUE!")]
    [Arguments("=SUMPRODUCT(Tabela1[Valor],Tabela1[#Data])", "#VALUE!")]
    [Arguments("=SUMPRODUCT(Tabela1[Valor]*1,Tabela1[Qtd]*1)", "0")]
    [Arguments("=SUMPRODUCT(Tabela1[#Data]*1,Tabela1[#Data]*1)", "0")]
    public async Task SumProduct_ComparesAZeroRowShape(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === The top-left of an empty binding (fix wave: NIT 4) ===============================================
    //
    // Not made consistent, because the oracle does not ask for it (primed, PLAIN / CSE, sentinel identical
    // for ROW): LET(x,ROW(T[Valor]),x) 2 / 1 — ROW's value is the reference's anchor row, a number whatever
    // the height, and MySheet's 2 is the plain column (the CSE 1 is the header row) — while
    // LET(x,T[Valor]*2,x) #VALUE! / 0 [#VALUE! / 14] has no element to read, so MySheet's top-left of the
    // empty operator array is #N/A (the CSE 0 / 14 reads the cell under the header, a registered defect).
    // SUM(LET(x,ROW(T[Valor]),x)) is 3 / 3 on the oracle — the row-number vector of the normalized inverted
    // rectangle {1,2}, a header read — and 0 here: the empty band has no row number to sum.
    [Test]
    [Arguments("=LET(x,ROW(Tabela1[Valor]),x)", "2")]
    [Arguments("=LET(x,Tabela1[Valor]*2,x)", "#N/A")]
    [Arguments("=SUM(LET(x,ROW(Tabela1[Valor]),x))", "0")]
    public async Task TheTopLeftOfAnEmptyBinding_FollowsWhatTheBindingHolds(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Final-review fix wave, minor M3 (Important, unpinned before this): the BARE (non-LET) forms of the same
    // shape as the NIT 4 pin above. =ROW(Tabela1[Valor]) alone is 2 (the anchor row, pinned elsewhere in this
    // file); by the ruling's own classification ("an answer reachable from the zero-row reference at its
    // anchor with the function's own semantics is MATCHED") that 2 is not a header read, so MAX/SUM of it
    // should carry through — but the branch answers 0 for both. This is a DIFFERENT shape from the LET one
    // above (SUM(LET(x,ROW(T),x)) 0, oracle 3/3, the inverted rectangle's header-row vector {1,2}): MAX/SUM
    // applied directly never go through a LET binding's array-eligible capture. Oracle (core fixture, primed):
    // MAX(ROW(Tabela1[Valor])) 2 in all four columns; SUM(ROW(Tabela1[Valor])) 2 PLAIN / 3 CSE. Registered for
    // the controller; not fixed here.
    [Test]
    [Arguments("=MAX(ROW(Tabela1[Valor]))", "0")]
    [Arguments("=SUM(ROW(Tabela1[Valor]))", "0")]
    public async Task ABareAggregateOfRowOverAnEmptyBand_IsZero_ARecordedDivergence(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Final-review fix wave, minor M4: unmeasured folds over the empty band, registered — no code changes.
    // None of the five crashes; none was in the fix wave's brief's own matrix. Oracle (core fixture, PLAIN =
    // CSE, primed): DEVSQ(Tabela1[Valor]) 0, XNPV(0.1,Tabela1[Valor],Tabela1[Qtd]) 0,
    // PERCENTRANK(Tabela1[Valor],1) #N/A, SUMX2MY2(Tabela1[Valor],Tabela1[Qtd]) #DIV/0!, SHEET(Tabela1[Valor])
    // 2 (pre-existing: SHEET does not resolve a structured reference over a NON-empty table on main either).
    [Test]
    [Arguments("=DEVSQ(Tabela1[Valor])", "#NUM!")]
    [Arguments("=XNPV(0.1,Tabela1[Valor],Tabela1[Qtd])", "#NUM!")]
    [Arguments("=PERCENTRANK(Tabela1[Valor],1)", "#NUM!")]
    [Arguments("=SUMX2MY2(Tabela1[Valor],Tabela1[Qtd])", "0")]
    [Arguments("=SHEET(Tabela1[Valor])", "#REF!")]
    public async Task AnUnmeasuredFoldOverAnEmptyBand_DoesNotCrash_ARecordedDivergence(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === A header-only table that grows (fix wave: RISK 2) ================================================
    //
    // A structured reference over an empty band has no cell dependency (DependencyExtractor keeps it
    // always-dirty), and a redefinition bumps the definitions version, so a dependent recomputes when the
    // table grows — through the recalculation engine (a full fallback) and through InvalidateCache (a
    // definition alone evicts nothing, so the memo is stale until then, as the Tables docs state).
    [Test]
    public async Task AHeaderOnlyTable_ThatGrows_RecomputesItsDependents_ThroughTheEngine()
    {
        var workbook = GrowingTable();
        workbook.ComputeAll();
        var engine = workbook.CreateRecalculationEngine();

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");
        await Assert.That(Display(workbook.GetCellValue("Main", "B2"))).IsEqualTo("0");

        workbook.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("42");
        await Assert.That(Display(workbook.GetCellValue("Main", "B2"))).IsEqualTo("2");
    }

    [Test]
    public async Task AHeaderOnlyTable_ThatGrows_RecomputesItsDependents_ThroughInvalidateCache()
    {
        var workbook = GrowingTable();

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");

        workbook.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);

        // A definition evicts nothing: the memoized 0 is still served until the cache is invalidated.
        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");

        workbook.InvalidateCache();

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("42");
        await Assert.That(Display(workbook.GetCellValue("Main", "B2"))).IsEqualTo("2");
    }

    // MySheet tables do not grow on their own: a value typed directly under the header of a table that is
    // still header-only is outside it, so SUM stays 0 through the engine and through a full invalidation. (The
    // oracle's read of that cell — SUM 7 over a value under the header — is the registered sentinel defect.)
    [Test]
    public async Task AValueTypedUnderTheHeader_OfAStillHeaderOnlyTable_LeavesItsDependentsAtZero()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        var main = workbook.Sheets.Add("Main");
        data["A1"] = new StringValue("Valor");
        workbook.DefineTable("Tabela1", "Data", "A1:A1", ["Valor"]);
        main["B1"] = ExpressionParser.Parse("=SUM(Tabela1[Valor])", main);
        workbook.ComputeAll();
        var engine = workbook.CreateRecalculationEngine();

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");

        data["A2"] = new NumberValue(10);
        engine.Recalculate([new CellRef("Data", "A2")]);

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");

        workbook.InvalidateCache();

        await Assert.That(Display(workbook.GetCellValue("Main", "B1"))).IsEqualTo("0");
    }

    // Data!A1 = "Valor" over A2 = 10 and A3 = 32, registered header-only (A1:A1); Main!B1 = SUM of the column,
    // B2 = ROWS of the data band.
    private static Workbook GrowingTable()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        var main = workbook.Sheets.Add("Main");
        data["A1"] = new StringValue("Valor");
        data["A2"] = new NumberValue(10);
        data["A3"] = new NumberValue(32);
        workbook.DefineTable("Tabela1", "Data", "A1:A1", ["Valor"]);
        main["B1"] = ExpressionParser.Parse("=SUM(Tabela1[Valor])", main);
        main["B2"] = ExpressionParser.Parse("=ROWS(Tabela1[#Data])", main);

        return workbook;
    }

    // === The ABSENT singleton still errors ================================================================
    //
    // [#Totals] on a table with no totals row is a region that does not exist, not an empty one. Measured on
    // the header-only table, both modes: SUM #REF!, ISREF FALSE; COUNT 0 and COUNTA 1 are ruling R2's shape
    // (measured on the three-row fixture) — the error value counted as one element.
    [Test]
    [Arguments("=SUM(Tabela1[#Totals])", "#REF!")]
    [Arguments("=ROWS(Tabela1[#Totals])", "#REF!")]
    [Arguments("=ISREF(Tabela1[#Totals])", "FALSE")]
    [Arguments("=COUNT(Tabela1[#Totals])", "0")]
    [Arguments("=COUNTA(Tabela1[#Totals])", "1")]
    [Arguments("=SUM(Tabela1[NoSuchColumn])", "#REF!")]
    public async Task OverAHeaderOnlyTable_AnAbsentRegionIsStillRef(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // === A header row and a totals row with no data between them ==========================================
    //
    // [#Data] is empty (anchored at the totals row), [#Totals] is a PRESENT row, and the pairs shrink to the
    // row they still have. Oracle, primed, both modes (ROW 2 plain / 1 CSE): SUM 0, ROWS 0, ROWS(T[#Data]) 0,
    // COUNTA(T[Valor]) 0, ISREF TRUE, COLUMNS(T[#Data]) 3, INDEX #REF!, SUM(T*2) #VALUE! / 0 (the mini-CSE
    // lift, as above), COUNTIF 0, ROWS(T[[#Data],[#Totals]]) 1, ROWS(T[[#Headers],[#Data]]) 1,
    // SUM(T[#Totals]) 0, SUM(Tabela1) 0. Registered defect: COUNTA(T[#Data]) and COUNTA(Tabela1) are 2 on
    // the oracle — they count the totals row's cells through the empty data band; MySheet answers 0.
    [Test]
    [Arguments("=SUM(Tabela1[Valor])", "0")]
    [Arguments("=ROWS(Tabela1[Valor])", "0")]
    [Arguments("=ROWS(Tabela1[#Data])", "0")]
    [Arguments("=COUNTA(Tabela1[Valor])", "0")]
    [Arguments("=ISREF(Tabela1[Valor])", "TRUE")]
    [Arguments("=ROW(Tabela1[Valor])", "2")]
    [Arguments("=COLUMNS(Tabela1[#Data])", "3")]
    [Arguments("=INDEX(Tabela1[Valor],1,1)", "#REF!")]
    [Arguments("=SUM(Tabela1[Valor]*2)", "0")]
    [Arguments("=COUNTIF(Tabela1[Valor],\">0\")", "0")]
    [Arguments("=ROWS(Tabela1[[#Data],[#Totals]])", "1")]
    [Arguments("=ROWS(Tabela1[[#Headers],[#Data]])", "1")]
    [Arguments("=SUM(Tabela1[#Totals])", "0")]
    [Arguments("=SUM(Tabela1)", "0")]
    [Arguments("=COUNTA(Tabela1[#Data])", "0")]
    [Arguments("=COUNTA(Tabela1)", "0")]
    public async Task AHeaderAndTotalsTable_HasAnEmptyDataBand_AndAPresentTotalsRow(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula, Shape.HeaderAndTotals)).IsEqualTo(expected);
    }

    // === On the table's own sheet =========================================================================
    //
    // The committed f7 layout (formulas beside the table). Oracle, primed, PLAIN: a bare =Tabela1[Valor] is
    // #VALUE! at E1, E2 and B5 (array-entered 0) — a zero-row reference covers no row, not even the anchor's,
    // so the cell boundary has nothing to intersect; ROW 2 (CSE 1); =Tabela1[Valor]*2 #VALUE! (CSE 0);
    // SUM 0 and ROWS(T[#Data]) 0 at the fixture's own D1/D2.
    [Test]
    [Arguments("=Tabela1[Valor]", "E1", "#VALUE!")]
    [Arguments("=Tabela1[Valor]", "E2", "#VALUE!")]
    [Arguments("=Tabela1[Valor]", "B5", "#VALUE!")]
    [Arguments("=ROW(Tabela1[Valor])", "E5", "2")]
    [Arguments("=Tabela1[Valor]*2", "B5", "#VALUE!")]
    [Arguments("=SUM(Tabela1[Valor])", "D1", "0")]
    [Arguments("=ROWS(Tabela1[#Data])", "D2", "0")]
    public async Task OnTheTablesOwnSheet_TheBareReferenceHasNoRowToIntersect(
        string formula,
        string id,
        string expected
    )
    {
        await Assert.That(InCell(formula, Shape.HeaderOnly, "Data", id)).IsEqualTo(expected);
    }

    // === Save / Load ======================================================================================
    //
    // The empty reference is a runtime value, never part of a stored tree: what travels is the TABLE (the
    // header-only geometry) and the structured-reference formulas, and they resolve to the empty reference
    // again after a round trip. Before this item every one of these answered #REF! (ISREF FALSE).
    [Test]
    public async Task AHeaderOnlyTable_SurvivesSaveAndLoad_AndStillResolvesToAnEmptyReference()
    {
        var workbook = Fixture(Shape.HeaderOnly);
        var main = workbook["Main"];
        main["H20"] = ExpressionParser.Parse("=SUM(Tabela1[Valor])", main);
        main["H21"] = ExpressionParser.Parse("=ROWS(Tabela1[#Data])", main);
        main["H22"] = ExpressionParser.Parse("=ISREF(Tabela1[Valor])", main);
        main["H23"] = ExpressionParser.Parse("=COUNTIF(Tabela1[Valor],\">0\")", main);
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var reloaded = Workbook.Load(path);

            await Assert.That(reloaded.Tables["Tabela1"].DataRowCount).IsEqualTo(0);
            await Assert.That(Display(reloaded.GetCellValue("Main", "H20"))).IsEqualTo("0");
            await Assert.That(Display(reloaded.GetCellValue("Main", "H21"))).IsEqualTo("0");
            await Assert.That(Display(reloaded.GetCellValue("Main", "H22"))).IsEqualTo("TRUE");
            await Assert.That(Display(reloaded.GetCellValue("Main", "H23"))).IsEqualTo("0");
        }
        finally
        {
            File.Delete(path);
        }
    }
}
