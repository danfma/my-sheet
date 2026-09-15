using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class LookupFunctionTests
{
    private static object? Calc(string formula, params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static object? CalcMixed(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value switch
            {
                string s => new Danfma.MySheet.Expressions.StringValue(s),
                double d => new NumberValue(d),
                int i => new NumberValue(i),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task XMatch_InvalidMode_PrecedesLookupArrayError()
    {
        // Aspose.Cells 26.7.0 CSE validates the mode before inspecting the erroring lookup array.
        await Assert
            .That(Calc("=XMATCH(\"a\",1/0,\"bad\")") as ErrorValue)
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task VLookup_ColumnIndexBelowOne_IsValueError()
    {
        // support.microsoft.com VLOOKUP: col_index_num < 1 -> #VALUE!; greater than the number of
        // columns in table_array -> #REF!. Mirrors the HLOOKUP row_index_num rule (wave 3 finding).
        await Assert
            .That(Calc("=VLOOKUP(1,A1:B2,0)", ("A1", 1), ("B1", 2)) as ErrorValue)
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=VLOOKUP(1,A1:B2,3)", ("A1", 1), ("B1", 2)) as ErrorValue)
            .IsEqualTo(ErrorValue.Reference);
    }

    // ================================================================================================
    // Sweep item 34(b) — a consumer that resolves its own reference argument reports the node's OWN error
    // when the node does not resolve (#NAME? for an unknown name, the node's #REF! for an unresolvable
    // structured reference) instead of its fallback code. Oracle Aspose.Cells 26.6.0, measured 2026-09-11,
    // PLAIN and array-entered agreeing on every row below. What each row WAS: the resolvers' #REF!
    // (VLOOKUP/HLOOKUP/INDEX/OFFSET) and the scans' not-found #N/A (MATCH/XMATCH) or #VALUE!-for-a-
    // non-reference (FORMULATEXT).
    // ================================================================================================

    [Test]
    public async Task AnUnresolvedNode_InAReferenceSlot_ReportsItsOwnError()
    {
        await Assert
            .That(Calc("=VLOOKUP(1,NoSuch,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=HLOOKUP(1,NoSuch,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=INDEX(NoSuch,1,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=MATCH(1,NoSuch,0)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=XMATCH(1,NoSuch)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=OFFSET(NoSuch,1,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=LOOKUP(1,NoSuch)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=FORMULATEXT(NoSuch)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);

        // The measured exception that keeps the rule honest: XLOOKUP's oracle answer for an unresolved
        // NAME in the lookup-array slot is its own #N/A (and #VALUE! for the return array), so XLOOKUP
        // keeps its code — pinned green here, with the shape's full story in
        // MissingSheetReferenceTests.XLookup_OverAnUnresolvedName_KeepsItsOwnCode_WhereTheOracleDoesToo.
        await Assert
            .That(Calc("=XLOOKUP(1,NoSuch,B1:B3)", ("A1", 1), ("B1", 2)) as ErrorValue)
            .IsEqualTo(ErrorValue.NotAvailable);

        // A value that is merely NOT a reference keeps the consumer's own not-found answer. Before item 41,
        // VLOOKUP(1,5,1) returned #REF!; the measured oracle answer is #N/A, like MATCH's existing result.
        await Assert
            .That(Calc("=VLOOKUP(1,5,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Calc("=MATCH(1,5,0)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task AnUnresolvedLookupValue_IsPropagated_ByTheScanFamily()
    {
        // The VALUE slot of the same family: the lookup's own error leads the scan. MATCH's approximate
        // path always propagated it; the exact path, XMATCH and XLOOKUP now match it (the oracle answers
        // #NAME? on every row here, both entry modes) — VLOOKUP/HLOOKUP/LOOKUP always did.
        await Assert
            .That(Calc("=MATCH(NoSuch,A1:A3,0)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=MATCH(NoSuch,A1:A3,1)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=XMATCH(NoSuch,A1:A3)", ("A1", 1)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=XLOOKUP(NoSuch,A1:A3,B1:B3)", ("A1", 1), ("B1", 2)) as ErrorValue)
            .IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task Rows_CountsRowsInRange()
    {
        await Assert.That(Calc("=ROWS(A1:A3)") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=ROWS(A1:C2)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Row_ReturnsRowOfReference()
    {
        await Assert.That(Calc("=ROW(A5)") as double?).IsEqualTo(5.0);
        await Assert.That(Calc("=ROW(B2:B4)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Match_Exact()
    {
        await Assert
            .That(Calc("=MATCH(20,A1:A3,0)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(2.0);
        await Assert
            .That(Calc("=MATCH(99,A1:A3,0)", ("A1", 10), ("A2", 20), ("A3", 30)))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task LookupWildcards_UseExcelTildeEscapes()
    {
        // Fixture A1:A4 = "a*", "ab", "a?", "a~". Aspose.Cells 26.7.0 PLAIN/CSE: MATCH and
        // XMATCH both return 1. Before this fix MATCH returned #N/A and XMATCH returned 4.
        var cells = new (string, object)[]
        {
            ("A1", "a*"),
            ("A2", "ab"),
            ("A3", "a?"),
            ("A4", "a~"),
        };
        await Assert.That(CalcMixed("=MATCH(\"a~*\",A1:A4,0)", cells) as double?).IsEqualTo(1.0);
        await Assert.That(CalcMixed("=XMATCH(\"a~*\",A1:A4,2)", cells) as double?).IsEqualTo(1.0);

        // SEARCH already had the same measured answer; pin the agreeing consumer beside the lookup rows.
        await Assert.That(CalcMixed("=SEARCH(\"~*\",\"a*b\")") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Match_ApproximateAscending()
    {
        // Largest value <= 25 is 20, at position 2.
        await Assert
            .That(Calc("=MATCH(25,A1:A3,1)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task VLookup_ApproximateTextKey()
    {
        // Sorted A->Z text table; the reported case: lookup IS the first row's key.
        (string, object)[] table =
        [
            ("A1", "Bradbury Creek"),
            ("B1", 49.35),
            ("A2", "Cedar Falls"),
            ("B2", 10.0),
            ("A3", "Dunmore"),
            ("B3", 20.0),
        ];

        // First row, exact text key under approximate mode -> 49.35 (was #N/A before the fix).
        await Assert
            .That(CalcMixed("=VLOOKUP(\"Bradbury Creek\",A1:B3,2,TRUE)", table) as double?)
            .IsEqualTo(49.35);

        // Exact middle key.
        await Assert
            .That(CalcMixed("=VLOOKUP(\"Cedar Falls\",A1:B3,2,TRUE)", table) as double?)
            .IsEqualTo(10.0);

        // Between names: "Cz" > "Cedar Falls" but < "Dunmore" -> largest key <= "Cz" is "Cedar Falls".
        await Assert
            .That(CalcMixed("=VLOOKUP(\"Cz\",A1:B3,2,TRUE)", table) as double?)
            .IsEqualTo(10.0);

        // Below the smallest key -> #N/A (Excel contract).
        await Assert
            .That(CalcMixed("=VLOOKUP(\"Aardvark\",A1:B3,2,TRUE)", table))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Match_ApproximateTextKey()
    {
        (string, object)[] column =
        [
            ("A1", "Bradbury Creek"),
            ("A2", "Cedar Falls"),
            ("A3", "Dunmore"),
        ];

        // Largest key <= "Cz" is "Cedar Falls" at position 2 (was #VALUE!/#N/A before the fix).
        await Assert.That(CalcMixed("=MATCH(\"Cz\",A1:A3,1)", column) as double?).IsEqualTo(2.0);

        // Exact text key under approximate mode.
        await Assert
            .That(CalcMixed("=MATCH(\"Bradbury Creek\",A1:A3,1)", column) as double?)
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task XLookup_ClosestTextKey()
    {
        (string, object)[] table =
        [
            ("A1", "Bradbury Creek"),
            ("B1", 49.35),
            ("A2", "Cedar Falls"),
            ("B2", 10.0),
            ("A3", "Dunmore"),
            ("B3", 20.0),
        ];

        // match_mode -1 = exact-or-next-smaller: "Cz" -> "Cedar Falls" -> 10.0 (was #N/A before the fix).
        await Assert
            .That(CalcMixed("=XLOOKUP(\"Cz\",A1:A3,B1:B3,,-1)", table) as double?)
            .IsEqualTo(10.0);

        // match_mode 1 = exact-or-next-larger: "Cz" -> "Dunmore" -> 20.0.
        await Assert
            .That(CalcMixed("=XLOOKUP(\"Cz\",A1:A3,B1:B3,,1)", table) as double?)
            .IsEqualTo(20.0);
    }

    [Test]
    public async Task Index_SingleColumn()
    {
        await Assert
            .That(Calc("=INDEX(A1:A3,2)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(20.0);
    }

    [Test]
    public async Task Index_SingleRow_TreatsArgAsColumn()
    {
        await Assert
            .That(Calc("=INDEX(A1:C1,2)", ("A1", 7), ("B1", 8), ("C1", 9)) as double?)
            .IsEqualTo(8.0);
    }

    [Test]
    public async Task Index_TwoDimensional()
    {
        await Assert
            .That(Calc("=INDEX(A1:B2,2,2)", ("A1", 1), ("B1", 2), ("A2", 3), ("B2", 4)) as double?)
            .IsEqualTo(4.0);
    }
}
