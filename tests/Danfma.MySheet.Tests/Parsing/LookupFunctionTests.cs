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
