using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ReferenceFunctionTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid(
        params (string Id, Expression Value)[] cells
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return (workbook, sheet);
    }

    private static Expression N(double v) => new NumberValue(v);

    private static Expression T(string v) => new Danfma.MySheet.Expressions.StringValue(v);

    // Anything involving INDIRECT (or the current cell) needs a real sheet name in the context: a bare
    // Parse(f, sheet).Evaluate(workbook) carries an EMPTY SheetName and Indirect.TryResolveReference bails
    // on it with #REF!. Same idiom as IndirectTests.
    private static object? Calc(Workbook workbook, Sheet sheet, string formula) =>
        ExpressionParser
            .Parse(formula, sheet)
            .Evaluate(new EvaluationContext(workbook, sheet.Name))
            .AsObject();

    // Fixture for the ROW/COLUMN resolution table. The CONTENTS are deliberate lies: A1=5, A2=0, A3=9 and
    // B1=7, C1=3 are neither the row nor the column of the cell holding them, so an assertion of 1/2 can
    // only come from the resolved reference's POSITION — never from a cell's value leaking through.
    private static (Workbook Workbook, Sheet Sheet) PositionGrid()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(5)),
            ("A2", N(0)),
            ("A3", N(9)),
            ("B1", N(7)),
            ("C1", N(3))
        );

        workbook.DefineName("MyName", "Sheet1!$A$1:$A$3");
        workbook.DefineName("MyCell", "Sheet1!$A$2");

        return (workbook, sheet);
    }

    [Test]
    public async Task Row_NoArgument_UsesCurrentCell()
    {
        var (workbook, sheet) = Grid();
        sheet["A5"] = ExpressionParser.Parse("=ROW()", sheet);

        // Reaching A5 through a reference sets the current cell to A5.
        await Assert
            .That(ExpressionParser.Parse("=A5", sheet).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(5.0);
    }

    // ROW/COLUMN of ANY reference-producing expression, not just a syntactic cell/range node: the top row
    // (leftmost column) of the reference the argument RESOLVES to, which is Excel's definition. Each row
    // below is Excel's answer for that formula.
    [Test]
    [Arguments("=ROW(A1:A3)", 1.0)] // syntactic fast path: top row of the range
    [Arguments("=ROW(A2)", 2.0)] // syntactic fast path: the cell itself
    [Arguments("=ROW(INDEX(A1:A3,2,1))", 2.0)] // INDEX resolves to the cell address A2
    [Arguments("=ROW(OFFSET(A1,1,0))", 2.0)] // 1x1 OFFSET resolves to A2
    [Arguments("=ROW(OFFSET(A1,1,0,2,1))", 2.0)] // resized OFFSET resolves to the range A2:A3
    [Arguments("=ROW(INDIRECT(\"A2\"))", 2.0)] // INDIRECT parses ref_text into A2
    [Arguments("=ROW(INDIRECT(\"A2:A3\"))", 2.0)] // …and into a range, whose top row is 2
    [Arguments("=ROW(MyName)", 1.0)] // defined name standing for Sheet1!$A$1:$A$3
    [Arguments("=ROW(MyCell)", 2.0)] // defined name standing for Sheet1!$A$2
    [Arguments("=ROW(CHOOSE(1,A2:A3,B1))", 2.0)] // CHOOSE resolves to its chosen range A2:A3
    [Arguments("=ROW(A:A)", 1.0)] // whole column: the DECLARED top row 1, not the first populated one
    [Arguments("=ROW(A2:A)", 2.0)] // open on the bottom only: the declared row 2
    [Arguments("=ROW(INDEX(A1:A3,2,1):A3)", 2.0)] // ':' with a reference-returning endpoint → A2:A3
    [Arguments("=COLUMN(A1:C1)", 1.0)] // syntactic fast path: leftmost column
    [Arguments("=COLUMN(INDEX(A1:C1,1,2))", 2.0)] // INDEX resolves to B1
    [Arguments("=COLUMN(MyName)", 1.0)] // defined name → column A
    [Arguments("=COLUMN(A:A)", 1.0)] // whole column: declared column A
    [Arguments("=COLUMN(1:1)", 1.0)] // whole row: no declared column → Excel's 1
    public async Task RowAndColumn_ResolveAnyReferenceProducingArgument(
        string formula,
        double expected
    )
    {
        var (workbook, sheet) = PositionGrid();

        await Assert.That(Calc(workbook, sheet, formula) as double?).IsEqualTo(expected);
    }

    // The two failure arms of the shared resolution: an argument that CANNOT be resolved to a reference
    // reports its OWN error when it has one, and #VALUE! only when it is simply not a reference.
    [Test]
    public async Task RowAndColumn_UnresolvableArgument_ReportTheArgumentsOwnError()
    {
        var (workbook, sheet) = PositionGrid();

        // An unknown name evaluates to #NAME? — the argument's own error, not a generic #VALUE!.
        await Assert.That(Calc(workbook, sheet, "=ROW(NoSuchName)")).IsEqualTo(ErrorValue.Name);
        await Assert.That(Calc(workbook, sheet, "=COLUMN(NoSuchName)")).IsEqualTo(ErrorValue.Name);

        // A reference to a sheet that does not exist is a structural #REF!, both written directly…
        await Assert
            .That(Calc(workbook, sheet, "=ROW(Ghost!A1:A3)"))
            .IsEqualTo(ErrorValue.Reference);
        // …and produced by a function, which the syntactic guard cannot see into: the resolved target is
        // re-checked, so a row number is never reported for a deleted sheet.
        await Assert
            .That(Calc(workbook, sheet, "=ROW(INDEX(Ghost!A1:A3,2,1))"))
            .IsEqualTo(ErrorValue.Reference);

        // COLUMN runs the SAME two passes (its own ReferenceGuard.MissingSheet over the arguments, then
        // ReferencePosition's re-check of the RESOLVED target), so both #REF! arms mirror on the column
        // axis: a resolved target on a deleted sheet, and an argument whose own failure IS #REF!.
        await Assert
            .That(Calc(workbook, sheet, "=COLUMN(INDEX(Ghost!A1:C1,1,2))"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc(workbook, sheet, "=COLUMN(INDIRECT(\"zz\"))"))
            .IsEqualTo(ErrorValue.Reference);

        // A scalar is not a reference at all: #VALUE!.
        await Assert.That(Calc(workbook, sheet, "=ROW(1)")).IsEqualTo(ErrorValue.NotValue);

        // A union resolves to a reference but has no single top row — deliberately unchanged: #VALUE!.
        await Assert
            .That(Calc(workbook, sheet, "=ROW((A1:A3,B1:B3))"))
            .IsEqualTo(ErrorValue.NotValue);

        // The propagation is observable in the harness itself: with NO current sheet (a bare
        // Evaluate(workbook), empty SheetName) INDIRECT cannot resolve an unqualified reference and reports
        // #REF! — which ROW now surfaces, where it used to flatten every failure to #VALUE!.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=ROW(INDIRECT(\"A2\"))", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task RowAndColumn_NoArgument_WithNoCurrentCell_IsValueError()
    {
        // ROW()/COLUMN() report the cell they are IN, and a ROOT evaluation has none: Calc passes a sheet
        // name but no cell id, so EvaluationContext.CellId is null and the zero-argument arm — guarded by
        // `when context.CellId is { } id` — is skipped. What is left is the terminal #VALUE!, never an
        // invented row 1. Row_NoArgument_UsesCurrentCell pins the other side of that guard.
        var (workbook, sheet) = PositionGrid();

        await Assert.That(Calc(workbook, sheet, "=ROW()")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc(workbook, sheet, "=COLUMN()")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task RowAndColumn_WithTwoArguments_AreRejectedAtParseTime()
    {
        // FunctionRegistry declares both as 0..1 arguments, so a second argument never reaches evaluation:
        // the parser refuses the formula, as Excel refuses it at entry. That is why the `_` arm of
        // Row/Column.Evaluate is NOT reachable through this door — the only ways in are the zero-argument
        // ROW() with no current cell (above) and a host-built node the parser never wrote.
        var (_, sheet) = PositionGrid();

        string[] formulas = ["=ROW(A1,A2)", "=COLUMN(A1,A2)"];

        foreach (var formula in formulas)
        {
            ParseException? thrown = null;

            try
            {
                ExpressionParser.Parse(formula, sheet);
            }
            catch (ParseException exception)
            {
                thrown = exception;
            }

            await Assert.That(thrown?.Kind).IsEqualTo(ParseErrorKind.InvalidArgumentCount);
        }
    }

    // ROWS/COLUMNS/AREAS share ROW/COLUMN's error recovery: a broken reference reports the argument's own
    // error instead of a plausible count. A non-reference VALUE keeps each function's own answer (Excel
    // treats a scalar as a 1x1 array for ROWS/COLUMNS, but AREAS strictly requires a reference).
    [Test]
    public async Task RowsColumnsAreas_UnresolvableArgument_ReportTheArgumentsOwnError()
    {
        var (workbook, sheet) = PositionGrid();

        await Assert.That(Calc(workbook, sheet, "=ROWS(NoSuchName)")).IsEqualTo(ErrorValue.Name);
        await Assert.That(Calc(workbook, sheet, "=COLUMNS(NoSuchName)")).IsEqualTo(ErrorValue.Name);
        await Assert.That(Calc(workbook, sheet, "=AREAS(NoSuchName)")).IsEqualTo(ErrorValue.Name);

        // INDIRECT of unparsable reference text is #REF!, and that is what ROWS must report.
        await Assert
            .That(Calc(workbook, sheet, "=ROWS(INDIRECT(\"zz\"))"))
            .IsEqualTo(ErrorValue.Reference);

        // Regression: the non-error fallbacks are untouched.
        await Assert.That(Calc(workbook, sheet, "=ROWS(5)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc(workbook, sheet, "=COLUMNS(5)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc(workbook, sheet, "=AREAS(5)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task VLookup_Exact()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("B1", T("a")),
            ("A2", N(2)),
            ("B2", T("b")),
            ("A3", N(3)),
            ("B3", T("c"))
        );

        await Assert
            .That(
                ExpressionParser
                    .Parse("=VLOOKUP(2,A1:B3,2,FALSE)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
        await Assert
            .That(
                ExpressionParser
                    .Parse("=VLOOKUP(99,A1:B3,2,FALSE)", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task VLookup_Approximate()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("B1", T("a")),
            ("A2", N(2)),
            ("B2", T("b")),
            ("A3", N(3)),
            ("B3", T("c"))
        );

        await Assert
            .That(
                ExpressionParser
                    .Parse("=VLOOKUP(2.5,A1:B3,2,TRUE)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
    }

    [Test]
    public async Task XLookup_ExactAndNotFound()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("B1", T("a")),
            ("A2", N(2)),
            ("B2", T("b")),
            ("A3", N(3)),
            ("B3", T("c"))
        );

        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2,A1:A3,B1:B3)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(99,A1:A3,B1:B3,\"none\")", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("none");
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(99,A1:A3,B1:B3)", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task XLookup_MismatchedArrayLengths_BoundToShorter()
    {
        // lookup_array (A1:A5) is longer than return_array (B1:B3). The non-admitted streaming path advances
        // both cursors in lockstep and stops at the shorter — reproducing the pre-refactor Math.Min(count)
        // bound exactly: a match WITHIN the shared prefix pairs with its return cell; a match only in the
        // uncovered tail is dropped (→ #N/A), never returning past the end of the return array.
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("A2", N(2)),
            ("A3", N(3)),
            ("A4", N(4)),
            ("A5", N(5)),
            ("B1", T("a")),
            ("B2", T("b")),
            ("B3", T("c"))
        );

        // Match at position 2 (within the [0,3) shared prefix) → the paired return cell.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2,A1:A5,B1:B3)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");

        // Match only in the uncovered tail (position 4 > return length) → dropped by the shorter bound.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(4,A1:A5,B1:B3)", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Offset_ScalarCell()
    {
        var (workbook, sheet) = Grid(("A1", N(10)), ("A2", N(20)), ("A3", N(30)), ("B1", N(5)));

        await Assert
            .That(
                ExpressionParser.Parse("=OFFSET(A1,2,0)", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(30.0);
        await Assert
            .That(
                ExpressionParser.Parse("=OFFSET(A1,0,1)", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Offset_MultiCell_FeedsAggregation()
    {
        var (workbook, sheet) = Grid(("A1", N(10)), ("A2", N(20)), ("A3", N(30)));

        // OFFSET(A1,0,0,3,1) is the range A1:A3, so SUM over it is 60.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=SUM(OFFSET(A1,0,0,3,1))", sheet)
                    .Evaluate(workbook)
                    .AsObject() as double?
            )
            .IsEqualTo(60.0);
    }

    [Test]
    public async Task XLookup_ApproximateModes()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("B1", T("a")),
            ("A2", N(2)),
            ("B2", T("b")),
            ("A3", N(3)),
            ("B3", T("c"))
        );

        // Omitted if_not_found (,,) then match_mode -1 (next smaller) / 1 (next larger).
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2.5,A1:A3,B1:B3,,-1)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2.5,A1:A3,B1:B3,,1)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("c");
    }

    [Test]
    public async Task XLookup_Wildcard()
    {
        var (workbook, sheet) = Grid(
            ("A1", T("apple")),
            ("A2", T("banana")),
            ("B1", N(1)),
            ("B2", N(2))
        );

        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(\"a*\",A1:A2,B1:B2,,2)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as double?
            )
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task XLookup_BinarySearchModesReturnCorrectResults()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("B1", T("a")),
            ("A2", N(2)),
            ("B2", T("b")),
            ("A3", N(3)),
            ("B3", T("c"))
        );

        // search_mode 2 / -2 (binary) are not optimized but must still return the correct match.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2,A1:A3,B1:B3,,0,2)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2,A1:A3,B1:B3,,0,-2)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("b");
    }

    [Test]
    public async Task XLookup_ReverseSearchFindsLast()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)),
            ("A2", N(2)),
            ("A3", N(2)),
            ("A4", N(3)),
            ("B1", T("a")),
            ("B2", T("b")),
            ("B3", T("c")),
            ("B4", T("d"))
        );

        await Assert
            .That(
                ExpressionParser
                    .Parse("=XLOOKUP(2,A1:A4,B1:B4,,0,-1)", sheet)
                    .Evaluate(workbook)
                    .AsObject() as string
            )
            .IsEqualTo("c");
    }
}
