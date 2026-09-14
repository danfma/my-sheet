using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Index(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            Arguments[0] is XLookup xlookup
            && xlookup.TryResolveReference(context, out var xlookupReference)
            && xlookupReference is not null
            && RangeBounds.TryFrom(xlookupReference, out var xlookupBounds)
        )
        {
            return IndexIntoRange(xlookupReference, xlookupBounds, context);
        }

        // Mini-CSE array-argument forms (the array is a computed vector, not a reference):
        //   • ROW of a whole/one-sided-open column is the IDENTITY row vector [top, top+1, …]; INDEX(…,n)
        //     returns its n-th worksheet row number WITHOUT materializing the grid-less column (the K1
        //     "INDEX(ROW($A:$A), SMALL(…))" idiom). Documented special case for Row([OpenRangeReference]).
        if (Arguments[0] is Row { Arguments: [OpenRangeReference openColumn] })
        {
            return IndexIntoOpenRowNumbers(openColumn, context);
        }

        //   • Any other array-eligible, non-reference first argument — ROW(B2:B5), ROW(name), IF(range=…,…)
        //     — is materialized row-major and indexed. References fall through to the concrete-range path
        //     below, so plain INDEX(A1:C10, r, c) is untouched.
        if (ArrayEvaluation.TryStream(Arguments[0], context, out var array))
        {
            return IndexIntoArray(array, context);
        }

        // The array may be a literal range or a defined name that stands for one. When it does not
        // resolve, the node's OWN error is the answer (sweep item 34(b): #NAME? for an unknown name, the
        // node's #REF! for an unresolvable structured reference) — INDEX's own #REF! stays only for an
        // argument that is merely not a range (a cell, a union).
        //
        // boundOpenRanges: false — sweep item 37 follow-up, ruling (a): an OPEN range (a whole row/column,
        // $5:$1000, …) stays open here, so IndexIntoOpenRange can address it by ABSOLUTE position (column
        // A / row 1, not the POPULATED bounding box's own corner that boundOpenRanges: true would collapse
        // it to — see IndexIntoOpenRange's own remarks). Every other shape (RangeReference,
        // EmptyRangeReference, a resolved TableReference/name) is unaffected: boundOpenRanges only changes
        // what happens to an OpenRangeReference.
        if (
            NamedReferences.TryResolveReference(
                Arguments[0],
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            if (reference is CellReference cell)
            {
                var value = cell.Evaluate(context);
                if (Arguments[0] is NameReference && value.TryGetError(out _))
                {
                    return value;
                }

                return IndexIntoScalar(value, context);
            }

            if (reference is OpenRangeReference open)
            {
                return IndexIntoOpenRange(open, context);
            }

            if (!RangeBounds.TryFrom(reference, out var bounds))
            {
                return ComputedValue.Error(Error.Ref);
            }

            return IndexIntoRange(reference, bounds, context);
        }

        if (ReferencePosition.TryUnresolvedError(Arguments[0], context, out var unresolved))
        {
            return unresolved;
        }

        return Arguments[0].Evaluate(context);
    }

    private ComputedValue IndexIntoScalar(ComputedValue value, EvaluationContext context)
    {
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        var row = Arguments.Length == 2 ? 1 : (int)first;
        var column = Arguments.Length == 2 ? (int)first : 1;

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out var third) is { } thirdError)
            {
                return ComputedValue.Error(thirdError);
            }

            column = (int)third;
        }

        return row is 0 or 1 && column is 0 or 1 ? value : ComputedValue.Error(Error.Ref);
    }

    // The concrete-range form, split out of Evaluate so the resolution arm above can hand the resolved
    // rectangle over. A zero row or column (sweep item 37) selects the whole column/row/area as a
    // REFERENCE rather than a value — mirrors Offset.Evaluate's own reference-vs-1x1-dereference split, so
    // every reference-aware consumer (SUM, ROWS, ROW, MATCH's lookup array, a nested INDEX, an OFFSET
    // base, a bare cell's implicit intersection, a ':' range endpoint, …) sees a real range instead of a
    // materialized value, with no per-consumer change needed.
    private ComputedValue IndexIntoRange(
        Reference reference,
        RangeBounds bounds,
        EvaluationContext context
    )
    {
        if (TryComputeAxes(bounds, context, out var row, out var column) is { } error)
        {
            return error;
        }

        if (row == 0 || column == 0)
        {
            if (ReferenceGuard.MissingSheet(Arguments[0], context) is { } missing)
            {
                return ComputedValue.Error(missing);
            }

            return ComputedValue.Reference(BuildAxisReference(reference, bounds, row, column));
        }

        // row and column are both >= 1 here, so reference is necessarily a RangeReference: an
        // EmptyRangeReference's RowCount is 0, and TryComputeAxes already rejected any row > 0 against it.
        return reference is RangeReference range
            ? range.CellComputedValueAt(context, row, column)
            : ComputedValue.Error(Error.Ref);
    }

    // The row_num/column_num resolution INDEX's two reference-consuming paths share (Evaluate's
    // concrete-range arm and TryResolveReference): argument coercion, the 2-arg axis rule (a single-row
    // area takes the lone index as a column, otherwise as a row), truncation to an integer (Excel
    // truncates toward zero, matching Offset's height/width), and the bounds check — now admitting exactly
    // 0 on either axis (sweep item 37) where it used to require >= 1. Returns the error RESULT to hand
    // back on failure, or null with (row, column) set on success.
    private ComputedValue? TryComputeAxes(
        RangeBounds bounds,
        EvaluationContext context,
        out int row,
        out int column
    )
    {
        row = 0;
        column = 0;

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        double rowValue;
        double columnValue;

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out columnValue) is { } columnError)
            {
                return ComputedValue.Error(columnError);
            }

            rowValue = first;
        }
        else if (bounds.RowCount == 1)
        {
            // A single-row range takes the lone index as a column.
            rowValue = 1;
            columnValue = first;
        }
        else
        {
            rowValue = first;
            columnValue = 1;
        }

        row = (int)rowValue;
        column = (int)columnValue;

        if (row < 0 || column < 0 || row > bounds.RowCount || column > bounds.ColumnCount)
        {
            return ComputedValue.Error(row < 0 || column < 0 ? Error.Value : Error.Ref);
        }

        return null;
    }

    // The reference a zero row/column selects: (0,0) is the WHOLE area — the same reference back, no new
    // node — and each single-zero form narrows one axis to the row_num/column_num given while the other
    // spans the area's full extent. An EmptyRangeReference only ever reaches the row == 0 (column select)
    // arm: TryComputeAxes already turned any row > 0 against its zero RowCount into #REF!, so there is no
    // row to select and the result stays a (narrower) EmptyRangeReference — still zero rows.
    private static Reference BuildAxisReference(
        Reference reference,
        RangeBounds bounds,
        int row,
        int column
    )
    {
        if (row == 0 && column == 0)
        {
            return reference;
        }

        var left = row == 0 ? bounds.LeftColumn + column - 1 : bounds.LeftColumn;
        var right = row == 0 ? left : bounds.RightColumn;
        var top = column == 0 ? bounds.TopRow + row - 1 : bounds.TopRow;
        var bottom = column == 0 ? top : bounds.BottomRow;

        return reference switch
        {
            EmptyRangeReference empty => new EmptyRangeReference(empty.SheetName, top, left, right),
            RangeReference range => new RangeReference(
                new CellAddress(left, top).ToId(),
                new CellAddress(right, bottom).ToId(),
                range.SheetName
            ),
            _ => reference,
        };
    }

    // The OPEN-base form, sweep item 37 follow-up ruling (a): row_num/column_num address ABSOLUTE grid
    // positions on an open axis (column A / row 1 when that side has no declared bound), mirroring
    // OpenRangeReference's own AbsoluteRow/AbsoluteColumn — never the POPULATED bounding box's own corner
    // ToBoundedRange collapses an open range to, which is a DIFFERENT cell whenever the first populated
    // column/row is not the sheet's own first one (INDEX($5:$1000,1,3) is C5, "a" — the box-relative
    // reading before this fix landed on whatever the first POPULATED column was instead). A zero row or
    // column still selects the whole row/column/area as a REFERENCE, exactly like IndexIntoRange, so
    // COLUMN/COUNTA/SUM/ROWS of the result all read the SAME absolute rectangle.
    private ComputedValue IndexIntoOpenRange(OpenRangeReference open, EvaluationContext context)
    {
        if (TryComputeOpenAxes(open, context, out var row, out var column) is { } error)
        {
            return error;
        }

        if (row == 0 || column == 0)
        {
            return ComputedValue.Reference(BuildOpenAxisReference(open, row, column));
        }

        return context.Workbook.GetCellValue(
            open.SheetName,
            new CellAddress(open.AbsoluteColumn(column), open.AbsoluteRow(row)).ToId()
        );
    }

    // The open-range twin of TryComputeAxes: the SAME argument coercion, 2-arg axis rule (IsSingleRow
    // stands in for bounds.RowCount == 1) and admits-zero bounds check, but validated against the
    // reference's OWN declared/grid limits (OpenRangeReference.IsRowIndexValid/IsColumnIndexValid) instead
    // of a closed RangeBounds' RowCount/ColumnCount.
    private ComputedValue? TryComputeOpenAxes(
        OpenRangeReference open,
        EvaluationContext context,
        out int row,
        out int column
    )
    {
        row = 0;
        column = 0;

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        double rowValue;
        double columnValue;

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out columnValue) is { } columnError)
            {
                return ComputedValue.Error(columnError);
            }

            rowValue = first;
        }
        else if (open.IsSingleRow)
        {
            // A single-row range takes the lone index as a column.
            rowValue = 1;
            columnValue = first;
        }
        else
        {
            rowValue = first;
            columnValue = 1;
        }

        row = (int)rowValue;
        column = (int)columnValue;

        if (row < 0 || column < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (
            (row != 0 && !open.IsRowIndexValid(row))
            || (column != 0 && !open.IsColumnIndexValid(column))
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

        return null;
    }

    // The open-range twin of BuildAxisReference: (0,0) is the reference itself (no new node); a single
    // zero narrows ONE axis to the absolute row/column given and leaves the OTHER axis exactly as the open
    // reference already declares it (its own bound, open or not) — so INDEX($5:$1000,0,3) keeps the
    // DECLARED row bound (5..1000) and narrows only the column, becoming the ordinary RangeReference C5:C1000,
    // while a zero on an axis that is itself still open after narrowing (e.g. INDEX($5:$1000,3,0), whose
    // column axis has no declared bound at all) stays an OpenRangeReference.
    private static Reference BuildOpenAxisReference(OpenRangeReference open, int row, int column)
    {
        if (row == 0 && column == 0)
        {
            return open;
        }

        int? rowMin;
        int? rowMax;
        int? colMin;
        int? colMax;

        if (row == 0)
        {
            rowMin = open.RowMin;
            rowMax = open.RowMax;
            colMin = colMax = open.AbsoluteColumn(column);
        }
        else
        {
            rowMin = rowMax = open.AbsoluteRow(row);
            colMin = open.ColMin;
            colMax = open.ColMax;
        }

        if (rowMin is { } r0 && rowMax is { } r1 && colMin is { } c0 && colMax is { } c1)
        {
            return new RangeReference(
                new CellAddress(c0, r0).ToId(),
                new CellAddress(c1, r1).ToId(),
                open.SheetName
            );
        }

        return new OpenRangeReference(colMin, colMax, rowMin, rowMax, open.SheetName);
    }

    // Indexes an element-wise vector row-major (the mini-CSE array form) through the LAZY stream: only the
    // selected element is computed (no ComputedValue[] materialized). Mirrors the concrete-range branch's
    // row/column resolution: a 3-arg call takes (row, column); a 2-arg call maps the lone index to the array's
    // only axis (column for a single row, row otherwise). Out of bounds → #REF!.
    private ComputedValue IndexIntoArray(
        ArrayEvaluation.ArrayStream array,
        EvaluationContext context
    )
    {
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        double row;
        double column;

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out column) is { } columnError)
            {
                return ComputedValue.Error(columnError);
            }

            row = first;
        }
        else if (array.Rows == 1)
        {
            row = 1;
            column = first;
        }
        else
        {
            row = first;
            column = 1;
        }

        var truncatedRow = (int)row;
        var truncatedColumn = (int)column;

        if (truncatedRow < 0 || truncatedColumn < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (
            truncatedRow < 1
            || truncatedColumn < 1
            || truncatedRow > array.Rows
            || truncatedColumn > array.Columns
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

        // Row-major layout (ArrayEvaluation lays out row-then-column): (r,c) 1-based → (r-1)·Columns + (c-1).
        return array.ElementAt((truncatedRow - 1) * array.Columns + (truncatedColumn - 1));
    }

    // ROW of an open column is the identity vector [top, top+1, …] with top = RowMin (or 1 when the top is
    // open). INDEX(ROW($A:$A), n) therefore returns top+n-1 directly — no column is materialized, honoring
    // the grid-less model (an open bottom has no fixed 1,048,576 cap). n < 1, or n past a bounded bottom,
    // is #REF!. A 3-arg form must select the single column (column 1).
    private ComputedValue IndexIntoOpenRowNumbers(
        OpenRangeReference open,
        EvaluationContext context
    )
    {
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out var column) is { } columnError)
            {
                return ComputedValue.Error(columnError);
            }

            if (column is < 1 or > 1)
            {
                return ComputedValue.Error(Error.Ref);
            }
        }

        var n = (int)Math.Truncate(first);
        var top = open.RowMin ?? 1;

        if (n < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (n < 1 || (open.RowMax is { } bottom && top + n - 1 > bottom))
        {
            return ComputedValue.Error(Error.Ref);
        }

        return ComputedValue.Number(top + n - 1);
    }

    // Mirrors the concrete-range branch of Evaluate (via the same TryComputeAxes/BuildAxisReference
    // helpers), but yields a Reference instead of reading a value: a CELL ADDRESS for a non-zero (row,
    // column), or the zero-axis REFERENCE itself (sweep item 37) for a zero row or column — which is what
    // lets a ':' range endpoint (INDEX(r,0,1):A3), an OFFSET base and every other reference-consuming
    // caller reach it without going through Evaluate's ComputedValue.Reference wrapper. Array forms (the
    // mini-CSE vector and the open-column ROW identity) have no cell address to hand back, so they return
    // false and fall through to normal evaluation elsewhere.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // Array forms (mini-CSE vector, open-column ROW identity) have no cell address. Deliberately NOT
        // ArrayEvaluation.TryStream: this probes only to REJECT those forms and never builds a stream, so
        // the third condition would be a wasted build (see IsArrayEligible's remark on this exact caller).
        // The first condition is TryStream's own predicate, shared so a bare NAME — array-eligible since
        // Phase 11a, but a reference at this top level — keeps resolving to its cell: without it, measured
        // on the prototype, SUM(A1:INDEX(Rng,3)) went 14 → #REF!, ROW(INDEX(Rng,2)) 2 → #VALUE!,
        // ISREF(INDEX(Rng,2)) TRUE → FALSE and OFFSET(INDEX(Rng,1),1,0) 0 → #REF!
        // (DefinedNameArrayEligibilityTests.Index_OverABareName_StillReturnsAReference). The context-aware
        // overload is what lets a name bound to an ARRAY by LET fall to the array path instead:
        // LET(f,FILTER(…),INDEX(f,2)) is 9 (ArrayBindingTests), where the reference path answered #REF!.
        if (
            Arguments[0] is Row { Arguments: [OpenRangeReference] }
            || (
                !ArrayEvaluation.IsBareReferenceNode(Arguments[0], context)
                && ArrayEvaluation.IsArrayEligible(Arguments[0], context)
            )
        )
        {
            return false;
        }

        if (
            !NamedReferences.TryResolveReference(
                Arguments[0],
                context,
                out var resolved,
                boundOpenRanges: false
            )
        )
        {
            return false;
        }

        if (resolved is OpenRangeReference open)
        {
            if (TryComputeOpenAxes(open, context, out var openRow, out var openColumn) is not null)
            {
                return false;
            }

            if (openRow == 0 || openColumn == 0)
            {
                reference = BuildOpenAxisReference(open, openRow, openColumn);
                return true;
            }

            // row and column are both valid here: TryComputeOpenAxes already checked
            // IsRowIndexValid/IsColumnIndexValid for a non-zero axis before returning null.
            var openTarget = new CellAddress(
                open.AbsoluteColumn(openColumn),
                open.AbsoluteRow(openRow)
            );
            reference = new CellReference(openTarget.ToId(), open.SheetName);
            return true;
        }

        if (!RangeBounds.TryFrom(resolved, out var bounds))
        {
            return false;
        }

        if (TryComputeAxes(bounds, context, out var row, out var column) is not null)
        {
            return false;
        }

        if (row == 0 || column == 0)
        {
            if (ReferenceGuard.MissingSheet(Arguments[0], context) is not null)
            {
                return false;
            }

            reference = BuildAxisReference(resolved, bounds, row, column);
            return true;
        }

        // row and column are both >= 1 here, so resolved is necessarily a RangeReference (see
        // IndexIntoRange's identical reasoning).
        if (resolved is not RangeReference range)
        {
            return false;
        }

        // Mirrors RangeReference.CellComputedValueAt's normalization: StartId is not guaranteed to be the
        // top-left corner (e.g. A3:A1 has StartId="A3"), so the origin is the range's normalized top-left.
        var target = new CellAddress(range.LeftColumn + column - 1, range.TopRow + row - 1);
        reference = new CellReference(target.ToId(), range.SheetName);
        return true;
    }
}
