namespace Danfma.MySheet.Expressions;

/// <summary>
/// Structural validation of a reference ARGUMENT before a function enumerates it. A reference to a sheet
/// that does not exist is a structural failure of the reference itself (<c>#REF!</c>), fiel ao Excel —
/// distinct from a VALUE error inside a cell of an existing sheet. The distinction matters because the
/// error-ignoring COUNT family would silently treat a missing sheet as an empty range (returning 0) if the
/// failure were only surfaced as a per-cell error in the value stream; it must instead SHORT-CIRCUIT the
/// whole function to <c>#REF!</c>. Consuming functions check this at their choke point, before enumeration.
/// </summary>
internal static class ReferenceGuard
{
    /// <summary>
    /// Returns <see cref="Error.Ref"/> when ANY argument is (or contains, for a union) a reference to a
    /// missing sheet; otherwise <c>null</c>. Non-reference arguments are ignored.
    /// </summary>
    public static Error? MissingSheet(Expression[] arguments, EvaluationContext context)
    {
        foreach (var argument in arguments)
        {
            if (MissingSheet(argument, context) is { } error)
            {
                return error;
            }
        }

        return null;
    }

    /// <summary>
    /// Returns <see cref="Error.Ref"/> when the argument is a reference (a cell, a range, an open range,
    /// any area of a union, or a defined name that stands for one) whose sheet does not exist; otherwise
    /// <c>null</c>. A reference produced by a FUNCTION (e.g. OFFSET) is not inspected here — such functions
    /// already resolve their base through <see cref="NamedReferences.TryResolveReference"/> and yield
    /// <c>#REF!</c> themselves when the base sheet is missing.
    /// </summary>
    public static Error? MissingSheet(Expression argument, EvaluationContext context)
    {
        switch (argument)
        {
            case CellReference cell:
                return Check(context, cell.SheetName);

            case RangeReference range:
                return Check(context, range.SheetName);

            // G3 spike (node-delta shared formulas): mirrors CellReference/RangeReference above so a
            // shared-formula slave's aggregate-function argument short-circuits to #REF! the same way when
            // its sheet is missing (SheetName is a literal component of these nodes, unaffected by delta).
            case AnchoredCellReference anchoredCell:
                return Check(context, anchoredCell.SheetName);

            case AnchoredRangeReference anchoredRange:
                return Check(context, anchoredRange.SheetName);

            case OpenRangeReference open:
                return Check(context, open.SheetName);

            case UnionReference union:
                foreach (var area in union.Areas)
                {
                    if (MissingSheet(area, context) is { } areaError)
                    {
                        return areaError;
                    }
                }

                return null;

            case NameReference:
                // A defined name may stand for a reference to a missing sheet. Resolve it WITHOUT bounding
                // the open range (bounding would scan the sheet); a resolved reference is re-checked.
                return NamedReferences.TryResolveReference(
                    argument,
                    context,
                    out var resolved,
                    boundOpenRanges: false
                )
                    ? MissingSheet(resolved, context)
                    : null;

            // Unary '+' is a transparent no-op (see UnaryOperation): SUM(+Ghost!A1:A3) must be the same
            // structural #REF! as SUM(Ghost!A1:A3), not an empty range.
            case UnaryOperation { Operator: UnaryOperator.Plus } plus:
                return MissingSheet(plus.Operand, context);

            case DynamicRange dynamic:
                // A ':' range with reference-returning endpoints (INDEX(...):A5). If it resolves to a concrete
                // range, re-check that range's sheet (same as NameReference). If it CANNOT form a concrete
                // range — cross-sheet endpoints, an open/array endpoint — that is a structural #REF! that the
                // error-ignoring family (COUNT/…) must respect up front, not silently treat as an empty range.
                return dynamic.TryResolveReference(context, out var resolvedRange)
                    ? MissingSheet(resolvedRange!, context)
                    : Error.Ref;

            // Phase 7: the three axis-selection producers stand for their SOURCE array the way the unary-plus
            // arm above stands for its operand — COUNT(FILTER(Ghost!A1:A3,…)) must be the same structural
            // #REF! as COUNT(Ghost!A1:A3), not the silent 0 an empty selection would count to. SEQUENCE has
            // no reference argument and needs no arm.
            //
            // THE SOURCE SLOT ONLY, and the phase's final review measured what that leaves open: a ghost sheet
            // in the INCLUDE slot is not seen structurally, so COUNT(FILTER(A1:A3,Ghost!B1:B3>0)) is 0 while
            // SUM of the same formula is #REF!. That is the pre-existing IF/Binary hole rather than anything
            // these three arms introduce — the same gap exists for COUNT(IF(Ghost!B1:B3>0,A1:A3,0)) — and the
            // oracle also answers 0 there, which is why it is recorded and not closed here.
            case Lookup.Filter filter:
                return MissingSheet(filter.Arguments[0], context);

            case Lookup.Sort sort:
                return MissingSheet(sort.Arguments[0], context);

            case Lookup.Unique unique:
                return MissingSheet(unique.Arguments[0], context);

            default:
                return null;
        }
    }

    private static Error? Check(EvaluationContext context, string sheetName) =>
        context.Workbook.Sheets.ContainsKey(sheetName) ? null : Error.Ref;
}
