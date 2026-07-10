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

            case DynamicRange dynamic:
                // A ':' range with reference-returning endpoints (INDEX(...):A5). If it resolves to a concrete
                // range, re-check that range's sheet (same as NameReference). If it CANNOT form a concrete
                // range — cross-sheet endpoints, an open/array endpoint — that is a structural #REF! that the
                // error-ignoring family (COUNT/…) must respect up front, not silently treat as an empty range.
                return dynamic.TryResolveReference(context, out var resolvedRange)
                    ? MissingSheet(resolvedRange!, context)
                    : Error.Ref;

            default:
                return null;
        }
    }

    private static Error? Check(EvaluationContext context, string sheetName) =>
        context.Workbook.Sheets.ContainsKey(sheetName) ? null : Error.Ref;
}
