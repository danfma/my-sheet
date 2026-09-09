using System.Diagnostics.CodeAnalysis;

namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// The position (top row / leftmost column) of the reference an argument RESOLVES to, plus the error
/// recovery that <c>ROWS</c>/<c>COLUMNS</c>/<c>AREAS</c> share with <c>ROW</c>/<c>COLUMN</c>.
///
/// <para>Excel defines <c>ROW</c>/<c>COLUMN</c> over "a reference", not over a reference NODE: anything that
/// denotes a reference at evaluation time qualifies — a defined name, <c>INDEX</c>, <c>OFFSET</c>,
/// <c>INDIRECT</c>, <c>CHOOSE</c>, a <c>':'</c> range with reference-returning endpoints. That is exactly
/// what <see cref="NamedReferences.TryResolveReference"/> already does for <c>ROWS</c>/<c>COLUMNS</c>/
/// <c>AREAS</c>/<c>ISREF</c>, so <c>ROW</c>/<c>COLUMN</c> use it as their TERMINAL fallback, after their
/// syntactic arms (which stay as the zero-resolution fast path and as the shared-formula delta path).</para>
///
/// <para>This lives in its own file rather than as a <c>file static class</c> (the local idiom, e.g.
/// <c>OrderSelection</c> in OrderStatistics.cs) because <c>ROW</c> lives in Row.cs while <c>COLUMN</c> and
/// the counting functions live in LookupFunctions.cs, and a file-local type is invisible across files.</para>
/// </summary>
internal static class ReferencePosition
{
    /// <summary>
    /// The row of the reference <paramref name="argument"/> resolves to: a cell's own row, a range's TOP
    /// row, and for an open reference the DECLARED lower row bound (<c>ROW(A:A)</c> = 1, Excel's answer),
    /// which is why resolution asks for <c>boundOpenRanges:false</c> — bounding would substitute the first
    /// POPULATED row. A union has no single top row and stays <c>#VALUE!</c> — a deliberate no-change:
    /// MySheet has never assigned a position to a multi-area reference and this fix does not settle what one
    /// should report (Excel's own answer there was not measured, so no divergence is guessed at either way).
    /// </summary>
    public static ComputedValue Row(Expression argument, EvaluationContext context) =>
        TryResolve(argument, context, out var reference, out var failure)
            ? reference switch
            {
                CellReference cell => ComputedValue.Number(CellAddress.Parse(cell.Id).Row),
                RangeReference range => ComputedValue.Number(range.TopRow),
                OpenRangeReference open => ComputedValue.Number(open.RowMin ?? 1),
                _ => ComputedValue.Error(Error.Value),
            }
            : failure;

    /// <summary>The mirror of <see cref="Row"/> on the column axis (<c>COLUMN(1:1)</c> = 1).</summary>
    public static ComputedValue Column(Expression argument, EvaluationContext context) =>
        TryResolve(argument, context, out var reference, out var failure)
            ? reference switch
            {
                CellReference cell => ComputedValue.Number(CellAddress.Parse(cell.Id).Column),
                RangeReference range => ComputedValue.Number(range.LeftColumn),
                OpenRangeReference open => ComputedValue.Number(open.ColMin ?? 1),
                _ => ComputedValue.Error(Error.Value),
            }
            : failure;

    /// <summary>
    /// What a reference-requiring function returns for an argument it could NOT resolve to a reference: the
    /// argument's OWN error when it has one (<c>#NAME?</c> for an unknown name, <c>#REF!</c> for a failed
    /// <c>INDIRECT</c>/<c>OFFSET</c>), otherwise <paramref name="fallback"/>. Excel propagates the argument's
    /// error rather than inventing an answer for a broken reference, and the caller keeps its own answer for
    /// an argument that is merely not a reference (<c>#VALUE!</c> for <c>ROW</c>/<c>COLUMN</c>/<c>AREAS</c>,
    /// <c>1</c> for <c>ROWS</c>/<c>COLUMNS</c>, which treat a scalar as a 1x1 array).
    /// </summary>
    /// <remarks>
    /// This re-evaluates the argument, which the failed resolution attempt may already have partly evaluated
    /// (INDIRECT's <c>ref_text</c>, OFFSET's displacements). That cost is paid only on the FAILURE path,
    /// where the alternative is losing the error the user needs to see.
    /// </remarks>
    public static ComputedValue Unresolved(
        Expression argument,
        EvaluationContext context,
        ComputedValue fallback
    )
    {
        var value = argument.Evaluate(context);

        return value.TryGetError(out _) ? value : fallback;
    }

    // Resolves the argument to a concrete reference, or produces the ComputedValue the caller must return.
    private static bool TryResolve(
        Expression argument,
        EvaluationContext context,
        [NotNullWhen(true)] out Reference? reference,
        out ComputedValue failure
    )
    {
        if (
            !NamedReferences.TryResolveReference(
                argument,
                context,
                out reference,
                boundOpenRanges: false
            )
        )
        {
            failure = Unresolved(argument, context, ComputedValue.Error(Error.Value));
            return false;
        }

        // The caller's syntactic ReferenceGuard.MissingSheet pass cannot see INTO a function argument (its
        // `default` arm ignores one), so the RESOLVED target is re-checked here — the same re-check
        // ReferenceGuard does for a NameReference and Subtotal does after its own re-dispatch. Without it
        // ROW(INDEX(Ghost!A1:A3,2,1)) would report a row number for a sheet that no longer exists.
        if (ReferenceGuard.MissingSheet(reference, context) is { } missing)
        {
            reference = null;
            failure = ComputedValue.Error(missing);
            return false;
        }

        // `default` rather than ComputedValue.Blank: the two are bit-identical (ComputedValueKind.Blank is 0),
        // so this is about intent, not behaviour — Blank would read as a deliberate blank RESULT, while there
        // is no failure to report on this path. `failure` is only meaningful when this returns false.
        failure = default;
        return true;
    }
}
