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
/// <c>CriteriaPairs</c> in ConditionalAggregates.cs) because <c>ROW</c> lives in Row.cs while
/// <c>COLUMN</c> and the counting functions live in LookupFunctions.cs, and a file-local type is invisible
/// across files.</para>
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
        TryResolve(argument, context, ValueFallback, out var reference, out var failure)
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
        TryResolve(argument, context, ValueFallback, out var reference, out var failure)
            ? reference switch
            {
                CellReference cell => ComputedValue.Number(CellAddress.Parse(cell.Id).Column),
                RangeReference range => ComputedValue.Number(range.LeftColumn),
                OpenRangeReference open => ComputedValue.Number(open.ColMin ?? 1),
                _ => ComputedValue.Error(Error.Value),
            }
            : failure;

    /// <summary>
    /// The mini-CSE half of <c>ROWS</c>/<c>COLUMNS</c>: the extent of a COMPUTED array on the asked axis —
    /// <c>ROWS(FILTER(A1:A3,A1:A3&gt;0))</c> = 2, <c>ROWS(A1:A3*2)</c> = 3, <c>COLUMNS(SEQUENCE(2,3))</c> = 3
    /// — except that a 1x1 array whose only element is an error IS that error, the way a scalar error
    /// already reports itself on the reference path (<c>ROWS(1/0)</c> = <c>#DIV/0!</c>). A producer's own
    /// failure — an empty <c>FILTER</c>'s <c>#CALC!</c>, a bad <c>SEQUENCE</c> size's <c>#VALUE!</c> — is
    /// exactly such a 1x1 singleton (<see cref="ArrayShaping"/>), and the engine cannot tell it from a kept
    /// error element, so the one rule serves both. The caller reaches this only through
    /// <see cref="ArrayEvaluation.TryStream"/>, which keeps a bare reference or name on the reference path.
    /// Measured on Aspose.Cells 26.6.0, 2026-09-10, plain and array-entered agreeing, and pinned with the
    /// rows the oracle splits on in <c>DynamicArrayTests</c>.
    /// </summary>
    public static ComputedValue ArrayExtent(ArrayEvaluation.ArrayStream array, int extent) =>
        array.Length == 1 && array.ElementAt(0).TryGetError(out var error)
            ? ComputedValue.Error(error)
            : ComputedValue.Number(extent);

    /// <summary>The fallback <c>ROW</c>/<c>COLUMN</c>/<c>AREAS</c> share: an argument that is not a
    /// reference at all is <c>#VALUE!</c>. <c>ROWS</c>/<c>COLUMNS</c> pass <c>1</c> instead, treating a
    /// scalar as a 1x1 array.</summary>
    private static ComputedValue ValueFallback => ComputedValue.Error(Error.Value);

    /// <summary>
    /// Resolves <paramref name="argument"/> to the reference it DENOTES, or hands back — in
    /// <paramref name="failure"/> — the value the caller must return in its place. This is the single rule
    /// the whole family shares (<c>ROW</c>/<c>COLUMN</c> for a position, <c>ROWS</c>/<c>COLUMNS</c>/
    /// <c>AREAS</c> for a count), and it has two failure arms:
    ///
    /// <list type="bullet">
    /// <item>the argument does not resolve to a reference at all — see <see cref="Unresolved"/>, which
    /// reports the argument's OWN error when it has one and <paramref name="fallback"/> otherwise;</item>
    /// <item>it DOES resolve, but onto a sheet that no longer exists — a structural <c>#REF!</c>.</item>
    /// </list>
    ///
    /// <para>The second arm is why every one of these functions routes through here instead of resolving on
    /// its own: a caller's syntactic <see cref="ReferenceGuard.MissingSheet"/> pass cannot see INTO a
    /// function argument (its <c>default</c> arm ignores one), so only the RESOLVED target exposes a ghost
    /// sheet reached through <c>INDEX</c>/<c>OFFSET</c>/<c>INDIRECT</c>. Without it
    /// <c>ROWS(INDEX(Ghost!A1:A3,2,1))</c> would count a row of a deleted sheet as a plausible <c>1</c>.</para>
    ///
    /// <para><c>boundOpenRanges:false</c> is uniform across the family: <c>ROW</c>/<c>COLUMN</c> need the
    /// DECLARED bound (<c>ROW(A:A)</c> = 1, not the first populated row) and <c>ROWS</c>/<c>COLUMNS</c> need
    /// the open reference itself to apply their populated-extent rule. It is neutral for <c>AREAS</c>, which
    /// counts an open range and its bounding box alike as one area, and saves that function the sheet scan
    /// bounding would cost.</para>
    /// </summary>
    public static bool TryResolve(
        Expression argument,
        EvaluationContext context,
        ComputedValue fallback,
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
            failure = Unresolved(argument, context, fallback);
            return false;
        }

        // The re-check the summary calls the second failure arm: the RESOLVED target's sheet, which no
        // syntactic pass over the argument NODE can reach — the same re-check ReferenceGuard does for a
        // NameReference and Subtotal does after its own re-dispatch.
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
    private static ComputedValue Unresolved(
        Expression argument,
        EvaluationContext context,
        ComputedValue fallback
    )
    {
        var value = argument.Evaluate(context);

        return value.TryGetError(out _) ? value : fallback;
    }
}
