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
                // Sweep item 33: a zero-row rectangle's top row is its anchor, the row after a header-only
                // table's header (oracle: ROW(Tabela1[Valor]) 2 typed plain; the array-entered 1 is the
                // header row, which only a normalized inverted rectangle would read).
                EmptyRangeReference empty => ComputedValue.Number(empty.TopRow),
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
                EmptyRangeReference empty => ComputedValue.Number(empty.LeftColumn),
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
    /// The NAME-CLASS rule the resolving consumers share (sweep item 34(b)): when
    /// <paramref name="argument"/> does not resolve to a reference AND its own value is an error, that
    /// error IS the answer — <c>#NAME?</c> for an unknown name, the node's own <c>#REF!</c> for an
    /// unresolvable structured reference, <c>#DIV/0!</c> for 1/0 — instead of the consumer's fallback
    /// code (VLOOKUP/HLOOKUP/INDEX/OFFSET answered their <c>#REF!</c>, MATCH/XMATCH/LOOKUP their
    /// not-found <c>#N/A</c>, FORMULATEXT its <c>#VALUE!</c>, where the oracle answers the node's error:
    /// Aspose.Cells 26.6.0, measured 2026-09-11, both entry modes). Returns <c>false</c> when the consumer
    /// must proceed as before — the argument RESOLVED, or is a value without an error (a scalar name
    /// keeps the consumer's own answer, measured: VLOOKUP over <c>5</c> is the oracle's <c>#N/A</c>, a
    /// fallback-code question this rule deliberately does not settle).
    ///
    /// <para>XLOOKUP's array slots are the measured exception and are NOT routed through here: the oracle
    /// answers its own <c>#N/A</c>/<c>#VALUE!</c> for an unresolved NAME there, so no one rule could also
    /// cover them without a per-shape arm.</para>
    /// </summary>
    /// <remarks>
    /// This evaluates the argument, but only after the resolution attempt FAILED — the same failure-path
    /// cost <see cref="Unresolved"/> already documented, where the alternative is losing the error the
    /// user needs to see.
    /// </remarks>
    public static bool TryUnresolvedError(
        Expression argument,
        EvaluationContext context,
        out ComputedValue error
    )
    {
        error = ComputedValue.Blank;

        if (NamedReferences.TryResolveReference(argument, context, out _, boundOpenRanges: false))
        {
            return false;
        }

        var value = argument.Evaluate(context);

        if (value.TryGetError(out var nodeError))
        {
            error = ComputedValue.Error(nodeError);
            return true;
        }

        return false;
    }

    /// <summary>
    /// The lookup VALUE slot's rule, the one site MATCH, XMATCH and XLOOKUP share: the lookup value's error
    /// leads the scan when it is the argument's OWN error (<see cref="PositionalRange.IsOwnSlotError"/> — an
    /// unresolved name's <c>#NAME?</c>, <c>1/0</c>'s <c>#DIV/0!</c>) OR the argument denotes a single CELL that
    /// holds it — a cell reference, an anchored cell, a name bound to one. A value slot is not a range slot:
    /// the cell's value IS the lookup value, so its error is not "content" the way an error cell inside a
    /// criteria range is (which is all <see cref="PositionalRange.IsOwnSlotError"/> was written for, and why
    /// reusing it alone lost the approximate path's propagation). Measured on Aspose.Cells 26.6.0 and 26.7.0
    /// (2026-09-14, PLAIN and array-entered, identical): over A1 = <c>=1/0</c>, <c>MATCH(A1,A1)</c> on every
    /// match type, <c>XMATCH(A1,B1:B3)</c> and <c>XLOOKUP(A1,B1:B3,C1:C3)</c> — even with an
    /// <c>if_not_found</c> — are <c>#DIV/0!</c>, as are VLOOKUP/HLOOKUP/LOOKUP, which check the value's error
    /// unconditionally. A RANGE in the value slot stays out: its <c>#VALUE!</c> is the collapse artifact
    /// (<c>SUM(MATCH(A1:A3,A1:A3,0))</c>, ElementwiseLiftingTests' known divergence).
    /// </summary>
    /// <remarks>The resolution runs only after the value turned out to be an error.</remarks>
    public static bool IsLookupValueError(
        Expression argument,
        ComputedValue lookup,
        EvaluationContext context
    ) =>
        lookup.TryGetError(out var error)
        && (
            PositionalRange.IsOwnSlotError(argument, error, context)
            || (
                NamedReferences.TryResolveReference(argument, context, out var reference)
                && reference is CellReference
            )
        );

    /// <summary>
    /// What a reference-requiring function returns for an argument it could NOT resolve to a reference: the
    /// argument's OWN error when it has one (<c>#NAME?</c> for an unknown name, <c>#REF!</c> for a failed
    /// <c>INDIRECT</c>/<c>OFFSET</c>), otherwise <paramref name="fallback"/>. Excel propagates the argument's
    /// error rather than inventing an answer for a broken reference, and the caller keeps its own answer for
    /// an argument that is merely not a reference (<c>#VALUE!</c> for <c>ROW</c>/<c>COLUMN</c>/<c>AREAS</c>,
    /// <c>1</c> for <c>ROWS</c>/<c>COLUMNS</c>, which treat a scalar as a 1x1 array).
    /// </summary>
    internal static ComputedValue Unresolved(
        Expression argument,
        EvaluationContext context,
        ComputedValue fallback
    ) => TryUnresolvedError(argument, context, out var error) ? error : fallback;
}
