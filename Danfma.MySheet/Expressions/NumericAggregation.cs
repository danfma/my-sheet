using System.Globalization;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Receives each numeric value gathered by <see cref="NumericAggregation"/>. Implemented by mutable
/// structs so the JIT specializes the fold per function and avoids heap allocation in the hot path.
/// </summary>
internal interface INumericFold
{
    void Accept(double value);
}

/// <summary>
/// Fold that materializes the gathered values, for aggregations that need more than a single pass
/// (GCD, LCM, MULTINOMIAL).
/// </summary>
internal struct NumericListFold : INumericFold
{
    public List<double> Values;

    public void Accept(double value) => Values.Add(value);
}

/// <summary>
/// Shared single-pass numeric gathering for aggregate functions (SUM, AVERAGE, MIN, MAX, COUNT), reading
/// each cell as a <see cref="ComputedValue"/> straight from the cache (no boxing). Mirrors Excel's rule:
/// numeric text and logicals passed <em>directly</em> as arguments are counted, but text/logicals/blanks
/// pulled from <em>referenced</em> cells are ignored. The first error encountered is returned; the caller
/// decides whether to propagate it (SUM…) or ignore it (COUNT).
/// </summary>
internal static class NumericAggregation
{
    public static Error? Fold<TFold>(
        Expression[] arguments,
        EvaluationContext context,
        ref TFold fold
    )
        where TFold : struct, INumericFold
    {
        // A reference argument to a missing sheet is a STRUCTURAL #REF! that short-circuits the whole
        // aggregation, before enumeration. Returned through the error channel, so functions that propagate
        // errors (SUM, AVERAGE, MIN, MAX, PRODUCT, STDEV, VAR, …) surface it. COUNT ignores this channel, so
        // it guards structurally on its own.
        if (ReferenceGuard.MissingSheet(arguments, context) is { } missingSheet)
        {
            return missingSheet;
        }

        Error? error = null;

        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                case OpenRangeReference open:
                    foreach (var value in open.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Evaluate(context), ref fold, ref error);
                    break;

                // G3 spike (node-delta shared formulas): mirrors the CellReference/RangeReference cases
                // above so a shared-formula slave's MAX(B2,C2)-shaped argument gets the REFERENCED-cell
                // aggregation rule (text/logicals/blanks pulled from a referenced cell are ignored) instead
                // of silently falling to the DIRECT-value rule in the `default` branch below (which would
                // happen if these fell through: they evaluate fine via Expression.Evaluate — that is not the
                // bug — but AddDirect's Excel-parity rule for a directly-passed value differs from
                // AddReferenced's rule for a value read through a reference).
                case AnchoredCellReference anchoredCell:
                    AddReferenced(anchoredCell.Evaluate(context), ref fold, ref error);
                    break;

                case AnchoredRangeReference anchoredRange:
                    foreach (var value in anchoredRange.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                // Phase 5 item 14: PERF/consistency only — the `default:` arm below already answers
                // correctly for a TableReference (it evaluates the node, gets a reference VALUE back for a
                // resolved table, and expands it through EnumerateValues' boxed iterator). This arm buys the
                // allocation-free struct RangeValueSequence enumerator instead, mirroring the
                // AnchoredRangeReference arm above for the identical reason.
                case TableReference table:
                    if (!table.TryResolve(context.Workbook, out var tableArea, out var tableError))
                    {
                        error ??= tableError;
                    }
                    else if (tableArea is RangeReference tableRange)
                    {
                        foreach (var value in tableRange.ExpandComputedValues(context))
                        {
                            AddReferenced(value, ref fold, ref error);
                        }
                    }

                    // An EMPTY area (sweep item 33, an EmptyRangeReference) has no cell to fold.
                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                default:
                    if (
                        argument is Lookup.XLookup computedXLookup
                        && computedXLookup.TryBuildSelection(context, out var selected)
                    )
                    {
                        var selectedStream = new ArrayEvaluation.ArrayStream(
                            selected,
                            selected.Rows,
                            selected.Columns
                        );
                        foreach (var value in selectedStream)
                        {
                            AddReferenced(value, ref fold, ref error);
                        }

                        break;
                    }

                    if (
                        argument is Lookup.XLookup xlookup
                        && xlookup.TryResolveReference(context, out var xlookupReference)
                    )
                    {
                        foreach (
                            var cellValue in ComputedValue
                                .Reference(xlookupReference!)
                                .EnumerateValues(context)
                        )
                        {
                            AddReferenced(cellValue, ref fold, ref error);
                        }

                        break;
                    }

                    // Mini-CSE: an array-eligible argument — IF(range=…,…), a range comparison,
                    // ROW/COLUMN of a range or of a name that stands for one, a name-built array — folds
                    // element-by-element with RANGE semantics (logicals/text ignored, so the FALSE of a
                    // branch-less IF drops out; the first cell error propagates). The shared gate keeps the
                    // scalar hot path below at a shallow type-walk (one reference resolution for the
                    // ROW/COLUMN and bare-name cases — never an evaluation of the node itself, though
                    // resolving a ':' range with reference-returning endpoints does evaluate those endpoints'
                    // own arguments, as IsArrayEligible's remark records) and avoids any double evaluation
                    // (IsArrayEligible ⇒ the stream build succeeds as the single evaluation).
                    //
                    // The gate's leading condition (ArrayEvaluation.IsBareReferenceNode) is LOAD-BEARING
                    // here for a bare NameReference, which the switch above does not peel off (it is not a
                    // Reference) and which IsArrayEligible answers TRUE for once it is bound to a rectangle.
                    // Without it the name would stream through RangeOperand instead of taking the
                    // referenced-value path below — measured on the prototype: SUM(Wide) over a 2-D name
                    // holding #N/A at B1 and #DIV/0! at A3 went #DIV/0! → #N/A, because the stream is
                    // row-major where this arm's referenced-value walk is column-major, so a different
                    // first error wins; the same flip hit MAX/SMALL/SUBTOTAL/AGGREGATE over the name
                    // (DefinedNameArrayEligibilityTests.TwoDimensionalBareName_OnTheErrorFixture_…).
                    // It is load-bearing for a unary '+' over a reference as well (Phase 11c): the '+' is
                    // transparent to the probe now, so without the condition SUM(+A1:A3) would leave the
                    // referenced-value path for the stream — the same column-major/row-major exposure, and
                    // the criteria family's twin of it is measured (CriteriaScan.RejectComputedArray).
                    // For a syntactic Reference the condition is moot rather than harmful: every reference
                    // node this arm could see is peeled off above — TableReference joined that set only
                    // with the arm above it (Phase 5 item 14) — except DynamicRange, which has no Probe arm
                    // and already takes the path below. Before that arm existed, a TableReference reached
                    // exactly this branch and DEPENDED on the condition precisely like the bare name and the
                    // unary '+' above: narrowing IsBareReferenceNode to answer false for a TableReference
                    // (measured on the mutation, then reverted) flips SUM over a 2-D table area with one
                    // error in each column from #DIV/0! to #N/A, the same row-major/column-major exposure —
                    // and that exposure is STILL live wherever a sibling gate shares the predicate with no
                    // arm of its own: PositionalRange.RejectComputedArray (CriteriaScan.cs) turns
                    // COUNTIF(Mat[#Data],">0") from counting normally (2) into #REF!, and
                    // OrderSelection.KthValue (SMALL/LARGE, outside this phase's files) inherits the same
                    // flip SUM no longer has (MiniCseConsumerTests.TableReference_OnTheTwoDimensionalErrorFixture_KeepsTheEnginesScanOrder).
                    if (ArrayEvaluation.TryStream(argument, context, out var array))
                    {
                        // Stream the element-wise vector: aggregate straight from the lazy view, allocating no
                        // ComputedValue[] (a 50k-row idiom drops from ~14MB to the handful of tree nodes). The
                        // row-major order is the materialized order, so the first cell error (scan order) still
                        // wins.
                        foreach (var element in array)
                        {
                            AddReferenced(element, ref fold, ref error);
                        }

                        break;
                    }

                    // Sweep item 32's C1 fix wave: a scalar-condition IF/CHOOSE whose taken branch is a
                    // SINGLE cell now hands a scalar consumer the cell's own VALUE (ArrayEvaluation.BranchValue
                    // / Choose.CaptureChosen), not a Reference-kind wrapper — so the Kind==Reference check
                    // below no longer catches it, and evaluating it directly would fold a referenced cell's
                    // TEXT through AddDirect's stricter rule (SUM(CHOOSE(1,MyCell)) over a text cell would
                    // become #VALUE! instead of the oracle's 0). Resolve the selector to its reference FIRST,
                    // through the very same reading ISREF/ROWS/INDEX already use
                    // (If.TryResolveReference / Choose's own), and gather it exactly like any other
                    // referenced cell — one element for a single cell, every cell for a range. A selector
                    // whose CONDITION itself errors is not a reference (TryResolveReference answers false)
                    // and falls through to the plain evaluation below, which surfaces that error unchanged.
                    if (
                        argument is If or Choose
                        && argument.TryResolveReference(context, out var selectorReference)
                    )
                    {
                        foreach (
                            var cellValue in ComputedValue
                                .Reference(selectorReference!)
                                .EnumerateValues(context)
                        )
                        {
                            AddReferenced(cellValue, ref fold, ref error);
                        }

                        break;
                    }

                    var argumentValue = argument.Evaluate(context);

                    // A function (e.g. OFFSET) may yield a range value; expand it as referenced cells.
                    if (argumentValue.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in argumentValue.EnumerateValues(context))
                        {
                            AddReferenced(cellValue, ref fold, ref error);
                        }
                    }
                    else
                    {
                        AddDirect(argumentValue, ref fold, ref error);
                    }

                    break;
            }
        }

        return error;
    }

    /// <summary>
    /// The A-variant gathering (AVERAGEA/MAXA/MINA/STDEVA/…): like <see cref="Fold{TFold}"/> except
    /// that referenced text counts as 0 and referenced logicals as 1/0 (blanks are still ignored),
    /// mirroring Excel's documented *A rule.
    /// </summary>
    public static Error? FoldA<TFold>(
        Expression[] arguments,
        EvaluationContext context,
        ref TFold fold
    )
        where TFold : struct, INumericFold
    {
        // See Fold: a missing-sheet reference is a structural #REF! that short-circuits the aggregation.
        if (ReferenceGuard.MissingSheet(arguments, context) is { } missingSheet)
        {
            return missingSheet;
        }

        Error? error = null;

        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        AddReferencedA(value, ref fold, ref error);
                    }

                    break;

                case OpenRangeReference open:
                    foreach (var value in open.ExpandComputedValues(context))
                    {
                        AddReferencedA(value, ref fold, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferencedA(cell.Evaluate(context), ref fold, ref error);
                    break;

                // G3 spike (node-delta shared formulas): see the matching cases in Fold above.
                case AnchoredCellReference anchoredCell:
                    AddReferencedA(anchoredCell.Evaluate(context), ref fold, ref error);
                    break;

                case AnchoredRangeReference anchoredRange:
                    foreach (var value in anchoredRange.ExpandComputedValues(context))
                    {
                        AddReferencedA(value, ref fold, ref error);
                    }

                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        AddReferencedA(value, ref fold, ref error);
                    }

                    break;

                default:
                    var argumentValue = argument.Evaluate(context);

                    if (argumentValue.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in argumentValue.EnumerateValues(context))
                        {
                            AddReferencedA(cellValue, ref fold, ref error);
                        }
                    }
                    else
                    {
                        AddDirect(argumentValue, ref fold, ref error);
                    }

                    break;
            }
        }

        return error;
    }

    private static void AddReferenced<TFold>(
        in ComputedValue value,
        ref TFold fold,
        ref Error? error
    )
        where TFold : struct, INumericFold
    {
        if (value.TryGetError(out var referencedError))
        {
            error ??= referencedError;
        }
        else if (value.TryGetNumber(out var number))
        {
            fold.Accept(number);
        }

        // Referenced text, logicals and blanks are ignored, matching Excel.
    }

    private static void AddReferencedA<TFold>(
        in ComputedValue value,
        ref TFold fold,
        ref Error? error
    )
        where TFold : struct, INumericFold
    {
        if (value.TryGetError(out var referencedError))
        {
            error ??= referencedError;
        }
        else if (value.TryGetNumber(out var number))
        {
            fold.Accept(number);
        }
        else if (value.TryGetBoolean(out var boolean))
        {
            fold.Accept(boolean ? 1 : 0);
        }
        else if (value.Kind == ComputedValueKind.Text)
        {
            // The *A rule: text in a referenced cell counts as 0 (even numeric-looking text).
            fold.Accept(0);
        }

        // Referenced blanks are still ignored.
    }

    private static void AddDirect<TFold>(in ComputedValue value, ref TFold fold, ref Error? error)
        where TFold : struct, INumericFold
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Error:
                value.TryGetError(out var directError);
                error ??= directError;
                break;

            case ComputedValueKind.Number:
                value.TryGetNumber(out var number);
                fold.Accept(number);
                break;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                fold.Accept(boolean ? 1 : 0);
                break;

            case ComputedValueKind.Blank:
                // Blank ignored.
                break;

            case ComputedValueKind.Text:
                value.TryGetText(out var text);
                if (
                    double.TryParse(
                        text,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var parsed
                    )
                )
                {
                    fold.Accept(parsed);
                }
                else
                {
                    error ??= Error.Value;
                }

                break;
        }
    }
}
