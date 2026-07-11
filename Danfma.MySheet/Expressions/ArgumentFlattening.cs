namespace Danfma.MySheet.Expressions;

/// <summary>
/// Flattens function arguments into a sequence of computed values, expanding range arguments
/// cell-by-cell. Used by variadic functions (CONCAT, TEXTJOIN, COUNTA, …). The <see cref="ComputedValue"/>
/// overloads read straight from the cache (no boxing); the <c>object?</c> ones are boxed views for interop.
/// </summary>
internal static class ArgumentFlattening
{
    public static IEnumerable<ComputedValue> FlattenComputedValues(
        Expression[] arguments,
        EvaluationContext context
    )
    {
        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        yield return value;
                    }

                    break;

                case OpenRangeReference open:
                    foreach (var value in open.ExpandComputedValues(context))
                    {
                        yield return value;
                    }

                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        yield return value;
                    }

                    break;

                // Phase 2 audit (shared-formula delta production): mirrors the RangeReference case above so a
                // range argument written INSIDE a shared-formula master (an anchored node, e.g.
                // CONCAT/TEXTJOIN/COUNTA over a relative range) expands its cells instead of falling to the
                // `default` branch's `argument.Evaluate(context)`, which for an AnchoredRangeReference always
                // yields #VALUE! (a range has no scalar value, mirroring RangeReference.Evaluate). A bare
                // AnchoredCellReference needs no case here: its Evaluate already dereferences correctly (same
                // as a plain CellReference, which likewise has no explicit case and falls to `default` below).
                case AnchoredRangeReference anchoredRange:
                    foreach (
                        var value in anchoredRange
                            .ToRangeReference(context)
                            .ExpandComputedValues(context)
                    )
                    {
                        yield return value;
                    }

                    break;

                default:
                    var computed = argument.Evaluate(context);

                    if (computed.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in computed.EnumerateValues(context))
                        {
                            yield return cellValue;
                        }
                    }
                    else
                    {
                        yield return computed;
                    }

                    break;
            }
        }
    }

    /// <summary>
    /// Expands a single argument into an ordered list of computed values (a range yields one entry per
    /// cell). Used by the conditional aggregations to align parallel ranges by position.
    /// </summary>
    public static List<ComputedValue> ExpandComputedValues(
        Expression argument,
        EvaluationContext context
    )
    {
        // Phase 2 audit (shared-formula delta production): a range argument written INSIDE a shared-formula
        // master is an anchored node (its own $-anchors preserved, delta applied at evaluation time), not a
        // plain RangeReference. Resolving it to its concrete, delta-applied twin UP FRONT — rather than adding
        // a mirrored `case AnchoredRangeReference` below — lets the sizing hint and the switch below treat it
        // exactly like an ordinary RangeReference, with no logic duplicated (a bare AnchoredCellReference needs
        // no such resolution: its Evaluate already dereferences correctly through the `default` branch below,
        // same as a plain CellReference).
        if (argument is AnchoredRangeReference anchoredRange)
        {
            argument = anchoredRange.ToRangeReference(context);
        }

        // A closed rectangle has known bounds: size the buffer to its exact cell count so the hot per-cell fill
        // never pays a List doubling + recopy. Open ranges/unions/scalars have no known count up front — keep the
        // default-capacity List (no heroics where the bound is unknown). GetBounds() is called ONCE here (rather
        // than once per RowCount/ColumnCount property read) so sizing the list costs a single corner parse.
        List<ComputedValue> values;

        if (argument is RangeReference closed)
        {
            var bounds = closed.GetBounds();
            values = new List<ComputedValue>(bounds.RowCount * bounds.ColumnCount);
        }
        else
        {
            values = new List<ComputedValue>();
        }

        switch (argument)
        {
            case RangeReference range:
                foreach (var value in range.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                break;

            case OpenRangeReference open:
                foreach (var value in open.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                break;

            case UnionReference union:
                foreach (var value in union.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                break;

            default:
                var computed = argument.Evaluate(context);

                if (computed.Kind == ComputedValueKind.Reference)
                {
                    foreach (var cellValue in computed.EnumerateValues(context))
                    {
                        values.Add(cellValue);
                    }
                }
                else
                {
                    values.Add(computed);
                }

                break;
        }

        return values;
    }

    /// <summary>
    /// The Layer-2 view of a single range/vector argument: the shared per-epoch snapshot's values when the
    /// argument is a big populated range (so the O(N) cell re-read is paid once for the whole epoch), else
    /// the ordinary per-call expansion. The returned list is identical, cell for cell, to
    /// <see cref="ExpandComputedValues(Expression, EvaluationContext)"/> — the snapshot is built through the
    /// same enumeration — so a linear scan over it stays bit-for-bit equivalent. <paramref name="snapshot"/>
    /// is non-null exactly when the cached derived accelerators (hash, sorted index, …) are available.
    /// </summary>
    public static IReadOnlyList<ComputedValue> ExpandCached(
        Expression argument,
        EvaluationContext context,
        out RangeSnapshot? snapshot
    )
    {
        snapshot = argument is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        return snapshot?.Values
            ?? (IReadOnlyList<ComputedValue>)ExpandComputedValues(argument, context);
    }
}
