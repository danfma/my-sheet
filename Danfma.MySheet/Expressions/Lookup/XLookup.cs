using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    // match_mode: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first-to-last, -1 last-to-first (binary modes not supported).
    // The match engine itself is shared with XMATCH and LOOKUP (see LookupMatching).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet lookup/return array is a structural #REF! — distinct from an empty array over an
        // existing sheet, which stays #N/A. Guard before enumerating so it is not swallowed as empty.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var lookup = Arguments[0].Evaluate(context);

        // Sweep item 34(b), the VALUE slot: the lookup's error leads the scan (the oracle answers #NAME? for
        // XLOOKUP(NoSuch,A1:A3,B1:B3), and a single error cell's own code even over an if_not_found —
        // ReferencePosition.IsLookupValueError, which also reads a 1x1 range's own cell directly since
        // finding I4: XLOOKUP(A1:A1,B1:B3,C1:C3) is #DIV/0! over A1 = =1/0) instead of the not-found #N/A.
        // The ARRAY slots are the measured exception and keep their own codes: the oracle itself answers
        // #N/A (lookup array) and #VALUE! (return array) for an unresolved name there — see
        // MissingSheetReferenceTests.XLookup_OverAnUnresolvedName_KeepsItsOwnCode_WhereTheOracleDoesToo.
        if (ReferencePosition.IsLookupValueError(Arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        if (
            Arguments[1] is not IArrayProducer
            && ArraySlotError(Arguments[1], context, Error.NA) is { } lookupArrayError
        )
        {
            return lookupArrayError;
        }

        if (
            Arguments[2] is not IArrayProducer
            && ArraySlotError(Arguments[2], context, Error.Value) is { } returnArrayError
        )
        {
            return returnArrayError;
        }

        if (!TryBindArray(Arguments[1], context, out var lookupArray))
        {
            return ComputedValue.Error(Error.Value);
        }

        if (!TryBindArray(Arguments[2], context, out var returnArray))
        {
            return ComputedValue.Error(Error.Value);
        }

        var lookupIsColumn = lookupArray.Columns == 1;
        var lookupIsRow = lookupArray.Rows == 1;
        if (
            (!lookupIsColumn && !lookupIsRow)
            || (
                lookupIsColumn
                    ? lookupArray.Rows != returnArray.Rows
                    : lookupArray.Columns != returnArray.Columns
            )
        )
        {
            return ComputedValue.Error(Error.Value);
        }

        var matchMode = 0.0;
        if (
            Arguments.Length >= 5
            && Arguments[4].Evaluate(context).CoerceToNumber(out matchMode) is { } matchError
        )
        {
            return ComputedValue.Error(matchError);
        }

        var searchMode = 1.0;
        if (
            Arguments.Length >= 6
            && Arguments[5].Evaluate(context).CoerceToNumber(out searchMode) is { } searchError
        )
        {
            return ComputedValue.Error(searchError);
        }

        if ((int)matchMode == 0 && searchMode >= 0)
        {
            using var lookupValues = lookupArray.Values().GetEnumerator();
            using var returnValues = returnArray.Values().GetEnumerator();
            while (lookupValues.MoveNext() && returnValues.MoveNext())
            {
                if (ValueCoercion.AreEqual(lookupValues.Current, lookup))
                {
                    return returnValues.Current;
                }
            }

            return NotFound(context);
        }

        var values = lookupArray.Values().ToArray();

        var match = LookupMatching.FindMatch(
            lookup,
            values,
            values.Length,
            (int)matchMode,
            reverse: searchMode < 0
        );

        if (match >= 0)
        {
            return returnArray.Values().ElementAt(match);
        }

        return NotFound(context);
    }

    // The not-found result: the caller-supplied [if_not_found] when present and not omitted, else #N/A.
    private ComputedValue NotFound(EvaluationContext context) =>
        Arguments.Length >= 4 && Arguments[3] is not BlankValue
            ? Arguments[3].Evaluate(context)
            : ComputedValue.Error(Error.NA);

    private static ComputedValue? ArraySlotError(
        Expression argument,
        EvaluationContext context,
        Error fallback
    )
    {
        if (NamedReferences.TryResolveReference(argument, context, out _, boundOpenRanges: false))
        {
            return null;
        }

        if (
            argument is TableReference
            && argument.Evaluate(context).TryGetError(out var tableError)
        )
        {
            return ComputedValue.Error(tableError);
        }

        return argument is NameReference ? ComputedValue.Error(fallback) : null;
    }

    private static bool TryBindArray(
        Expression argument,
        EvaluationContext context,
        out LookupArray array
    )
    {
        if (argument is IArrayProducer producer)
        {
            producer.TryBuildArrayOperand(context, out var operand);
            var producerStream = new ArrayEvaluation.ArrayStream(
                operand,
                operand.Rows,
                operand.Columns
            );
            array = new LookupArray(
                null,
                producerStream,
                producerStream.Rows,
                producerStream.Columns,
                context
            );
            return true;
        }

        if (ArrayEvaluation.TryEvaluateStream(argument, context, out var stream))
        {
            array = new LookupArray(null, stream, stream.Rows, stream.Columns, context);
            return true;
        }

        if (
            NamedReferences.TryResolveReference(
                argument,
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            if (reference is OpenRangeReference open)
            {
                var rows = (open.RowMax ?? OpenRangeReference.GridMaxRow) - (open.RowMin ?? 1) + 1;
                var columns =
                    (open.ColMax ?? OpenRangeReference.GridMaxColumn) - (open.ColMin ?? 1) + 1;
                array = new LookupArray(argument, default, rows, columns, context);
                return true;
            }

            if (RangeBounds.TryFrom(reference, out var bounds))
            {
                array = new LookupArray(
                    argument,
                    default,
                    bounds.RowCount,
                    bounds.ColumnCount,
                    context
                );
                return true;
            }
        }

        array = default;
        return false;
    }

    private readonly struct LookupArray(
        Expression? reference,
        ArrayEvaluation.ArrayStream stream,
        int rows,
        int columns,
        EvaluationContext context
    )
    {
        public int Rows { get; } = rows;
        public int Columns { get; } = columns;
        public int Length => Rows * Columns;

        public IEnumerable<ComputedValue> Values()
        {
            if (reference is null)
            {
                foreach (var value in stream)
                {
                    yield return value;
                }

                yield break;
            }

            var cursor = RangeValueCursor.Open(reference, context);
            while (cursor.MoveNext(out var value))
            {
                yield return value;
            }
        }
    }
}
