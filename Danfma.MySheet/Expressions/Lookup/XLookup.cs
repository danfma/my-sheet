using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function, IArrayProducer
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    // match_mode: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first-to-last, -1 last-to-first (binary modes not supported).
    // The match engine itself is shared with XMATCH and LOOKUP (see LookupMatching).
    private sealed record SelectionMemo(
        Reference? Reference,
        ArrayOperand? Operand,
        ComputedValue Value,
        bool Matched
    );

    public override ComputedValue Evaluate(EvaluationContext context) =>
        GetSelection(context).Value;

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        return TryResolveReferenceResult(context, out reference, out _, out _);
    }

    internal bool TryResolveReferenceResult(
        EvaluationContext context,
        out Reference reference,
        out ComputedValue unresolvedValue,
        out bool matched
    )
    {
        if (!ReturnsReference(context))
        {
            reference = null!;
            unresolvedValue = default;
            matched = false;
            return false;
        }

        var result = GetSelection(context);
        matched = result.Matched;
        if (result.Reference is { } resolvedReference)
        {
            reference = resolvedReference!;
            unresolvedValue = default;
            return true;
        }

        reference = null!;
        unresolvedValue = result.Value;
        return false;
    }

    internal bool ReturnsReference(EvaluationContext context)
    {
        var returnArgument = Arguments[2];
        return (
                returnArgument is not NameReference name
                || (
                    !context.TryGetName(name.Name, out _)
                    && !context.TryGetArrayBinding(name.Name, out _)
                )
            ) && ArrayEvaluation.IsBareReferenceNode(returnArgument, context);
    }

    internal bool TryBuildSelection(EvaluationContext context, out ArrayOperand operand)
    {
        var selection = GetSelection(context);
        operand = selection.Operand!;
        return ReturnsReference(context) && selection.Operand is not null;
    }

    (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
        ArrayEvaluation.Probe(Arguments[1], context).Succeeds
        && ArrayEvaluation.Probe(Arguments[2], context).Succeeds
            ? (true, true)
            : (false, false);

    bool IArrayProducer.TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand)
    {
        if (!ReturnsReference(context))
        {
            return TryBuildArrayValue(context, out operand);
        }

        var selection = GetSelection(context);
        if (selection.Operand is { } selected)
        {
            operand = selected;
            return true;
        }

        if (selection.Reference is { } reference)
        {
            operand = reference switch
            {
                RangeReference range => ArrayEvaluation.BuildRange(range, context),
                EmptyRangeReference empty => ArrayEvaluation.BuildEmpty(empty),
                CellReference cell => new SingletonArrayOperand(cell.Evaluate(context)),
                _ => new SingletonArrayOperand(ComputedValue.Error(Error.Value)),
            };
            return true;
        }

        operand = new SingletonArrayOperand(selection.Value);
        return true;
    }

    private bool TryBuildArrayValue(EvaluationContext context, out ArrayOperand operand)
    {
        var selection = GetSelection(context);
        operand = selection.Operand ?? new SingletonArrayOperand(selection.Value);
        return true;
    }

    private SelectionMemo GetSelection(EvaluationContext context)
    {
        if (context.TryGetNodeMemo<SelectionMemo>(this, out var cached))
        {
            return cached;
        }

        var selection = BuildSelection(context);
        context.SetNodeMemo(this, selection);
        return selection;
    }

    private SelectionMemo BuildSelection(EvaluationContext context)
    {
        SelectionMemo Result(ComputedValue value, bool matched = false) =>
            new(null, null, value, matched);

        // A missing-sheet lookup/return array is a structural #REF! — distinct from an empty array over an
        // existing sheet, which stays #N/A. Guard before enumerating so it is not swallowed as empty.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return Result(ComputedValue.Error(missing));
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
            return Result(valueError);
        }

        if (
            Arguments[1] is NameReference or TableReference
            && !IsBoundArray(Arguments[1], context)
            && ArraySlotError(Arguments[1], context, Error.NA) is { } lookupArrayError
        )
        {
            return Result(lookupArrayError);
        }

        if (
            Arguments[2] is NameReference or TableReference
            && !IsBoundArray(Arguments[2], context)
            && ArraySlotError(Arguments[2], context, Error.Value) is { } returnArrayError
        )
        {
            return Result(returnArrayError);
        }

        if (!TryBindArray(Arguments[1], context, out var lookupArray))
        {
            return Result(ComputedValue.Error(Error.Value));
        }

        if (!TryBindArray(Arguments[2], context, out var returnArray))
        {
            return Result(ComputedValue.Error(Error.Value));
        }

        var lookupIsColumn = lookupArray.Columns == 1;
        var lookupIsRow = lookupArray.Rows == 1;
        var lookupAxis = lookupIsRow ? ArrayAxis.Columns : ArrayAxis.Rows;
        if (
            (!lookupIsColumn && !lookupIsRow)
            || (
                lookupAxis is ArrayAxis.Rows
                    ? lookupArray.Rows != returnArray.Rows
                    : lookupArray.Columns != returnArray.Columns
            )
        )
        {
            return Result(ComputedValue.Error(Error.Value));
        }

        var matchMode = 0.0;
        if (
            Arguments.Length >= 5
            && Arguments[4].Evaluate(context).CoerceToNumber(out matchMode) is { } matchError
        )
        {
            return Result(ComputedValue.Error(matchError));
        }

        var searchMode = 1.0;
        if (
            Arguments.Length >= 6
            && Arguments[5].Evaluate(context).CoerceToNumber(out searchMode) is { } searchError
        )
        {
            return Result(ComputedValue.Error(searchError));
        }

        if ((int)matchMode == 0 && searchMode >= 0)
        {
            using var lookupValues = lookupArray.Entries().GetEnumerator();
            while (lookupValues.MoveNext())
            {
                if (ValueCoercion.AreEqual(lookupValues.Current.Value, lookup))
                {
                    return returnArray.Select(lookupValues.Current.Position, lookupAxis);
                }
            }

            return Result(NotFound(context));
        }

        var entries = lookupArray.Entries().ToArray();
        var values = entries.Select(entry => entry.Value).ToArray();

        var match = LookupMatching.FindMatch(
            lookup,
            values,
            values.Length,
            (int)matchMode,
            reverse: searchMode < 0
        );

        if (match >= 0)
        {
            return returnArray.Select(entries[match].Position, lookupAxis);
        }

        return Result(NotFound(context));
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

    private static bool IsBoundArray(Expression argument, EvaluationContext context) =>
        argument is NameReference name
        && (
            context.TryGetArrayBinding(name.Name, out _)
            || (
                context.Workbook.DefinedNames.TryGetValue(name.Name, out var definition)
                && ArrayEvaluation.IsArrayEligible(definition, context)
            )
        );

    private static bool TryBindArray(
        Expression argument,
        EvaluationContext context,
        out LookupArray array
    )
    {
        if (
            NamedReferences.TryResolveReference(
                argument,
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            if (reference is CellReference cell)
            {
                var cellStream = new ArrayEvaluation.ArrayStream(
                    new SingletonArrayOperand(cell.Evaluate(context)),
                    1,
                    1
                );
                array = new LookupArray(null, cellStream, 1, 1, context);
                return true;
            }

            if (reference is OpenRangeReference open)
            {
                var rows = (open.RowMax ?? OpenRangeReference.GridMaxRow) - (open.RowMin ?? 1) + 1;
                var columns =
                    (open.ColMax ?? OpenRangeReference.GridMaxColumn) - (open.ColMin ?? 1) + 1;
                array = new LookupArray(reference, default, rows, columns, context);
                return true;
            }

            if (RangeBounds.TryFrom(reference, out var bounds))
            {
                array = new LookupArray(
                    reference,
                    default,
                    bounds.RowCount,
                    bounds.ColumnCount,
                    context
                );
                return true;
            }
        }

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

        array = default;
        return false;
    }

    private readonly struct LookupArray(
        Reference? reference,
        ArrayEvaluation.ArrayStream stream,
        int rows,
        int columns,
        EvaluationContext context
    )
    {
        public int Rows { get; } = rows;
        public int Columns { get; } = columns;
        public int Length => Rows * Columns;

        public IEnumerable<(ComputedValue Value, int Position)> Entries()
        {
            if (reference is null)
            {
                var position = 0;
                foreach (var value in stream)
                {
                    yield return (value, position++);
                }

                yield break;
            }

            if (reference is OpenRangeReference open)
            {
                var workbook = context.Workbook;
                var handle = workbook.ResolveDenseHandle(open.SheetName);
                foreach (var (column, row) in open.PopulatedCells(context))
                {
                    var position = open.IsSingleRow
                        ? open.ColumnPosition(column) - 1
                        : open.RowPosition(row) - 1;
                    yield return (
                        workbook.GetCellValueDense(handle, open.SheetName, column, row),
                        position
                    );
                }

                yield break;
            }

            var cursor = RangeValueCursor.Open(reference, context);
            var referencePosition = 0;
            while (cursor.MoveNext(out var value))
            {
                yield return (value, referencePosition++);
            }
        }

        public SelectionMemo Select(int position, ArrayAxis axis)
        {
            if (reference is not null)
            {
                return SelectReference(reference, position, axis);
            }

            var selected = new AxisSelectionOperand(stream.Operand, axis, [position]);
            return new SelectionMemo(
                null,
                selected,
                selected.At(0, selected.Rows, selected.Columns),
                true
            );
        }

        private SelectionMemo SelectReference(Reference resolved, int position, ArrayAxis axis)
        {
            if (resolved is OpenRangeReference open)
            {
                var row = axis is ArrayAxis.Rows ? open.AbsoluteRow(position + 1) : 0;
                var column = axis is ArrayAxis.Columns ? open.AbsoluteColumn(position + 1) : 0;
                var selectedOpen = SelectOpenReference(open, row, column);
                var value = selectedOpen switch
                {
                    RangeReference range => range.CellComputedValueAt(context, 1, 1),
                    OpenRangeReference selectedRange => selectedRange
                        .ExpandComputedValues(context)
                        .FirstOrDefault(ComputedValue.Blank),
                    _ => ComputedValue.Error(Error.Value),
                };
                return new SelectionMemo(
                    IsMultiCell(selectedOpen) ? selectedOpen : null,
                    null,
                    value,
                    true
                );
            }

            if (!RangeBounds.TryFrom(resolved, out var bounds))
            {
                return new SelectionMemo(null, null, ComputedValue.Error(Error.Value), true);
            }

            var top = axis is ArrayAxis.Rows ? bounds.TopRow + position : bounds.TopRow;
            var bottom = axis is ArrayAxis.Rows ? top : bounds.BottomRow;
            var left = axis is ArrayAxis.Columns ? bounds.LeftColumn + position : bounds.LeftColumn;
            var right = axis is ArrayAxis.Columns ? left : bounds.RightColumn;
            var sheetName = resolved switch
            {
                RangeReference range => range.SheetName,
                EmptyRangeReference empty => empty.SheetName,
                CellReference cell => cell.SheetName,
                _ => context.SheetName ?? string.Empty,
            };
            var selected = new RangeReference(
                new CellAddress(left, top).ToId(),
                new CellAddress(right, bottom).ToId(),
                sheetName
            );

            return new SelectionMemo(
                selected.RowCount > 1 || selected.ColumnCount > 1 ? selected : null,
                null,
                selected.CellComputedValueAt(context, 1, 1),
                true
            );
        }

        private static bool IsMultiCell(Reference reference) =>
            reference switch
            {
                RangeReference range => range.RowCount > 1 || range.ColumnCount > 1,
                OpenRangeReference => true,
                _ => false,
            };

        private static Reference SelectOpenReference(OpenRangeReference open, int row, int column)
        {
            int? rowMin = row == 0 ? open.RowMin : row;
            int? rowMax = row == 0 ? open.RowMax : row;
            int? colMin = column == 0 ? open.ColMin : column;
            int? colMax = column == 0 ? open.ColMax : column;

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
    }
}
