using System.Diagnostics.CodeAnalysis;

namespace Danfma.MySheet.Expressions.Lookup;

internal static class LookupTable
{
    public static bool TryResolveTable(
        Expression[] arguments,
        EvaluationContext context,
        [NotNullWhen(true)] out Reference? reference,
        out ComputedValue answer
    )
    {
        if (!NamedReferences.TryResolveReference(arguments[1], context, out reference))
        {
            var tableValue = arguments[1].Evaluate(context);

            if (
                arguments[1] is not NameReference and not TableReference
                && !ArrayEvaluation.IsArrayEligible(arguments[1], context)
                && tableValue.TryGetError(out _)
            )
            {
                answer = ComputedValue.Error(Error.NA);
                return false;
            }

            answer = tableValue is { Kind: not ComputedValueKind.Error }
                ? LookupScalarTable(arguments, tableValue, context)
                : ReferencePosition.Unresolved(
                    arguments[1],
                    context,
                    ComputedValue.Error(Error.Ref)
                );
            return false;
        }

        if (reference is not CellReference cell)
        {
            answer = default;
            return true;
        }

        var value = cell.Evaluate(context);
        if (arguments[1] is NameReference && value.TryGetError(out _))
        {
            answer = value;
            return false;
        }

        answer = value.TryGetError(out _)
            ? DirectErrorCellTable(arguments, context)
            : LookupScalarTable(arguments, value, context);
        reference = null;
        return false;
    }

    public static ComputedValue LookupScalarTable(
        Expression[] arguments,
        ComputedValue tableValue,
        EvaluationContext context
    )
    {
        var lookup = arguments[0].Evaluate(context);
        if (ReferencePosition.IsLookupValueError(arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        if (arguments[2].Evaluate(context).CoerceToNumber(out var index) is { } indexError)
        {
            return ComputedValue.Error(indexError);
        }

        if (index < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (index > 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var approximate = true;
        if (
            arguments.Length == 4
            && arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        return lookup.TryGetError(out _) || !ValueCoercion.AreEqual(tableValue, lookup)
            ? ComputedValue.Error(Error.NA)
            : tableValue;
    }

    public static ComputedValue DirectErrorCellTable(
        Expression[] arguments,
        EvaluationContext context
    )
    {
        var approximate = true;
        if (
            arguments.Length == 4
            && arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        return approximate ? arguments[1].Evaluate(context) : ComputedValue.Error(Error.NA);
    }
}

internal readonly struct LookupGrid
{
    private readonly ArrayEvaluation.ArrayStream _array;
    private readonly RangeReference? _range;
    private readonly RangeBounds _bounds;
    private readonly Workbook? _workbook;
    private readonly int _handle;

    public LookupGrid(ArrayEvaluation.ArrayStream array)
    {
        _array = array;
        _range = null;
        _bounds = default;
        _workbook = null;
        _handle = 0;
        Rows = array.Rows;
        Columns = array.Columns;
    }

    public LookupGrid(RangeReference? range, RangeBounds bounds, EvaluationContext context)
    {
        _array = default;
        _range = range;
        _bounds = bounds;
        _workbook = context.Workbook;
        _handle = range is null ? 0 : context.Workbook.ResolveDenseHandle(range.SheetName);
        Rows = range is null ? 0 : bounds.RowCount;
        Columns = bounds.ColumnCount;
    }

    public int Rows { get; }
    public int Columns { get; }

    public static bool TryCreate(
        Expression[] arguments,
        EvaluationContext context,
        out LookupGrid grid,
        out ComputedValue answer
    )
    {
        if (
            arguments[1] is not Reference
            && ArrayEvaluation.TryStream(arguments[1], context, out var array)
        )
        {
            grid = new LookupGrid(array);
            answer = default;
            return true;
        }

        if (!LookupTable.TryResolveTable(arguments, context, out var reference, out answer))
        {
            grid = default;
            return false;
        }

        if (!RangeBounds.TryFrom(reference, out var bounds))
        {
            grid = default;
            answer = ComputedValue.Error(Error.Ref);
            return false;
        }

        grid = new LookupGrid(reference as RangeReference, bounds, context);
        answer = default;
        return true;
    }

    public ComputedValue At(int row, int column) =>
        _range is null
            ? _array.ElementAt((row - 1) * Columns + column - 1)
            : _range.CellComputedValueAt(_workbook!, _handle, _bounds, row, column);

    public RangeSnapshot? TryGetKeySnapshot(EvaluationContext context, bool vertical)
    {
        var count = vertical ? Rows : Columns;
        if (
            _range is null
            || _workbook!.RangeCacheDisabled
            || count < Workbook.RangeCacheMinimumCells
        )
        {
            return null;
        }

        var keys = vertical
            ? new RangeReference(
                new CellAddress(_bounds.LeftColumn, _bounds.TopRow).ToId(),
                new CellAddress(_bounds.LeftColumn, _bounds.BottomRow).ToId(),
                _range.SheetName
            )
            : new RangeReference(
                new CellAddress(_bounds.LeftColumn, _bounds.TopRow).ToId(),
                new CellAddress(_bounds.RightColumn, _bounds.TopRow).ToId(),
                _range.SheetName
            );
        return _workbook.TryGetRangeSnapshot(keys, context);
    }
}
