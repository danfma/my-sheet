using System.Text;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

// Onda 3 — lookup & reference escalar: CHOOSE (lazy), HLOOKUP (espelho horizontal do VLOOKUP),
// LOOKUP (formas vetor e array), COLUMN/COLUMNS (espelhos de ROW/ROWS), XMATCH (mesmo engine de
// match do XLOOKUP), ADDRESS, AREAS (checagem sintática) e FORMULATEXT (reusa o FormulaWriter).

[MemoryPackable]
public sealed partial record Choose(Expression[] Arguments) : Function
{
    // CHOOSE(index_num, value1, [value2], …) — lazy like IF: the index is evaluated and truncated,
    // then ONLY the chosen value argument is evaluated. Out of range -> #VALUE! (per the docs).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var index) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var position = (int)Math.Truncate(index);

        if (position < 1 || position > Arguments.Length - 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        var chosen = Arguments[position];

        // A chosen range stays a reference, so range-aware consumers (SUM(CHOOSE(2,A1:A10,B1:B10)))
        // expand it — the same technique OFFSET uses for its multi-cell results.
        return chosen is RangeReference or OpenRangeReference or UnionReference
            ? ComputedValue.Reference((Reference)chosen)
            : chosen.Evaluate(context);
    }
}

[MemoryPackable]
public sealed partial record HLookup(Expression[] Arguments) : Function
{
    // HLOOKUP(lookup, table, row_index, [range_lookup]) — the horizontal mirror of VLOOKUP:
    // searches the table's first ROW and returns from row_index in the matching column.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet table is a structural #REF! — a BOUNDED ghost range would otherwise scan its cells,
        // skip the per-cell #REF! keys, and degrade to #N/A. Guard before the table is inspected.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // The table may be written directly or through a defined name that stands for a range.
        if (
            !NamedReferences.TryResolveReference(Arguments[1], context, out var reference)
            || reference is not RangeReference table
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

        var lookup = Arguments[0].Evaluate(context);

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var rowIndex) is { } rowError)
        {
            return ComputedValue.Error(rowError);
        }

        // Per the docs: row_index_num < 1 -> #VALUE!, greater than the table's rows -> #REF!.
        if (rowIndex < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (rowIndex > table.RowCount)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var approximate = true;

        if (
            Arguments.Length == 4
            && Arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        // The first row is a sub-range of the table; its per-epoch snapshot serves the key search O(1)
        // (exact) / O(log n) (approximate). A 1-based snapshot position IS the 1-based table column, because
        // the key row enumerates left-to-right. A small table keeps the linear scan.
        var keyRow = new RangeReference(
            new CellAddress(table.LeftColumn, table.TopRow).ToId(),
            new CellAddress(table.LeftColumn + table.ColumnCount - 1, table.TopRow).ToId(),
            table.SheetName
        );
        var keySnapshot = context.Workbook.TryGetRangeSnapshot(keyRow, context);

        var matchColumn = -1;

        if (approximate)
        {
            // Largest first-row key <= lookup, assuming the row is sorted ascending. Cross-type
            // ordering (ValueCoercion.Compare) lets text keys sort lexicographically, exactly like
            // the <= operator — not only numeric keys.
            if (keySnapshot is not null)
            {
                var position = keySnapshot.ApproximateAscendingPosition(lookup);
                matchColumn = position >= 1 ? position : -1;
            }
            else
            {
                for (var column = 1; column <= table.ColumnCount; column++)
                {
                    var key = table.CellComputedValueAt(context, 1, column);
                    if (key.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
                    {
                        continue;
                    }

                    if (ValueCoercion.Compare(key, lookup) <= 0)
                    {
                        matchColumn = column;
                    }
                }
            }
        }
        else
        {
            if (keySnapshot is not null)
            {
                switch (keySnapshot.TryExactPosition(lookup, out var position))
                {
                    case ExactMatchOutcome.Found:
                        matchColumn = position;
                        break;
                    case ExactMatchOutcome.NotFound:
                        return ComputedValue.Error(Error.NA);
                }
            }

            if (matchColumn < 1)
            {
                for (var column = 1; column <= table.ColumnCount; column++)
                {
                    if (ValueCoercion.AreEqual(table.CellComputedValueAt(context, 1, column), lookup))
                    {
                        matchColumn = column;
                        break;
                    }
                }
            }
        }

        return matchColumn >= 1
            ? table.CellComputedValueAt(context, (int)rowIndex, matchColumn)
            : ComputedValue.Error(Error.NA);
    }
}

[MemoryPackable]
public sealed partial record Lookup(Expression[] Arguments) : Function
{
    // LOOKUP(value, lookup_vector, [result_vector]) — vector form — and LOOKUP(value, array) —
    // array form: an area wider than it is tall searches the first ROW and returns from the last
    // ROW; square or taller searches the first COLUMN and returns from the last COLUMN (per the
    // docs). Always approximate: exact match first, otherwise the largest value <= lookup
    // (cross-type ordering, shared with XLOOKUP's -1 mode); below the smallest -> #N/A.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet vector/array is a structural #REF!, before any match is attempted.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var lookup = Arguments[0].Evaluate(context);

        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        if (Arguments.Length == 2 && Arguments[1] is RangeReference array)
        {
            var byRow = array.ColumnCount > array.RowCount;
            var count = byRow ? array.ColumnCount : array.RowCount;
            var keys = new List<ComputedValue>(count);
            var results = new List<ComputedValue>(count);

            for (var i = 1; i <= count; i++)
            {
                keys.Add(
                    byRow
                        ? array.CellComputedValueAt(context, 1, i)
                        : array.CellComputedValueAt(context, i, 1)
                );
                results.Add(
                    byRow
                        ? array.CellComputedValueAt(context, array.RowCount, i)
                        : array.CellComputedValueAt(context, i, array.ColumnCount)
                );
            }

            return Find(lookup, keys, results);
        }

        var lookupVector = ArgumentFlattening.ExpandCached(Arguments[1], context, out _);
        var resultVector = Arguments.Length == 3
            ? ArgumentFlattening.ExpandCached(Arguments[2], context, out _)
            : lookupVector;

        return Find(lookup, lookupVector, resultVector);
    }

    private static ComputedValue Find(
        in ComputedValue lookup,
        IReadOnlyList<ComputedValue> keys,
        IReadOnlyList<ComputedValue> results
    )
    {
        var count = Math.Min(keys.Count, results.Count);
        var match = LookupMatching.FindMatch(lookup, keys, count, matchMode: -1, reverse: false);

        return match >= 0 ? results[match] : ComputedValue.Error(Error.NA);
    }
}

[MemoryPackable]
public sealed partial record Column(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        // A reference to a missing sheet is a structural #REF!, not a column position.
        ReferenceGuard.MissingSheet(Arguments, context) is { } missing
            ? ComputedValue.Error(missing)
            : Arguments switch
            {
                [CellReference cell] => ComputedValue.Number(CellAddress.Parse(cell.Id).Column),
                [RangeReference range] => ComputedValue.Number(range.LeftColumn),
                // COLUMN() with no argument uses the cell currently being evaluated, when one is known.
                [] when context.CellId is { } id => ComputedValue.Number(CellAddress.Parse(id).Column),
                _ => ComputedValue.Error(Error.Value),
            };
}

[MemoryPackable]
public sealed partial record Columns(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its columns; a whole-column/row reference uses the
    // exact structural count on a bounded column axis (COLUMNS(A:C) = 3) and the populated extent on an
    // open one; anything else is 1.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        // A reference to a missing sheet is a structural #REF!, not an empty (0-column) extent.
        ReferenceGuard.MissingSheet(Arguments[0], context) is { } missing
            ? ComputedValue.Error(missing)
            : ComputedValue.Number(
                NamedReferences.TryResolveReference(Arguments[0], context, out var reference, boundOpenRanges: false)
                    ? reference switch
                    {
                        RangeReference range => range.ColumnCount,
                        OpenRangeReference open => open.ColumnExtent(context),
                        _ => 1.0,
                    }
                    : 1.0
            );
}

[MemoryPackable]
public sealed partial record XMatch(Expression[] Arguments) : Function
{
    // XMATCH(lookup, array, [match_mode], [search_mode]) — the 1-based POSITION of the match, with
    // the same mode semantics as XLOOKUP (shared LookupMatching engine): match_mode 0 exact
    // (default), -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard; search_mode 1
    // first-to-last (default), -1 last-to-first (binary modes not supported). No match -> #N/A.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet array is a structural #REF! — distinct from an empty existing array (still #N/A).
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var lookup = Arguments[0].Evaluate(context);
        var array = ArgumentFlattening.ExpandCached(Arguments[1], context, out var snapshot);

        var matchMode = 0.0;
        if (
            Arguments.Length >= 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out matchMode) is { } matchError
        )
        {
            return ComputedValue.Error(matchError);
        }

        var searchMode = 1.0;
        if (
            Arguments.Length >= 4
            && Arguments[3].Evaluate(context).CoerceToNumber(out searchMode) is { } searchError
        )
        {
            return ComputedValue.Error(searchError);
        }

        // Forward exact (the default) → O(1) via the value→first-position hash; every other mode (reverse,
        // approximate, wildcard) keeps the shared LookupMatching engine over the cached array.
        if ((int)matchMode == 0 && searchMode >= 0 && snapshot is not null)
        {
            switch (snapshot.TryExactPosition(lookup, out var hashPosition))
            {
                case ExactMatchOutcome.Found:
                    return ComputedValue.Number(hashPosition);
                case ExactMatchOutcome.NotFound:
                    return ComputedValue.Error(Error.NA);
            }
        }

        var match = LookupMatching.FindMatch(
            lookup,
            array,
            array.Count,
            (int)matchMode,
            reverse: searchMode < 0
        );

        return match >= 0 ? ComputedValue.Number(match + 1) : ComputedValue.Error(Error.NA);
    }
}

[MemoryPackable]
public sealed partial record Address(Expression[] Arguments) : Function
{
    // ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]) — the cell address as TEXT.
    // abs_num: 1 $C$2 (default), 2 C$2, 3 $C2, 4 C2. a1 = FALSE renders the documented absolute
    // R1C1 form (R2C3); the relative R1C1 forms (R2C[3]) are not modeled, so abs_num != 1 with
    // a1 = FALSE -> #VALUE! (declared limitation). sheet_text is prefixed with the same quoting
    // rule the FormulaWriter uses ('EXCEL SHEET'!R2C3).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var row) is { } rowError)
        {
            return ComputedValue.Error(rowError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var column) is { } columnError)
        {
            return ComputedValue.Error(columnError);
        }

        var absNum = 1.0;
        if (
            Arguments.Length >= 3
            && Arguments[2] is not BlankValue
            && Arguments[2].Evaluate(context).CoerceToNumber(out absNum) is { } absError
        )
        {
            return ComputedValue.Error(absError);
        }

        var a1 = true;
        if (
            Arguments.Length >= 4
            && Arguments[3] is not BlankValue
            && Arguments[3].Evaluate(context).CoerceToBool(out a1) is { } a1Error
        )
        {
            return ComputedValue.Error(a1Error);
        }

        var rowNumber = (int)Math.Truncate(row);
        var columnNumber = (int)Math.Truncate(column);
        var abs = (int)Math.Truncate(absNum);

        if (rowNumber < 1 || columnNumber < 1 || abs is < 1 or > 4)
        {
            return ComputedValue.Error(Error.Value);
        }

        string body;

        if (a1)
        {
            var columnDollar = abs is 1 or 3 ? "$" : string.Empty;
            var rowDollar = abs is 1 or 2 ? "$" : string.Empty;

            body = $"{columnDollar}{ColumnLetters(columnNumber)}{rowDollar}{rowNumber}";
        }
        else if (abs == 1)
        {
            body = $"R{rowNumber}C{columnNumber}";
        }
        else
        {
            // Relative R1C1 (R2C[3]) needs an origin the text form does not carry — not modeled.
            return ComputedValue.Error(Error.Value);
        }

        if (Arguments.Length < 5 || Arguments[4] is BlankValue)
        {
            return ComputedValue.Text(body);
        }

        if (Arguments[4].Evaluate(context).CoerceToText(out var sheetText) is { } sheetError)
        {
            return ComputedValue.Error(sheetError);
        }

        var prefix = FormulaWriter.IsSimpleSheetName(sheetText)
            ? sheetText
            : "'" + sheetText.Replace("'", "''") + "'";

        return ComputedValue.Text(prefix + "!" + body);
    }

    private static string ColumnLetters(int column)
    {
        var builder = new StringBuilder();

        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            column = (column - 1) / 26;
        }

        return builder.ToString();
    }
}

[MemoryPackable]
public sealed partial record Areas(Expression[] Arguments) : Function
{
    // AREAS(reference) — the number of areas (contiguous ranges or single cells) in the reference.
    // A syntactic check on the argument node (a defined name resolves to the reference it stands for),
    // like ISREF: a union counts its areas (recursively, for nested unions), any other reference is one
    // area, a non-reference -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        NamedReferences.TryResolveReference(Arguments[0], context, out var reference)
            ? reference switch
            {
                UnionReference union => ComputedValue.Number(CountAreas(union)),
                _ => ComputedValue.Number(1),
            }
            : ComputedValue.Error(Error.Value);

    private static int CountAreas(UnionReference union)
    {
        var count = 0;

        foreach (var area in union.Areas)
        {
            count += area is UnionReference nested ? CountAreas(nested) : 1;
        }

        return count;
    }
}

[MemoryPackable]
public sealed partial record FormulaText(Expression[] Arguments) : Function
{
    // FORMULATEXT(reference) — the referenced cell's formula as TEXT, "=" included, un-parsed by
    // the FormulaWriter in the REFERENCED cell's sheet context (its local references stay
    // unqualified). A range reads its top-left cell. A cell holding a plain literal
    // (ValueExpression) or nothing -> #N/A; a non-reference argument -> #VALUE! (per the docs).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var (sheetName, cellId) = Arguments[0] switch
        {
            CellReference cell => (cell.SheetName, cell.Id),
            RangeReference range => (range.SheetName, range.StartId),
            _ => (null, null),
        };

        if (sheetName is null || cellId is null)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (
            !context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
            || !sheet.TryGetValue(cellId, out var expression)
            || expression is ValueExpression
        )
        {
            return ComputedValue.Error(Error.NA);
        }

        return ComputedValue.Text("=" + expression.ToFormula(sheetName));
    }
}
