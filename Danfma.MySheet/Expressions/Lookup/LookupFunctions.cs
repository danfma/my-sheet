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
    //
    // The chosen argument is a BINDING SITE (Phase 11c, ArrayBindings.Capture): a chosen range stays a
    // reference, so range-aware consumers (SUM(CHOOSE(2,A1:A10,B1:B10))) expand it — the same technique
    // OFFSET uses for its multi-cell results — and a chosen computed ARRAY is built once and read here as
    // its TOP-LEFT, the @ rule a bare producer follows: =CHOOSE(1,FILTER(A1:A3,A1:A3>0)) is 5 and
    // =CHOOSE(1,A1:A3*2) is 10 (Aspose.Cells 26.6.0, 2026-09-11, H20 — the second in the array-entered
    // column, where plain entry answers #VALUE!), against the #VALUE! CaptureValue's own reading gave the
    // operator. A CONSUMER of the same node streams the whole array instead, through
    // ArrayEvaluation.TryBuildChoose. Since sweep item 32 the chosen BARE-REFERENCE branch carries the
    // reference it resolves to for a RANGE (<see cref="CaptureChosen"/>); a SINGLE-CELL branch instead hands
    // this scalar reading the cell's own VALUE (the fix wave's C1 finding), exactly as
    // INDEX(A1:A3,2)+1/OFFSET(A1,0,0)+1 already read a producer's single-cell result — the referenced-cell
    // rule (text skipped: SUM(CHOOSE(1,MyCell)) 0) still applies because NumericAggregation resolves the
    // node to its reference first (NumericAggregation.Fold's If/Choose arm) rather than reading a
    // Reference-kind scalar here.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryChoose(context, out var chosen) ?? CaptureChosen(chosen, context);

    /// <summary>
    /// The chosen branch's value, shared by <see cref="Evaluate"/> and the mini-CSE's <c>Choose</c> arm's
    /// scalar (<c>!isArray</c>) wrap (<c>ArrayEvaluation.TryBuildChoose</c>) so the two cannot drift: a
    /// bare-reference branch resolves to the reference it denotes (boundOpenRanges:false, so an open range
    /// stays itself and the consumers' expansion rules apply); a SINGLE CELL among those (a
    /// <see cref="CellReference"/>) instead hands back the cell's own value, so a scalar consumer
    /// (<c>IF(A1&gt;0,CHOOSE(1,B1),C1)*2</c>) sees a number, not a reference wrapper it cannot coerce — the
    /// fix wave's C1 finding, sweep item 32; anything else goes through
    /// <see cref="ArrayBindings.Capture"/>'s top-left exactly as before — a chosen PRODUCER still reads as
    /// its first element. The array-ELIGIBLE build (<c>TryBuildChoose</c>'s <c>isArray</c> branch) does not
    /// call this method — it reads <see cref="ArrayBindings.Capture"/> directly, so a single-cell chosen
    /// branch under a computed sibling still carries its REFERENCE for the range-aware consumer that made
    /// the node eligible.
    /// </summary>
    internal static ComputedValue CaptureChosen(Expression chosen, EvaluationContext context) =>
        ArrayEvaluation.IsBareReferenceNode(chosen, context)
        && NamedReferences.TryResolveReference(
            chosen,
            context,
            out var resolved,
            boundOpenRanges: false
        )
            ? resolved is CellReference cell
                ? cell.Evaluate(context)
                : ComputedValue.Reference(resolved)
            : ArrayBindings.Capture(chosen, context).TopLeft;

    /// <summary>
    /// The index rule, shared by <see cref="Evaluate"/> and the mini-CSE's <c>Choose</c> arm
    /// (<c>ArrayEvaluation.TryBuildChoose</c>) so the two cannot drift: <c>index_num</c> is evaluated ONCE
    /// and truncated; an uncoercible index propagates its own error and an index outside
    /// <c>1..Arguments.Length - 1</c> is <c>#VALUE!</c> (the CHOOSE page's rule). Returns that error, or
    /// <c>null</c> with <paramref name="chosen"/> set to the chosen argument — which neither this method nor
    /// its callers' shared code evaluates, so CHOOSE stays lazy.
    /// </summary>
    internal ComputedValue? TryChoose(EvaluationContext context, out Expression chosen)
    {
        chosen = null!;

        // Final-review fix wave, finding I3: EvaluateConditionOnce (shared with If's condition — see its
        // remarks on EvaluationContext) memoizes the index draw for THIS node within one cell's evaluation,
        // so TryResolveReference below (ArrayBindings.Shape's probe) and this method (ArrayBindings.Capture's
        // build) agree on which branch a VOLATILE index picks instead of drawing it twice, independently.
        if (context.EvaluateConditionOnce(Arguments[0]).CoerceToNumber(out var index) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var position = (int)Math.Truncate(index);

        if (position < 1 || position > Arguments.Length - 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        chosen = Arguments[position];
        return null;
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;
        if (context.EvaluateConditionOnce(Arguments[0]).CoerceToNumber(out var index) is not null)
        {
            return false;
        }

        var slot = (int)index;
        if (slot < 1 || slot >= Arguments.Length)
        {
            return false;
        }

        return Arguments[slot].TryResolveReference(context, out reference);
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

        // The table may be written directly or through a defined name that stands for a range. When it
        // does not resolve, the node's OWN error is the answer (sweep item 34(b), the mirror of
        // VLookup's arm). A plain scalar is the measured exception: it is a 1x1 table, while a directly
        // computed error follows the CSE not-found reading.
        if (!NamedReferences.TryResolveReference(Arguments[1], context, out var reference))
        {
            var tableValue = Arguments[1].Evaluate(context);

            if (
                Arguments[1] is not NameReference and not TableReference
                && !ArrayEvaluation.IsArrayEligible(Arguments[1], context)
                && tableValue.TryGetError(out _)
            )
            {
                return ComputedValue.Error(Error.NA);
            }

            if (tableValue is { Kind: not ComputedValueKind.Error })
            {
                return LookupScalarTable(tableValue, context);
            }

            return ReferencePosition.Unresolved(
                Arguments[1],
                context,
                ComputedValue.Error(Error.Ref)
            );
        }

        if (reference is CellReference cell)
        {
            var value = cell.Evaluate(context);
            if (Arguments[1] is NameReference && value.TryGetError(out _))
            {
                return value;
            }

            return LookupScalarTable(value, context);
        }

        // Bounds are resolved ONCE here, not re-parsed on every column of the linear fallback scan below. This
        // is a pure, side-effect-free read of the table's own corners, so hoisting it ahead of the argument
        // evaluation below does not change Arguments' evaluation order. A zero-row rectangle (sweep item 33)
        // has bounds too; it leaves the function at the not-found check below.
        var workbook = context.Workbook;

        if (!RangeBounds.TryFrom(reference, out var bounds))
        {
            return ComputedValue.Error(Error.Ref);
        }

        var lookup = Arguments[0].Evaluate(context);

        if (
            (
                Arguments[0] is not RangeReference range
                || (range.RowCount == 1 && range.ColumnCount == 1)
            )
            && lookup.Kind == ComputedValueKind.Error
            && ReferencePosition.IsLookupValueError(
                Arguments[0],
                lookup,
                context,
                out var valueError
            )
        )
        {
            return valueError;
        }

        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var rowIndex) is { } rowError)
        {
            return ComputedValue.Error(rowError);
        }

        // Per the docs: row_index_num < 1 -> #VALUE!, greater than the table's rows -> #REF!.
        if (rowIndex < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        // Sweep item 33: a zero-row rectangle has no first ROW to search, so the lookup is not found before
        // its row index is ever measured against a height of 0 (oracle, array-entered:
        // HLOOKUP(1,<the empty band>,1,FALSE) is #N/A; typed plain, Aspose throws inside its calculation).
        if (reference is not RangeReference table)
        {
            return ComputedValue.Error(Error.NA);
        }

        var handle = workbook.ResolveDenseHandle(table.SheetName);

        if (rowIndex > bounds.RowCount)
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

        // The first row is a sub-range of the table; its per-epoch snapshot serves the key search O(1)
        // (exact) / O(log n) (approximate). A 1-based snapshot position IS the 1-based table column, because
        // the key row enumerates left-to-right. Built LAZILY: TryGetRangeSnapshot would reject a table below
        // the cache's own size threshold anyway, so a small table (the common case) skips the RangeReference +
        // two CellAddress.ToId allocations entirely and goes straight to the linear scan.
        RangeSnapshot? keySnapshot = null;

        if (!workbook.RangeCacheDisabled && bounds.ColumnCount >= Workbook.RangeCacheMinimumCells)
        {
            var keyRow = new RangeReference(
                new CellAddress(bounds.LeftColumn, bounds.TopRow).ToId(),
                new CellAddress(bounds.RightColumn, bounds.TopRow).ToId(),
                table.SheetName
            );
            keySnapshot = workbook.TryGetRangeSnapshot(keyRow, context);
        }

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
                for (var column = 1; column <= bounds.ColumnCount; column++)
                {
                    var key = table.CellComputedValueAt(workbook, handle, bounds, 1, column);
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
                for (var column = 1; column <= bounds.ColumnCount; column++)
                {
                    if (
                        ValueCoercion.AreEqual(
                            table.CellComputedValueAt(workbook, handle, bounds, 1, column),
                            lookup
                        )
                    )
                    {
                        matchColumn = column;
                        break;
                    }
                }
            }
        }

        return matchColumn >= 1
            ? table.CellComputedValueAt(workbook, handle, bounds, (int)rowIndex, matchColumn)
            : ComputedValue.Error(Error.NA);
    }

    private ComputedValue LookupScalarTable(ComputedValue tableValue, EvaluationContext context)
    {
        if (Arguments[2].Evaluate(context).CoerceToNumber(out var rowIndex) is { } rowError)
        {
            return ComputedValue.Error(rowError);
        }

        if (rowIndex < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (rowIndex > 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var lookup = Arguments[0].Evaluate(context);

        var approximate = true;
        if (
            Arguments.Length == 4
            && Arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        return lookup.TryGetError(out _) || !ValueCoercion.AreEqual(tableValue, lookup)
            ? ComputedValue.Error(Error.NA)
            : tableValue;
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

        // Sweep item 34(b): the vector slot — an unresolved node reports its own error (the oracle
        // answers #NAME? for LOOKUP(1,NoSuch)) instead of streaming it as the one key the scan discards
        // into #N/A.
        if (ReferencePosition.TryUnresolvedError(Arguments[1], context, out var unresolved))
        {
            return unresolved;
        }

        var lookupVector = ArgumentFlattening.ExpandCached(Arguments[1], context, out _);
        var resultVector =
            Arguments.Length == 3
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
        ReferenceGuard.MissingSheet(Arguments, context)
            is { } missing
            ? ComputedValue.Error(missing)
            : Arguments switch
            {
                [CellReference cell] => ComputedValue.Number(CellAddress.Parse(cell.Id).Column),
                [RangeReference range] => ComputedValue.Number(range.LeftColumn),
                // Phase 2 audit (shared-formula delta production): mirrors Row's identical fix — see that
                // file for the full rationale (a COLUMN(ref) argument inside a shared-formula master is an
                // anchored node, not a plain CellReference/RangeReference).
                [AnchoredCellReference anchoredCell] => ComputedValue.Number(
                    anchoredCell.Effective(context).Column
                ),
                [AnchoredRangeReference anchoredRange] => ComputedValue.Number(
                    anchoredRange.ToRangeReference(context).LeftColumn
                ),
                // COLUMN() with no argument uses the cell currently being evaluated, when one is known.
                [] when context.CellId is { } id => ComputedValue.Number(
                    CellAddress.Parse(id).Column
                ),
                // Terminal fallback over any reference-producing argument — the mirror of Row.cs's arm; see
                // that file (and ReferencePosition) for the rationale. Placed after the single-argument
                // syntactic arms above, which it would otherwise subsume (the compiler rejects that
                // ordering); a zero-argument COLUMN() cannot reach it — [var only] requires exactly one
                // argument.
                [var only] => ReferencePosition.Column(only, context),
                _ => ComputedValue.Error(Error.Value),
            };
}

[MemoryPackable]
public sealed partial record Columns(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its columns; a whole-column/row reference uses the
    // exact structural count on a bounded column axis (COLUMNS(A:C) = 3) and the populated extent on an
    // open one; anything else is 1 — except a reference that FAILED to resolve, which reports its own error.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF!, not an empty (0-column) extent — the same
        // two passes ROWS runs (syntactic here, resolved target inside TryResolve); see Rows.cs.
        if (ReferenceGuard.MissingSheet(Arguments[0], context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // The computed-array gate, in the same position as in ROWS (after the syntactic guard, before the
        // resolution whose failure arm would collapse the array); see Rows.cs.
        if (ArrayEvaluation.TryStream(Arguments[0], context, out var array))
        {
            return ReferencePosition.ArrayExtent(array, array.Columns);
        }

        // Same shared resolution as ROWS, fallback included: the argument's own error instead of a
        // plausible 1.
        if (
            !ReferencePosition.TryResolve(
                Arguments[0],
                context,
                ComputedValue.Number(1),
                out var reference,
                out var failure
            )
        )
        {
            return failure;
        }

        return ComputedValue.Number(
            reference switch
            {
                RangeReference range => range.ColumnCount,
                EmptyRangeReference empty => empty.ColumnCount,
                OpenRangeReference open => open.ColumnExtent(context),
                _ => 1.0,
            }
        );
    }
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

        // Sweep item 34(b), the VALUE slot: the lookup's error leads the scan (the oracle answers #NAME? for
        // XMATCH(NoSuch,A1:A3), and #DIV/0! for XMATCH(A1,B1:B3) over an error cell) instead of the not-found
        // #N/A. ReferencePosition.IsLookupValueError keeps a MULTI-cell range lookup's collapse #VALUE! (a
        // range has no scalar value) out of the rule — that artifact is content, and the scan answers #N/A on
        // it exactly as before — but reads a 1x1 range's own cell directly (finding I4): XMATCH(A1:A1,B1:B3)
        // is #DIV/0! over A1 = =1/0, the oracle's answer (26.7.0, both modes), not the not-found #N/A a
        // #VALUE!-literal scan gave before.
        if (ReferencePosition.IsLookupValueError(Arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        // ... and the ARRAY slot: an unresolved node reports its own error the same way (ReferencePosition's
        // shared rule) instead of streaming it as the one element the scan discards.
        if (ReferencePosition.TryUnresolvedError(Arguments[1], context, out var unresolved))
        {
            return unresolved;
        }

        OpenRangeReference? open = null;
        if (
            NamedReferences.TryResolveReference(
                Arguments[1],
                context,
                out var arrayReference,
                boundOpenRanges: false
            ) && arrayReference is OpenRangeReference openReference
        )
        {
            open = openReference;
        }

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
                    return ComputedValue.Number(snapshot.SourcePosition(hashPosition));
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

        if (match < 0)
        {
            return ComputedValue.Error(Error.NA);
        }

        var populatedPosition = match + 1;
        if (snapshot is not null)
        {
            return ComputedValue.Number(snapshot.SourcePosition(populatedPosition));
        }

        if (open is not null)
        {
            var (column, row) = open.PopulatedCells(context).ElementAt(match);
            return ComputedValue.Number(
                open.IsSingleRow ? open.ColumnPosition(column) : open.RowPosition(row)
            );
        }

        return ComputedValue.Number(populatedPosition);
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
    // AREAS(reference) — the number of areas (contiguous ranges or single cells) in the reference the
    // argument resolves to (a defined name stands for one, like ISREF): a union counts its areas
    // (recursively, for nested unions), any other reference is one area.
    //
    // The shared resolution carries both failure arms: a broken reference reports its own error (#NAME? for
    // an unknown name, #REF! for a failed INDIRECT/OFFSET), a plain non-reference value stays #VALUE!, and a
    // reference whose SHEET is gone — written literally (AREAS(Ghost!A1:A3)) or reached through a function
    // (AREAS(INDEX(Ghost!A1:A3,2,1))) — is the structural #REF! Excel answers, where counting it as one area
    // used to hand back a confident 1 for a sheet that no longer exists. AREAS needs no syntactic
    // ReferenceGuard pass of its own for that: unlike ROWS/COLUMNS it has no fast path that skips the
    // resolution, so the resolved-target re-check sees every argument that reaches it.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ReferencePosition.TryResolve(
            Arguments[0],
            context,
            ComputedValue.Error(Error.Value),
            out var reference,
            out var failure
        )
            ? reference switch
            {
                UnionReference union => ComputedValue.Number(CountAreas(union)),
                _ => ComputedValue.Number(1),
            }
            : failure;

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
        // Sweep item 34(b): an unresolved node reports its OWN error (#NAME? for an unknown name) instead
        // of the switch's uniform #VALUE! for a non-reference — the same rule the TableReference arm below
        // already carries for an unresolvable structured reference (Phase 5 item 17), now shared through
        // ReferencePosition. A resolved non-reference (a name bound to a cell) still falls through.
        if (ReferencePosition.TryUnresolvedError(Arguments[0], context, out var unresolved))
        {
            return unresolved;
        }

        // Phase 5 item 17. A structured reference cannot join the (sheetName, cellId) switch below as an
        // ordinary arm: an unresolvable one must report its OWN error (measured, Aspose.Cells 26.6.0,
        // 2026-09-11 — FORMULATEXT(Tabela1[#Totals]) is #REF! over a table with no totals row), never the
        // switch's uniform #VALUE! for a non-reference argument — mirroring ReferenceGuard's TableReference
        // arm (R2): the failure IS a value, not a short-circuit to a different code.
        if (Arguments[0] is TableReference table)
        {
            if (!table.TryResolve(context.Workbook, out var tableArea, out var error))
            {
                return ComputedValue.Error(error);
            }

            // Sweep item 33: a zero-row rectangle has no top-left cell and so no formula to show — the #N/A
            // a plain-literal target gets (oracle: FORMULATEXT over a header-only table's data band #N/A).
            return tableArea is RangeReference tableRange
                ? EvaluateAt(context, tableRange.SheetName, tableRange.StartId)
                : ComputedValue.Error(Error.NA);
        }

        var (sheetName, cellId) = Arguments[0] switch
        {
            CellReference cell => (cell.SheetName, cell.Id),
            RangeReference range => (range.SheetName, range.StartId),
            // Phase 2 audit (shared-formula delta production): a FORMULATEXT(ref) argument written INSIDE a
            // shared-formula master is an anchored node — mirrors the two cases above, applying the ambient
            // delta before resolving the target cell/range's un-parsed formula text.
            AnchoredCellReference anchoredCell => (
                anchoredCell.SheetName,
                new CellAddress(
                    anchoredCell.Effective(context).Column,
                    anchoredCell.Effective(context).Row
                ).ToId()
            ),
            AnchoredRangeReference anchoredRange => (
                anchoredRange.SheetName,
                anchoredRange.ToRangeReference(context).StartId
            ),
            _ => (null, null),
        };

        if (sheetName is null || cellId is null)
        {
            return ComputedValue.Error(Error.Value);
        }

        return EvaluateAt(context, sheetName, cellId);
    }

    // The shared tail once (sheetName, cellId) is known good, whether it came from a plain reference node
    // or a resolved structured reference: a plain-literal or missing target is #N/A, otherwise the target
    // expression's un-parsed formula text.
    private static ComputedValue EvaluateAt(
        EvaluationContext context,
        string sheetName,
        string cellId
    ) =>
        !context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
        || !sheet.TryGetValue(cellId, out var expression)
        || expression is ValueExpression
            ? ComputedValue.Error(Error.NA)
            : ComputedValue.Text("=" + expression.ToFormula(sheetName));
}
