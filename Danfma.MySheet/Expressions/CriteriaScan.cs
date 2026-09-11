namespace Danfma.MySheet.Expressions;

/// <summary>
/// A forward-only positional cursor over a single range argument of a criteria aggregate (the SUMIFS
/// family). It NEVER materializes an intermediate list for a closed rectangle: when the shared per-epoch
/// snapshot is admitted it indexes the snapshot's <see cref="RangeSnapshot.Values"/> array directly
/// (zero-copy); otherwise it streams the memoized cells through <see cref="RangeValueSequence"/>'s
/// dispatch-free struct enumerator (dense positional reads, no allocation). Only the uncommon non-rectangle
/// shapes (open ranges/unions/scalars without a snapshot) fall back to the ordinary materialized list.
///
/// <para><see cref="OpenArrayOrRange"/> adds a fourth, OPT-IN backing for SUMPRODUCT alone: the element-wise
/// mini-CSE <see cref="ArrayEvaluation.ArrayStream"/>, so a computed array argument is a first-class vector
/// instead of one collapsed scalar. The *IFS family keeps <see cref="Open(Expression, EvaluationContext)"/>,
/// where Excel requires real ranges: a computed array in one of those slots is REJECTED up front by
/// <see cref="RejectComputedArray"/> rather than read — within the limits
/// <see cref="Open(Expression, EvaluationContext)"/>'s own doc names.</para>
/// </summary>
internal struct PositionalRange
{
    // Exactly one backing is live: a list (snapshot array or the materialized fallback), the dense rectangle
    // stream, OR the element-wise mini-CSE array — the latter discriminated by _streamRows > 0 (a live array
    // always has at least one row), so no extra flag field is needed.
    private readonly IReadOnlyList<ComputedValue>? _list;
    private readonly ArrayEvaluation.ArrayStream _stream;
    private readonly int _streamRows;
    private readonly int _streamColumns;
    private RangeValueSequence.Enumerator _cursor;
    private int _index;

    /// <summary>The cell count, known up front for every backing (array length, rectangle area, or list
    /// count) so the *IFS length validation never forces a materialization just to measure.</summary>
    public readonly int Count;

    /// <summary>The argument's rectangular shape — <c>0</c>/<c>0</c> when it HAS none (an open range, a
    /// union, a name or a scalar), whichever backing serves it. Needed because <see cref="Count"/> alone
    /// cannot tell a 3x1 column from a 1x3 row, and Excel's SUMPRODUCT dimension rule is about the
    /// dimensions, not the count.</summary>
    public readonly int Rows;

    /// <summary><see cref="Rows"/>' column twin: <c>0</c> when the argument is not a rectangle.</summary>
    public readonly int Columns;

    private PositionalRange(IReadOnlyList<ComputedValue> list)
        : this(list, rows: 0, columns: 0) { }

    private PositionalRange(IReadOnlyList<ComputedValue> list, int rows, int columns)
    {
        _list = list;
        _stream = default;
        _streamRows = 0;
        _streamColumns = 0;
        _cursor = default;
        _index = 0;
        Count = list.Count;
        Rows = rows;
        Columns = columns;
    }

    private PositionalRange(RangeValueSequence.Enumerator cursor, int rows, int columns)
    {
        _list = null;
        _stream = default;
        _streamRows = 0;
        _streamColumns = 0;
        _cursor = cursor;
        _index = 0;
        Count = rows * columns;
        Rows = rows;
        Columns = columns;
    }

    private PositionalRange(ArrayEvaluation.ArrayStream stream)
    {
        _list = null;
        _stream = stream;
        _streamRows = stream.Rows;
        _streamColumns = stream.Columns;
        _cursor = default;
        _index = 0;
        Count = stream.Length;
        Rows = stream.Rows;
        Columns = stream.Columns;
    }

    /// <summary>
    /// Opens a cursor over one argument, preferring the cheapest backing: the admitted per-epoch snapshot
    /// (zero-copy over <see cref="RangeSnapshot.Values"/>) → the dense positional stream for a closed
    /// rectangle (no allocation) → a materialized list for open ranges/unions/scalars.
    ///
    /// <para>What it does not read is a RECOGNISED computed array, which is narrower than "an array": every
    /// caller in the criteria family runs <see cref="RejectComputedArray"/> on the same argument first, and
    /// that predicate is exactly <c>!IsBareReferenceNode(argument, context) &amp;&amp;
    /// IsArrayEligible(argument, context)</c>, so the materialized fallback below STILL produces the
    /// one-element <c>#VALUE!</c> collapse whenever an array reaches the slot behind a node the predicate
    /// excludes. One such shape remains, PRE-EXISTING (measured identical on <c>main</c>) and pinned as a
    /// known divergence in <c>CriteriaComputedArgumentTests</c>: an operand the mini-CSE cost guard REFUSES,
    /// which is therefore not array-eligible, so <c>SUMIF(A:A*1,"&gt;0")</c> is <c>0</c> where the oracle
    /// answers <c>#REF!</c> in both entry modes (Aspose.Cells 26.6.0, measured 2026-09-10). Five more did
    /// until Phase 11c (array bindings) and are now refused: a <c>Let</c> NODE, once an opaque scalar to the
    /// shape probe and now array-eligible when its body is (<see cref="ArrayEvaluation.Probe"/>'s <c>Let</c>
    /// arm), so <c>COUNTIF(LET(r,A1:A3,r*1),"&gt;0")</c> is <c>#REF!</c>; a LET-BOUND NAME, once a
    /// reference node to the predicate while <see cref="NamedReferences.CaptureValue"/> had already collapsed
    /// its binding to one scalar, now an array binding the context-aware
    /// <see cref="ArrayEvaluation.IsBareReferenceNode(Expression, EvaluationContext)"/> does not admit, so
    /// <c>LET(r,A1:A3*1,COUNTIF(r,"&gt;0"))</c> is <c>#REF!</c> (both rows in
    /// <c>CriteriaComputedArgumentTests.LetBoundComputedArray_InARangeSlot_IsRefused</c>); and the other three
    /// binding sites, each array-eligible now that it carries its array — <c>COUNTIF(CHOOSE(1,FILTER(…)),
    /// "&gt;0")</c>, <c>COUNTIF(+FILTER(…),"&gt;0")</c> and <c>COUNTIF(ProdName,"&gt;0")</c> for a defined
    /// name whose definition is a computed array (ArrayBindingTests). So the fallback
    /// sees the shapes Excel accepts as a range (a reference, a name bound to one, a scalar, a refused open
    /// range) AND the collapse the refused shape still delivers; the gate removes the arrays the mini-CSE
    /// can recognise, not every array.</para>
    /// </summary>
    public static PositionalRange Open(Expression argument, EvaluationContext context)
    {
        var snapshot = argument is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        return Open(argument, context, snapshot);
    }

    /// <summary>
    /// Same backing preference as <see cref="Open(Expression, EvaluationContext)"/>, but takes an
    /// ALREADY-resolved snapshot probe instead of running its own. <see cref="Workbook.TryGetRangeSnapshot"/>
    /// is stateful (second-use admission): calling it a SECOND time for the same range within one function
    /// evaluation — e.g. once for a numeric-equality fast-path check, then again here — would itself count as
    /// the admitting "second read" and eagerly build the snapshot on what should still be the range's first,
    /// streaming read. A caller that already probed the snapshot (SUMIF/AVERAGEIF's fast-path check) MUST
    /// route that same result through here rather than calling the no-snapshot overload again.
    /// </summary>
    public static PositionalRange Open(
        Expression argument,
        EvaluationContext context,
        RangeSnapshot? snapshot
    )
    {
        // Phase 2 audit (shared-formula delta production): a *IFS criteria/value range written INSIDE a
        // shared-formula master (SUMIF(A1:A3,">0",...) as a slave) is an anchored node, not a plain
        // RangeReference — resolving it UP FRONT to its concrete, delta-applied twin lets the dense-rectangle
        // fast path below (and the ArgumentFlattening fallback) treat it exactly like an ordinary range, with
        // no logic duplicated. Falling through to ArgumentFlattening.ExpandComputedValues unresolved would hit
        // AnchoredRangeReference.Evaluate's unconditional #VALUE! (a range has no scalar value) before this fix.
        // It runs BEFORE the snapshot branch so that branch can read the resolved rectangle's shape too.
        if (argument is AnchoredRangeReference anchoredRange)
        {
            argument = anchoredRange.ToRangeReference(context);
        }

        if (snapshot is not null)
        {
            // The snapshot hands its values over as a FLAT list, but the SHAPE still has to travel with them:
            // Excel's SUMPRODUCT dimension rule reads Rows/Columns, and the snapshot is admitted only on a
            // range's SECOND read of the epoch — so leaving the shape unknown here would make an
            // Excel-documented rule depend on cache-admission order, and two cells holding the identical
            // formula would disagree. A closed rectangle's snapshot covers exactly its area (column-major), so
            // its bounds ARE the shape. An open range's snapshot holds only its POPULATED cells and has no
            // rectangle at all, so it keeps 0/0 — shape unknown — like the materialized fallback below.
            if (argument is RangeReference snapshotRectangle)
            {
                var snapshotBounds = snapshotRectangle.GetBounds();

                return new PositionalRange(
                    snapshot.Values,
                    snapshotBounds.RowCount,
                    snapshotBounds.ColumnCount
                );
            }

            return new PositionalRange(snapshot.Values);
        }

        if (argument is RangeReference rectangle)
        {
            // A single GetBounds() call sizes the cursor without paying one corner parse per property read.
            var bounds = rectangle.GetBounds();

            return new PositionalRange(
                rectangle.ExpandComputedValues(context).GetEnumerator(),
                bounds.RowCount,
                bounds.ColumnCount
            );
        }

        return new PositionalRange(ArgumentFlattening.ExpandComputedValues(argument, context));
    }

    /// <summary>
    /// SUMPRODUCT-ONLY factory: the backings of <see cref="Open(Expression, EvaluationContext)"/> plus the
    /// element-wise mini-CSE array, for an argument that is not a reference yet computes an array
    /// (<c>(A1:A3&lt;&gt;0)*1</c>, <c>ROW(A1:A3)</c>, <c>IF(A1:A3&gt;0,1,0)</c>).
    ///
    /// <para>Deliberately an opt-in factory rather than a gate folded into <see cref="Open"/>: <c>Open</c> is
    /// shared with the *IFS criteria family (SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS), where Excel REQUIRES
    /// real ranges — <c>SUMIFS((A1:A3)*1, …)</c> is <c>#REF!</c> array-entered and <c>#VALUE!</c> entered
    /// plainly (Aspose.Cells 26.6.0, measured 2026-09-10) — so those slots REJECT a computed array through
    /// <see cref="RejectComputedArray"/> instead of reading it.</para>
    /// </summary>
    public static PositionalRange OpenArrayOrRange(Expression argument, EvaluationContext context)
    {
        // The shared mini-CSE gate (ArrayEvaluation.TryStream): references keep the range path
        // (snapshot/dense stream) untouched, the cheap structural pre-check keeps the scalar path a mere
        // type-walk, and the build that follows is the argument's SINGLE evaluation — no double-eval, so a
        // volatile operand still draws exactly once.
        if (ArrayEvaluation.TryStream(argument, context, out var stream))
        {
            return new PositionalRange(stream);
        }

        return Open(argument, context);
    }

    /// <summary>
    /// The criteria family's range-slot gate: <c>#REF!</c> for an argument that is not a bare reference NODE
    /// yet the mini-CSE would stream as an array, and <c>null</c> (open the cursor as usual) for everything
    /// else. The predicate is exactly the first two conditions of
    /// <see cref="ArrayEvaluation.TryStream"/> — the context-aware
    /// <see cref="ArrayEvaluation.IsBareReferenceNode(Expression, EvaluationContext)"/> then
    /// <see cref="ArrayEvaluation.IsArrayEligible"/> — so an <c>OFFSET</c>, a single cell, a bare scalar, a
    /// cost-guard-REFUSED open-range expression, and every shape whose branches or definition hold no
    /// computed array (a scalar-conditioned <c>IF</c> over references, a <c>CHOOSE</c> between them, a name
    /// bound to a range, a <c>+</c> over any of those) are outside it and keep the paths they have today.
    /// What is INSIDE it, since Phase 11c, is every binding site carrying an array — a LET-bound name
    /// (<c>LET(f,FILTER(…),COUNTIF(f,"&gt;0"))</c>), a <c>Let</c> node, a <c>CHOOSE</c> with an array branch,
    /// a <c>+</c> over a producer, and a defined name whose definition is a computed array
    /// (<c>COUNTIF(ProdName,"&gt;0")</c>) — each <c>#REF!</c> like the producer written in the slot, which is
    /// the oracle's answer for all of them (ArrayBindingTests). A <c>+</c> over a bare REFERENCE stays
    /// outside, measured: <c>COUNTIF(+A1:A3,"&gt;0")</c> is 2 and <c>SUMIF(+A1:A3,"&gt;0")</c> 14 in both
    /// entry modes (26.6.0, 2026-09-11), which is what <see cref="ArrayEvaluation.IsBareReferenceNode"/>'s
    /// Plus arm is for.
    ///
    /// <para>The rule: a range slot of <c>SUMIF</c>/<c>SUMIFS</c>/<c>COUNTIF</c>/<c>COUNTIFS</c>/
    /// <c>AVERAGEIF</c>/<c>AVERAGEIFS</c>/<c>MAXIFS</c>/<c>MINIFS</c> — criteria range and
    /// sum/average/max/min range alike — takes a reference, so a computed array there is rejected rather
    /// than collapsed and scanned. Oracle Aspose.Cells 26.6.0, measured 2026-09-10 in both entry modes: a
    /// COMPUTED array (<c>COUNTIF(A1:A3*1,"&gt;0")</c>, <c>COUNTIF(ROW(A1:A3),"&gt;1")</c>) is
    /// <c>#VALUE!</c> entered plainly and <c>#REF!</c> array-entered (<c>Cell.SetArrayFormula</c>); a
    /// dynamic-array PRODUCER (<c>COUNTIF(FILTER(A1:A3,A1:A3&gt;0),"&gt;5")</c>,
    /// <c>COUNTIF(SEQUENCE(5),"&gt;3")</c>, <c>SUMIF(SORT(A1:A3),"&gt;0")</c>) is <c>#REF!</c> in BOTH
    /// columns. The mini-CSE implements the array-entered rule, and <c>#REF!</c> is the one value both
    /// producer columns agree on — which is why the answer is <c>#REF!</c> and not <c>#VALUE!</c>. Pinned by
    /// CriteriaComputedArgumentTests.</para>
    ///
    /// <para><see cref="OpenArrayOrRange"/> — SUMPRODUCT's opt-in factory — is deliberately NOT gated: Excel
    /// reads a computed array there (<c>SUMPRODUCT((A1:A3&lt;&gt;0)*1)</c> = 2 and
    /// <c>SUMPRODUCT(A1:A3*1,B1:B3)</c> = 32, both entry modes, same measurement), so this gate guards the
    /// eight <see cref="Open"/>/<c>RangeValueCursor.Open</c> call sites of the criteria family alone.</para>
    /// </summary>
    public static Error? RejectComputedArray(Expression argument, EvaluationContext context) =>
        !ArrayEvaluation.IsBareReferenceNode(argument, context)
        && ArrayEvaluation.IsArrayEligible(argument, context)
            ? Error.Ref
            : null;

    /// <summary>The next cell in position order (column-major, matching the materialized expansion exactly —
    /// the row-major array backing is transposed here to agree). Every parallel cursor is advanced once per
    /// position so they stay aligned.</summary>
    public ComputedValue Next()
    {
        if (_list is { } list)
        {
            return list[_index++];
        }

        if (_streamRows > 0)
        {
            // TRANSPOSE on read — load-bearing, not defensive. Every other backing yields COLUMN-major
            // positions (RangeValueSequence.Enumerator walks column-outer/row-inner, and so does the
            // snapshot's view), while ArrayStream.ElementAt is indexed ROW-major. Mapping the column-major
            // position through (row, column) is what keeps the parallel cursors of a mixed range+array
            // argument list aligned: without it a 2-D SUMPRODUCT(A1:B2,(A1:B2)*1) pairs A2 with B1 and
            // silently answers 29 instead of 30.
            var position = _index++;
            var column = position / _streamRows;
            var row = position % _streamRows;

            return _stream.ElementAt((row * _streamColumns) + column);
        }

        _cursor.MoveNext();
        return _cursor.Current;
    }
}

/// <summary>
/// The shared position-by-position scan of the *IFS criteria aggregates (SUMIFS, COUNTIFS, AVERAGEIFS,
/// MAXIFS, MINIFS). It walks the criteria ranges (and the optional value range) as parallel forward
/// cursors — no intermediate <c>List&lt;ComputedValue&gt;</c> per range — after enforcing equal lengths up
/// front (<c>#VALUE!</c> on mismatch, like the pre-refactor code). Each <see cref="MoveNext"/> reports
/// whether every criterion matched at that position and, for the value families, hands back the
/// value-range cell so the caller's own accumulator (sum/count/average/extreme) stays in the function body.
/// </summary>
internal struct CriteriaScan
{
    private readonly PositionalRange[] _ranges;
    private readonly Criteria[] _criterias;
    private PositionalRange _valueRange;
    private readonly bool _hasValue;
    private readonly int _length;
    private int _index;

    private CriteriaScan(
        PositionalRange[] ranges,
        Criteria[] criterias,
        PositionalRange valueRange,
        bool hasValue,
        int length
    )
    {
        _ranges = ranges;
        _criterias = criterias;
        _valueRange = valueRange;
        _hasValue = hasValue;
        _length = length;
        _index = 0;
    }

    /// <summary>
    /// Builds a scan whose length is the VALUE range's cell count — the SUMIFS/AVERAGEIFS/MAXIFS/MINIFS
    /// shape where <c>arguments[0]</c> is the value range and the (criteria_range, criteria) pairs follow.
    /// Returns <c>#REF!</c> for a missing sheet or for a computed array in ANY range slot (value range or
    /// criteria range — <see cref="PositionalRange.RejectComputedArray"/>), and <c>#VALUE!</c> when a
    /// criteria range's length differs.
    /// </summary>
    public static Error? CreateWithValue(
        Expression[] arguments,
        EvaluationContext context,
        out CriteriaScan scan
    )
    {
        scan = default;

        if (ReferenceGuard.MissingSheet(arguments, context) is { } missing)
        {
            return missing;
        }

        if (PositionalRange.RejectComputedArray(arguments[0], context) is { } computedValueRange)
        {
            return computedValueRange;
        }

        var valueRange = PositionalRange.Open(arguments[0], context);
        var length = valueRange.Count;
        var pairCount = (arguments.Length - 1) / 2;
        var ranges = new PositionalRange[pairCount];
        var criterias = new Criteria[pairCount];

        for (var p = 0; p < pairCount; p++)
        {
            if (
                PositionalRange.RejectComputedArray(arguments[1 + (p * 2)], context) is
                { } computedRange
            )
            {
                return computedRange;
            }

            var range = PositionalRange.Open(arguments[1 + (p * 2)], context);

            if (range.Count != length)
            {
                return Error.Value;
            }

            ranges[p] = range;
            criterias[p] = Criteria.Parse(arguments[2 + (p * 2)].Evaluate(context));
        }

        scan = new CriteriaScan(ranges, criterias, valueRange, hasValue: true, length);

        return null;
    }

    /// <summary>
    /// Builds a value-less scan — the COUNTIFS shape where the (criteria_range, criteria) pairs start at
    /// <c>arguments[0]</c> and the length is the first criteria range's cell count. Returns <c>#REF!</c>
    /// for a missing sheet or for a computed array in any criteria range
    /// (<see cref="PositionalRange.RejectComputedArray"/>) and <c>#VALUE!</c> when a later criteria range's
    /// length differs.
    /// </summary>
    public static Error? CreateCountOnly(
        Expression[] arguments,
        EvaluationContext context,
        out CriteriaScan scan
    )
    {
        scan = default;

        if (ReferenceGuard.MissingSheet(arguments, context) is { } missing)
        {
            return missing;
        }

        var pairCount = arguments.Length / 2;
        var ranges = new PositionalRange[pairCount];
        var criterias = new Criteria[pairCount];
        var length = 0;

        for (var p = 0; p < pairCount; p++)
        {
            if (PositionalRange.RejectComputedArray(arguments[p * 2], context) is { } computedRange)
            {
                return computedRange;
            }

            var range = PositionalRange.Open(arguments[p * 2], context);

            if (p == 0)
            {
                length = range.Count;
            }
            else if (range.Count != length)
            {
                return Error.Value;
            }

            ranges[p] = range;
            criterias[p] = Criteria.Parse(arguments[(p * 2) + 1].Evaluate(context));
        }

        scan = new CriteriaScan(ranges, criterias, default, hasValue: false, length);

        return null;
    }

    /// <summary>
    /// Advances every cursor by one position (keeping them aligned) and returns <c>false</c> once the range
    /// is exhausted. On <c>true</c>, <paramref name="matched"/> tells whether every criterion matched at
    /// this position; <paramref name="value"/> is the value-range cell (default for COUNTIFS). The criteria
    /// checks short-circuit — the cursors still all advance, but the expensive <see cref="Criteria.Matches"/>
    /// is skipped once a mismatch is known, exactly like the old <c>AllMatch</c>.
    /// </summary>
    public bool MoveNext(out bool matched, out ComputedValue value)
    {
        if (_index >= _length)
        {
            matched = false;
            value = default;

            return false;
        }

        matched = true;
        var ranges = _ranges;
        var criterias = _criterias;

        for (var j = 0; j < ranges.Length; j++)
        {
            var cell = ranges[j].Next();

            if (matched && !criterias[j].Matches(cell))
            {
                matched = false;
            }
        }

        value = _hasValue ? _valueRange.Next() : default;
        _index++;

        return true;
    }
}
