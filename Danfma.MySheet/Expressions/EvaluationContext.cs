namespace Danfma.MySheet.Expressions;

using Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// Evaluation context threaded through <see cref="Expression.Evaluate(EvaluationContext)"/>: the workbook,
/// the cell currently being evaluated (null at the root), and local LET name bindings. A readonly struct so
/// threading it through the recursion allocates nothing (only the LET name scope, when used, lives on the
/// heap — one small node per binding, not a whole map).
/// </summary>
public readonly struct EvaluationContext
{
    // Immutable singly-linked list of LET bindings, newest first. LET(name1, value1, …) chains one node per
    // binding via WithName instead of copying a dictionary: a k-binding LET used to do a
    // `new Dictionary(_names, …)` copy of every prior name on EACH binding (O(k^2) copies plus k dictionary
    // allocations); this is O(1) per WithName and O(k) worst case per lookup. LET bindings are few (a handful
    // per formula), so the linear walk beats the hashing + copying overhead. Walking newest-first also gives
    // shadowing for free: a rebind of the same name (LET(x,1,LET(x,x+1,x))) pushes a new node in front, so the
    // most recent one is found first and the outer binding is simply never reached.
    //
    // A binding holds one of TWO forms (Phase 11c, array bindings): a ComputedValue — a scalar, or the
    // reference VALUE a range binding captures (NamedReferences.CaptureValue) — or an ArrayOperand built
    // ONCE at binding time for an array-eligible binding expression (ArrayBindings.Capture). One node type
    // carries both so the walk, the shadowing and the cost above are unchanged; which form a node holds is
    // what TryGetName and TryGetArrayBinding disagree about, on purpose. A name answers exactly one of them,
    // and the NEAREST binding of a name is the answer for both, so it shadows an outer binding of the OTHER
    // form too: LET(x,1,LET(x,FILTER(…),SUM(x))) sums the array and never reaches the 1, and
    // LET(x,FILTER(…),LET(x,1,x)) is 1.
    private sealed class NameScope
    {
        public readonly string Name;
        public readonly ComputedValue Value;
        public readonly ArrayOperand? Operand;
        public readonly NameScope? Parent;

        public NameScope(string name, ComputedValue value, ArrayOperand? operand, NameScope? parent)
        {
            Name = name;
            Value = value;
            Operand = operand;
            Parent = parent;
        }
    }

    private readonly NameScope? _names;

    // Final-review fix wave, finding I3: a scalar-condition selector (IF/CHOOSE) bound to a LET name is
    // resolved TWICE — ArrayBindings.Shape (the probe, via If.TryResolveReference/Choose's own resolution)
    // and ArrayBindings.Capture (the build, via If.Evaluate/Choose.Evaluate) each evaluate the CONDITION
    // independently, because they are two SEPARATE calls to Let.TryBind with no state shared between them.
    // For a volatile condition (RAND()<0.5) the two draws can disagree, so the probe's shape (built from one
    // branch) and the build's actual value (drawn from the OTHER branch) mismatch — measured over 200 seeds,
    // SUM(LET(a,IF(RAND()<0.5,Rng,0),b,a,b*1)) answered #VALUE! in a quarter of them, a third bucket neither
    // draw alone would ever give. This cache makes the CONDITION sub-expression's evaluation idempotent
    // WITHIN one cell's evaluation (see If.cs/Choose's TryChoose): the first reader draws it and records the
    // value here; every later reader of the SAME condition node reuses it instead of redrawing. Reference
    // identity is the right key — the same source `IF(...)`/`CHOOSE(...)` occurrence is one parsed node
    // reused by every reader that reaches it during ONE evaluation, and a SharedFormulaSlave's anchored
    // master tree is likewise one shared node per formula group, not duplicated per slave.
    //
    // Scope is exactly what WithCell already establishes for _names: created FRESH per top-level cell
    // evaluation (Workbook.EvaluateCell's constructor call, and WithCell's own reset — a formula's local
    // state, including this cache, does not leak into a cell it references) and threaded BY REFERENCE
    // through every WithName/WithDelta derivation of THAT SAME evaluation, so ArrayBindings.Shape's probe
    // pass and ArrayBindings.Capture's later build pass — both reached from the SAME top-level Evaluate call
    // — see the identical cache instance. Allocated eagerly (a small, field-less-until-used wrapper, not a
    // Dictionary) rather than lazily: a struct copy of a null reference can never become non-null in a
    // SIBLING copy, so a cache created on first use inside ArrayBindings.Shape would be invisible to the
    // later, separately-constructed ArrayBindings.Capture call — the exact bug this class exists to close.
    private readonly ConditionCache _conditions;

    /// <summary>See the remarks on <see cref="EvaluationContext"/>'s <c>_conditions</c> field.</summary>
    private sealed class ConditionCache
    {
        private Dictionary<Expression, ComputedValue>? _values;
        private Dictionary<Expression, object>? _nodeMemos;

        public bool TryGet(Expression condition, out ComputedValue value)
        {
            if (_values is not null && _values.TryGetValue(condition, out var cached))
            {
                value = cached;
                return true;
            }

            value = default;
            return false;
        }

        public void Set(Expression condition, ComputedValue value) =>
            (
                _values ??= new Dictionary<Expression, ComputedValue>(
                    ReferenceEqualityComparer.Instance
                )
            )[condition] = value;

        public bool TryGetNodeMemo<T>(Expression node, out T value)
            where T : class
        {
            if (_nodeMemos is not null && _nodeMemos.TryGetValue(node, out var cached))
            {
                value = (T)cached;
                return true;
            }

            value = null!;
            return false;
        }

        public void SetNodeMemo(Expression node, object value) =>
            (_nodeMemos ??= new Dictionary<Expression, object>(ReferenceEqualityComparer.Instance))[
                node
            ] = value;
    }

    /// <summary>
    /// The first reader of <paramref name="condition"/> within this cell's evaluation draws it (its own
    /// <c>Evaluate</c>, against THIS context) and every later reader of the SAME node — a selector's
    /// condition resolved once by the probe and again by the build — reuses that draw. See the remarks on
    /// <see cref="EvaluationContext"/>'s <c>_conditions</c> field.
    /// </summary>
    internal ComputedValue EvaluateConditionOnce(Expression condition)
    {
        if (_conditions.TryGet(condition, out var cached))
        {
            return cached;
        }

        var value = condition.Evaluate(this);
        _conditions.Set(condition, value);
        return value;
    }

    internal bool TryGetEvaluatedCondition(Expression condition, out ComputedValue value) =>
        _conditions.TryGet(condition, out value);

    internal bool TryGetNodeMemo<T>(Expression node, out T value)
        where T : class => _conditions.TryGetNodeMemo(node, out value);

    internal void SetNodeMemo(Expression node, object value) =>
        _conditions.SetNodeMemo(node, value);

    public Workbook Workbook { get; }
    public string? SheetName { get; }
    public string? CellId { get; }

    /// <summary>
    /// G3 spike (node-delta shared formulas): the (row, column) offset a <see cref="SharedFormulaSlave"/>
    /// pushes for the duration of evaluating its shared <see cref="SharedFormulaSlave.Master"/> tree. Zero
    /// everywhere else (default), so every other node's evaluation is byte-for-byte unaffected. Not
    /// serialized — this is transient evaluation STATE, not workbook data — and it never survives a
    /// <c>GetCellValue</c>/<c>GetCellValueDense</c> cell-boundary crossing: <see cref="Workbook.EvaluateCell"/>
    /// always constructs a fresh <see cref="EvaluationContext"/> (see that method), which defaults both
    /// fields back to 0.
    /// </summary>
    public int DeltaRow { get; }
    public int DeltaColumn { get; }

    // Internal route diagnostics let tests prove that a consumer retained a resolved reference instead of
    // materializing it. They belong to the top-level evaluation context, just like the condition cache.
    private readonly XMatchRouteDiagnostics? _routeDiagnostics;

    internal void RecordArrayMaterialization() => _routeDiagnostics?.RecordArrayMaterialization();

    internal void RecordArrayStream() => _routeDiagnostics?.RecordArrayStream();

    internal void RecordReferenceExpansion() => _routeDiagnostics?.RecordReferenceExpansion();

    internal EvaluationContext(
        Workbook workbook,
        string? sheetName,
        XMatchRouteDiagnostics routeDiagnostics
    )
        : this(workbook, sheetName, null, null, 0, 0, null, routeDiagnostics) { }

    public EvaluationContext(Workbook workbook, string? sheetName = null, string? cellId = null)
        : this(
            workbook,
            sheetName,
            cellId,
            names: null,
            deltaRow: 0,
            deltaColumn: 0,
            conditions: null,
            routeDiagnostics: null
        ) { }

    private EvaluationContext(
        Workbook workbook,
        string? sheetName,
        string? cellId,
        NameScope? names,
        int deltaRow,
        int deltaColumn,
        ConditionCache? conditions,
        XMatchRouteDiagnostics? routeDiagnostics
    )
    {
        Workbook = workbook;
        SheetName = sheetName;
        CellId = cellId;
        _names = names;
        DeltaRow = deltaRow;
        DeltaColumn = deltaColumn;
        // A null `conditions` means "start of a fresh top-level evaluation" (the public constructor and
        // WithCell both pass null), so a new cache is made HERE — the one point that must run before any
        // WithName/WithDelta derivation copies the reference onward. See the remarks on the field above.
        _conditions = conditions ?? new ConditionCache();
        _routeDiagnostics = routeDiagnostics;
    }

    // LET names are local to a formula and do not leak into referenced cells, so they are dropped here. The
    // shared-formula delta is likewise a property of the ORIGINATING slave cell, not of whatever cell it
    // references, so it resets to 0 here too (this mirrors EvaluateCell's fresh-context behavior for the
    // rare direct caller of WithCell — the normal GetCellValue/GetCellValueDense path never routes through
    // this method at all, it always goes through EvaluateCell). The condition cache resets for the same
    // reason as _names: a referenced cell's own volatile selectors are its own evaluation's business.
    public EvaluationContext WithCell(string sheetName, string cellId) =>
        new(
            Workbook,
            sheetName,
            cellId,
            names: null,
            deltaRow: 0,
            deltaColumn: 0,
            conditions: null,
            routeDiagnostics: null
        );

    public EvaluationContext WithName(string name, ComputedValue value) =>
        new(
            Workbook,
            SheetName,
            CellId,
            new NameScope(name, value, operand: null, _names),
            DeltaRow,
            DeltaColumn,
            _conditions,
            _routeDiagnostics
        );

    /// <summary>
    /// The ARRAY-binding form (Phase 11c): <paramref name="name"/> bound to an operand that was built once,
    /// at binding time, and is read — never rebuilt — by every consumer of the name for the scope's lifetime
    /// (one evaluation). Found by <see cref="TryGetArrayBinding"/>, invisible to <see cref="TryGetName"/>:
    /// an array binding is not a <see cref="ComputedValue"/>.
    /// </summary>
    internal EvaluationContext WithName(string name, ArrayOperand operand) =>
        new(
            Workbook,
            SheetName,
            CellId,
            new NameScope(name, value: default, operand, _names),
            DeltaRow,
            DeltaColumn,
            _conditions,
            _routeDiagnostics
        );

    /// <summary>Binds whichever form <paramref name="binding"/> holds — the one call a binding site makes.</summary>
    internal EvaluationContext WithName(string name, ArrayBindings.Binding binding) =>
        binding.Operand is { } operand ? WithName(name, operand) : WithName(name, binding.Value);

    /// <summary>
    /// G3 spike: pushes a shared-formula delta for the duration of evaluating a
    /// <see cref="SharedFormulaSlave.Master"/> tree. LET bindings are preserved (a LET inside a shared-formula
    /// master is still local to THAT evaluation), <see cref="SheetName"/>/<see cref="CellId"/> are unchanged
    /// (still the slave's own cell — only the anchored nodes inside Master read the delta).
    /// </summary>
    public EvaluationContext WithDelta(int deltaRow, int deltaColumn) =>
        new(
            Workbook,
            SheetName,
            CellId,
            _names,
            deltaRow,
            deltaColumn,
            _conditions,
            _routeDiagnostics
        );

    /// <summary>
    /// The nearest SCALAR-form binding of <paramref name="name"/> — a scalar or a captured reference value.
    /// False when the name is unbound, and false when its nearest binding is an array (see
    /// <see cref="TryGetArrayBinding"/>): the two lookups answer for the same node and never both.
    /// </summary>
    public bool TryGetName(string name, out ComputedValue value)
    {
        if (Find(name) is { Operand: null } node)
        {
            value = node.Value;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// The nearest ARRAY-form binding of <paramref name="name"/> (Phase 11c): the operand
    /// <see cref="WithName(string, ArrayOperand)"/> bound. False when the name is unbound or its nearest
    /// binding is a scalar.
    /// </summary>
    internal bool TryGetArrayBinding(string name, out ArrayOperand operand)
    {
        if (Find(name) is { Operand: { } bound })
        {
            operand = bound;
            return true;
        }

        operand = null!;
        return false;
    }

    // The nearest binding of a name, whichever form it holds. Names are case-insensitive, matching the
    // OrdinalIgnoreCase comparer the old dictionary used (and Workbook.DefinedNames still uses for the layer
    // this falls back to — see NameReference).
    private NameScope? Find(string name)
    {
        for (var node = _names; node is not null; node = node.Parent)
        {
            if (string.Equals(node.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return node;
            }
        }

        return null;
    }
}
