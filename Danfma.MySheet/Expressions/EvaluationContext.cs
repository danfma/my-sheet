namespace Danfma.MySheet.Expressions;

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

    public EvaluationContext(Workbook workbook, string? sheetName = null, string? cellId = null)
        : this(workbook, sheetName, cellId, names: null, deltaRow: 0, deltaColumn: 0) { }

    private EvaluationContext(
        Workbook workbook,
        string? sheetName,
        string? cellId,
        NameScope? names,
        int deltaRow,
        int deltaColumn
    )
    {
        Workbook = workbook;
        SheetName = sheetName;
        CellId = cellId;
        _names = names;
        DeltaRow = deltaRow;
        DeltaColumn = deltaColumn;
    }

    // LET names are local to a formula and do not leak into referenced cells, so they are dropped here. The
    // shared-formula delta is likewise a property of the ORIGINATING slave cell, not of whatever cell it
    // references, so it resets to 0 here too (this mirrors EvaluateCell's fresh-context behavior for the
    // rare direct caller of WithCell — the normal GetCellValue/GetCellValueDense path never routes through
    // this method at all, it always goes through EvaluateCell).
    public EvaluationContext WithCell(string sheetName, string cellId) =>
        new(Workbook, sheetName, cellId, names: null, deltaRow: 0, deltaColumn: 0);

    public EvaluationContext WithName(string name, ComputedValue value) =>
        new(
            Workbook,
            SheetName,
            CellId,
            new NameScope(name, value, operand: null, _names),
            DeltaRow,
            DeltaColumn
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
            DeltaColumn
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
        new(Workbook, SheetName, CellId, _names, deltaRow, deltaColumn);

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
