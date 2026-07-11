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
    private sealed class NameScope
    {
        public readonly string Name;
        public readonly ComputedValue Value;
        public readonly NameScope? Parent;

        public NameScope(string name, ComputedValue value, NameScope? parent)
        {
            Name = name;
            Value = value;
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
        new(Workbook, SheetName, CellId, new NameScope(name, value, _names), DeltaRow, DeltaColumn);

    /// <summary>
    /// G3 spike: pushes a shared-formula delta for the duration of evaluating a
    /// <see cref="SharedFormulaSlave.Master"/> tree. LET bindings are preserved (a LET inside a shared-formula
    /// master is still local to THAT evaluation), <see cref="SheetName"/>/<see cref="CellId"/> are unchanged
    /// (still the slave's own cell — only the anchored nodes inside Master read the delta).
    /// </summary>
    public EvaluationContext WithDelta(int deltaRow, int deltaColumn) =>
        new(Workbook, SheetName, CellId, _names, deltaRow, deltaColumn);

    // Names are case-insensitive, matching the OrdinalIgnoreCase comparer the old dictionary used (and
    // Workbook.DefinedNames still uses for the layer this falls back to — see NameReference).
    public bool TryGetName(string name, out ComputedValue value)
    {
        for (var node = _names; node is not null; node = node.Parent)
        {
            if (string.Equals(node.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = node.Value;
                return true;
            }
        }

        value = default;
        return false;
    }
}
