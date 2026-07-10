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

    public EvaluationContext(Workbook workbook, string? sheetName = null, string? cellId = null)
        : this(workbook, sheetName, cellId, names: null) { }

    private EvaluationContext(
        Workbook workbook,
        string? sheetName,
        string? cellId,
        NameScope? names
    )
    {
        Workbook = workbook;
        SheetName = sheetName;
        CellId = cellId;
        _names = names;
    }

    // LET names are local to a formula and do not leak into referenced cells, so they are dropped here.
    public EvaluationContext WithCell(string sheetName, string cellId) =>
        new(Workbook, sheetName, cellId, names: null);

    public EvaluationContext WithName(string name, ComputedValue value) =>
        new(Workbook, SheetName, CellId, new NameScope(name, value, _names));

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
