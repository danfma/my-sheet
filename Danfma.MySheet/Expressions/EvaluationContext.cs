namespace Danfma.MySheet.Expressions;

/// <summary>
/// Evaluation context threaded through Compute: the workbook, the cell currently being evaluated
/// (null at the root), and local LET name bindings. A readonly struct so threading it through the
/// recursion allocates nothing (only the LET name map, when used, lives on the heap).
/// </summary>
public readonly struct EvaluationContext
{
    private readonly IReadOnlyDictionary<string, object?>? _names;

    public Workbook Workbook { get; }
    public string? SheetName { get; }
    public string? CellId { get; }

    public EvaluationContext(Workbook workbook, string? sheetName = null, string? cellId = null)
        : this(workbook, sheetName, cellId, names: null) { }

    private EvaluationContext(
        Workbook workbook,
        string? sheetName,
        string? cellId,
        IReadOnlyDictionary<string, object?>? names
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

    public EvaluationContext WithName(string name, object? value)
    {
        var names = _names is null
            ? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, object?>(_names, StringComparer.OrdinalIgnoreCase);

        names[name] = value;

        return new EvaluationContext(Workbook, SheetName, CellId, names);
    }

    public bool TryGetName(string name, out object? value)
    {
        if (_names is not null && _names.TryGetValue(name, out value))
        {
            return true;
        }

        value = null;
        return false;
    }
}
