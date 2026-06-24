namespace MySheet.Expressions;

/// <summary>
/// Evaluation context threaded through Compute: the workbook plus the cell currently being evaluated
/// (null at the root). Enables functions like ROW()/SHEET() and future LET scopes/memoization.
/// </summary>
public sealed class EvaluationContext
{
    public Workbook Workbook { get; }
    public string? SheetName { get; }
    public string? CellId { get; }

    public EvaluationContext(Workbook workbook, string? sheetName = null, string? cellId = null)
    {
        Workbook = workbook;
        SheetName = sheetName;
        CellId = cellId;
    }

    public EvaluationContext WithCell(string sheetName, string cellId) => new(Workbook, sheetName, cellId);
}
