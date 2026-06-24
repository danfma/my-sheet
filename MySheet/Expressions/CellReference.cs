using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CellReference(string Id, string SheetName) : Reference
{
    public override object? Compute(EvaluationContext context)
    {
        // Sheet indexer returns BlankValue for missing cells, so referencing an empty cell is blank
        // rather than throwing.
        var cell = context.Workbook.Sheets[SheetName][Id];

        return cell.Compute(context.WithCell(SheetName, Id));
    }
}
