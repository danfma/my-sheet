using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CellReference(string Id, string SheetName) : Reference
{
    public override object? Compute(Workbook workbook)
    {
        // Sheet indexer returns BlankValue for missing cells, so referencing an empty cell is blank
        // rather than throwing.
        var cell = workbook.Sheets[SheetName][Id];

        return cell.Compute(workbook);
    }
}
