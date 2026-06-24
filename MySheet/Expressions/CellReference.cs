using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CellReference(string Id, string SheetName) : Reference
{
    public override object? Compute(Workbook workbook)
    {
        var cell = workbook.Sheets[SheetName].Cells[Id];

        return cell.Compute(workbook);
    }

    public Expression? Resolve(Workbook workbook) => workbook.Sheets[SheetName].Cells[Id];
}