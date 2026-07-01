using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CellReference(string Id, string SheetName) : Reference
{
    // Resolution goes through the workbook's memoized cell cache, so a cell referenced by many formulas
    // is computed once. GetCellValue evaluates the cell with itself as the current cell.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        context.Workbook.GetCellComputedValue(SheetName, Id);

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
