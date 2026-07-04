using MemoryPack;

namespace Danfma.MySheet.Expressions;

// The wire stays a plain string: InternStringFormatter serializes with the default WriteString (byte-identical
// to an un-annotated member) and only the READ side changes — it interns the deserialized name via
// string.Intern, the same pool the parser uses. So N cross-sheet references to one sheet share a single
// SheetName instance after Load, exactly as after a parse (the ~24MB duplicate-"Data" lever at K1 scale).
[MemoryPackable]
public sealed partial record CellReference(
    string Id,
    [property: InternStringFormatter] string SheetName
) : Reference
{
    // Resolution goes through the workbook's memoized cell cache, so a cell referenced by many formulas
    // is computed once. GetCellValue evaluates the cell with itself as the current cell.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        context.Workbook.GetCellValue(SheetName, Id);

}
