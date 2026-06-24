using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
[MemoryPackUnion(0, typeof(StringValue))]
[MemoryPackUnion(1, typeof(NumberValue))]
[MemoryPackUnion(2, typeof(BooleanValue))]
[MemoryPackUnion(3, typeof(BlankValue))]
[MemoryPackUnion(4, typeof(CellReference))]
[MemoryPackUnion(5, typeof(RangeReference))]
[MemoryPackUnion(6, typeof(SumOperator))]
public abstract partial record Expression
{
    public abstract object? Compute(Workbook workbook);

    public static NumberValue Number(double value) => new(value);

    public static CellReference Cell(string id, Sheet sheet) => new(id, sheet.Name);

    public static SumOperator Sum(params Expression[] expressions) => new(expressions);
}