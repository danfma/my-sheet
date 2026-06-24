using System.Runtime.CompilerServices;
using MemoryPack;

namespace MySheet.Expressions;

// MemoryPackUnion tags are APPEND-ONLY: never renumber, reorder or reuse an existing tag,
// or previously serialized data (and the WorkbookTests round-trip) will break. Add new tags at 7+.
[MemoryPackable]
[MemoryPackUnion(0, typeof(StringValue))]
[MemoryPackUnion(1, typeof(NumberValue))]
[MemoryPackUnion(2, typeof(BooleanValue))]
[MemoryPackUnion(3, typeof(BlankValue))]
[MemoryPackUnion(4, typeof(CellReference))]
[MemoryPackUnion(5, typeof(RangeReference))]
[MemoryPackUnion(6, typeof(Sum))]
[MemoryPackUnion(7, typeof(ErrorValue))]
[MemoryPackUnion(8, typeof(BinaryOperation))]
[MemoryPackUnion(9, typeof(UnaryOperation))]
[MemoryPackUnion(10, typeof(Average))]
[MemoryPackUnion(11, typeof(Min))]
[MemoryPackUnion(12, typeof(Max))]
[MemoryPackUnion(13, typeof(Count))]
[MemoryPackUnion(14, typeof(If))]
[MemoryPackUnion(15, typeof(And))]
[MemoryPackUnion(16, typeof(Or))]
[MemoryPackUnion(17, typeof(Not))]
[MemoryPackUnion(18, typeof(IfError))]
public abstract partial record Expression
{
    public abstract object? Compute(Workbook workbook);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static NumberValue Number(double value) => new(value);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static CellReference Cell(string id, Sheet sheet) => Cell(id, sheet.Name);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static CellReference Cell(string id, string sheetName) => new(id, sheetName);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static Sum Sum(params Expression[] expressions) => new(expressions);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static Average Average(params Expression[] expressions) => new(expressions);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static Min Min(params Expression[] expressions) => new(expressions);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static Max Max(params Expression[] expressions) => new(expressions);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static Count Count(params Expression[] expressions) => new(expressions);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static RangeReference Range(string startId, string endId, Sheet sheet) => new(startId, endId, sheet.Name);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Add(Expression left, Expression right) => new(BinaryOperator.Add, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Subtract(Expression left, Expression right) => new(BinaryOperator.Subtract, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Divide(Expression left, Expression right) => new(BinaryOperator.Divide, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Power(Expression left, Expression right) => new(BinaryOperator.Power, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation GreaterThan(Expression left, Expression right) => new(BinaryOperator.GreaterThan, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static UnaryOperation Negate(Expression operand) => new(UnaryOperator.Negate, operand);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static UnaryOperation Plus(Expression operand) => new(UnaryOperator.Plus, operand);
}
