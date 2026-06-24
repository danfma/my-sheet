using System.Runtime.CompilerServices;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

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
[MemoryPackUnion(19, typeof(FunctionCall))]
[MemoryPackUnion(20, typeof(Int))]
[MemoryPackUnion(21, typeof(Round))]
[MemoryPackUnion(22, typeof(RoundUp))]
[MemoryPackUnion(23, typeof(Abs))]
[MemoryPackUnion(24, typeof(IsNumber))]
[MemoryPackUnion(25, typeof(IsBlank))]
[MemoryPackUnion(26, typeof(IfNa))]
[MemoryPackUnion(27, typeof(Upper))]
[MemoryPackUnion(28, typeof(Lower))]
[MemoryPackUnion(29, typeof(Trim))]
[MemoryPackUnion(30, typeof(Len))]
[MemoryPackUnion(31, typeof(Left))]
[MemoryPackUnion(32, typeof(Mid))]
[MemoryPackUnion(33, typeof(Value))]
[MemoryPackUnion(34, typeof(Concat))]
[MemoryPackUnion(35, typeof(Concatenate))]
[MemoryPackUnion(36, typeof(TextJoin))]
[MemoryPackUnion(37, typeof(CountA))]
[MemoryPackUnion(38, typeof(CountBlank))]
[MemoryPackUnion(39, typeof(CountIf))]
[MemoryPackUnion(40, typeof(CountIfs))]
[MemoryPackUnion(41, typeof(SumIf))]
[MemoryPackUnion(42, typeof(SumIfs))]
[MemoryPackUnion(43, typeof(Rows))]
[MemoryPackUnion(44, typeof(Row))]
[MemoryPackUnion(45, typeof(Match))]
[MemoryPackUnion(46, typeof(Index))]
[MemoryPackUnion(47, typeof(VLookup))]
[MemoryPackUnion(48, typeof(XLookup))]
[MemoryPackUnion(49, typeof(Offset))]
[MemoryPackUnion(50, typeof(NameReference))]
[MemoryPackUnion(51, typeof(Let))]
[MemoryPackUnion(52, typeof(Text))]
[MemoryPackUnion(53, typeof(SheetNumber))]
[MemoryPackUnion(54, typeof(UnionReference))]
public abstract partial record Expression
{
    public abstract object? Compute(EvaluationContext context);

    // Backwards-compatible entry point used by tests/benchmark and external callers.
    public object? Compute(Workbook workbook) => Compute(new EvaluationContext(workbook));

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
    public static RangeReference Range(string startId, string endId, Sheet sheet) =>
        new(startId, endId, sheet.Name);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Add(Expression left, Expression right) =>
        new(BinaryOperator.Add, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Subtract(Expression left, Expression right) =>
        new(BinaryOperator.Subtract, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Divide(Expression left, Expression right) =>
        new(BinaryOperator.Divide, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation Power(Expression left, Expression right) =>
        new(BinaryOperator.Power, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static BinaryOperation GreaterThan(Expression left, Expression right) =>
        new(BinaryOperator.GreaterThan, left, right);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static UnaryOperation Negate(Expression operand) => new(UnaryOperator.Negate, operand);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static UnaryOperation Plus(Expression operand) => new(UnaryOperator.Plus, operand);
}
