using System.Runtime.CompilerServices;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

// MemoryPackUnion tags are APPEND-ONLY: never renumber, reorder or reuse an existing tag,
// or previously serialized data (and the WorkbookTests round-trip) will break. Add new tags at 124+.
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
[MemoryPackUnion(55, typeof(Pmt))]
[MemoryPackUnion(56, typeof(Pv))]
[MemoryPackUnion(57, typeof(Fv))]
[MemoryPackUnion(58, typeof(Nper))]
[MemoryPackUnion(59, typeof(Ipmt))]
[MemoryPackUnion(60, typeof(Ppmt))]
[MemoryPackUnion(61, typeof(Npv))]
[MemoryPackUnion(62, typeof(Rate))]
[MemoryPackUnion(63, typeof(Irr))]
[MemoryPackUnion(64, typeof(Sqrt))]
[MemoryPackUnion(65, typeof(Power))]
[MemoryPackUnion(66, typeof(Exp))]
[MemoryPackUnion(67, typeof(Ln))]
[MemoryPackUnion(68, typeof(Log))]
[MemoryPackUnion(69, typeof(Log10))]
[MemoryPackUnion(70, typeof(SqrtPi))]
[MemoryPackUnion(71, typeof(RoundDown))]
[MemoryPackUnion(72, typeof(Trunc))]
[MemoryPackUnion(73, typeof(MRound))]
[MemoryPackUnion(74, typeof(Ceiling))]
[MemoryPackUnion(75, typeof(CeilingMath))]
[MemoryPackUnion(76, typeof(CeilingPrecise))]
[MemoryPackUnion(77, typeof(IsoCeiling))]
[MemoryPackUnion(78, typeof(Floor))]
[MemoryPackUnion(79, typeof(FloorMath))]
[MemoryPackUnion(80, typeof(FloorPrecise))]
[MemoryPackUnion(81, typeof(Even))]
[MemoryPackUnion(82, typeof(Odd))]
[MemoryPackUnion(83, typeof(Mod))]
[MemoryPackUnion(84, typeof(Quotient))]
[MemoryPackUnion(85, typeof(Sign))]
[MemoryPackUnion(86, typeof(Pi))]
[MemoryPackUnion(87, typeof(Product))]
[MemoryPackUnion(88, typeof(SumSq))]
[MemoryPackUnion(89, typeof(Multinomial))]
[MemoryPackUnion(90, typeof(SeriesSum))]
[MemoryPackUnion(91, typeof(Fact))]
[MemoryPackUnion(92, typeof(FactDouble))]
[MemoryPackUnion(93, typeof(Combin))]
[MemoryPackUnion(94, typeof(CombinA))]
[MemoryPackUnion(95, typeof(Gcd))]
[MemoryPackUnion(96, typeof(Lcm))]
[MemoryPackUnion(97, typeof(Sin))]
[MemoryPackUnion(98, typeof(Cos))]
[MemoryPackUnion(99, typeof(Tan))]
[MemoryPackUnion(100, typeof(Cot))]
[MemoryPackUnion(101, typeof(Sec))]
[MemoryPackUnion(102, typeof(Csc))]
[MemoryPackUnion(103, typeof(Asin))]
[MemoryPackUnion(104, typeof(Acos))]
[MemoryPackUnion(105, typeof(Atan))]
[MemoryPackUnion(106, typeof(Atan2))]
[MemoryPackUnion(107, typeof(Acot))]
[MemoryPackUnion(108, typeof(Sinh))]
[MemoryPackUnion(109, typeof(Cosh))]
[MemoryPackUnion(110, typeof(Tanh))]
[MemoryPackUnion(111, typeof(Coth))]
[MemoryPackUnion(112, typeof(Sech))]
[MemoryPackUnion(113, typeof(Csch))]
[MemoryPackUnion(114, typeof(Asinh))]
[MemoryPackUnion(115, typeof(Acosh))]
[MemoryPackUnion(116, typeof(Atanh))]
[MemoryPackUnion(117, typeof(Acoth))]
[MemoryPackUnion(118, typeof(Degrees))]
[MemoryPackUnion(119, typeof(Radians))]
[MemoryPackUnion(120, typeof(Base))]
[MemoryPackUnion(121, typeof(DecimalNumber))]
[MemoryPackUnion(122, typeof(Roman))]
[MemoryPackUnion(123, typeof(Arabic))]
[MemoryPackUnion(124, typeof(TrueFunction))]
[MemoryPackUnion(125, typeof(FalseFunction))]
[MemoryPackUnion(126, typeof(Xor))]
[MemoryPackUnion(127, typeof(Ifs))]
[MemoryPackUnion(128, typeof(Switch))]
[MemoryPackUnion(129, typeof(Na))]
[MemoryPackUnion(130, typeof(IsError))]
[MemoryPackUnion(131, typeof(IsErr))]
[MemoryPackUnion(132, typeof(IsNa))]
[MemoryPackUnion(133, typeof(IsText))]
[MemoryPackUnion(134, typeof(IsNonText))]
[MemoryPackUnion(135, typeof(IsLogical))]
[MemoryPackUnion(136, typeof(IsEven))]
[MemoryPackUnion(137, typeof(IsOdd))]
[MemoryPackUnion(138, typeof(IsRef))]
[MemoryPackUnion(139, typeof(IsFormula))]
[MemoryPackUnion(140, typeof(N))]
[MemoryPackUnion(141, typeof(T))]
[MemoryPackUnion(142, typeof(TypeFunction))]
[MemoryPackUnion(143, typeof(ErrorType))]
[MemoryPackUnion(144, typeof(SheetsCount))]
[MemoryPackUnion(145, typeof(Right))]
[MemoryPackUnion(146, typeof(Find))]
[MemoryPackUnion(147, typeof(Search))]
[MemoryPackUnion(148, typeof(Replace))]
[MemoryPackUnion(149, typeof(Substitute))]
[MemoryPackUnion(150, typeof(Rept))]
[MemoryPackUnion(151, typeof(Proper))]
[MemoryPackUnion(152, typeof(Exact))]
[MemoryPackUnion(153, typeof(CharFunction))]
[MemoryPackUnion(154, typeof(Code))]
[MemoryPackUnion(155, typeof(UniChar))]
[MemoryPackUnion(156, typeof(Unicode))]
[MemoryPackUnion(157, typeof(Clean))]
[MemoryPackUnion(158, typeof(Fixed))]
[MemoryPackUnion(159, typeof(Dollar))]
[MemoryPackUnion(160, typeof(NumberValueFunction))]
[MemoryPackUnion(161, typeof(TextBefore))]
[MemoryPackUnion(162, typeof(TextAfter))]
[MemoryPackUnion(163, typeof(ValueToText))]
[MemoryPackUnion(164, typeof(RegexTest))]
[MemoryPackUnion(165, typeof(RegexExtract))]
[MemoryPackUnion(166, typeof(RegexReplace))]
public abstract partial record Expression
{
    // The one evaluation contract: evaluate the node to a value type, with no boxing. Callers that want a
    // loosely-typed value call `.AsObject()` on the result.
    public abstract ComputedValue Evaluate(EvaluationContext context);

    public ComputedValue Evaluate(Workbook workbook) => Evaluate(new EvaluationContext(workbook));

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static NumberValue Number(double value) => new(value);

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public static StringValue String(string value) => new(value);

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
