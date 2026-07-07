using System.Runtime.CompilerServices;
using Danfma.MySheet.Expressions.Dates;
using Danfma.MySheet.Expressions.Financial;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using Danfma.MySheet.Expressions.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

// MemoryPackUnion tags are APPEND-ONLY: never renumber, reorder or reuse an existing tag,
// or previously serialized data (and the WorkbookTests round-trip) will break. Add new tags at 316+.
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
[MemoryPackUnion(46, typeof(Lookup.Index))]
[MemoryPackUnion(47, typeof(VLookup))]
[MemoryPackUnion(48, typeof(XLookup))]
[MemoryPackUnion(49, typeof(Offset))]
[MemoryPackUnion(50, typeof(NameReference))]
[MemoryPackUnion(51, typeof(Let))]
// `Text.Text`/`Lookup.Lookup`: inside this namespace the simple names `Text` and `Lookup` bind to the
// child NAMESPACES (namespace members win over using-imports), so those two records need qualification.
[MemoryPackUnion(52, typeof(Text.Text))]
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
[MemoryPackUnion(167, typeof(Choose))]
[MemoryPackUnion(168, typeof(HLookup))]
[MemoryPackUnion(169, typeof(Lookup.Lookup))]
[MemoryPackUnion(170, typeof(Column))]
[MemoryPackUnion(171, typeof(Columns))]
[MemoryPackUnion(172, typeof(XMatch))]
[MemoryPackUnion(173, typeof(Address))]
[MemoryPackUnion(174, typeof(Areas))]
[MemoryPackUnion(175, typeof(FormulaText))]
[MemoryPackUnion(176, typeof(AverageIf))]
[MemoryPackUnion(177, typeof(AverageIfs))]
[MemoryPackUnion(178, typeof(MaxIfs))]
[MemoryPackUnion(179, typeof(MinIfs))]
[MemoryPackUnion(180, typeof(AverageA))]
[MemoryPackUnion(181, typeof(MaxA))]
[MemoryPackUnion(182, typeof(MinA))]
[MemoryPackUnion(183, typeof(SumProduct))]
[MemoryPackUnion(184, typeof(SumX2MY2))]
[MemoryPackUnion(185, typeof(SumX2PY2))]
[MemoryPackUnion(186, typeof(SumXMY2))]
[MemoryPackUnion(187, typeof(Subtotal))]
[MemoryPackUnion(188, typeof(Median))]
[MemoryPackUnion(189, typeof(ModeSngl))]
[MemoryPackUnion(190, typeof(Large))]
[MemoryPackUnion(191, typeof(Small))]
[MemoryPackUnion(192, typeof(RankEq))]
[MemoryPackUnion(193, typeof(RankAvg))]
[MemoryPackUnion(194, typeof(PercentileInc))]
[MemoryPackUnion(195, typeof(PercentileExc))]
[MemoryPackUnion(196, typeof(PercentRankInc))]
[MemoryPackUnion(197, typeof(PercentRankExc))]
[MemoryPackUnion(198, typeof(QuartileInc))]
[MemoryPackUnion(199, typeof(QuartileExc))]
[MemoryPackUnion(200, typeof(TrimMean))]
[MemoryPackUnion(201, typeof(StDevS))]
[MemoryPackUnion(202, typeof(StDevP))]
[MemoryPackUnion(203, typeof(StDevA))]
[MemoryPackUnion(204, typeof(StDevPA))]
[MemoryPackUnion(205, typeof(VarS))]
[MemoryPackUnion(206, typeof(VarP))]
[MemoryPackUnion(207, typeof(VarA))]
[MemoryPackUnion(208, typeof(VarPA))]
[MemoryPackUnion(209, typeof(AveDev))]
[MemoryPackUnion(210, typeof(DevSq))]
[MemoryPackUnion(211, typeof(GeoMean))]
[MemoryPackUnion(212, typeof(HarMean))]
[MemoryPackUnion(213, typeof(Skew))]
[MemoryPackUnion(214, typeof(SkewP))]
[MemoryPackUnion(215, typeof(Kurt))]
[MemoryPackUnion(216, typeof(Standardize))]
[MemoryPackUnion(217, typeof(Correl))]
[MemoryPackUnion(218, typeof(Pearson))]
[MemoryPackUnion(219, typeof(CovarianceP))]
[MemoryPackUnion(220, typeof(CovarianceS))]
[MemoryPackUnion(221, typeof(Rsq))]
[MemoryPackUnion(222, typeof(Slope))]
[MemoryPackUnion(223, typeof(Intercept))]
[MemoryPackUnion(224, typeof(Steyx))]
[MemoryPackUnion(225, typeof(ForecastLinear))]
[MemoryPackUnion(226, typeof(Fisher))]
[MemoryPackUnion(227, typeof(FisherInv))]
[MemoryPackUnion(228, typeof(Phi))]
[MemoryPackUnion(229, typeof(Permut))]
[MemoryPackUnion(230, typeof(PermutationA))]
[MemoryPackUnion(231, typeof(Prob))]
// Compatibility aliases: distinct nodes (never the modern record) so the un-parse preserves the
// legacy spelling the user wrote. Qualified like `Text.Text` above (child namespace wins).
[MemoryPackUnion(232, typeof(Compatibility.Mode))]
[MemoryPackUnion(233, typeof(Compatibility.StDev))]
[MemoryPackUnion(234, typeof(Compatibility.StDevP))]
[MemoryPackUnion(235, typeof(Compatibility.Var))]
[MemoryPackUnion(236, typeof(Compatibility.VarP))]
[MemoryPackUnion(237, typeof(Compatibility.Rank))]
[MemoryPackUnion(238, typeof(Compatibility.Percentile))]
[MemoryPackUnion(239, typeof(Compatibility.PercentRank))]
[MemoryPackUnion(240, typeof(Compatibility.Quartile))]
[MemoryPackUnion(241, typeof(Compatibility.Covar))]
[MemoryPackUnion(242, typeof(Compatibility.Forecast))]
// Wave 5 — Date and time. `Dates` (not `DateTime`) avoids colliding with System.DateTime.
[MemoryPackUnion(243, typeof(Date))]
[MemoryPackUnion(244, typeof(Time))]
[MemoryPackUnion(245, typeof(DateValue))]
[MemoryPackUnion(246, typeof(TimeValue))]
[MemoryPackUnion(247, typeof(Year))]
[MemoryPackUnion(248, typeof(Month))]
[MemoryPackUnion(249, typeof(Day))]
[MemoryPackUnion(250, typeof(Hour))]
[MemoryPackUnion(251, typeof(Minute))]
[MemoryPackUnion(252, typeof(Second))]
[MemoryPackUnion(253, typeof(Days))]
[MemoryPackUnion(254, typeof(Days360))]
[MemoryPackUnion(255, typeof(EDate))]
[MemoryPackUnion(256, typeof(EoMonth))]
[MemoryPackUnion(257, typeof(Weekday))]
[MemoryPackUnion(258, typeof(WeekNum))]
[MemoryPackUnion(259, typeof(IsoWeekNum))]
[MemoryPackUnion(260, typeof(DateDif))]
[MemoryPackUnion(261, typeof(YearFrac))]
[MemoryPackUnion(262, typeof(NetworkDays))]
[MemoryPackUnion(263, typeof(NetworkDaysIntl))]
[MemoryPackUnion(264, typeof(Workday))]
[MemoryPackUnion(265, typeof(WorkdayIntl))]
[MemoryPackUnion(266, typeof(Sln))]
[MemoryPackUnion(267, typeof(Syd))]
[MemoryPackUnion(268, typeof(Db))]
[MemoryPackUnion(269, typeof(Ddb))]
[MemoryPackUnion(270, typeof(Vdb))]
[MemoryPackUnion(271, typeof(AmorLinc))]
[MemoryPackUnion(272, typeof(AmorDegrc))]
[MemoryPackUnion(273, typeof(Effect))]
[MemoryPackUnion(274, typeof(Nominal))]
[MemoryPackUnion(275, typeof(Mirr))]
[MemoryPackUnion(276, typeof(Rri))]
[MemoryPackUnion(277, typeof(PDuration))]
[MemoryPackUnion(278, typeof(ISPmt))]
[MemoryPackUnion(279, typeof(CumIPmt))]
[MemoryPackUnion(280, typeof(CumPrinc))]
[MemoryPackUnion(281, typeof(FvSchedule))]
[MemoryPackUnion(282, typeof(DollarDe))]
[MemoryPackUnion(283, typeof(DollarFr))]
[MemoryPackUnion(284, typeof(XNpv))]
[MemoryPackUnion(285, typeof(XIrr))]
[MemoryPackUnion(286, typeof(AccrInt))]
[MemoryPackUnion(287, typeof(AccrIntM))]
[MemoryPackUnion(288, typeof(Disc))]
[MemoryPackUnion(289, typeof(IntRate))]
[MemoryPackUnion(290, typeof(Received))]
[MemoryPackUnion(291, typeof(PriceDisc))]
[MemoryPackUnion(292, typeof(PriceMat))]
[MemoryPackUnion(293, typeof(YieldDisc))]
[MemoryPackUnion(294, typeof(YieldMat))]
[MemoryPackUnion(295, typeof(TBillEq))]
[MemoryPackUnion(296, typeof(TBillPrice))]
[MemoryPackUnion(297, typeof(TBillYield))]
[MemoryPackUnion(298, typeof(CoupPcd))]
[MemoryPackUnion(299, typeof(CoupNcd))]
[MemoryPackUnion(300, typeof(CoupNum))]
[MemoryPackUnion(301, typeof(CoupDays))]
[MemoryPackUnion(302, typeof(CoupDayBs))]
[MemoryPackUnion(303, typeof(CoupDaysNc))]
[MemoryPackUnion(304, typeof(Price))]
[MemoryPackUnion(305, typeof(Yield))]
[MemoryPackUnion(306, typeof(Duration))]
[MemoryPackUnion(307, typeof(MDuration))]
[MemoryPackUnion(308, typeof(OddFPrice))]
[MemoryPackUnion(309, typeof(OddFYield))]
[MemoryPackUnion(310, typeof(OddLPrice))]
[MemoryPackUnion(311, typeof(OddLYield))]
// F1 — volatile functions. NOW/TODAY read the clock; RAND/RANDBETWEEN the RNG.
[MemoryPackUnion(312, typeof(Now))]
[MemoryPackUnion(313, typeof(Today))]
[MemoryPackUnion(314, typeof(Rand))]
[MemoryPackUnion(315, typeof(RandBetween))]
// Whole-column / whole-row / one-sided open references (A:A, 1:1, A2:A, A:A10, A1:C).
[MemoryPackUnion(316, typeof(OpenRangeReference))]
[MemoryPackUnion(317, typeof(DynamicRange))]
[MemoryPackUnion(318, typeof(Lookup.Indirect))]
public abstract partial record Expression
{
    // The one evaluation contract: evaluate the node to a value type, with no boxing. Callers that want a
    // loosely-typed value call `.AsObject()` on the result.
    public abstract ComputedValue Evaluate(EvaluationContext context);

    /// <summary>
    /// Resolves this expression to its target <see cref="Reference"/> WITHOUT dereferencing a single cell
    /// to its value (unlike <see cref="Evaluate"/>). Used by the ':' range operator and reference-context
    /// consumers. Default: not a reference-producing expression.
    /// </summary>
    public virtual bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;
        return false;
    }

    /// <summary>
    /// Whether this node is intrinsically volatile — a clock or random source (<c>NOW</c>, <c>TODAY</c>,
    /// <c>RAND</c>, <c>RANDBETWEEN</c>) whose value can change between epochs. Pure introspection: the actual
    /// cache/refresh contagion is driven at evaluation time by <see cref="Workbook.MarkVolatileTouched"/>,
    /// which volatile nodes call from <see cref="Evaluate"/> (so dependents become volatile transitively).
    /// Not serialized (behavior, not state).
    /// </summary>
    [MemoryPackIgnore]
    public virtual bool IsVolatile => false;

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
