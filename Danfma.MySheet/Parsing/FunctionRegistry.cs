using System.Collections.Frozen;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Dates;
using Danfma.MySheet.Expressions.Financial;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using Danfma.MySheet.Expressions.Text;
using Compat = Danfma.MySheet.Expressions.Compatibility;

namespace Danfma.MySheet.Parsing;

/// <summary>
/// Whether the mini-CSE (<see cref="Danfma.MySheet.Expressions.ArrayEvaluation"/>) may lift a built-in over
/// an array argument element by element.
/// </summary>
/// <remarks>
/// <see cref="Consumes"/> is the ZERO value on purpose: lifting is opt-in, so a default-constructed entry —
/// or a future entry whose author forgets the flag — keeps today's behaviour (the function is evaluated once,
/// as an opaque scalar) instead of silently answering from a single element of the rectangle it was meant to
/// consume whole. The two failure modes are not symmetric: forgetting the flag on a new scalar function only
/// loses the feature, while forgetting it on a new range-aware one produces wrong numbers with no error.
/// </remarks>
internal enum ArrayLifting : byte
{
    /// <summary>
    /// The function reads ranges, arrays or references itself (an aggregate, a criteria scan, a lookup, a
    /// shape or position query). It is NEVER lifted: it is handed the whole argument, as today.
    /// </summary>
    Consumes,

    /// <summary>
    /// The function is pure-scalar — every argument is a single value — so the mini-CSE may evaluate its node
    /// once per element of an array argument, broadcasting the scalar arguments.
    /// </summary>
    Elementwise,
}

/// <summary>
/// The single source of truth for every built-in Excel function: name, arity, the factory that builds its
/// AST node (used by <see cref="Parser"/>), the accessor that reads its argument array back out (used by
/// <see cref="FormulaWriter"/>), and its <see cref="ArrayLifting"/> classification. One entry per function
/// serves all of those, so adding a function is one line here — <c>Entry&lt;T&gt;</c> for a function that
/// consumes ranges/arrays itself, <c>Elementwise&lt;T&gt;</c> for a pure-scalar one the mini-CSE may lift per
/// element — plus the node itself and its <see cref="Expression"/> union tag, with no hand-syncing of a
/// parser table and a writer table. Lifting is opt-in (<see cref="ArrayLifting.Consumes"/> is the default)
/// because a forgotten flag on a scalar function only loses the feature, while a forgotten flag on a
/// range-aware one would answer from one element and be silently wrong.
/// </summary>
/// <remarks>
/// Excludes <see cref="FunctionCall"/>: that node is not a built-in, it is the runtime fallback the Parser
/// produces for an unrecognized name (resolved later against the workbook's custom-function registry), so it
/// has no fixed arity or table entry here. The Parser and FormulaWriter each keep their own one-line special
/// case for it.
/// </remarks>
internal static class FunctionRegistry
{
    /// <summary>
    /// One built-in function: its Excel name, argument-count bounds, the factory that builds its AST node
    /// from parsed arguments, its CLR node type, the accessor that reads the arguments back out of an
    /// instance of that node (almost always its <c>Arguments</c> property — <see cref="Sum"/> is the one
    /// exception, whose parameter is named <c>Expressions</c>), and whether the mini-CSE may lift it over an
    /// array element by element.
    /// </summary>
    internal readonly record struct RegistryEntry(
        string Name,
        int MinArgs,
        int MaxArgs,
        Func<Expression[], Expression> Create,
        Type NodeType,
        Func<Function, Expression[]> GetArguments,
        ArrayLifting Lifting
    );

    /// <summary>
    /// Throws when an entry flags an <see cref="IArrayProducer"/> node
    /// <see cref="ArrayLifting.Elementwise"/>. Run over the whole table at type initialization (see the
    /// static constructor), so the offender is named before any formula is parsed.
    /// </summary>
    /// <remarks>
    /// A producer flagged Elementwise does not merely answer from one cell — it takes the HOST DOWN. The
    /// lift arm sits BEFORE the producer arm in both <see cref="ArrayEvaluation"/> switches, so the lift asks
    /// the node for its scalar, the producer's scalar rule is <c>FirstElement</c>, <c>FirstElement</c> builds
    /// the array operand, that build re-enters <c>TryGetLift</c>, and the cycle recurses until the stack
    /// overflows — a crash with no failing test, which is strictly worse than a wrong number a guard can
    /// name. Measured by Phase 7's final review, which mis-flagged <c>SEQUENCE</c> the way a contributor
    /// would and got a stack overflow instead of a red suite. C# cannot express "T does not implement
    /// IArrayProducer" as a generic constraint, which is why this is a runtime check and not a
    /// <c>where</c> clause on <see cref="Elementwise{T}"/>.
    /// </remarks>
    internal static void RequireProducerIsConsumes(in RegistryEntry entry)
    {
        if (
            entry.Lifting is ArrayLifting.Elementwise
            && typeof(IArrayProducer).IsAssignableFrom(entry.NodeType)
        )
        {
            throw new InvalidOperationException(
                $"{entry.Name} ({entry.NodeType.Name}) implements IArrayProducer and must be registered "
                    + "with Entry<T> (ArrayLifting.Consumes), never Elementwise<T>. The lift arm precedes "
                    + "the producer arm in ArrayEvaluation's two switches, so the producer's FirstElement "
                    + "rule re-enters the lift and recurses until the stack overflows."
            );
        }
    }

    static FunctionRegistry()
    {
        foreach (var entry in Entries)
        {
            RequireProducerIsConsumes(entry);
        }
    }

    // Ordinal by declaration order below (mirrors the historical Parser table / FormulaWriter dictionary
    // grouping — arithmetic/logical/text/lookup, financial, math, statistical, compatibility aliases, then
    // date/time and the F1 volatile functions).
    private static readonly RegistryEntry[] Entries =
    [
        Entry<Sum>(
            "SUM",
            0,
            int.MaxValue,
            static arguments => new Sum(arguments),
            static f => ((Sum)f).Expressions
        ),
        Entry<Average>(
            "AVERAGE",
            0,
            int.MaxValue,
            static arguments => new Average(arguments),
            static f => ((Average)f).Arguments
        ),
        Entry<Min>(
            "MIN",
            0,
            int.MaxValue,
            static arguments => new Min(arguments),
            static f => ((Min)f).Arguments
        ),
        Entry<Max>(
            "MAX",
            0,
            int.MaxValue,
            static arguments => new Max(arguments),
            static f => ((Max)f).Arguments
        ),
        Entry<Count>(
            "COUNT",
            0,
            int.MaxValue,
            static arguments => new Count(arguments),
            static f => ((Count)f).Arguments
        ),
        Entry<If>("IF", 2, 3, static arguments => new If(arguments), static f => ((If)f).Arguments),
        Entry<And>(
            "AND",
            1,
            int.MaxValue,
            static arguments => new And(arguments),
            static f => ((And)f).Arguments
        ),
        Entry<Or>(
            "OR",
            1,
            int.MaxValue,
            static arguments => new Or(arguments),
            static f => ((Or)f).Arguments
        ),
        Elementwise<Not>(
            "NOT",
            1,
            1,
            static arguments => new Not(arguments),
            static f => ((Not)f).Arguments
        ),
        Elementwise<IfError>(
            "IFERROR",
            2,
            2,
            static arguments => new IfError(arguments),
            static f => ((IfError)f).Arguments
        ),
        Elementwise<Int>(
            "INT",
            1,
            1,
            static arguments => new Int(arguments),
            static f => ((Int)f).Arguments
        ),
        Elementwise<Round>(
            "ROUND",
            2,
            2,
            static arguments => new Round(arguments),
            static f => ((Round)f).Arguments
        ),
        Elementwise<RoundUp>(
            "ROUNDUP",
            2,
            2,
            static arguments => new RoundUp(arguments),
            static f => ((RoundUp)f).Arguments
        ),
        Elementwise<Abs>(
            "ABS",
            1,
            1,
            static arguments => new Abs(arguments),
            static f => ((Abs)f).Arguments
        ),
        Elementwise<IsNumber>(
            "ISNUMBER",
            1,
            1,
            static arguments => new IsNumber(arguments),
            static f => ((IsNumber)f).Arguments
        ),
        Elementwise<IsBlank>(
            "ISBLANK",
            1,
            1,
            static arguments => new IsBlank(arguments),
            static f => ((IsBlank)f).Arguments
        ),
        Elementwise<IfNa>(
            "IFNA",
            2,
            2,
            static arguments => new IfNa(arguments),
            static f => ((IfNa)f).Arguments
        ),
        Elementwise<Upper>(
            "UPPER",
            1,
            1,
            static arguments => new Upper(arguments),
            static f => ((Upper)f).Arguments
        ),
        Elementwise<Lower>(
            "LOWER",
            1,
            1,
            static arguments => new Lower(arguments),
            static f => ((Lower)f).Arguments
        ),
        Elementwise<Trim>(
            "TRIM",
            1,
            1,
            static arguments => new Trim(arguments),
            static f => ((Trim)f).Arguments
        ),
        Elementwise<Len>(
            "LEN",
            1,
            1,
            static arguments => new Len(arguments),
            static f => ((Len)f).Arguments
        ),
        Elementwise<Left>(
            "LEFT",
            1,
            2,
            static arguments => new Left(arguments),
            static f => ((Left)f).Arguments
        ),
        Elementwise<Mid>(
            "MID",
            3,
            3,
            static arguments => new Mid(arguments),
            static f => ((Mid)f).Arguments
        ),
        Elementwise<Value>(
            "VALUE",
            1,
            1,
            static arguments => new Value(arguments),
            static f => ((Value)f).Arguments
        ),
        Entry<Concat>(
            "CONCAT",
            1,
            int.MaxValue,
            static arguments => new Concat(arguments),
            static f => ((Concat)f).Arguments
        ),
        Entry<Concatenate>(
            "CONCATENATE",
            1,
            int.MaxValue,
            static arguments => new Concatenate(arguments),
            static f => ((Concatenate)f).Arguments
        ),
        Entry<TextJoin>(
            "TEXTJOIN",
            3,
            int.MaxValue,
            static arguments => new TextJoin(arguments),
            static f => ((TextJoin)f).Arguments
        ),
        Entry<CountA>(
            "COUNTA",
            1,
            int.MaxValue,
            static arguments => new CountA(arguments),
            static f => ((CountA)f).Arguments
        ),
        Entry<CountBlank>(
            "COUNTBLANK",
            1,
            int.MaxValue,
            static arguments => new CountBlank(arguments),
            static f => ((CountBlank)f).Arguments
        ),
        Entry<CountIf>(
            "COUNTIF",
            2,
            2,
            static arguments => new CountIf(arguments),
            static f => ((CountIf)f).Arguments
        ),
        Entry<CountIfs>(
            "COUNTIFS",
            2,
            int.MaxValue,
            static arguments => new CountIfs(arguments),
            static f => ((CountIfs)f).Arguments
        ),
        Entry<SumIf>(
            "SUMIF",
            2,
            3,
            static arguments => new SumIf(arguments),
            static f => ((SumIf)f).Arguments
        ),
        Entry<SumIfs>(
            "SUMIFS",
            3,
            int.MaxValue,
            static arguments => new SumIfs(arguments),
            static f => ((SumIfs)f).Arguments
        ),
        Entry<Rows>(
            "ROWS",
            1,
            1,
            static arguments => new Rows(arguments),
            static f => ((Rows)f).Arguments
        ),
        Entry<Row>(
            "ROW",
            0,
            1,
            static arguments => new Row(arguments),
            static f => ((Row)f).Arguments
        ),
        Entry<Match>(
            "MATCH",
            2,
            3,
            static arguments => new Match(arguments),
            static f => ((Match)f).Arguments
        ),
        Entry<MySheet.Expressions.Lookup.Index>(
            "INDEX",
            2,
            3,
            static arguments => new MySheet.Expressions.Lookup.Index(arguments),
            static f => ((MySheet.Expressions.Lookup.Index)f).Arguments
        ),
        Entry<VLookup>(
            "VLOOKUP",
            3,
            4,
            static arguments => new VLookup(arguments),
            static f => ((VLookup)f).Arguments
        ),
        Entry<XLookup>(
            "XLOOKUP",
            3,
            6,
            static arguments => new XLookup(arguments),
            static f => ((XLookup)f).Arguments
        ),
        Entry<Offset>(
            "OFFSET",
            3,
            5,
            static arguments => new Offset(arguments),
            static f => ((Offset)f).Arguments
        ),
        Entry<MySheet.Expressions.Lookup.Indirect>(
            "INDIRECT",
            1,
            2,
            static arguments => new MySheet.Expressions.Lookup.Indirect(arguments),
            static f => ((MySheet.Expressions.Lookup.Indirect)f).Arguments
        ),
        Entry<Let>(
            "LET",
            3,
            int.MaxValue,
            static arguments => new Let(arguments),
            static f => ((Let)f).Arguments
        ),
        Elementwise<Text>(
            "TEXT",
            2,
            2,
            static arguments => new Text(arguments),
            static f => ((Text)f).Arguments
        ),
        Entry<SheetNumber>(
            "SHEET",
            0,
            1,
            static arguments => new SheetNumber(arguments),
            static f => ((SheetNumber)f).Arguments
        ),
        Elementwise<Pmt>(
            "PMT",
            3,
            5,
            static arguments => new Pmt(arguments),
            static f => ((Pmt)f).Arguments
        ),
        Elementwise<Pv>(
            "PV",
            3,
            5,
            static arguments => new Pv(arguments),
            static f => ((Pv)f).Arguments
        ),
        Elementwise<Fv>(
            "FV",
            3,
            5,
            static arguments => new Fv(arguments),
            static f => ((Fv)f).Arguments
        ),
        Elementwise<Nper>(
            "NPER",
            3,
            5,
            static arguments => new Nper(arguments),
            static f => ((Nper)f).Arguments
        ),
        Elementwise<Ipmt>(
            "IPMT",
            4,
            6,
            static arguments => new Ipmt(arguments),
            static f => ((Ipmt)f).Arguments
        ),
        Elementwise<Ppmt>(
            "PPMT",
            4,
            6,
            static arguments => new Ppmt(arguments),
            static f => ((Ppmt)f).Arguments
        ),
        Entry<Npv>(
            "NPV",
            2,
            int.MaxValue,
            static arguments => new Npv(arguments),
            static f => ((Npv)f).Arguments
        ),
        Elementwise<Rate>(
            "RATE",
            3,
            6,
            static arguments => new Rate(arguments),
            static f => ((Rate)f).Arguments
        ),
        Entry<Irr>(
            "IRR",
            1,
            2,
            static arguments => new Irr(arguments),
            static f => ((Irr)f).Arguments
        ),
        Elementwise<Sln>(
            "SLN",
            3,
            3,
            static arguments => new Sln(arguments),
            static f => ((Sln)f).Arguments
        ),
        Elementwise<Syd>(
            "SYD",
            4,
            4,
            static arguments => new Syd(arguments),
            static f => ((Syd)f).Arguments
        ),
        Elementwise<Db>(
            "DB",
            4,
            5,
            static arguments => new Db(arguments),
            static f => ((Db)f).Arguments
        ),
        Elementwise<Ddb>(
            "DDB",
            4,
            5,
            static arguments => new Ddb(arguments),
            static f => ((Ddb)f).Arguments
        ),
        Elementwise<Vdb>(
            "VDB",
            5,
            7,
            static arguments => new Vdb(arguments),
            static f => ((Vdb)f).Arguments
        ),
        Elementwise<AmorLinc>(
            "AMORLINC",
            6,
            7,
            static arguments => new AmorLinc(arguments),
            static f => ((AmorLinc)f).Arguments
        ),
        Elementwise<AmorDegrc>(
            "AMORDEGRC",
            6,
            7,
            static arguments => new AmorDegrc(arguments),
            static f => ((AmorDegrc)f).Arguments
        ),
        Elementwise<Effect>(
            "EFFECT",
            2,
            2,
            static arguments => new Effect(arguments),
            static f => ((Effect)f).Arguments
        ),
        Elementwise<Nominal>(
            "NOMINAL",
            2,
            2,
            static arguments => new Nominal(arguments),
            static f => ((Nominal)f).Arguments
        ),
        Entry<Mirr>(
            "MIRR",
            3,
            3,
            static arguments => new Mirr(arguments),
            static f => ((Mirr)f).Arguments
        ),
        Elementwise<Rri>(
            "RRI",
            3,
            3,
            static arguments => new Rri(arguments),
            static f => ((Rri)f).Arguments
        ),
        Elementwise<PDuration>(
            "PDURATION",
            3,
            3,
            static arguments => new PDuration(arguments),
            static f => ((PDuration)f).Arguments
        ),
        Elementwise<ISPmt>(
            "ISPMT",
            4,
            4,
            static arguments => new ISPmt(arguments),
            static f => ((ISPmt)f).Arguments
        ),
        Elementwise<CumIPmt>(
            "CUMIPMT",
            6,
            6,
            static arguments => new CumIPmt(arguments),
            static f => ((CumIPmt)f).Arguments
        ),
        Elementwise<CumPrinc>(
            "CUMPRINC",
            6,
            6,
            static arguments => new CumPrinc(arguments),
            static f => ((CumPrinc)f).Arguments
        ),
        Entry<FvSchedule>(
            "FVSCHEDULE",
            2,
            2,
            static arguments => new FvSchedule(arguments),
            static f => ((FvSchedule)f).Arguments
        ),
        Elementwise<DollarDe>(
            "DOLLARDE",
            2,
            2,
            static arguments => new DollarDe(arguments),
            static f => ((DollarDe)f).Arguments
        ),
        Elementwise<DollarFr>(
            "DOLLARFR",
            2,
            2,
            static arguments => new DollarFr(arguments),
            static f => ((DollarFr)f).Arguments
        ),
        Entry<XNpv>(
            "XNPV",
            3,
            3,
            static arguments => new XNpv(arguments),
            static f => ((XNpv)f).Arguments
        ),
        Entry<XIrr>(
            "XIRR",
            2,
            3,
            static arguments => new XIrr(arguments),
            static f => ((XIrr)f).Arguments
        ),
        Elementwise<AccrInt>(
            "ACCRINT",
            6,
            8,
            static arguments => new AccrInt(arguments),
            static f => ((AccrInt)f).Arguments
        ),
        Elementwise<AccrIntM>(
            "ACCRINTM",
            4,
            5,
            static arguments => new AccrIntM(arguments),
            static f => ((AccrIntM)f).Arguments
        ),
        Elementwise<Disc>(
            "DISC",
            4,
            5,
            static arguments => new Disc(arguments),
            static f => ((Disc)f).Arguments
        ),
        Elementwise<IntRate>(
            "INTRATE",
            4,
            5,
            static arguments => new IntRate(arguments),
            static f => ((IntRate)f).Arguments
        ),
        Elementwise<Received>(
            "RECEIVED",
            4,
            5,
            static arguments => new Received(arguments),
            static f => ((Received)f).Arguments
        ),
        Elementwise<PriceDisc>(
            "PRICEDISC",
            4,
            5,
            static arguments => new PriceDisc(arguments),
            static f => ((PriceDisc)f).Arguments
        ),
        Elementwise<PriceMat>(
            "PRICEMAT",
            5,
            6,
            static arguments => new PriceMat(arguments),
            static f => ((PriceMat)f).Arguments
        ),
        Elementwise<YieldDisc>(
            "YIELDDISC",
            4,
            5,
            static arguments => new YieldDisc(arguments),
            static f => ((YieldDisc)f).Arguments
        ),
        Elementwise<YieldMat>(
            "YIELDMAT",
            5,
            6,
            static arguments => new YieldMat(arguments),
            static f => ((YieldMat)f).Arguments
        ),
        Elementwise<TBillEq>(
            "TBILLEQ",
            3,
            3,
            static arguments => new TBillEq(arguments),
            static f => ((TBillEq)f).Arguments
        ),
        Elementwise<TBillPrice>(
            "TBILLPRICE",
            3,
            3,
            static arguments => new TBillPrice(arguments),
            static f => ((TBillPrice)f).Arguments
        ),
        Elementwise<TBillYield>(
            "TBILLYIELD",
            3,
            3,
            static arguments => new TBillYield(arguments),
            static f => ((TBillYield)f).Arguments
        ),
        Elementwise<CoupPcd>(
            "COUPPCD",
            3,
            4,
            static arguments => new CoupPcd(arguments),
            static f => ((CoupPcd)f).Arguments
        ),
        Elementwise<CoupNcd>(
            "COUPNCD",
            3,
            4,
            static arguments => new CoupNcd(arguments),
            static f => ((CoupNcd)f).Arguments
        ),
        Elementwise<CoupNum>(
            "COUPNUM",
            3,
            4,
            static arguments => new CoupNum(arguments),
            static f => ((CoupNum)f).Arguments
        ),
        Elementwise<CoupDays>(
            "COUPDAYS",
            3,
            4,
            static arguments => new CoupDays(arguments),
            static f => ((CoupDays)f).Arguments
        ),
        Elementwise<CoupDayBs>(
            "COUPDAYBS",
            3,
            4,
            static arguments => new CoupDayBs(arguments),
            static f => ((CoupDayBs)f).Arguments
        ),
        Elementwise<CoupDaysNc>(
            "COUPDAYSNC",
            3,
            4,
            static arguments => new CoupDaysNc(arguments),
            static f => ((CoupDaysNc)f).Arguments
        ),
        Elementwise<Price>(
            "PRICE",
            6,
            7,
            static arguments => new Price(arguments),
            static f => ((Price)f).Arguments
        ),
        Elementwise<Yield>(
            "YIELD",
            6,
            7,
            static arguments => new Yield(arguments),
            static f => ((Yield)f).Arguments
        ),
        Elementwise<Duration>(
            "DURATION",
            5,
            6,
            static arguments => new Duration(arguments),
            static f => ((Duration)f).Arguments
        ),
        Elementwise<MDuration>(
            "MDURATION",
            5,
            6,
            static arguments => new MDuration(arguments),
            static f => ((MDuration)f).Arguments
        ),
        Elementwise<OddFPrice>(
            "ODDFPRICE",
            8,
            9,
            static arguments => new OddFPrice(arguments),
            static f => ((OddFPrice)f).Arguments
        ),
        Elementwise<OddFYield>(
            "ODDFYIELD",
            8,
            9,
            static arguments => new OddFYield(arguments),
            static f => ((OddFYield)f).Arguments
        ),
        Elementwise<OddLPrice>(
            "ODDLPRICE",
            7,
            8,
            static arguments => new OddLPrice(arguments),
            static f => ((OddLPrice)f).Arguments
        ),
        Elementwise<OddLYield>(
            "ODDLYIELD",
            7,
            8,
            static arguments => new OddLYield(arguments),
            static f => ((OddLYield)f).Arguments
        ),
        Elementwise<Sqrt>(
            "SQRT",
            1,
            1,
            static arguments => new Sqrt(arguments),
            static f => ((Sqrt)f).Arguments
        ),
        Elementwise<Power>(
            "POWER",
            2,
            2,
            static arguments => new Power(arguments),
            static f => ((Power)f).Arguments
        ),
        Elementwise<Exp>(
            "EXP",
            1,
            1,
            static arguments => new Exp(arguments),
            static f => ((Exp)f).Arguments
        ),
        Elementwise<Ln>(
            "LN",
            1,
            1,
            static arguments => new Ln(arguments),
            static f => ((Ln)f).Arguments
        ),
        Elementwise<Log>(
            "LOG",
            1,
            2,
            static arguments => new Log(arguments),
            static f => ((Log)f).Arguments
        ),
        Elementwise<Log10>(
            "LOG10",
            1,
            1,
            static arguments => new Log10(arguments),
            static f => ((Log10)f).Arguments
        ),
        Elementwise<SqrtPi>(
            "SQRTPI",
            1,
            1,
            static arguments => new SqrtPi(arguments),
            static f => ((SqrtPi)f).Arguments
        ),
        Elementwise<RoundDown>(
            "ROUNDDOWN",
            2,
            2,
            static arguments => new RoundDown(arguments),
            static f => ((RoundDown)f).Arguments
        ),
        Elementwise<Trunc>(
            "TRUNC",
            1,
            2,
            static arguments => new Trunc(arguments),
            static f => ((Trunc)f).Arguments
        ),
        Elementwise<MRound>(
            "MROUND",
            2,
            2,
            static arguments => new MRound(arguments),
            static f => ((MRound)f).Arguments
        ),
        Elementwise<Ceiling>(
            "CEILING",
            2,
            2,
            static arguments => new Ceiling(arguments),
            static f => ((Ceiling)f).Arguments
        ),
        Elementwise<CeilingMath>(
            "CEILING.MATH",
            1,
            3,
            static arguments => new CeilingMath(arguments),
            static f => ((CeilingMath)f).Arguments
        ),
        Elementwise<CeilingPrecise>(
            "CEILING.PRECISE",
            1,
            2,
            static arguments => new CeilingPrecise(arguments),
            static f => ((CeilingPrecise)f).Arguments
        ),
        Elementwise<IsoCeiling>(
            "ISO.CEILING",
            1,
            2,
            static arguments => new IsoCeiling(arguments),
            static f => ((IsoCeiling)f).Arguments
        ),
        Elementwise<Floor>(
            "FLOOR",
            2,
            2,
            static arguments => new Floor(arguments),
            static f => ((Floor)f).Arguments
        ),
        Elementwise<FloorMath>(
            "FLOOR.MATH",
            1,
            3,
            static arguments => new FloorMath(arguments),
            static f => ((FloorMath)f).Arguments
        ),
        Elementwise<FloorPrecise>(
            "FLOOR.PRECISE",
            1,
            2,
            static arguments => new FloorPrecise(arguments),
            static f => ((FloorPrecise)f).Arguments
        ),
        Elementwise<Even>(
            "EVEN",
            1,
            1,
            static arguments => new Even(arguments),
            static f => ((Even)f).Arguments
        ),
        Elementwise<Odd>(
            "ODD",
            1,
            1,
            static arguments => new Odd(arguments),
            static f => ((Odd)f).Arguments
        ),
        Elementwise<Mod>(
            "MOD",
            2,
            2,
            static arguments => new Mod(arguments),
            static f => ((Mod)f).Arguments
        ),
        Elementwise<Quotient>(
            "QUOTIENT",
            2,
            2,
            static arguments => new Quotient(arguments),
            static f => ((Quotient)f).Arguments
        ),
        Elementwise<Sign>(
            "SIGN",
            1,
            1,
            static arguments => new Sign(arguments),
            static f => ((Sign)f).Arguments
        ),
        Entry<Pi>("PI", 0, 0, static arguments => new Pi(arguments), static f => ((Pi)f).Arguments),
        Entry<Product>(
            "PRODUCT",
            1,
            int.MaxValue,
            static arguments => new Product(arguments),
            static f => ((Product)f).Arguments
        ),
        Entry<SumSq>(
            "SUMSQ",
            1,
            int.MaxValue,
            static arguments => new SumSq(arguments),
            static f => ((SumSq)f).Arguments
        ),
        Entry<Multinomial>(
            "MULTINOMIAL",
            1,
            int.MaxValue,
            static arguments => new Multinomial(arguments),
            static f => ((Multinomial)f).Arguments
        ),
        Entry<SeriesSum>(
            "SERIESSUM",
            4,
            4,
            static arguments => new SeriesSum(arguments),
            static f => ((SeriesSum)f).Arguments
        ),
        Elementwise<Fact>(
            "FACT",
            1,
            1,
            static arguments => new Fact(arguments),
            static f => ((Fact)f).Arguments
        ),
        Elementwise<FactDouble>(
            "FACTDOUBLE",
            1,
            1,
            static arguments => new FactDouble(arguments),
            static f => ((FactDouble)f).Arguments
        ),
        Elementwise<Combin>(
            "COMBIN",
            2,
            2,
            static arguments => new Combin(arguments),
            static f => ((Combin)f).Arguments
        ),
        Elementwise<CombinA>(
            "COMBINA",
            2,
            2,
            static arguments => new CombinA(arguments),
            static f => ((CombinA)f).Arguments
        ),
        Entry<Gcd>(
            "GCD",
            1,
            int.MaxValue,
            static arguments => new Gcd(arguments),
            static f => ((Gcd)f).Arguments
        ),
        Entry<Lcm>(
            "LCM",
            1,
            int.MaxValue,
            static arguments => new Lcm(arguments),
            static f => ((Lcm)f).Arguments
        ),
        Elementwise<Sin>(
            "SIN",
            1,
            1,
            static arguments => new Sin(arguments),
            static f => ((Sin)f).Arguments
        ),
        Elementwise<Cos>(
            "COS",
            1,
            1,
            static arguments => new Cos(arguments),
            static f => ((Cos)f).Arguments
        ),
        Elementwise<Tan>(
            "TAN",
            1,
            1,
            static arguments => new Tan(arguments),
            static f => ((Tan)f).Arguments
        ),
        Elementwise<Cot>(
            "COT",
            1,
            1,
            static arguments => new Cot(arguments),
            static f => ((Cot)f).Arguments
        ),
        Elementwise<Sec>(
            "SEC",
            1,
            1,
            static arguments => new Sec(arguments),
            static f => ((Sec)f).Arguments
        ),
        Elementwise<Csc>(
            "CSC",
            1,
            1,
            static arguments => new Csc(arguments),
            static f => ((Csc)f).Arguments
        ),
        Elementwise<Asin>(
            "ASIN",
            1,
            1,
            static arguments => new Asin(arguments),
            static f => ((Asin)f).Arguments
        ),
        Elementwise<Acos>(
            "ACOS",
            1,
            1,
            static arguments => new Acos(arguments),
            static f => ((Acos)f).Arguments
        ),
        Elementwise<Atan>(
            "ATAN",
            1,
            1,
            static arguments => new Atan(arguments),
            static f => ((Atan)f).Arguments
        ),
        Elementwise<Atan2>(
            "ATAN2",
            2,
            2,
            static arguments => new Atan2(arguments),
            static f => ((Atan2)f).Arguments
        ),
        Elementwise<Acot>(
            "ACOT",
            1,
            1,
            static arguments => new Acot(arguments),
            static f => ((Acot)f).Arguments
        ),
        Elementwise<Sinh>(
            "SINH",
            1,
            1,
            static arguments => new Sinh(arguments),
            static f => ((Sinh)f).Arguments
        ),
        Elementwise<Cosh>(
            "COSH",
            1,
            1,
            static arguments => new Cosh(arguments),
            static f => ((Cosh)f).Arguments
        ),
        Elementwise<Tanh>(
            "TANH",
            1,
            1,
            static arguments => new Tanh(arguments),
            static f => ((Tanh)f).Arguments
        ),
        Elementwise<Coth>(
            "COTH",
            1,
            1,
            static arguments => new Coth(arguments),
            static f => ((Coth)f).Arguments
        ),
        Elementwise<Sech>(
            "SECH",
            1,
            1,
            static arguments => new Sech(arguments),
            static f => ((Sech)f).Arguments
        ),
        Elementwise<Csch>(
            "CSCH",
            1,
            1,
            static arguments => new Csch(arguments),
            static f => ((Csch)f).Arguments
        ),
        Elementwise<Asinh>(
            "ASINH",
            1,
            1,
            static arguments => new Asinh(arguments),
            static f => ((Asinh)f).Arguments
        ),
        Elementwise<Acosh>(
            "ACOSH",
            1,
            1,
            static arguments => new Acosh(arguments),
            static f => ((Acosh)f).Arguments
        ),
        Elementwise<Atanh>(
            "ATANH",
            1,
            1,
            static arguments => new Atanh(arguments),
            static f => ((Atanh)f).Arguments
        ),
        Elementwise<Acoth>(
            "ACOTH",
            1,
            1,
            static arguments => new Acoth(arguments),
            static f => ((Acoth)f).Arguments
        ),
        Elementwise<Degrees>(
            "DEGREES",
            1,
            1,
            static arguments => new Degrees(arguments),
            static f => ((Degrees)f).Arguments
        ),
        Elementwise<Radians>(
            "RADIANS",
            1,
            1,
            static arguments => new Radians(arguments),
            static f => ((Radians)f).Arguments
        ),
        Elementwise<Base>(
            "BASE",
            2,
            3,
            static arguments => new Base(arguments),
            static f => ((Base)f).Arguments
        ),
        Elementwise<DecimalNumber>(
            "DECIMAL",
            2,
            2,
            static arguments => new DecimalNumber(arguments),
            static f => ((DecimalNumber)f).Arguments
        ),
        Elementwise<Roman>(
            "ROMAN",
            1,
            2,
            static arguments => new Roman(arguments),
            static f => ((Roman)f).Arguments
        ),
        Elementwise<Arabic>(
            "ARABIC",
            1,
            1,
            static arguments => new Arabic(arguments),
            static f => ((Arabic)f).Arguments
        ),
        Entry<TrueFunction>(
            "TRUE",
            0,
            0,
            static arguments => new TrueFunction(arguments),
            static f => ((TrueFunction)f).Arguments
        ),
        Entry<FalseFunction>(
            "FALSE",
            0,
            0,
            static arguments => new FalseFunction(arguments),
            static f => ((FalseFunction)f).Arguments
        ),
        Entry<Xor>(
            "XOR",
            1,
            int.MaxValue,
            static arguments => new Xor(arguments),
            static f => ((Xor)f).Arguments
        ),
        Elementwise<Ifs>(
            "IFS",
            2,
            int.MaxValue,
            static arguments => new Ifs(arguments),
            static f => ((Ifs)f).Arguments
        ),
        Elementwise<Switch>(
            "SWITCH",
            3,
            int.MaxValue,
            static arguments => new Switch(arguments),
            static f => ((Switch)f).Arguments
        ),
        Entry<Na>("NA", 0, 0, static arguments => new Na(arguments), static f => ((Na)f).Arguments),
        Elementwise<IsError>(
            "ISERROR",
            1,
            1,
            static arguments => new IsError(arguments),
            static f => ((IsError)f).Arguments
        ),
        Elementwise<IsErr>(
            "ISERR",
            1,
            1,
            static arguments => new IsErr(arguments),
            static f => ((IsErr)f).Arguments
        ),
        Elementwise<IsNa>(
            "ISNA",
            1,
            1,
            static arguments => new IsNa(arguments),
            static f => ((IsNa)f).Arguments
        ),
        Elementwise<IsText>(
            "ISTEXT",
            1,
            1,
            static arguments => new IsText(arguments),
            static f => ((IsText)f).Arguments
        ),
        Elementwise<IsNonText>(
            "ISNONTEXT",
            1,
            1,
            static arguments => new IsNonText(arguments),
            static f => ((IsNonText)f).Arguments
        ),
        Elementwise<IsLogical>(
            "ISLOGICAL",
            1,
            1,
            static arguments => new IsLogical(arguments),
            static f => ((IsLogical)f).Arguments
        ),
        Elementwise<IsEven>(
            "ISEVEN",
            1,
            1,
            static arguments => new IsEven(arguments),
            static f => ((IsEven)f).Arguments
        ),
        Elementwise<IsOdd>(
            "ISODD",
            1,
            1,
            static arguments => new IsOdd(arguments),
            static f => ((IsOdd)f).Arguments
        ),
        Entry<IsRef>(
            "ISREF",
            1,
            1,
            static arguments => new IsRef(arguments),
            static f => ((IsRef)f).Arguments
        ),
        Entry<IsFormula>(
            "ISFORMULA",
            1,
            1,
            static arguments => new IsFormula(arguments),
            static f => ((IsFormula)f).Arguments
        ),
        Elementwise<N>(
            "N",
            1,
            1,
            static arguments => new N(arguments),
            static f => ((N)f).Arguments
        ),
        Elementwise<T>(
            "T",
            1,
            1,
            static arguments => new T(arguments),
            static f => ((T)f).Arguments
        ),
        Entry<TypeFunction>(
            "TYPE",
            1,
            1,
            static arguments => new TypeFunction(arguments),
            static f => ((TypeFunction)f).Arguments
        ),
        Elementwise<ErrorType>(
            "ERROR.TYPE",
            1,
            1,
            static arguments => new ErrorType(arguments),
            static f => ((ErrorType)f).Arguments
        ),
        Entry<SheetsCount>(
            "SHEETS",
            0,
            0,
            static arguments => new SheetsCount(arguments),
            static f => ((SheetsCount)f).Arguments
        ),
        Elementwise<Right>(
            "RIGHT",
            1,
            2,
            static arguments => new Right(arguments),
            static f => ((Right)f).Arguments
        ),
        Elementwise<Find>(
            "FIND",
            2,
            3,
            static arguments => new Find(arguments),
            static f => ((Find)f).Arguments
        ),
        Elementwise<Search>(
            "SEARCH",
            2,
            3,
            static arguments => new Search(arguments),
            static f => ((Search)f).Arguments
        ),
        Elementwise<Replace>(
            "REPLACE",
            4,
            4,
            static arguments => new Replace(arguments),
            static f => ((Replace)f).Arguments
        ),
        Elementwise<Substitute>(
            "SUBSTITUTE",
            3,
            4,
            static arguments => new Substitute(arguments),
            static f => ((Substitute)f).Arguments
        ),
        Elementwise<Rept>(
            "REPT",
            2,
            2,
            static arguments => new Rept(arguments),
            static f => ((Rept)f).Arguments
        ),
        Elementwise<Proper>(
            "PROPER",
            1,
            1,
            static arguments => new Proper(arguments),
            static f => ((Proper)f).Arguments
        ),
        Elementwise<Exact>(
            "EXACT",
            2,
            2,
            static arguments => new Exact(arguments),
            static f => ((Exact)f).Arguments
        ),
        Elementwise<CharFunction>(
            "CHAR",
            1,
            1,
            static arguments => new CharFunction(arguments),
            static f => ((CharFunction)f).Arguments
        ),
        Elementwise<Code>(
            "CODE",
            1,
            1,
            static arguments => new Code(arguments),
            static f => ((Code)f).Arguments
        ),
        Elementwise<UniChar>(
            "UNICHAR",
            1,
            1,
            static arguments => new UniChar(arguments),
            static f => ((UniChar)f).Arguments
        ),
        Elementwise<Unicode>(
            "UNICODE",
            1,
            1,
            static arguments => new Unicode(arguments),
            static f => ((Unicode)f).Arguments
        ),
        Elementwise<Clean>(
            "CLEAN",
            1,
            1,
            static arguments => new Clean(arguments),
            static f => ((Clean)f).Arguments
        ),
        Elementwise<Fixed>(
            "FIXED",
            1,
            3,
            static arguments => new Fixed(arguments),
            static f => ((Fixed)f).Arguments
        ),
        Elementwise<Dollar>(
            "DOLLAR",
            1,
            2,
            static arguments => new Dollar(arguments),
            static f => ((Dollar)f).Arguments
        ),
        Elementwise<NumberValueFunction>(
            "NUMBERVALUE",
            1,
            3,
            static arguments => new NumberValueFunction(arguments),
            static f => ((NumberValueFunction)f).Arguments
        ),
        Elementwise<TextBefore>(
            "TEXTBEFORE",
            2,
            6,
            static arguments => new TextBefore(arguments),
            static f => ((TextBefore)f).Arguments
        ),
        Elementwise<TextAfter>(
            "TEXTAFTER",
            2,
            6,
            static arguments => new TextAfter(arguments),
            static f => ((TextAfter)f).Arguments
        ),
        Elementwise<ValueToText>(
            "VALUETOTEXT",
            1,
            2,
            static arguments => new ValueToText(arguments),
            static f => ((ValueToText)f).Arguments
        ),
        Elementwise<RegexTest>(
            "REGEXTEST",
            2,
            3,
            static arguments => new RegexTest(arguments),
            static f => ((RegexTest)f).Arguments
        ),
        Elementwise<RegexExtract>(
            "REGEXEXTRACT",
            2,
            4,
            static arguments => new RegexExtract(arguments),
            static f => ((RegexExtract)f).Arguments
        ),
        Elementwise<RegexReplace>(
            "REGEXREPLACE",
            3,
            5,
            static arguments => new RegexReplace(arguments),
            static f => ((RegexReplace)f).Arguments
        ),
        Entry<Choose>(
            "CHOOSE",
            2,
            int.MaxValue,
            static arguments => new Choose(arguments),
            static f => ((Choose)f).Arguments
        ),
        Entry<HLookup>(
            "HLOOKUP",
            3,
            4,
            static arguments => new HLookup(arguments),
            static f => ((HLookup)f).Arguments
        ),
        Entry<Lookup>(
            "LOOKUP",
            2,
            3,
            static arguments => new Lookup(arguments),
            static f => ((Lookup)f).Arguments
        ),
        Entry<Column>(
            "COLUMN",
            0,
            1,
            static arguments => new Column(arguments),
            static f => ((Column)f).Arguments
        ),
        Entry<Columns>(
            "COLUMNS",
            1,
            1,
            static arguments => new Columns(arguments),
            static f => ((Columns)f).Arguments
        ),
        Entry<XMatch>(
            "XMATCH",
            2,
            4,
            static arguments => new XMatch(arguments),
            static f => ((XMatch)f).Arguments
        ),
        // Phase 7 (dynamic arrays): the three axis-selection producers. Entry<T>, never Elementwise<T> — each
        // reads its whole source and answers a computed ARRAY through IArrayProducer, so a lift would hand it
        // one element and shadow the producer arm in ArrayEvaluation (the Consumes half is what TryGetLift
        // refuses). Arity per the Microsoft pages: FILTER(array, include, [if_empty]), SORT(array,
        // [sort_index], [sort_order], [by_col]), UNIQUE(array, [by_col], [exactly_once]).
        Entry<Filter>(
            "FILTER",
            2,
            3,
            static arguments => new Filter(arguments),
            static f => ((Filter)f).Arguments
        ),
        Entry<Sort>(
            "SORT",
            1,
            4,
            static arguments => new Sort(arguments),
            static f => ((Sort)f).Arguments
        ),
        Entry<Unique>(
            "UNIQUE",
            1,
            3,
            static arguments => new Unique(arguments),
            static f => ((Unique)f).Arguments
        ),
        Elementwise<Address>(
            "ADDRESS",
            2,
            5,
            static arguments => new Address(arguments),
            static f => ((Address)f).Arguments
        ),
        Entry<Areas>(
            "AREAS",
            1,
            1,
            static arguments => new Areas(arguments),
            static f => ((Areas)f).Arguments
        ),
        Entry<FormulaText>(
            "FORMULATEXT",
            1,
            1,
            static arguments => new FormulaText(arguments),
            static f => ((FormulaText)f).Arguments
        ),
        Entry<AverageIf>(
            "AVERAGEIF",
            2,
            3,
            static arguments => new AverageIf(arguments),
            static f => ((AverageIf)f).Arguments
        ),
        Entry<AverageIfs>(
            "AVERAGEIFS",
            3,
            int.MaxValue,
            static arguments => new AverageIfs(arguments),
            static f => ((AverageIfs)f).Arguments
        ),
        Entry<MaxIfs>(
            "MAXIFS",
            3,
            int.MaxValue,
            static arguments => new MaxIfs(arguments),
            static f => ((MaxIfs)f).Arguments
        ),
        Entry<MinIfs>(
            "MINIFS",
            3,
            int.MaxValue,
            static arguments => new MinIfs(arguments),
            static f => ((MinIfs)f).Arguments
        ),
        Entry<AverageA>(
            "AVERAGEA",
            1,
            int.MaxValue,
            static arguments => new AverageA(arguments),
            static f => ((AverageA)f).Arguments
        ),
        Entry<MaxA>(
            "MAXA",
            1,
            int.MaxValue,
            static arguments => new MaxA(arguments),
            static f => ((MaxA)f).Arguments
        ),
        Entry<MinA>(
            "MINA",
            1,
            int.MaxValue,
            static arguments => new MinA(arguments),
            static f => ((MinA)f).Arguments
        ),
        Entry<SumProduct>(
            "SUMPRODUCT",
            1,
            int.MaxValue,
            static arguments => new SumProduct(arguments),
            static f => ((SumProduct)f).Arguments
        ),
        Entry<SumX2MY2>(
            "SUMX2MY2",
            2,
            2,
            static arguments => new SumX2MY2(arguments),
            static f => ((SumX2MY2)f).Arguments
        ),
        Entry<SumX2PY2>(
            "SUMX2PY2",
            2,
            2,
            static arguments => new SumX2PY2(arguments),
            static f => ((SumX2PY2)f).Arguments
        ),
        Entry<SumXMY2>(
            "SUMXMY2",
            2,
            2,
            static arguments => new SumXMY2(arguments),
            static f => ((SumXMY2)f).Arguments
        ),
        Entry<Subtotal>(
            "SUBTOTAL",
            2,
            int.MaxValue,
            static arguments => new Subtotal(arguments),
            static f => ((Subtotal)f).Arguments
        ),
        // MinArgs = 3 makes AGGREGATE(15,6) a PARSE-time error, as Excel rejects it at entry; MaxArgs is
        // unbounded because the reference form takes ref1, [ref2], … — the array form's "exactly 4" rule
        // depends on function_num and therefore belongs to Evaluate, which is also where Excel's own
        // #VALUE! for a missing k lives.
        Entry<Aggregate>(
            "AGGREGATE",
            3,
            int.MaxValue,
            static arguments => new Aggregate(arguments),
            static f => ((Aggregate)f).Arguments
        ),
        // Phase 7 (dynamic arrays): SEQUENCE(rows, [columns], [start], [step]) — a producer with no range
        // argument at all, so the three-rectangle sweep in ElementwiseLiftingTests is blind to it and it is
        // named there by hand. Entry<T> for the same reason as FILTER/SORT/UNIQUE above.
        Entry<Sequence>(
            "SEQUENCE",
            1,
            4,
            static arguments => new Sequence(arguments),
            static f => ((Sequence)f).Arguments
        ),
        Entry<Median>(
            "MEDIAN",
            1,
            int.MaxValue,
            static arguments => new Median(arguments),
            static f => ((Median)f).Arguments
        ),
        Entry<ModeSngl>(
            "MODE.SNGL",
            1,
            int.MaxValue,
            static arguments => new ModeSngl(arguments),
            static f => ((ModeSngl)f).Arguments
        ),
        Entry<Large>(
            "LARGE",
            2,
            2,
            static arguments => new Large(arguments),
            static f => ((Large)f).Arguments
        ),
        Entry<Small>(
            "SMALL",
            2,
            2,
            static arguments => new Small(arguments),
            static f => ((Small)f).Arguments
        ),
        Entry<RankEq>(
            "RANK.EQ",
            2,
            3,
            static arguments => new RankEq(arguments),
            static f => ((RankEq)f).Arguments
        ),
        Entry<RankAvg>(
            "RANK.AVG",
            2,
            3,
            static arguments => new RankAvg(arguments),
            static f => ((RankAvg)f).Arguments
        ),
        Entry<PercentileInc>(
            "PERCENTILE.INC",
            2,
            2,
            static arguments => new PercentileInc(arguments),
            static f => ((PercentileInc)f).Arguments
        ),
        Entry<PercentileExc>(
            "PERCENTILE.EXC",
            2,
            2,
            static arguments => new PercentileExc(arguments),
            static f => ((PercentileExc)f).Arguments
        ),
        Entry<PercentRankInc>(
            "PERCENTRANK.INC",
            2,
            3,
            static arguments => new PercentRankInc(arguments),
            static f => ((PercentRankInc)f).Arguments
        ),
        Entry<PercentRankExc>(
            "PERCENTRANK.EXC",
            2,
            3,
            static arguments => new PercentRankExc(arguments),
            static f => ((PercentRankExc)f).Arguments
        ),
        Entry<QuartileInc>(
            "QUARTILE.INC",
            2,
            2,
            static arguments => new QuartileInc(arguments),
            static f => ((QuartileInc)f).Arguments
        ),
        Entry<QuartileExc>(
            "QUARTILE.EXC",
            2,
            2,
            static arguments => new QuartileExc(arguments),
            static f => ((QuartileExc)f).Arguments
        ),
        Entry<TrimMean>(
            "TRIMMEAN",
            2,
            2,
            static arguments => new TrimMean(arguments),
            static f => ((TrimMean)f).Arguments
        ),
        Entry<StDevS>(
            "STDEV.S",
            1,
            int.MaxValue,
            static arguments => new StDevS(arguments),
            static f => ((StDevS)f).Arguments
        ),
        Entry<StDevP>(
            "STDEV.P",
            1,
            int.MaxValue,
            static arguments => new StDevP(arguments),
            static f => ((StDevP)f).Arguments
        ),
        Entry<StDevA>(
            "STDEVA",
            1,
            int.MaxValue,
            static arguments => new StDevA(arguments),
            static f => ((StDevA)f).Arguments
        ),
        Entry<StDevPA>(
            "STDEVPA",
            1,
            int.MaxValue,
            static arguments => new StDevPA(arguments),
            static f => ((StDevPA)f).Arguments
        ),
        Entry<VarS>(
            "VAR.S",
            1,
            int.MaxValue,
            static arguments => new VarS(arguments),
            static f => ((VarS)f).Arguments
        ),
        Entry<VarP>(
            "VAR.P",
            1,
            int.MaxValue,
            static arguments => new VarP(arguments),
            static f => ((VarP)f).Arguments
        ),
        Entry<VarA>(
            "VARA",
            1,
            int.MaxValue,
            static arguments => new VarA(arguments),
            static f => ((VarA)f).Arguments
        ),
        Entry<VarPA>(
            "VARPA",
            1,
            int.MaxValue,
            static arguments => new VarPA(arguments),
            static f => ((VarPA)f).Arguments
        ),
        Entry<AveDev>(
            "AVEDEV",
            1,
            int.MaxValue,
            static arguments => new AveDev(arguments),
            static f => ((AveDev)f).Arguments
        ),
        Entry<DevSq>(
            "DEVSQ",
            1,
            int.MaxValue,
            static arguments => new DevSq(arguments),
            static f => ((DevSq)f).Arguments
        ),
        Entry<GeoMean>(
            "GEOMEAN",
            1,
            int.MaxValue,
            static arguments => new GeoMean(arguments),
            static f => ((GeoMean)f).Arguments
        ),
        Entry<HarMean>(
            "HARMEAN",
            1,
            int.MaxValue,
            static arguments => new HarMean(arguments),
            static f => ((HarMean)f).Arguments
        ),
        Entry<Skew>(
            "SKEW",
            1,
            int.MaxValue,
            static arguments => new Skew(arguments),
            static f => ((Skew)f).Arguments
        ),
        Entry<SkewP>(
            "SKEW.P",
            1,
            int.MaxValue,
            static arguments => new SkewP(arguments),
            static f => ((SkewP)f).Arguments
        ),
        Entry<Kurt>(
            "KURT",
            1,
            int.MaxValue,
            static arguments => new Kurt(arguments),
            static f => ((Kurt)f).Arguments
        ),
        Elementwise<Standardize>(
            "STANDARDIZE",
            3,
            3,
            static arguments => new Standardize(arguments),
            static f => ((Standardize)f).Arguments
        ),
        Entry<Correl>(
            "CORREL",
            2,
            2,
            static arguments => new Correl(arguments),
            static f => ((Correl)f).Arguments
        ),
        Entry<Pearson>(
            "PEARSON",
            2,
            2,
            static arguments => new Pearson(arguments),
            static f => ((Pearson)f).Arguments
        ),
        Entry<CovarianceP>(
            "COVARIANCE.P",
            2,
            2,
            static arguments => new CovarianceP(arguments),
            static f => ((CovarianceP)f).Arguments
        ),
        Entry<CovarianceS>(
            "COVARIANCE.S",
            2,
            2,
            static arguments => new CovarianceS(arguments),
            static f => ((CovarianceS)f).Arguments
        ),
        Entry<Rsq>(
            "RSQ",
            2,
            2,
            static arguments => new Rsq(arguments),
            static f => ((Rsq)f).Arguments
        ),
        Entry<Slope>(
            "SLOPE",
            2,
            2,
            static arguments => new Slope(arguments),
            static f => ((Slope)f).Arguments
        ),
        Entry<Intercept>(
            "INTERCEPT",
            2,
            2,
            static arguments => new Intercept(arguments),
            static f => ((Intercept)f).Arguments
        ),
        Entry<Steyx>(
            "STEYX",
            2,
            2,
            static arguments => new Steyx(arguments),
            static f => ((Steyx)f).Arguments
        ),
        Entry<ForecastLinear>(
            "FORECAST.LINEAR",
            3,
            3,
            static arguments => new ForecastLinear(arguments),
            static f => ((ForecastLinear)f).Arguments
        ),
        Elementwise<Fisher>(
            "FISHER",
            1,
            1,
            static arguments => new Fisher(arguments),
            static f => ((Fisher)f).Arguments
        ),
        Elementwise<FisherInv>(
            "FISHERINV",
            1,
            1,
            static arguments => new FisherInv(arguments),
            static f => ((FisherInv)f).Arguments
        ),
        Elementwise<Phi>(
            "PHI",
            1,
            1,
            static arguments => new Phi(arguments),
            static f => ((Phi)f).Arguments
        ),
        Elementwise<Permut>(
            "PERMUT",
            2,
            2,
            static arguments => new Permut(arguments),
            static f => ((Permut)f).Arguments
        ),
        Elementwise<PermutationA>(
            "PERMUTATIONA",
            2,
            2,
            static arguments => new PermutationA(arguments),
            static f => ((PermutationA)f).Arguments
        ),
        Entry<Prob>(
            "PROB",
            3,
            4,
            static arguments => new Prob(arguments),
            static f => ((Prob)f).Arguments
        ),
        // Compatibility aliases: distinct nodes so the un-parse keeps the legacy spelling.
        Entry<Compat.Mode>(
            "MODE",
            1,
            int.MaxValue,
            static arguments => new Compat.Mode(arguments),
            static f => ((Compat.Mode)f).Arguments
        ),
        Entry<Compat.StDev>(
            "STDEV",
            1,
            int.MaxValue,
            static arguments => new Compat.StDev(arguments),
            static f => ((Compat.StDev)f).Arguments
        ),
        Entry<Compat.StDevP>(
            "STDEVP",
            1,
            int.MaxValue,
            static arguments => new Compat.StDevP(arguments),
            static f => ((Compat.StDevP)f).Arguments
        ),
        Entry<Compat.Var>(
            "VAR",
            1,
            int.MaxValue,
            static arguments => new Compat.Var(arguments),
            static f => ((Compat.Var)f).Arguments
        ),
        Entry<Compat.VarP>(
            "VARP",
            1,
            int.MaxValue,
            static arguments => new Compat.VarP(arguments),
            static f => ((Compat.VarP)f).Arguments
        ),
        Entry<Compat.Rank>(
            "RANK",
            2,
            3,
            static arguments => new Compat.Rank(arguments),
            static f => ((Compat.Rank)f).Arguments
        ),
        Entry<Compat.Percentile>(
            "PERCENTILE",
            2,
            2,
            static arguments => new Compat.Percentile(arguments),
            static f => ((Compat.Percentile)f).Arguments
        ),
        Entry<Compat.PercentRank>(
            "PERCENTRANK",
            2,
            3,
            static arguments => new Compat.PercentRank(arguments),
            static f => ((Compat.PercentRank)f).Arguments
        ),
        Entry<Compat.Quartile>(
            "QUARTILE",
            2,
            2,
            static arguments => new Compat.Quartile(arguments),
            static f => ((Compat.Quartile)f).Arguments
        ),
        Entry<Compat.Covar>(
            "COVAR",
            2,
            2,
            static arguments => new Compat.Covar(arguments),
            static f => ((Compat.Covar)f).Arguments
        ),
        Entry<Compat.Forecast>(
            "FORECAST",
            3,
            3,
            static arguments => new Compat.Forecast(arguments),
            static f => ((Compat.Forecast)f).Arguments
        ),
        // Wave 5 — Date and time (TODAY/NOW deferred to F1: volatile).
        Elementwise<Date>(
            "DATE",
            3,
            3,
            static arguments => new Date(arguments),
            static f => ((Date)f).Arguments
        ),
        Elementwise<Time>(
            "TIME",
            3,
            3,
            static arguments => new Time(arguments),
            static f => ((Time)f).Arguments
        ),
        Elementwise<DateValue>(
            "DATEVALUE",
            1,
            1,
            static arguments => new DateValue(arguments),
            static f => ((DateValue)f).Arguments
        ),
        Elementwise<TimeValue>(
            "TIMEVALUE",
            1,
            1,
            static arguments => new TimeValue(arguments),
            static f => ((TimeValue)f).Arguments
        ),
        Elementwise<Year>(
            "YEAR",
            1,
            1,
            static arguments => new Year(arguments),
            static f => ((Year)f).Arguments
        ),
        Elementwise<Month>(
            "MONTH",
            1,
            1,
            static arguments => new Month(arguments),
            static f => ((Month)f).Arguments
        ),
        Elementwise<Day>(
            "DAY",
            1,
            1,
            static arguments => new Day(arguments),
            static f => ((Day)f).Arguments
        ),
        Elementwise<Hour>(
            "HOUR",
            1,
            1,
            static arguments => new Hour(arguments),
            static f => ((Hour)f).Arguments
        ),
        Elementwise<Minute>(
            "MINUTE",
            1,
            1,
            static arguments => new Minute(arguments),
            static f => ((Minute)f).Arguments
        ),
        Elementwise<Second>(
            "SECOND",
            1,
            1,
            static arguments => new Second(arguments),
            static f => ((Second)f).Arguments
        ),
        Elementwise<Days>(
            "DAYS",
            2,
            2,
            static arguments => new Days(arguments),
            static f => ((Days)f).Arguments
        ),
        Elementwise<Days360>(
            "DAYS360",
            2,
            3,
            static arguments => new Days360(arguments),
            static f => ((Days360)f).Arguments
        ),
        Elementwise<EDate>(
            "EDATE",
            2,
            2,
            static arguments => new EDate(arguments),
            static f => ((EDate)f).Arguments
        ),
        Elementwise<EoMonth>(
            "EOMONTH",
            2,
            2,
            static arguments => new EoMonth(arguments),
            static f => ((EoMonth)f).Arguments
        ),
        Elementwise<Weekday>(
            "WEEKDAY",
            1,
            2,
            static arguments => new Weekday(arguments),
            static f => ((Weekday)f).Arguments
        ),
        Elementwise<WeekNum>(
            "WEEKNUM",
            1,
            2,
            static arguments => new WeekNum(arguments),
            static f => ((WeekNum)f).Arguments
        ),
        Elementwise<IsoWeekNum>(
            "ISOWEEKNUM",
            1,
            1,
            static arguments => new IsoWeekNum(arguments),
            static f => ((IsoWeekNum)f).Arguments
        ),
        Elementwise<DateDif>(
            "DATEDIF",
            3,
            3,
            static arguments => new DateDif(arguments),
            static f => ((DateDif)f).Arguments
        ),
        Elementwise<YearFrac>(
            "YEARFRAC",
            2,
            3,
            static arguments => new YearFrac(arguments),
            static f => ((YearFrac)f).Arguments
        ),
        Entry<NetworkDays>(
            "NETWORKDAYS",
            2,
            3,
            static arguments => new NetworkDays(arguments),
            static f => ((NetworkDays)f).Arguments
        ),
        Entry<NetworkDaysIntl>(
            "NETWORKDAYS.INTL",
            2,
            4,
            static arguments => new NetworkDaysIntl(arguments),
            static f => ((NetworkDaysIntl)f).Arguments
        ),
        Entry<Workday>(
            "WORKDAY",
            2,
            3,
            static arguments => new Workday(arguments),
            static f => ((Workday)f).Arguments
        ),
        Entry<WorkdayIntl>(
            "WORKDAY.INTL",
            2,
            4,
            static arguments => new WorkdayIntl(arguments),
            static f => ((WorkdayIntl)f).Arguments
        ),
        // F1 — volatile clock functions (both take no arguments).
        Entry<Now>(
            "NOW",
            0,
            0,
            static arguments => new Now(arguments),
            static f => ((Now)f).Arguments
        ),
        Entry<Today>(
            "TODAY",
            0,
            0,
            static arguments => new Today(arguments),
            static f => ((Today)f).Arguments
        ),
        // F1 — volatile RNG functions.
        Entry<Rand>(
            "RAND",
            0,
            0,
            static arguments => new Rand(arguments),
            static f => ((Rand)f).Arguments
        ),
        Entry<RandBetween>(
            "RANDBETWEEN",
            2,
            2,
            static arguments => new RandBetween(arguments),
            static f => ((RandBetween)f).Arguments
        ),
    ];

    // Frozen for read-optimized lookups: built once, read on every function-call parse (Parser) and every
    // node-to-formula-text render (FormulaWriter's hot path — Formulas-mode export, FORMULATEXT).
    internal static readonly FrozenDictionary<string, RegistryEntry> ByName =
        Entries.ToFrozenDictionary(static e => e.Name, StringComparer.OrdinalIgnoreCase);

    internal static readonly FrozenDictionary<Type, RegistryEntry> ByType =
        Entries.ToFrozenDictionary(static e => e.NodeType);

    // Generic type inference lets each entry below read as "function name, arity, factory, arguments
    // accessor" without repeating the node type a third time.
    private static RegistryEntry Entry<T>(
        string name,
        int minArgs,
        int maxArgs,
        Func<Expression[], Expression> create,
        Func<Function, Expression[]> getArguments
    )
        where T : Function =>
        new(name, minArgs, maxArgs, create, typeof(T), getArguments, ArrayLifting.Consumes);

    // The sibling of Entry<T> for a pure-scalar function: same five arguments, so a liftable entry above
    // still reads as one line, and the classification is greppable — one factory call per lifted entry.
    private static RegistryEntry Elementwise<T>(
        string name,
        int minArgs,
        int maxArgs,
        Func<Expression[], Expression> create,
        Func<Function, Expression[]> getArguments
    )
        where T : Function =>
        new(name, minArgs, maxArgs, create, typeof(T), getArguments, ArrayLifting.Elementwise);
}
