using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Dates;
using Danfma.MySheet.Expressions.Financial;
using Compat = Danfma.MySheet.Expressions.Compatibility;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Expressions.Statistical;
using Danfma.MySheet.Expressions.Text;

namespace Danfma.MySheet.Parsing;

/// <summary>
/// A Pratt (top-down operator precedence) parser turning a token stream into an
/// <see cref="Expression"/> tree. Cell references are resolved against <c>sheetName</c>.
/// </summary>
internal sealed class Parser(List<Token> tokens, string sheetName)
{
    // Binding powers (higher binds tighter). Unary prefix binds tighter than '^' so that
    // '-2^2' parses as '(-2)^2' == 4, matching Excel.
    private const int ComparisonBindingPower = 10;
    private const int ConcatBindingPower = 15; // '&' binds below + - and above the comparators
    private const int AdditiveBindingPower = 20;
    private const int MultiplicativeBindingPower = 30;
    private const int PowerBindingPower = 40;
    private const int PercentBindingPower = 44; // postfix '%' binds above '^', below unary minus
    private const int PrefixBindingPower = 45;
    private const int RangeBindingPower = 50;

    private readonly record struct FunctionSpec(
        int MinArgs,
        int MaxArgs,
        Func<Expression[], Expression> Create
    );

    private static readonly Dictionary<string, FunctionSpec> Functions = new(
        StringComparer.OrdinalIgnoreCase
    )
    {
        ["SUM"] = new(0, int.MaxValue, arguments => new Sum(arguments)),
        ["AVERAGE"] = new(0, int.MaxValue, arguments => new Average(arguments)),
        ["MIN"] = new(0, int.MaxValue, arguments => new Min(arguments)),
        ["MAX"] = new(0, int.MaxValue, arguments => new Max(arguments)),
        ["COUNT"] = new(0, int.MaxValue, arguments => new Count(arguments)),
        ["IF"] = new(2, 3, arguments => new If(arguments)),
        ["AND"] = new(1, int.MaxValue, arguments => new And(arguments)),
        ["OR"] = new(1, int.MaxValue, arguments => new Or(arguments)),
        ["NOT"] = new(1, 1, arguments => new Not(arguments)),
        ["IFERROR"] = new(2, 2, arguments => new IfError(arguments)),
        ["INT"] = new(1, 1, arguments => new Int(arguments)),
        ["ROUND"] = new(2, 2, arguments => new Round(arguments)),
        ["ROUNDUP"] = new(2, 2, arguments => new RoundUp(arguments)),
        ["ABS"] = new(1, 1, arguments => new Abs(arguments)),
        ["ISNUMBER"] = new(1, 1, arguments => new IsNumber(arguments)),
        ["ISBLANK"] = new(1, 1, arguments => new IsBlank(arguments)),
        ["IFNA"] = new(2, 2, arguments => new IfNa(arguments)),
        ["UPPER"] = new(1, 1, arguments => new Upper(arguments)),
        ["LOWER"] = new(1, 1, arguments => new Lower(arguments)),
        ["TRIM"] = new(1, 1, arguments => new Trim(arguments)),
        ["LEN"] = new(1, 1, arguments => new Len(arguments)),
        ["LEFT"] = new(1, 2, arguments => new Left(arguments)),
        ["MID"] = new(3, 3, arguments => new Mid(arguments)),
        ["VALUE"] = new(1, 1, arguments => new Value(arguments)),
        ["CONCAT"] = new(1, int.MaxValue, arguments => new Concat(arguments)),
        ["CONCATENATE"] = new(1, int.MaxValue, arguments => new Concatenate(arguments)),
        ["TEXTJOIN"] = new(3, int.MaxValue, arguments => new TextJoin(arguments)),
        ["COUNTA"] = new(1, int.MaxValue, arguments => new CountA(arguments)),
        ["COUNTBLANK"] = new(1, int.MaxValue, arguments => new CountBlank(arguments)),
        ["COUNTIF"] = new(2, 2, arguments => new CountIf(arguments)),
        ["COUNTIFS"] = new(2, int.MaxValue, arguments => new CountIfs(arguments)),
        ["SUMIF"] = new(2, 3, arguments => new SumIf(arguments)),
        ["SUMIFS"] = new(3, int.MaxValue, arguments => new SumIfs(arguments)),
        ["ROWS"] = new(1, 1, arguments => new Rows(arguments)),
        ["ROW"] = new(0, 1, arguments => new Row(arguments)),
        ["MATCH"] = new(2, 3, arguments => new Match(arguments)),
        ["INDEX"] = new(2, 3, arguments => new MySheet.Expressions.Lookup.Index(arguments)),
        ["VLOOKUP"] = new(3, 4, arguments => new VLookup(arguments)),
        ["XLOOKUP"] = new(3, 6, arguments => new XLookup(arguments)),
        ["OFFSET"] = new(3, 5, arguments => new Offset(arguments)),
        ["LET"] = new(3, int.MaxValue, arguments => new Let(arguments)),
        ["TEXT"] = new(2, 2, arguments => new Text(arguments)),
        ["SHEET"] = new(0, 1, arguments => new SheetNumber(arguments)),
        ["PMT"] = new(3, 5, arguments => new Pmt(arguments)),
        ["PV"] = new(3, 5, arguments => new Pv(arguments)),
        ["FV"] = new(3, 5, arguments => new Fv(arguments)),
        ["NPER"] = new(3, 5, arguments => new Nper(arguments)),
        ["IPMT"] = new(4, 6, arguments => new Ipmt(arguments)),
        ["PPMT"] = new(4, 6, arguments => new Ppmt(arguments)),
        ["NPV"] = new(2, int.MaxValue, arguments => new Npv(arguments)),
        ["RATE"] = new(3, 6, arguments => new Rate(arguments)),
        ["IRR"] = new(1, 2, arguments => new Irr(arguments)),
        ["SLN"] = new(3, 3, arguments => new Sln(arguments)),
        ["SYD"] = new(4, 4, arguments => new Syd(arguments)),
        ["DB"] = new(4, 5, arguments => new Db(arguments)),
        ["DDB"] = new(4, 5, arguments => new Ddb(arguments)),
        ["VDB"] = new(5, 7, arguments => new Vdb(arguments)),
        ["AMORLINC"] = new(6, 7, arguments => new AmorLinc(arguments)),
        ["AMORDEGRC"] = new(6, 7, arguments => new AmorDegrc(arguments)),
        ["EFFECT"] = new(2, 2, arguments => new Effect(arguments)),
        ["NOMINAL"] = new(2, 2, arguments => new Nominal(arguments)),
        ["MIRR"] = new(3, 3, arguments => new Mirr(arguments)),
        ["RRI"] = new(3, 3, arguments => new Rri(arguments)),
        ["PDURATION"] = new(3, 3, arguments => new PDuration(arguments)),
        ["ISPMT"] = new(4, 4, arguments => new ISPmt(arguments)),
        ["CUMIPMT"] = new(6, 6, arguments => new CumIPmt(arguments)),
        ["CUMPRINC"] = new(6, 6, arguments => new CumPrinc(arguments)),
        ["FVSCHEDULE"] = new(2, 2, arguments => new FvSchedule(arguments)),
        ["DOLLARDE"] = new(2, 2, arguments => new DollarDe(arguments)),
        ["DOLLARFR"] = new(2, 2, arguments => new DollarFr(arguments)),
        ["XNPV"] = new(3, 3, arguments => new XNpv(arguments)),
        ["XIRR"] = new(2, 3, arguments => new XIrr(arguments)),
        ["ACCRINT"] = new(6, 8, arguments => new AccrInt(arguments)),
        ["ACCRINTM"] = new(4, 5, arguments => new AccrIntM(arguments)),
        ["DISC"] = new(4, 5, arguments => new Disc(arguments)),
        ["INTRATE"] = new(4, 5, arguments => new IntRate(arguments)),
        ["RECEIVED"] = new(4, 5, arguments => new Received(arguments)),
        ["PRICEDISC"] = new(4, 5, arguments => new PriceDisc(arguments)),
        ["PRICEMAT"] = new(5, 6, arguments => new PriceMat(arguments)),
        ["YIELDDISC"] = new(4, 5, arguments => new YieldDisc(arguments)),
        ["YIELDMAT"] = new(5, 6, arguments => new YieldMat(arguments)),
        ["TBILLEQ"] = new(3, 3, arguments => new TBillEq(arguments)),
        ["TBILLPRICE"] = new(3, 3, arguments => new TBillPrice(arguments)),
        ["TBILLYIELD"] = new(3, 3, arguments => new TBillYield(arguments)),
        ["COUPPCD"] = new(3, 4, arguments => new CoupPcd(arguments)),
        ["COUPNCD"] = new(3, 4, arguments => new CoupNcd(arguments)),
        ["COUPNUM"] = new(3, 4, arguments => new CoupNum(arguments)),
        ["COUPDAYS"] = new(3, 4, arguments => new CoupDays(arguments)),
        ["COUPDAYBS"] = new(3, 4, arguments => new CoupDayBs(arguments)),
        ["COUPDAYSNC"] = new(3, 4, arguments => new CoupDaysNc(arguments)),
        ["PRICE"] = new(6, 7, arguments => new Price(arguments)),
        ["YIELD"] = new(6, 7, arguments => new Yield(arguments)),
        ["DURATION"] = new(5, 6, arguments => new Duration(arguments)),
        ["MDURATION"] = new(5, 6, arguments => new MDuration(arguments)),
        ["ODDFPRICE"] = new(8, 9, arguments => new OddFPrice(arguments)),
        ["ODDFYIELD"] = new(8, 9, arguments => new OddFYield(arguments)),
        ["ODDLPRICE"] = new(7, 8, arguments => new OddLPrice(arguments)),
        ["ODDLYIELD"] = new(7, 8, arguments => new OddLYield(arguments)),
        ["SQRT"] = new(1, 1, arguments => new Sqrt(arguments)),
        ["POWER"] = new(2, 2, arguments => new Power(arguments)),
        ["EXP"] = new(1, 1, arguments => new Exp(arguments)),
        ["LN"] = new(1, 1, arguments => new Ln(arguments)),
        ["LOG"] = new(1, 2, arguments => new Log(arguments)),
        ["LOG10"] = new(1, 1, arguments => new Log10(arguments)),
        ["SQRTPI"] = new(1, 1, arguments => new SqrtPi(arguments)),
        ["ROUNDDOWN"] = new(2, 2, arguments => new RoundDown(arguments)),
        ["TRUNC"] = new(1, 2, arguments => new Trunc(arguments)),
        ["MROUND"] = new(2, 2, arguments => new MRound(arguments)),
        ["CEILING"] = new(2, 2, arguments => new Ceiling(arguments)),
        ["CEILING.MATH"] = new(1, 3, arguments => new CeilingMath(arguments)),
        ["CEILING.PRECISE"] = new(1, 2, arguments => new CeilingPrecise(arguments)),
        ["ISO.CEILING"] = new(1, 2, arguments => new IsoCeiling(arguments)),
        ["FLOOR"] = new(2, 2, arguments => new Floor(arguments)),
        ["FLOOR.MATH"] = new(1, 3, arguments => new FloorMath(arguments)),
        ["FLOOR.PRECISE"] = new(1, 2, arguments => new FloorPrecise(arguments)),
        ["EVEN"] = new(1, 1, arguments => new Even(arguments)),
        ["ODD"] = new(1, 1, arguments => new Odd(arguments)),
        ["MOD"] = new(2, 2, arguments => new Mod(arguments)),
        ["QUOTIENT"] = new(2, 2, arguments => new Quotient(arguments)),
        ["SIGN"] = new(1, 1, arguments => new Sign(arguments)),
        ["PI"] = new(0, 0, arguments => new Pi(arguments)),
        ["PRODUCT"] = new(1, int.MaxValue, arguments => new Product(arguments)),
        ["SUMSQ"] = new(1, int.MaxValue, arguments => new SumSq(arguments)),
        ["MULTINOMIAL"] = new(1, int.MaxValue, arguments => new Multinomial(arguments)),
        ["SERIESSUM"] = new(4, 4, arguments => new SeriesSum(arguments)),
        ["FACT"] = new(1, 1, arguments => new Fact(arguments)),
        ["FACTDOUBLE"] = new(1, 1, arguments => new FactDouble(arguments)),
        ["COMBIN"] = new(2, 2, arguments => new Combin(arguments)),
        ["COMBINA"] = new(2, 2, arguments => new CombinA(arguments)),
        ["GCD"] = new(1, int.MaxValue, arguments => new Gcd(arguments)),
        ["LCM"] = new(1, int.MaxValue, arguments => new Lcm(arguments)),
        ["SIN"] = new(1, 1, arguments => new Sin(arguments)),
        ["COS"] = new(1, 1, arguments => new Cos(arguments)),
        ["TAN"] = new(1, 1, arguments => new Tan(arguments)),
        ["COT"] = new(1, 1, arguments => new Cot(arguments)),
        ["SEC"] = new(1, 1, arguments => new Sec(arguments)),
        ["CSC"] = new(1, 1, arguments => new Csc(arguments)),
        ["ASIN"] = new(1, 1, arguments => new Asin(arguments)),
        ["ACOS"] = new(1, 1, arguments => new Acos(arguments)),
        ["ATAN"] = new(1, 1, arguments => new Atan(arguments)),
        ["ATAN2"] = new(2, 2, arguments => new Atan2(arguments)),
        ["ACOT"] = new(1, 1, arguments => new Acot(arguments)),
        ["SINH"] = new(1, 1, arguments => new Sinh(arguments)),
        ["COSH"] = new(1, 1, arguments => new Cosh(arguments)),
        ["TANH"] = new(1, 1, arguments => new Tanh(arguments)),
        ["COTH"] = new(1, 1, arguments => new Coth(arguments)),
        ["SECH"] = new(1, 1, arguments => new Sech(arguments)),
        ["CSCH"] = new(1, 1, arguments => new Csch(arguments)),
        ["ASINH"] = new(1, 1, arguments => new Asinh(arguments)),
        ["ACOSH"] = new(1, 1, arguments => new Acosh(arguments)),
        ["ATANH"] = new(1, 1, arguments => new Atanh(arguments)),
        ["ACOTH"] = new(1, 1, arguments => new Acoth(arguments)),
        ["DEGREES"] = new(1, 1, arguments => new Degrees(arguments)),
        ["RADIANS"] = new(1, 1, arguments => new Radians(arguments)),
        ["BASE"] = new(2, 3, arguments => new Base(arguments)),
        ["DECIMAL"] = new(2, 2, arguments => new DecimalNumber(arguments)),
        ["ROMAN"] = new(1, 2, arguments => new Roman(arguments)),
        ["ARABIC"] = new(1, 1, arguments => new Arabic(arguments)),
        ["TRUE"] = new(0, 0, arguments => new TrueFunction(arguments)),
        ["FALSE"] = new(0, 0, arguments => new FalseFunction(arguments)),
        ["XOR"] = new(1, int.MaxValue, arguments => new Xor(arguments)),
        ["IFS"] = new(2, int.MaxValue, arguments => new Ifs(arguments)),
        ["SWITCH"] = new(3, int.MaxValue, arguments => new Switch(arguments)),
        ["NA"] = new(0, 0, arguments => new Na(arguments)),
        ["ISERROR"] = new(1, 1, arguments => new IsError(arguments)),
        ["ISERR"] = new(1, 1, arguments => new IsErr(arguments)),
        ["ISNA"] = new(1, 1, arguments => new IsNa(arguments)),
        ["ISTEXT"] = new(1, 1, arguments => new IsText(arguments)),
        ["ISNONTEXT"] = new(1, 1, arguments => new IsNonText(arguments)),
        ["ISLOGICAL"] = new(1, 1, arguments => new IsLogical(arguments)),
        ["ISEVEN"] = new(1, 1, arguments => new IsEven(arguments)),
        ["ISODD"] = new(1, 1, arguments => new IsOdd(arguments)),
        ["ISREF"] = new(1, 1, arguments => new IsRef(arguments)),
        ["ISFORMULA"] = new(1, 1, arguments => new IsFormula(arguments)),
        ["N"] = new(1, 1, arguments => new N(arguments)),
        ["T"] = new(1, 1, arguments => new T(arguments)),
        ["TYPE"] = new(1, 1, arguments => new TypeFunction(arguments)),
        ["ERROR.TYPE"] = new(1, 1, arguments => new ErrorType(arguments)),
        ["SHEETS"] = new(0, 0, arguments => new SheetsCount(arguments)),
        ["RIGHT"] = new(1, 2, arguments => new Right(arguments)),
        ["FIND"] = new(2, 3, arguments => new Find(arguments)),
        ["SEARCH"] = new(2, 3, arguments => new Search(arguments)),
        ["REPLACE"] = new(4, 4, arguments => new Replace(arguments)),
        ["SUBSTITUTE"] = new(3, 4, arguments => new Substitute(arguments)),
        ["REPT"] = new(2, 2, arguments => new Rept(arguments)),
        ["PROPER"] = new(1, 1, arguments => new Proper(arguments)),
        ["EXACT"] = new(2, 2, arguments => new Exact(arguments)),
        ["CHAR"] = new(1, 1, arguments => new CharFunction(arguments)),
        ["CODE"] = new(1, 1, arguments => new Code(arguments)),
        ["UNICHAR"] = new(1, 1, arguments => new UniChar(arguments)),
        ["UNICODE"] = new(1, 1, arguments => new Unicode(arguments)),
        ["CLEAN"] = new(1, 1, arguments => new Clean(arguments)),
        ["FIXED"] = new(1, 3, arguments => new Fixed(arguments)),
        ["DOLLAR"] = new(1, 2, arguments => new Dollar(arguments)),
        ["NUMBERVALUE"] = new(1, 3, arguments => new NumberValueFunction(arguments)),
        ["TEXTBEFORE"] = new(2, 6, arguments => new TextBefore(arguments)),
        ["TEXTAFTER"] = new(2, 6, arguments => new TextAfter(arguments)),
        ["VALUETOTEXT"] = new(1, 2, arguments => new ValueToText(arguments)),
        ["REGEXTEST"] = new(2, 3, arguments => new RegexTest(arguments)),
        ["REGEXEXTRACT"] = new(2, 4, arguments => new RegexExtract(arguments)),
        ["REGEXREPLACE"] = new(3, 5, arguments => new RegexReplace(arguments)),
        ["CHOOSE"] = new(2, int.MaxValue, arguments => new Choose(arguments)),
        ["HLOOKUP"] = new(3, 4, arguments => new HLookup(arguments)),
        ["LOOKUP"] = new(2, 3, arguments => new Lookup(arguments)),
        ["COLUMN"] = new(0, 1, arguments => new Column(arguments)),
        ["COLUMNS"] = new(1, 1, arguments => new Columns(arguments)),
        ["XMATCH"] = new(2, 4, arguments => new XMatch(arguments)),
        ["ADDRESS"] = new(2, 5, arguments => new Address(arguments)),
        ["AREAS"] = new(1, 1, arguments => new Areas(arguments)),
        ["FORMULATEXT"] = new(1, 1, arguments => new FormulaText(arguments)),
        ["AVERAGEIF"] = new(2, 3, arguments => new AverageIf(arguments)),
        ["AVERAGEIFS"] = new(3, int.MaxValue, arguments => new AverageIfs(arguments)),
        ["MAXIFS"] = new(3, int.MaxValue, arguments => new MaxIfs(arguments)),
        ["MINIFS"] = new(3, int.MaxValue, arguments => new MinIfs(arguments)),
        ["AVERAGEA"] = new(1, int.MaxValue, arguments => new AverageA(arguments)),
        ["MAXA"] = new(1, int.MaxValue, arguments => new MaxA(arguments)),
        ["MINA"] = new(1, int.MaxValue, arguments => new MinA(arguments)),
        ["SUMPRODUCT"] = new(1, int.MaxValue, arguments => new SumProduct(arguments)),
        ["SUMX2MY2"] = new(2, 2, arguments => new SumX2MY2(arguments)),
        ["SUMX2PY2"] = new(2, 2, arguments => new SumX2PY2(arguments)),
        ["SUMXMY2"] = new(2, 2, arguments => new SumXMY2(arguments)),
        ["SUBTOTAL"] = new(2, int.MaxValue, arguments => new Subtotal(arguments)),
        ["MEDIAN"] = new(1, int.MaxValue, arguments => new Median(arguments)),
        ["MODE.SNGL"] = new(1, int.MaxValue, arguments => new ModeSngl(arguments)),
        ["LARGE"] = new(2, 2, arguments => new Large(arguments)),
        ["SMALL"] = new(2, 2, arguments => new Small(arguments)),
        ["RANK.EQ"] = new(2, 3, arguments => new RankEq(arguments)),
        ["RANK.AVG"] = new(2, 3, arguments => new RankAvg(arguments)),
        ["PERCENTILE.INC"] = new(2, 2, arguments => new PercentileInc(arguments)),
        ["PERCENTILE.EXC"] = new(2, 2, arguments => new PercentileExc(arguments)),
        ["PERCENTRANK.INC"] = new(2, 3, arguments => new PercentRankInc(arguments)),
        ["PERCENTRANK.EXC"] = new(2, 3, arguments => new PercentRankExc(arguments)),
        ["QUARTILE.INC"] = new(2, 2, arguments => new QuartileInc(arguments)),
        ["QUARTILE.EXC"] = new(2, 2, arguments => new QuartileExc(arguments)),
        ["TRIMMEAN"] = new(2, 2, arguments => new TrimMean(arguments)),
        ["STDEV.S"] = new(1, int.MaxValue, arguments => new StDevS(arguments)),
        ["STDEV.P"] = new(1, int.MaxValue, arguments => new StDevP(arguments)),
        ["STDEVA"] = new(1, int.MaxValue, arguments => new StDevA(arguments)),
        ["STDEVPA"] = new(1, int.MaxValue, arguments => new StDevPA(arguments)),
        ["VAR.S"] = new(1, int.MaxValue, arguments => new VarS(arguments)),
        ["VAR.P"] = new(1, int.MaxValue, arguments => new VarP(arguments)),
        ["VARA"] = new(1, int.MaxValue, arguments => new VarA(arguments)),
        ["VARPA"] = new(1, int.MaxValue, arguments => new VarPA(arguments)),
        ["AVEDEV"] = new(1, int.MaxValue, arguments => new AveDev(arguments)),
        ["DEVSQ"] = new(1, int.MaxValue, arguments => new DevSq(arguments)),
        ["GEOMEAN"] = new(1, int.MaxValue, arguments => new GeoMean(arguments)),
        ["HARMEAN"] = new(1, int.MaxValue, arguments => new HarMean(arguments)),
        ["SKEW"] = new(1, int.MaxValue, arguments => new Skew(arguments)),
        ["SKEW.P"] = new(1, int.MaxValue, arguments => new SkewP(arguments)),
        ["KURT"] = new(1, int.MaxValue, arguments => new Kurt(arguments)),
        ["STANDARDIZE"] = new(3, 3, arguments => new Standardize(arguments)),
        ["CORREL"] = new(2, 2, arguments => new Correl(arguments)),
        ["PEARSON"] = new(2, 2, arguments => new Pearson(arguments)),
        ["COVARIANCE.P"] = new(2, 2, arguments => new CovarianceP(arguments)),
        ["COVARIANCE.S"] = new(2, 2, arguments => new CovarianceS(arguments)),
        ["RSQ"] = new(2, 2, arguments => new Rsq(arguments)),
        ["SLOPE"] = new(2, 2, arguments => new Slope(arguments)),
        ["INTERCEPT"] = new(2, 2, arguments => new Intercept(arguments)),
        ["STEYX"] = new(2, 2, arguments => new Steyx(arguments)),
        ["FORECAST.LINEAR"] = new(3, 3, arguments => new ForecastLinear(arguments)),
        ["FISHER"] = new(1, 1, arguments => new Fisher(arguments)),
        ["FISHERINV"] = new(1, 1, arguments => new FisherInv(arguments)),
        ["PHI"] = new(1, 1, arguments => new Phi(arguments)),
        ["PERMUT"] = new(2, 2, arguments => new Permut(arguments)),
        ["PERMUTATIONA"] = new(2, 2, arguments => new PermutationA(arguments)),
        ["PROB"] = new(3, 4, arguments => new Prob(arguments)),
        // Compatibility aliases: distinct nodes so the un-parse keeps the legacy spelling.
        ["MODE"] = new(1, int.MaxValue, arguments => new Compat.Mode(arguments)),
        ["STDEV"] = new(1, int.MaxValue, arguments => new Compat.StDev(arguments)),
        ["STDEVP"] = new(1, int.MaxValue, arguments => new Compat.StDevP(arguments)),
        ["VAR"] = new(1, int.MaxValue, arguments => new Compat.Var(arguments)),
        ["VARP"] = new(1, int.MaxValue, arguments => new Compat.VarP(arguments)),
        ["RANK"] = new(2, 3, arguments => new Compat.Rank(arguments)),
        ["PERCENTILE"] = new(2, 2, arguments => new Compat.Percentile(arguments)),
        ["PERCENTRANK"] = new(2, 3, arguments => new Compat.PercentRank(arguments)),
        ["QUARTILE"] = new(2, 2, arguments => new Compat.Quartile(arguments)),
        ["COVAR"] = new(2, 2, arguments => new Compat.Covar(arguments)),
        ["FORECAST"] = new(3, 3, arguments => new Compat.Forecast(arguments)),
        // Wave 5 — Date and time (TODAY/NOW deferred to F1: volatile).
        ["DATE"] = new(3, 3, arguments => new Date(arguments)),
        ["TIME"] = new(3, 3, arguments => new Time(arguments)),
        ["DATEVALUE"] = new(1, 1, arguments => new DateValue(arguments)),
        ["TIMEVALUE"] = new(1, 1, arguments => new TimeValue(arguments)),
        ["YEAR"] = new(1, 1, arguments => new Year(arguments)),
        ["MONTH"] = new(1, 1, arguments => new Month(arguments)),
        ["DAY"] = new(1, 1, arguments => new Day(arguments)),
        ["HOUR"] = new(1, 1, arguments => new Hour(arguments)),
        ["MINUTE"] = new(1, 1, arguments => new Minute(arguments)),
        ["SECOND"] = new(1, 1, arguments => new Second(arguments)),
        ["DAYS"] = new(2, 2, arguments => new Days(arguments)),
        ["DAYS360"] = new(2, 3, arguments => new Days360(arguments)),
        ["EDATE"] = new(2, 2, arguments => new EDate(arguments)),
        ["EOMONTH"] = new(2, 2, arguments => new EoMonth(arguments)),
        ["WEEKDAY"] = new(1, 2, arguments => new Weekday(arguments)),
        ["WEEKNUM"] = new(1, 2, arguments => new WeekNum(arguments)),
        ["ISOWEEKNUM"] = new(1, 1, arguments => new IsoWeekNum(arguments)),
        ["DATEDIF"] = new(3, 3, arguments => new DateDif(arguments)),
        ["YEARFRAC"] = new(2, 3, arguments => new YearFrac(arguments)),
        ["NETWORKDAYS"] = new(2, 3, arguments => new NetworkDays(arguments)),
        ["NETWORKDAYS.INTL"] = new(2, 4, arguments => new NetworkDaysIntl(arguments)),
        ["WORKDAY"] = new(2, 3, arguments => new Workday(arguments)),
        ["WORKDAY.INTL"] = new(2, 4, arguments => new WorkdayIntl(arguments)),
        // F1 — volatile clock functions (both take no arguments).
        ["NOW"] = new(0, 0, arguments => new Now(arguments)),
        ["TODAY"] = new(0, 0, arguments => new Today(arguments)),
        // F1 — volatile RNG functions.
        ["RAND"] = new(0, 0, arguments => new Rand(arguments)),
        ["RANDBETWEEN"] = new(2, 2, arguments => new RandBetween(arguments)),
    };

    private int _index;

    private Token Current => tokens[_index];

    public Expression ParseFormula()
    {
        var expression = ParseExpression(0);

        if (Current.Type != TokenType.EndOfInput)
        {
            throw new ParseException($"Unexpected token '{Current.Text}'", Current.Position);
        }

        return expression;
    }

    private Expression ParseExpression(int rightBindingPower)
    {
        var left = ParsePrefix(Advance());

        while (rightBindingPower < LeftBindingPower(Current.Type))
        {
            left = ParseInfix(Advance(), left);
        }

        return left;
    }

    private Expression ParsePrefix(Token token)
    {
        switch (token.Type)
        {
            case TokenType.Number:
                return new NumberValue(double.Parse(token.Text, CultureInfo.InvariantCulture));

            case TokenType.String:
                return new StringValue(token.Text);

            case TokenType.Identifier:
                return ParseIdentifier(token);

            case TokenType.Minus:
                return new UnaryOperation(
                    UnaryOperator.Negate,
                    ParseExpression(PrefixBindingPower)
                );

            case TokenType.Plus:
                return new UnaryOperation(UnaryOperator.Plus, ParseExpression(PrefixBindingPower));

            case TokenType.LParen:
                var inner = ParseExpression(0);

                // A comma inside parentheses is the reference-union operator: (A1:A3, C1:C3).
                if (Current.Type == TokenType.Comma)
                {
                    var areas = new List<Expression> { inner };

                    while (Current.Type == TokenType.Comma)
                    {
                        Advance();
                        areas.Add(ParseExpression(0));
                    }

                    Expect(TokenType.RParen);
                    return new UnionReference(areas.ToArray());
                }

                Expect(TokenType.RParen);
                return inner;

            default:
                throw new ParseException($"Unexpected token '{token.Text}'", token.Position);
        }
    }

    private Expression ParseInfix(Token op, Expression left)
    {
        if (op.Type == TokenType.Colon)
        {
            return ParseRange(op, left);
        }

        // '%' is postfix: it divides the value to its left by 100, with no right operand.
        if (op.Type == TokenType.Percent)
        {
            return new UnaryOperation(UnaryOperator.Percent, left);
        }

        var bindingPower = LeftBindingPower(op.Type);
        var rightAssociative = op.Type == TokenType.Caret;
        var right = ParseExpression(rightAssociative ? bindingPower - 1 : bindingPower);

        return new BinaryOperation(ToBinaryOperator(op.Type), left, right);
    }

    private Expression ParseRange(Token colon, Expression left)
    {
        var right = ParseExpression(RangeBindingPower);

        if (left is CellReference start && right is CellReference end)
        {
            // The range lives on the start cell's sheet (e.g. Sheet2!A1:B2 is all on Sheet2).
            return new RangeReference(start.Id, end.Id, start.SheetName);
        }

        throw new ParseException("The ':' range operator requires cell references", colon.Position);
    }

    private Expression ParseIdentifier(Token token)
    {
        // A name before '!' is a sheet qualifier: Sheet2!A1, 'My Sheet'!A1.
        if (Current.Type == TokenType.Bang)
        {
            return ParseQualifiedReference(token.Text);
        }

        if (Current.Type == TokenType.LParen)
        {
            return ParseFunctionCall(token);
        }

        if (IsBoolean(token.Text, out var boolean))
        {
            return new BooleanValue(boolean);
        }

        if (IsCellReference(token.Text))
        {
            return new CellReference(NormalizeCellId(token.Text), sheetName);
        }

        // A bare name: a LET-bound name resolved at evaluation time (#NAME? if unbound).
        return new NameReference(token.Text);
    }

    private Expression ParseQualifiedReference(string sheet)
    {
        Expect(TokenType.Bang);
        var token = Advance();

        if (token.Type != TokenType.Identifier || !IsCellReference(token.Text))
        {
            throw new ParseException("Expected a cell reference after '!'", token.Position);
        }

        return new CellReference(NormalizeCellId(token.Text), sheet);
    }

    // Strips absolute markers ('$') and upper-cases, e.g. $A$1 -> A1. The reference identifies the same
    // cell regardless of '$'; absolute/relative only matters for Excel copy/fill, which we do not do.
    private static string NormalizeCellId(string text) => StripDollars(text).ToUpperInvariant();

    private static string StripDollars(string text) =>
        text.Contains('$') ? text.Replace("$", string.Empty) : text;

    private Expression ParseFunctionCall(Token name)
    {
        Expect(TokenType.LParen);

        var arguments = new List<Expression>();

        if (Current.Type != TokenType.RParen)
        {
            arguments.Add(ParseArgument());

            while (Current.Type == TokenType.Comma)
            {
                Advance();
                arguments.Add(ParseArgument());
            }
        }

        Expect(TokenType.RParen);

        var functionName = NormalizeFunctionName(name.Text);

        // Built-in: typed record with parse-time arity validation (a wrong count throws, like Excel
        // rejecting it at entry). Otherwise a generic call resolved at runtime against the workbook's
        // custom-function registry (#NAME? if never registered).
        if (!Functions.TryGetValue(functionName, out var spec))
        {
            return new FunctionCall(functionName, arguments.ToArray());
        }

        if (arguments.Count < spec.MinArgs || arguments.Count > spec.MaxArgs)
        {
            throw new ParseException(
                $"Function '{functionName}' does not accept {arguments.Count} argument(s)",
                name.Position
            );
        }

        return spec.Create(arguments.ToArray());
    }

    private Expression ParseArgument()
    {
        // An omitted argument (e.g. XLOOKUP(a,b,c,,2) or a trailing comma) is treated as blank.
        if (Current.Type is TokenType.Comma or TokenType.RParen)
        {
            return BlankValue.Instance;
        }

        return ParseExpression(0);
    }

    // Excel stores newer functions with an "_xlfn." prefix; normalize it (and the bare "XLFN.") away.
    private static string NormalizeFunctionName(string name)
    {
        if (name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase))
        {
            return name["_xlfn.".Length..];
        }

        if (name.StartsWith("xlfn.", StringComparison.OrdinalIgnoreCase))
        {
            return name["xlfn.".Length..];
        }

        return name;
    }

    private Token Advance()
    {
        var token = tokens[_index];

        if (token.Type != TokenType.EndOfInput)
        {
            _index++;
        }

        return token;
    }

    private void Expect(TokenType type)
    {
        if (Current.Type != type)
        {
            throw new ParseException(
                $"Expected {type} but found '{Current.Text}'",
                Current.Position
            );
        }

        Advance();
    }

    private static int LeftBindingPower(TokenType type) =>
        type switch
        {
            TokenType.Equal
            or TokenType.NotEqual
            or TokenType.Less
            or TokenType.Greater
            or TokenType.LessEqual
            or TokenType.GreaterEqual => ComparisonBindingPower,
            TokenType.Ampersand => ConcatBindingPower,
            TokenType.Plus or TokenType.Minus => AdditiveBindingPower,
            TokenType.Star or TokenType.Slash => MultiplicativeBindingPower,
            TokenType.Caret => PowerBindingPower,
            TokenType.Percent => PercentBindingPower,
            TokenType.Colon => RangeBindingPower,
            _ => 0,
        };

    private static BinaryOperator ToBinaryOperator(TokenType type) =>
        type switch
        {
            TokenType.Plus => BinaryOperator.Add,
            TokenType.Minus => BinaryOperator.Subtract,
            TokenType.Star => BinaryOperator.Multiply,
            TokenType.Slash => BinaryOperator.Divide,
            TokenType.Caret => BinaryOperator.Power,
            TokenType.Equal => BinaryOperator.Equal,
            TokenType.NotEqual => BinaryOperator.NotEqual,
            TokenType.Less => BinaryOperator.LessThan,
            TokenType.Greater => BinaryOperator.GreaterThan,
            TokenType.LessEqual => BinaryOperator.LessThanOrEqual,
            TokenType.GreaterEqual => BinaryOperator.GreaterThanOrEqual,
            TokenType.Ampersand => BinaryOperator.Concat,
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, null),
        };

    private static bool IsBoolean(string text, out bool value)
    {
        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        value = false;
        return false;
    }

    // Internal: the defined-name validator reuses this as the single source of truth for "looks like a
    // cell reference" (an A1-shaped name is reserved and cannot be a defined name).
    internal static bool IsCellReference(string text)
    {
        text = StripDollars(text);

        var letters = 0;
        while (letters < text.Length && char.IsLetter(text[letters]))
        {
            letters++;
        }

        if (letters == 0 || letters == text.Length)
        {
            return false;
        }

        for (var i = letters; i < text.Length; i++)
        {
            if (!char.IsDigit(text[i]))
            {
                return false;
            }
        }

        return true;
    }
}
