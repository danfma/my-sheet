using System.Globalization;
using System.Text;
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
/// The inverse of <see cref="ExpressionParser"/>: renders an <see cref="Expression"/> tree back to Excel
/// formula text (without the leading <c>=</c>), emitting the minimal parentheses that re-parse to the
/// same tree.
/// </summary>
public static class FormulaWriter
{
    // Mirror of the Parser's Pratt binding powers, so parenthesization decisions here agree exactly
    // with how the text will be re-parsed.
    private const int ComparisonPrecedence = 10;
    private const int ConcatPrecedence = 15;
    private const int AdditivePrecedence = 20;
    private const int MultiplicativePrecedence = 30;
    private const int PowerPrecedence = 40;
    private const int PercentPrecedence = 44;
    private const int PrefixPrecedence = 45;
    private const int AtomPrecedence = int.MaxValue;

    /// <summary>
    /// Renders the expression as Excel formula text. References on <paramref name="contextSheetName"/>
    /// stay unqualified (<c>A1</c>); references to other sheets are qualified (<c>Sheet2!A1</c>, quoted
    /// when the name needs it).
    /// </summary>
    public static string ToFormula(this Expression expression, string contextSheetName)
    {
        var builder = new StringBuilder();

        Write(builder, expression, contextSheetName, minPrecedence: 0);

        return builder.ToString();
    }

    private static void Write(
        StringBuilder builder,
        Expression expression,
        string context,
        int minPrecedence
    )
    {
        var parenthesize = Precedence(expression) < minPrecedence;

        if (parenthesize)
        {
            builder.Append('(');
        }

        switch (expression)
        {
            case NumberValue number:
                builder.Append(number.Value.ToString(CultureInfo.InvariantCulture));
                break;

            case StringValue text:
                builder.Append('"').Append(text.Value.Replace("\"", "\"\"")).Append('"');
                break;

            case BooleanValue boolean:
                builder.Append(boolean.Value ? "TRUE" : "FALSE");
                break;

            case BlankValue:
                break; // an omitted argument renders as nothing

            case ErrorValue error:
                builder.Append(error.ErrorCode);
                break;

            case CellReference cell:
                WriteSheetQualifier(builder, cell.SheetName, context);
                builder.Append(cell.Id);
                break;

            case RangeReference range:
                WriteSheetQualifier(builder, range.SheetName, context);
                builder.Append(range.StartId).Append(':').Append(range.EndId);
                break;

            case OpenRangeReference open:
                WriteSheetQualifier(builder, open.SheetName, context);
                WriteOpenEndpoint(builder, open.ColMin, open.RowMin);
                builder.Append(':');
                WriteOpenEndpoint(builder, open.ColMax, open.RowMax);
                break;

            case UnionReference union:
                builder.Append('(');
                WriteList(builder, union.Areas, context);
                builder.Append(')');
                break;

            case NameReference name:
                builder.Append(name.Name);
                break;

            case DynamicRange dyn:
                Write(builder, dyn.Start, context, AtomPrecedence);
                builder.Append(':');
                Write(builder, dyn.End, context, AtomPrecedence);
                break;

            case UnaryOperation { Operator: UnaryOperator.Percent } percent:
                Write(builder, percent.Operand, context, PercentPrecedence);
                builder.Append('%');
                break;

            case UnaryOperation prefix:
                builder.Append(prefix.Operator == UnaryOperator.Negate ? '-' : '+');
                Write(builder, prefix.Operand, context, PrefixPrecedence);
                break;

            case BinaryOperation binary:
                WriteBinary(builder, binary, context);
                break;

            case Function function:
                var (functionName, arguments) = Call(function);
                builder.Append(functionName).Append('(');
                WriteList(builder, arguments, context);
                builder.Append(')');
                break;

            default:
                throw new NotSupportedException(
                    $"No Excel formula rendering for node '{expression.GetType().Name}'."
                );
        }

        if (parenthesize)
        {
            builder.Append(')');
        }
    }

    private static void WriteBinary(StringBuilder builder, BinaryOperation binary, string context)
    {
        var precedence = Precedence(binary);

        // '^' is right-associative (like the Parser); every other operator is left-associative, so the
        // tighter minimum goes on the opposite side to force parentheses only where re-parsing needs them.
        var rightAssociative = binary.Operator == BinaryOperator.Power;

        Write(builder, binary.Left, context, rightAssociative ? precedence + 1 : precedence);
        builder.Append(Token(binary.Operator));
        Write(builder, binary.Right, context, rightAssociative ? precedence : precedence + 1);
    }

    private static void WriteList(StringBuilder builder, Expression[] items, string context)
    {
        for (var i = 0; i < items.Length; i++)
        {
            if (i > 0)
            {
                builder.Append(',');
            }

            Write(builder, items[i], context, minPrecedence: 0);
        }
    }

    // Renders one endpoint of an open range: the column letters (when the column is known) followed by the
    // row number (when the row is known). A column-only endpoint prints "A", a row-only endpoint "1", a
    // both-known endpoint the full cell id "A1" — so A:A, 1:5, A1:C etc. all re-parse to the same tree.
    private static void WriteOpenEndpoint(StringBuilder builder, int? column, int? row)
    {
        if (column is { } columnNumber)
        {
            var start = builder.Length;
            var value = columnNumber;

            while (value > 0)
            {
                var remainder = (value - 1) % 26;
                builder.Insert(start, (char)('A' + remainder));
                value = (value - 1) / 26;
            }
        }

        if (row is { } rowNumber)
        {
            builder.Append(rowNumber);
        }
    }

    private static void WriteSheetQualifier(StringBuilder builder, string sheetName, string context)
    {
        if (string.Equals(sheetName, context, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (IsSimpleSheetName(sheetName))
        {
            builder.Append(sheetName);
        }
        else
        {
            builder.Append('\'').Append(sheetName.Replace("'", "''")).Append('\'');
        }

        builder.Append('!');
    }

    // A name that tokenizes as a single identifier needs no quotes; anything else (spaces, symbols,
    // leading digit) is single-quoted with '' escaping, exactly what the Tokenizer reads back.
    // Internal because ADDRESS shares this quoting rule for its sheet_text prefix.
    internal static bool IsSimpleSheetName(string name)
    {
        if (name.Length == 0 || char.IsDigit(name[0]))
        {
            return false;
        }

        foreach (var c in name)
        {
            if (!char.IsLetterOrDigit(c) && c != '_')
            {
                return false;
            }
        }

        return true;
    }

    private static int Precedence(Expression expression) =>
        expression switch
        {
            BinaryOperation binary => binary.Operator switch
            {
                BinaryOperator.Equal
                or BinaryOperator.NotEqual
                or BinaryOperator.LessThan
                or BinaryOperator.GreaterThan
                or BinaryOperator.LessThanOrEqual
                or BinaryOperator.GreaterThanOrEqual => ComparisonPrecedence,
                BinaryOperator.Concat => ConcatPrecedence,
                BinaryOperator.Add or BinaryOperator.Subtract => AdditivePrecedence,
                BinaryOperator.Multiply or BinaryOperator.Divide => MultiplicativePrecedence,
                BinaryOperator.Power => PowerPrecedence,
                _ => AtomPrecedence,
            },
            UnaryOperation { Operator: UnaryOperator.Percent } => PercentPrecedence,
            UnaryOperation => PrefixPrecedence,
            _ => AtomPrecedence,
        };

    private static string Token(BinaryOperator op) =>
        op switch
        {
            BinaryOperator.Add => "+",
            BinaryOperator.Subtract => "-",
            BinaryOperator.Multiply => "*",
            BinaryOperator.Divide => "/",
            BinaryOperator.Power => "^",
            BinaryOperator.Equal => "=",
            BinaryOperator.NotEqual => "<>",
            BinaryOperator.LessThan => "<",
            BinaryOperator.GreaterThan => ">",
            BinaryOperator.LessThanOrEqual => "<=",
            BinaryOperator.GreaterThanOrEqual => ">=",
            BinaryOperator.Concat => "&",
            _ => throw new NotSupportedException($"Unknown binary operator '{op}'."),
        };

    // Central node-type → Excel-function-name map, the mirror of the Parser's Functions table. A custom
    // FunctionCall keeps the name it was parsed/registered with. Internal so the defined-name validator can
    // reuse it as the single source of a function node's argument list when walking for unqualified refs.
    internal static (string Name, Expression[] Arguments) Call(Function function) =>
        function switch
        {
            FunctionCall f => (f.Name, f.Arguments),
            Sum f => ("SUM", f.Expressions), // Sum is the one function whose parameter is not named Arguments
            Average f => ("AVERAGE", f.Arguments),
            Min f => ("MIN", f.Arguments),
            Max f => ("MAX", f.Arguments),
            Count f => ("COUNT", f.Arguments),
            If f => ("IF", f.Arguments),
            And f => ("AND", f.Arguments),
            Or f => ("OR", f.Arguments),
            Not f => ("NOT", f.Arguments),
            IfError f => ("IFERROR", f.Arguments),
            IfNa f => ("IFNA", f.Arguments),
            Int f => ("INT", f.Arguments),
            Round f => ("ROUND", f.Arguments),
            RoundUp f => ("ROUNDUP", f.Arguments),
            Abs f => ("ABS", f.Arguments),
            IsNumber f => ("ISNUMBER", f.Arguments),
            IsBlank f => ("ISBLANK", f.Arguments),
            Upper f => ("UPPER", f.Arguments),
            Lower f => ("LOWER", f.Arguments),
            Trim f => ("TRIM", f.Arguments),
            Len f => ("LEN", f.Arguments),
            Left f => ("LEFT", f.Arguments),
            Mid f => ("MID", f.Arguments),
            Value f => ("VALUE", f.Arguments),
            Text f => ("TEXT", f.Arguments),
            Concat f => ("CONCAT", f.Arguments),
            Concatenate f => ("CONCATENATE", f.Arguments),
            TextJoin f => ("TEXTJOIN", f.Arguments),
            CountA f => ("COUNTA", f.Arguments),
            CountBlank f => ("COUNTBLANK", f.Arguments),
            CountIf f => ("COUNTIF", f.Arguments),
            CountIfs f => ("COUNTIFS", f.Arguments),
            SumIf f => ("SUMIF", f.Arguments),
            SumIfs f => ("SUMIFS", f.Arguments),
            Rows f => ("ROWS", f.Arguments),
            Row f => ("ROW", f.Arguments),
            Match f => ("MATCH", f.Arguments),
            Expressions.Lookup.Index f => ("INDEX", f.Arguments),
            VLookup f => ("VLOOKUP", f.Arguments),
            XLookup f => ("XLOOKUP", f.Arguments),
            Offset f => ("OFFSET", f.Arguments),
            Indirect f => ("INDIRECT", f.Arguments),
            Let f => ("LET", f.Arguments),
            SheetNumber f => ("SHEET", f.Arguments),
            Pmt f => ("PMT", f.Arguments),
            Pv f => ("PV", f.Arguments),
            Fv f => ("FV", f.Arguments),
            Nper f => ("NPER", f.Arguments),
            Ipmt f => ("IPMT", f.Arguments),
            Ppmt f => ("PPMT", f.Arguments),
            Npv f => ("NPV", f.Arguments),
            Rate f => ("RATE", f.Arguments),
            Irr f => ("IRR", f.Arguments),
            Sln f => ("SLN", f.Arguments),
            Syd f => ("SYD", f.Arguments),
            Db f => ("DB", f.Arguments),
            Ddb f => ("DDB", f.Arguments),
            Vdb f => ("VDB", f.Arguments),
            AmorLinc f => ("AMORLINC", f.Arguments),
            AmorDegrc f => ("AMORDEGRC", f.Arguments),
            Effect f => ("EFFECT", f.Arguments),
            Nominal f => ("NOMINAL", f.Arguments),
            Mirr f => ("MIRR", f.Arguments),
            Rri f => ("RRI", f.Arguments),
            PDuration f => ("PDURATION", f.Arguments),
            ISPmt f => ("ISPMT", f.Arguments),
            CumIPmt f => ("CUMIPMT", f.Arguments),
            CumPrinc f => ("CUMPRINC", f.Arguments),
            FvSchedule f => ("FVSCHEDULE", f.Arguments),
            DollarDe f => ("DOLLARDE", f.Arguments),
            DollarFr f => ("DOLLARFR", f.Arguments),
            XNpv f => ("XNPV", f.Arguments),
            XIrr f => ("XIRR", f.Arguments),
            AccrInt f => ("ACCRINT", f.Arguments),
            AccrIntM f => ("ACCRINTM", f.Arguments),
            Disc f => ("DISC", f.Arguments),
            IntRate f => ("INTRATE", f.Arguments),
            Received f => ("RECEIVED", f.Arguments),
            PriceDisc f => ("PRICEDISC", f.Arguments),
            PriceMat f => ("PRICEMAT", f.Arguments),
            YieldDisc f => ("YIELDDISC", f.Arguments),
            YieldMat f => ("YIELDMAT", f.Arguments),
            TBillEq f => ("TBILLEQ", f.Arguments),
            TBillPrice f => ("TBILLPRICE", f.Arguments),
            TBillYield f => ("TBILLYIELD", f.Arguments),
            CoupPcd f => ("COUPPCD", f.Arguments),
            CoupNcd f => ("COUPNCD", f.Arguments),
            CoupNum f => ("COUPNUM", f.Arguments),
            CoupDays f => ("COUPDAYS", f.Arguments),
            CoupDayBs f => ("COUPDAYBS", f.Arguments),
            CoupDaysNc f => ("COUPDAYSNC", f.Arguments),
            Price f => ("PRICE", f.Arguments),
            Yield f => ("YIELD", f.Arguments),
            Duration f => ("DURATION", f.Arguments),
            MDuration f => ("MDURATION", f.Arguments),
            OddFPrice f => ("ODDFPRICE", f.Arguments),
            OddFYield f => ("ODDFYIELD", f.Arguments),
            OddLPrice f => ("ODDLPRICE", f.Arguments),
            OddLYield f => ("ODDLYIELD", f.Arguments),
            Sqrt f => ("SQRT", f.Arguments),
            Power f => ("POWER", f.Arguments),
            Exp f => ("EXP", f.Arguments),
            Ln f => ("LN", f.Arguments),
            Log f => ("LOG", f.Arguments),
            Log10 f => ("LOG10", f.Arguments),
            SqrtPi f => ("SQRTPI", f.Arguments),
            RoundDown f => ("ROUNDDOWN", f.Arguments),
            Trunc f => ("TRUNC", f.Arguments),
            MRound f => ("MROUND", f.Arguments),
            Ceiling f => ("CEILING", f.Arguments),
            CeilingMath f => ("CEILING.MATH", f.Arguments),
            CeilingPrecise f => ("CEILING.PRECISE", f.Arguments),
            IsoCeiling f => ("ISO.CEILING", f.Arguments),
            Floor f => ("FLOOR", f.Arguments),
            FloorMath f => ("FLOOR.MATH", f.Arguments),
            FloorPrecise f => ("FLOOR.PRECISE", f.Arguments),
            Even f => ("EVEN", f.Arguments),
            Odd f => ("ODD", f.Arguments),
            Mod f => ("MOD", f.Arguments),
            Quotient f => ("QUOTIENT", f.Arguments),
            Sign f => ("SIGN", f.Arguments),
            Pi f => ("PI", f.Arguments),
            Product f => ("PRODUCT", f.Arguments),
            SumSq f => ("SUMSQ", f.Arguments),
            Multinomial f => ("MULTINOMIAL", f.Arguments),
            SeriesSum f => ("SERIESSUM", f.Arguments),
            Fact f => ("FACT", f.Arguments),
            FactDouble f => ("FACTDOUBLE", f.Arguments),
            Combin f => ("COMBIN", f.Arguments),
            CombinA f => ("COMBINA", f.Arguments),
            Gcd f => ("GCD", f.Arguments),
            Lcm f => ("LCM", f.Arguments),
            Sin f => ("SIN", f.Arguments),
            Cos f => ("COS", f.Arguments),
            Tan f => ("TAN", f.Arguments),
            Cot f => ("COT", f.Arguments),
            Sec f => ("SEC", f.Arguments),
            Csc f => ("CSC", f.Arguments),
            Asin f => ("ASIN", f.Arguments),
            Acos f => ("ACOS", f.Arguments),
            Atan f => ("ATAN", f.Arguments),
            Atan2 f => ("ATAN2", f.Arguments),
            Acot f => ("ACOT", f.Arguments),
            Sinh f => ("SINH", f.Arguments),
            Cosh f => ("COSH", f.Arguments),
            Tanh f => ("TANH", f.Arguments),
            Coth f => ("COTH", f.Arguments),
            Sech f => ("SECH", f.Arguments),
            Csch f => ("CSCH", f.Arguments),
            Asinh f => ("ASINH", f.Arguments),
            Acosh f => ("ACOSH", f.Arguments),
            Atanh f => ("ATANH", f.Arguments),
            Acoth f => ("ACOTH", f.Arguments),
            Degrees f => ("DEGREES", f.Arguments),
            Radians f => ("RADIANS", f.Arguments),
            Base f => ("BASE", f.Arguments),
            DecimalNumber f => ("DECIMAL", f.Arguments),
            Roman f => ("ROMAN", f.Arguments),
            Arabic f => ("ARABIC", f.Arguments),
            TrueFunction f => ("TRUE", f.Arguments),
            FalseFunction f => ("FALSE", f.Arguments),
            Xor f => ("XOR", f.Arguments),
            Ifs f => ("IFS", f.Arguments),
            Switch f => ("SWITCH", f.Arguments),
            Na f => ("NA", f.Arguments),
            IsError f => ("ISERROR", f.Arguments),
            IsErr f => ("ISERR", f.Arguments),
            IsNa f => ("ISNA", f.Arguments),
            IsText f => ("ISTEXT", f.Arguments),
            IsNonText f => ("ISNONTEXT", f.Arguments),
            IsLogical f => ("ISLOGICAL", f.Arguments),
            IsEven f => ("ISEVEN", f.Arguments),
            IsOdd f => ("ISODD", f.Arguments),
            IsRef f => ("ISREF", f.Arguments),
            IsFormula f => ("ISFORMULA", f.Arguments),
            N f => ("N", f.Arguments),
            T f => ("T", f.Arguments),
            TypeFunction f => ("TYPE", f.Arguments),
            ErrorType f => ("ERROR.TYPE", f.Arguments),
            SheetsCount f => ("SHEETS", f.Arguments),
            Right f => ("RIGHT", f.Arguments),
            Find f => ("FIND", f.Arguments),
            Search f => ("SEARCH", f.Arguments),
            Replace f => ("REPLACE", f.Arguments),
            Substitute f => ("SUBSTITUTE", f.Arguments),
            Rept f => ("REPT", f.Arguments),
            Proper f => ("PROPER", f.Arguments),
            Exact f => ("EXACT", f.Arguments),
            CharFunction f => ("CHAR", f.Arguments),
            Code f => ("CODE", f.Arguments),
            UniChar f => ("UNICHAR", f.Arguments),
            Unicode f => ("UNICODE", f.Arguments),
            Clean f => ("CLEAN", f.Arguments),
            Fixed f => ("FIXED", f.Arguments),
            Dollar f => ("DOLLAR", f.Arguments),
            NumberValueFunction f => ("NUMBERVALUE", f.Arguments),
            TextBefore f => ("TEXTBEFORE", f.Arguments),
            TextAfter f => ("TEXTAFTER", f.Arguments),
            ValueToText f => ("VALUETOTEXT", f.Arguments),
            RegexTest f => ("REGEXTEST", f.Arguments),
            RegexExtract f => ("REGEXEXTRACT", f.Arguments),
            RegexReplace f => ("REGEXREPLACE", f.Arguments),
            Choose f => ("CHOOSE", f.Arguments),
            HLookup f => ("HLOOKUP", f.Arguments),
            Lookup f => ("LOOKUP", f.Arguments),
            Column f => ("COLUMN", f.Arguments),
            Columns f => ("COLUMNS", f.Arguments),
            XMatch f => ("XMATCH", f.Arguments),
            Address f => ("ADDRESS", f.Arguments),
            Areas f => ("AREAS", f.Arguments),
            FormulaText f => ("FORMULATEXT", f.Arguments),
            AverageIf f => ("AVERAGEIF", f.Arguments),
            AverageIfs f => ("AVERAGEIFS", f.Arguments),
            MaxIfs f => ("MAXIFS", f.Arguments),
            MinIfs f => ("MINIFS", f.Arguments),
            AverageA f => ("AVERAGEA", f.Arguments),
            MaxA f => ("MAXA", f.Arguments),
            MinA f => ("MINA", f.Arguments),
            SumProduct f => ("SUMPRODUCT", f.Arguments),
            SumX2MY2 f => ("SUMX2MY2", f.Arguments),
            SumX2PY2 f => ("SUMX2PY2", f.Arguments),
            SumXMY2 f => ("SUMXMY2", f.Arguments),
            Subtotal f => ("SUBTOTAL", f.Arguments),
            Median f => ("MEDIAN", f.Arguments),
            ModeSngl f => ("MODE.SNGL", f.Arguments),
            Large f => ("LARGE", f.Arguments),
            Small f => ("SMALL", f.Arguments),
            RankEq f => ("RANK.EQ", f.Arguments),
            RankAvg f => ("RANK.AVG", f.Arguments),
            PercentileInc f => ("PERCENTILE.INC", f.Arguments),
            PercentileExc f => ("PERCENTILE.EXC", f.Arguments),
            PercentRankInc f => ("PERCENTRANK.INC", f.Arguments),
            PercentRankExc f => ("PERCENTRANK.EXC", f.Arguments),
            QuartileInc f => ("QUARTILE.INC", f.Arguments),
            QuartileExc f => ("QUARTILE.EXC", f.Arguments),
            TrimMean f => ("TRIMMEAN", f.Arguments),
            StDevS f => ("STDEV.S", f.Arguments),
            StDevP f => ("STDEV.P", f.Arguments),
            StDevA f => ("STDEVA", f.Arguments),
            StDevPA f => ("STDEVPA", f.Arguments),
            VarS f => ("VAR.S", f.Arguments),
            VarP f => ("VAR.P", f.Arguments),
            VarA f => ("VARA", f.Arguments),
            VarPA f => ("VARPA", f.Arguments),
            AveDev f => ("AVEDEV", f.Arguments),
            DevSq f => ("DEVSQ", f.Arguments),
            GeoMean f => ("GEOMEAN", f.Arguments),
            HarMean f => ("HARMEAN", f.Arguments),
            Skew f => ("SKEW", f.Arguments),
            SkewP f => ("SKEW.P", f.Arguments),
            Kurt f => ("KURT", f.Arguments),
            Standardize f => ("STANDARDIZE", f.Arguments),
            Correl f => ("CORREL", f.Arguments),
            Pearson f => ("PEARSON", f.Arguments),
            CovarianceP f => ("COVARIANCE.P", f.Arguments),
            CovarianceS f => ("COVARIANCE.S", f.Arguments),
            Rsq f => ("RSQ", f.Arguments),
            Slope f => ("SLOPE", f.Arguments),
            Intercept f => ("INTERCEPT", f.Arguments),
            Steyx f => ("STEYX", f.Arguments),
            ForecastLinear f => ("FORECAST.LINEAR", f.Arguments),
            Fisher f => ("FISHER", f.Arguments),
            FisherInv f => ("FISHERINV", f.Arguments),
            Phi f => ("PHI", f.Arguments),
            Permut f => ("PERMUT", f.Arguments),
            PermutationA f => ("PERMUTATIONA", f.Arguments),
            Prob f => ("PROB", f.Arguments),
            // Compatibility aliases: the distinct node preserves the legacy spelling on un-parse.
            Compat.Mode f => ("MODE", f.Arguments),
            Compat.StDev f => ("STDEV", f.Arguments),
            Compat.StDevP f => ("STDEVP", f.Arguments),
            Compat.Var f => ("VAR", f.Arguments),
            Compat.VarP f => ("VARP", f.Arguments),
            Compat.Rank f => ("RANK", f.Arguments),
            Compat.Percentile f => ("PERCENTILE", f.Arguments),
            Compat.PercentRank f => ("PERCENTRANK", f.Arguments),
            Compat.Quartile f => ("QUARTILE", f.Arguments),
            Compat.Covar f => ("COVAR", f.Arguments),
            Compat.Forecast f => ("FORECAST", f.Arguments),
            // Wave 5 — Date and time.
            Date f => ("DATE", f.Arguments),
            Time f => ("TIME", f.Arguments),
            DateValue f => ("DATEVALUE", f.Arguments),
            TimeValue f => ("TIMEVALUE", f.Arguments),
            Year f => ("YEAR", f.Arguments),
            Month f => ("MONTH", f.Arguments),
            Day f => ("DAY", f.Arguments),
            Hour f => ("HOUR", f.Arguments),
            Minute f => ("MINUTE", f.Arguments),
            Second f => ("SECOND", f.Arguments),
            Days f => ("DAYS", f.Arguments),
            Days360 f => ("DAYS360", f.Arguments),
            EDate f => ("EDATE", f.Arguments),
            EoMonth f => ("EOMONTH", f.Arguments),
            Weekday f => ("WEEKDAY", f.Arguments),
            WeekNum f => ("WEEKNUM", f.Arguments),
            IsoWeekNum f => ("ISOWEEKNUM", f.Arguments),
            DateDif f => ("DATEDIF", f.Arguments),
            YearFrac f => ("YEARFRAC", f.Arguments),
            NetworkDays f => ("NETWORKDAYS", f.Arguments),
            NetworkDaysIntl f => ("NETWORKDAYS.INTL", f.Arguments),
            Workday f => ("WORKDAY", f.Arguments),
            WorkdayIntl f => ("WORKDAY.INTL", f.Arguments),
            // F1 — volatile clock functions.
            Now f => ("NOW", f.Arguments),
            Today f => ("TODAY", f.Arguments),
            // F1 — volatile RNG functions.
            Rand f => ("RAND", f.Arguments),
            RandBetween f => ("RANDBETWEEN", f.Arguments),
            _ => throw new NotSupportedException(
                $"No Excel function name registered for node '{function.GetType().Name}'."
            ),
        };
}
