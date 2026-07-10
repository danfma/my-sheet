using System.Collections.Frozen;
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

    // Mirrors the Parser's MaxDepth: guards against unbounded recursion (and an uncatchable
    // StackOverflowException) when rendering an expression tree built PROGRAMMATICALLY rather than parsed —
    // a parsed tree can never exceed this because the Parser enforces the same limit at parse time, but a
    // caller assembling nodes by hand (e.g. a loop of nested UnaryOperation) has no such guard, so the writer
    // enforces its own. 256 matches the parser's limit; there is no ParseException here (this is not parsing
    // a formula, it is un-parsing an already-built tree), so an illegal tree throws InvalidOperationException.
    private const int MaxDepth = 256;

    /// <summary>
    /// Renders the expression as Excel formula text. References on <paramref name="contextSheetName"/>
    /// stay unqualified (<c>A1</c>); references to other sheets are qualified (<c>Sheet2!A1</c>, quoted
    /// when the name needs it).
    /// </summary>
    public static string ToFormula(this Expression expression, string contextSheetName)
    {
        var builder = new StringBuilder();

        Write(builder, expression, contextSheetName, minPrecedence: 0, depth: 0);

        return builder.ToString();
    }

    // depth counts Write's own recursion; WriteBinary and WriteList are plain helpers invoked FROM here
    // (never independent recursion entry points), so they just pass the incremented depth through to their
    // own Write calls instead of tracking it themselves.
    private static void Write(
        StringBuilder builder,
        Expression expression,
        string context,
        int minPrecedence,
        int depth
    )
    {
        if (depth > MaxDepth)
        {
            throw new InvalidOperationException(
                "Formula nesting is too deep: the expression tree exceeds the maximum supported depth "
                    + $"({MaxDepth}). This tree was built programmatically — a parsed formula can never "
                    + "hit this limit."
            );
        }

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
                WriteList(builder, union.Areas, context, depth + 1);
                builder.Append(')');
                break;

            case NameReference name:
                builder.Append(name.Name);
                break;

            case DynamicRange dyn:
                Write(builder, dyn.Start, context, AtomPrecedence, depth + 1);
                builder.Append(':');
                Write(builder, dyn.End, context, AtomPrecedence, depth + 1);
                break;

            case UnaryOperation { Operator: UnaryOperator.Percent } percent:
                Write(builder, percent.Operand, context, PercentPrecedence, depth + 1);
                builder.Append('%');
                break;

            case UnaryOperation prefix:
                builder.Append(prefix.Operator == UnaryOperator.Negate ? '-' : '+');
                Write(builder, prefix.Operand, context, PrefixPrecedence, depth + 1);
                break;

            case BinaryOperation binary:
                WriteBinary(builder, binary, context, depth + 1);
                break;

            case Function function:
                var (functionName, arguments) = Call(function);
                builder.Append(functionName).Append('(');
                WriteList(builder, arguments, context, depth + 1);
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

    private static void WriteBinary(
        StringBuilder builder,
        BinaryOperation binary,
        string context,
        int depth
    )
    {
        var precedence = Precedence(binary);

        // '^' is right-associative (like the Parser); every other operator is left-associative, so the
        // tighter minimum goes on the opposite side to force parentheses only where re-parsing needs them.
        var rightAssociative = binary.Operator == BinaryOperator.Power;

        Write(builder, binary.Left, context, rightAssociative ? precedence + 1 : precedence, depth);
        builder.Append(Token(binary.Operator));
        Write(
            builder,
            binary.Right,
            context,
            rightAssociative ? precedence : precedence + 1,
            depth
        );
    }

    private static void WriteList(
        StringBuilder builder,
        Expression[] items,
        string context,
        int depth
    )
    {
        for (var i = 0; i < items.Length; i++)
        {
            if (i > 0)
            {
                builder.Append(',');
            }

            Write(builder, items[i], context, minPrecedence: 0, depth);
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

    // Central node-type -> Excel-function-name map, the mirror of the Parser's Functions table. A custom
    // FunctionCall keeps the name it was parsed/registered with. Internal so the defined-name validator can
    // reuse it as the single source of a function node's argument list when walking for unqualified refs.
    //
    // The overwhelming majority of node types (304 of 306) share one shape: `Type => ("NAME", node.Arguments)`
    // -- nothing but a literal Excel name and the node's own Arguments array. Those are served by UniformCalls,
    // a FrozenDictionary<Type, ...> keyed on the CLR type, so the per-call cost is one hash lookup + one
    // delegate invocation instead of what used to be a ~300-deep isinst chain (Call runs on EVERY cell of a
    // Formulas-mode export and every FORMULATEXT call, so the switch's worst case was ~300 type tests per
    // cell). Only the two IRREGULAR shapes stay in a residual switch: FunctionCall (the name comes from the
    // node itself, not a literal) and Sum (the one function whose argument array is named Expressions, not
    // Arguments).
    internal static (string Name, Expression[] Arguments) Call(Function function) =>
        function switch
        {
            FunctionCall f => (f.Name, f.Arguments),
            Sum f => ("SUM", f.Expressions), // Sum is the one function whose parameter is not named Arguments
            _ when UniformCalls.TryGetValue(function.GetType(), out var entry) => (
                entry.Name,
                entry.Arguments(function)
            ),
            _ => throw new NotSupportedException(
                $"No Excel function name registered for node '{function.GetType().Name}'."
            ),
        };

    // Every entry's Arguments delegate is a direct cast-and-read (e.g. `f => ((Average)f).Arguments`): every
    // node type mapped here is `sealed`, so the cast can never fail for a function actually stored under that
    // dictionary key. Built once as a static field initializer (not per call) -- the switch this replaced ran
    // its isinst chain fresh on every invocation.
    private static readonly FrozenDictionary<
        Type,
        (string Name, Func<Function, Expression[]> Arguments)
    > UniformCalls = new Dictionary<Type, (string, Func<Function, Expression[]>)>
    {
        [typeof(Average)] = ("AVERAGE", static f => ((Average)f).Arguments),
        [typeof(Min)] = ("MIN", static f => ((Min)f).Arguments),
        [typeof(Max)] = ("MAX", static f => ((Max)f).Arguments),
        [typeof(Count)] = ("COUNT", static f => ((Count)f).Arguments),
        [typeof(If)] = ("IF", static f => ((If)f).Arguments),
        [typeof(And)] = ("AND", static f => ((And)f).Arguments),
        [typeof(Or)] = ("OR", static f => ((Or)f).Arguments),
        [typeof(Not)] = ("NOT", static f => ((Not)f).Arguments),
        [typeof(IfError)] = ("IFERROR", static f => ((IfError)f).Arguments),
        [typeof(IfNa)] = ("IFNA", static f => ((IfNa)f).Arguments),
        [typeof(Int)] = ("INT", static f => ((Int)f).Arguments),
        [typeof(Round)] = ("ROUND", static f => ((Round)f).Arguments),
        [typeof(RoundUp)] = ("ROUNDUP", static f => ((RoundUp)f).Arguments),
        [typeof(Abs)] = ("ABS", static f => ((Abs)f).Arguments),
        [typeof(IsNumber)] = ("ISNUMBER", static f => ((IsNumber)f).Arguments),
        [typeof(IsBlank)] = ("ISBLANK", static f => ((IsBlank)f).Arguments),
        [typeof(Upper)] = ("UPPER", static f => ((Upper)f).Arguments),
        [typeof(Lower)] = ("LOWER", static f => ((Lower)f).Arguments),
        [typeof(Trim)] = ("TRIM", static f => ((Trim)f).Arguments),
        [typeof(Len)] = ("LEN", static f => ((Len)f).Arguments),
        [typeof(Left)] = ("LEFT", static f => ((Left)f).Arguments),
        [typeof(Mid)] = ("MID", static f => ((Mid)f).Arguments),
        [typeof(Value)] = ("VALUE", static f => ((Value)f).Arguments),
        [typeof(Text)] = ("TEXT", static f => ((Text)f).Arguments),
        [typeof(Concat)] = ("CONCAT", static f => ((Concat)f).Arguments),
        [typeof(Concatenate)] = ("CONCATENATE", static f => ((Concatenate)f).Arguments),
        [typeof(TextJoin)] = ("TEXTJOIN", static f => ((TextJoin)f).Arguments),
        [typeof(CountA)] = ("COUNTA", static f => ((CountA)f).Arguments),
        [typeof(CountBlank)] = ("COUNTBLANK", static f => ((CountBlank)f).Arguments),
        [typeof(CountIf)] = ("COUNTIF", static f => ((CountIf)f).Arguments),
        [typeof(CountIfs)] = ("COUNTIFS", static f => ((CountIfs)f).Arguments),
        [typeof(SumIf)] = ("SUMIF", static f => ((SumIf)f).Arguments),
        [typeof(SumIfs)] = ("SUMIFS", static f => ((SumIfs)f).Arguments),
        [typeof(Rows)] = ("ROWS", static f => ((Rows)f).Arguments),
        [typeof(Row)] = ("ROW", static f => ((Row)f).Arguments),
        [typeof(Match)] = ("MATCH", static f => ((Match)f).Arguments),
        [typeof(Expressions.Lookup.Index)] = (
            "INDEX",
            static f => ((Expressions.Lookup.Index)f).Arguments
        ),
        [typeof(VLookup)] = ("VLOOKUP", static f => ((VLookup)f).Arguments),
        [typeof(XLookup)] = ("XLOOKUP", static f => ((XLookup)f).Arguments),
        [typeof(Offset)] = ("OFFSET", static f => ((Offset)f).Arguments),
        [typeof(Indirect)] = ("INDIRECT", static f => ((Indirect)f).Arguments),
        [typeof(Let)] = ("LET", static f => ((Let)f).Arguments),
        [typeof(SheetNumber)] = ("SHEET", static f => ((SheetNumber)f).Arguments),
        [typeof(Pmt)] = ("PMT", static f => ((Pmt)f).Arguments),
        [typeof(Pv)] = ("PV", static f => ((Pv)f).Arguments),
        [typeof(Fv)] = ("FV", static f => ((Fv)f).Arguments),
        [typeof(Nper)] = ("NPER", static f => ((Nper)f).Arguments),
        [typeof(Ipmt)] = ("IPMT", static f => ((Ipmt)f).Arguments),
        [typeof(Ppmt)] = ("PPMT", static f => ((Ppmt)f).Arguments),
        [typeof(Npv)] = ("NPV", static f => ((Npv)f).Arguments),
        [typeof(Rate)] = ("RATE", static f => ((Rate)f).Arguments),
        [typeof(Irr)] = ("IRR", static f => ((Irr)f).Arguments),
        [typeof(Sln)] = ("SLN", static f => ((Sln)f).Arguments),
        [typeof(Syd)] = ("SYD", static f => ((Syd)f).Arguments),
        [typeof(Db)] = ("DB", static f => ((Db)f).Arguments),
        [typeof(Ddb)] = ("DDB", static f => ((Ddb)f).Arguments),
        [typeof(Vdb)] = ("VDB", static f => ((Vdb)f).Arguments),
        [typeof(AmorLinc)] = ("AMORLINC", static f => ((AmorLinc)f).Arguments),
        [typeof(AmorDegrc)] = ("AMORDEGRC", static f => ((AmorDegrc)f).Arguments),
        [typeof(Effect)] = ("EFFECT", static f => ((Effect)f).Arguments),
        [typeof(Nominal)] = ("NOMINAL", static f => ((Nominal)f).Arguments),
        [typeof(Mirr)] = ("MIRR", static f => ((Mirr)f).Arguments),
        [typeof(Rri)] = ("RRI", static f => ((Rri)f).Arguments),
        [typeof(PDuration)] = ("PDURATION", static f => ((PDuration)f).Arguments),
        [typeof(ISPmt)] = ("ISPMT", static f => ((ISPmt)f).Arguments),
        [typeof(CumIPmt)] = ("CUMIPMT", static f => ((CumIPmt)f).Arguments),
        [typeof(CumPrinc)] = ("CUMPRINC", static f => ((CumPrinc)f).Arguments),
        [typeof(FvSchedule)] = ("FVSCHEDULE", static f => ((FvSchedule)f).Arguments),
        [typeof(DollarDe)] = ("DOLLARDE", static f => ((DollarDe)f).Arguments),
        [typeof(DollarFr)] = ("DOLLARFR", static f => ((DollarFr)f).Arguments),
        [typeof(XNpv)] = ("XNPV", static f => ((XNpv)f).Arguments),
        [typeof(XIrr)] = ("XIRR", static f => ((XIrr)f).Arguments),
        [typeof(AccrInt)] = ("ACCRINT", static f => ((AccrInt)f).Arguments),
        [typeof(AccrIntM)] = ("ACCRINTM", static f => ((AccrIntM)f).Arguments),
        [typeof(Disc)] = ("DISC", static f => ((Disc)f).Arguments),
        [typeof(IntRate)] = ("INTRATE", static f => ((IntRate)f).Arguments),
        [typeof(Received)] = ("RECEIVED", static f => ((Received)f).Arguments),
        [typeof(PriceDisc)] = ("PRICEDISC", static f => ((PriceDisc)f).Arguments),
        [typeof(PriceMat)] = ("PRICEMAT", static f => ((PriceMat)f).Arguments),
        [typeof(YieldDisc)] = ("YIELDDISC", static f => ((YieldDisc)f).Arguments),
        [typeof(YieldMat)] = ("YIELDMAT", static f => ((YieldMat)f).Arguments),
        [typeof(TBillEq)] = ("TBILLEQ", static f => ((TBillEq)f).Arguments),
        [typeof(TBillPrice)] = ("TBILLPRICE", static f => ((TBillPrice)f).Arguments),
        [typeof(TBillYield)] = ("TBILLYIELD", static f => ((TBillYield)f).Arguments),
        [typeof(CoupPcd)] = ("COUPPCD", static f => ((CoupPcd)f).Arguments),
        [typeof(CoupNcd)] = ("COUPNCD", static f => ((CoupNcd)f).Arguments),
        [typeof(CoupNum)] = ("COUPNUM", static f => ((CoupNum)f).Arguments),
        [typeof(CoupDays)] = ("COUPDAYS", static f => ((CoupDays)f).Arguments),
        [typeof(CoupDayBs)] = ("COUPDAYBS", static f => ((CoupDayBs)f).Arguments),
        [typeof(CoupDaysNc)] = ("COUPDAYSNC", static f => ((CoupDaysNc)f).Arguments),
        [typeof(Price)] = ("PRICE", static f => ((Price)f).Arguments),
        [typeof(Yield)] = ("YIELD", static f => ((Yield)f).Arguments),
        [typeof(Duration)] = ("DURATION", static f => ((Duration)f).Arguments),
        [typeof(MDuration)] = ("MDURATION", static f => ((MDuration)f).Arguments),
        [typeof(OddFPrice)] = ("ODDFPRICE", static f => ((OddFPrice)f).Arguments),
        [typeof(OddFYield)] = ("ODDFYIELD", static f => ((OddFYield)f).Arguments),
        [typeof(OddLPrice)] = ("ODDLPRICE", static f => ((OddLPrice)f).Arguments),
        [typeof(OddLYield)] = ("ODDLYIELD", static f => ((OddLYield)f).Arguments),
        [typeof(Sqrt)] = ("SQRT", static f => ((Sqrt)f).Arguments),
        [typeof(Power)] = ("POWER", static f => ((Power)f).Arguments),
        [typeof(Exp)] = ("EXP", static f => ((Exp)f).Arguments),
        [typeof(Ln)] = ("LN", static f => ((Ln)f).Arguments),
        [typeof(Log)] = ("LOG", static f => ((Log)f).Arguments),
        [typeof(Log10)] = ("LOG10", static f => ((Log10)f).Arguments),
        [typeof(SqrtPi)] = ("SQRTPI", static f => ((SqrtPi)f).Arguments),
        [typeof(RoundDown)] = ("ROUNDDOWN", static f => ((RoundDown)f).Arguments),
        [typeof(Trunc)] = ("TRUNC", static f => ((Trunc)f).Arguments),
        [typeof(MRound)] = ("MROUND", static f => ((MRound)f).Arguments),
        [typeof(Ceiling)] = ("CEILING", static f => ((Ceiling)f).Arguments),
        [typeof(CeilingMath)] = ("CEILING.MATH", static f => ((CeilingMath)f).Arguments),
        [typeof(CeilingPrecise)] = ("CEILING.PRECISE", static f => ((CeilingPrecise)f).Arguments),
        [typeof(IsoCeiling)] = ("ISO.CEILING", static f => ((IsoCeiling)f).Arguments),
        [typeof(Floor)] = ("FLOOR", static f => ((Floor)f).Arguments),
        [typeof(FloorMath)] = ("FLOOR.MATH", static f => ((FloorMath)f).Arguments),
        [typeof(FloorPrecise)] = ("FLOOR.PRECISE", static f => ((FloorPrecise)f).Arguments),
        [typeof(Even)] = ("EVEN", static f => ((Even)f).Arguments),
        [typeof(Odd)] = ("ODD", static f => ((Odd)f).Arguments),
        [typeof(Mod)] = ("MOD", static f => ((Mod)f).Arguments),
        [typeof(Quotient)] = ("QUOTIENT", static f => ((Quotient)f).Arguments),
        [typeof(Sign)] = ("SIGN", static f => ((Sign)f).Arguments),
        [typeof(Pi)] = ("PI", static f => ((Pi)f).Arguments),
        [typeof(Product)] = ("PRODUCT", static f => ((Product)f).Arguments),
        [typeof(SumSq)] = ("SUMSQ", static f => ((SumSq)f).Arguments),
        [typeof(Multinomial)] = ("MULTINOMIAL", static f => ((Multinomial)f).Arguments),
        [typeof(SeriesSum)] = ("SERIESSUM", static f => ((SeriesSum)f).Arguments),
        [typeof(Fact)] = ("FACT", static f => ((Fact)f).Arguments),
        [typeof(FactDouble)] = ("FACTDOUBLE", static f => ((FactDouble)f).Arguments),
        [typeof(Combin)] = ("COMBIN", static f => ((Combin)f).Arguments),
        [typeof(CombinA)] = ("COMBINA", static f => ((CombinA)f).Arguments),
        [typeof(Gcd)] = ("GCD", static f => ((Gcd)f).Arguments),
        [typeof(Lcm)] = ("LCM", static f => ((Lcm)f).Arguments),
        [typeof(Sin)] = ("SIN", static f => ((Sin)f).Arguments),
        [typeof(Cos)] = ("COS", static f => ((Cos)f).Arguments),
        [typeof(Tan)] = ("TAN", static f => ((Tan)f).Arguments),
        [typeof(Cot)] = ("COT", static f => ((Cot)f).Arguments),
        [typeof(Sec)] = ("SEC", static f => ((Sec)f).Arguments),
        [typeof(Csc)] = ("CSC", static f => ((Csc)f).Arguments),
        [typeof(Asin)] = ("ASIN", static f => ((Asin)f).Arguments),
        [typeof(Acos)] = ("ACOS", static f => ((Acos)f).Arguments),
        [typeof(Atan)] = ("ATAN", static f => ((Atan)f).Arguments),
        [typeof(Atan2)] = ("ATAN2", static f => ((Atan2)f).Arguments),
        [typeof(Acot)] = ("ACOT", static f => ((Acot)f).Arguments),
        [typeof(Sinh)] = ("SINH", static f => ((Sinh)f).Arguments),
        [typeof(Cosh)] = ("COSH", static f => ((Cosh)f).Arguments),
        [typeof(Tanh)] = ("TANH", static f => ((Tanh)f).Arguments),
        [typeof(Coth)] = ("COTH", static f => ((Coth)f).Arguments),
        [typeof(Sech)] = ("SECH", static f => ((Sech)f).Arguments),
        [typeof(Csch)] = ("CSCH", static f => ((Csch)f).Arguments),
        [typeof(Asinh)] = ("ASINH", static f => ((Asinh)f).Arguments),
        [typeof(Acosh)] = ("ACOSH", static f => ((Acosh)f).Arguments),
        [typeof(Atanh)] = ("ATANH", static f => ((Atanh)f).Arguments),
        [typeof(Acoth)] = ("ACOTH", static f => ((Acoth)f).Arguments),
        [typeof(Degrees)] = ("DEGREES", static f => ((Degrees)f).Arguments),
        [typeof(Radians)] = ("RADIANS", static f => ((Radians)f).Arguments),
        [typeof(Base)] = ("BASE", static f => ((Base)f).Arguments),
        [typeof(DecimalNumber)] = ("DECIMAL", static f => ((DecimalNumber)f).Arguments),
        [typeof(Roman)] = ("ROMAN", static f => ((Roman)f).Arguments),
        [typeof(Arabic)] = ("ARABIC", static f => ((Arabic)f).Arguments),
        [typeof(TrueFunction)] = ("TRUE", static f => ((TrueFunction)f).Arguments),
        [typeof(FalseFunction)] = ("FALSE", static f => ((FalseFunction)f).Arguments),
        [typeof(Xor)] = ("XOR", static f => ((Xor)f).Arguments),
        [typeof(Ifs)] = ("IFS", static f => ((Ifs)f).Arguments),
        [typeof(Switch)] = ("SWITCH", static f => ((Switch)f).Arguments),
        [typeof(Na)] = ("NA", static f => ((Na)f).Arguments),
        [typeof(IsError)] = ("ISERROR", static f => ((IsError)f).Arguments),
        [typeof(IsErr)] = ("ISERR", static f => ((IsErr)f).Arguments),
        [typeof(IsNa)] = ("ISNA", static f => ((IsNa)f).Arguments),
        [typeof(IsText)] = ("ISTEXT", static f => ((IsText)f).Arguments),
        [typeof(IsNonText)] = ("ISNONTEXT", static f => ((IsNonText)f).Arguments),
        [typeof(IsLogical)] = ("ISLOGICAL", static f => ((IsLogical)f).Arguments),
        [typeof(IsEven)] = ("ISEVEN", static f => ((IsEven)f).Arguments),
        [typeof(IsOdd)] = ("ISODD", static f => ((IsOdd)f).Arguments),
        [typeof(IsRef)] = ("ISREF", static f => ((IsRef)f).Arguments),
        [typeof(IsFormula)] = ("ISFORMULA", static f => ((IsFormula)f).Arguments),
        [typeof(N)] = ("N", static f => ((N)f).Arguments),
        [typeof(T)] = ("T", static f => ((T)f).Arguments),
        [typeof(TypeFunction)] = ("TYPE", static f => ((TypeFunction)f).Arguments),
        [typeof(ErrorType)] = ("ERROR.TYPE", static f => ((ErrorType)f).Arguments),
        [typeof(SheetsCount)] = ("SHEETS", static f => ((SheetsCount)f).Arguments),
        [typeof(Right)] = ("RIGHT", static f => ((Right)f).Arguments),
        [typeof(Find)] = ("FIND", static f => ((Find)f).Arguments),
        [typeof(Search)] = ("SEARCH", static f => ((Search)f).Arguments),
        [typeof(Replace)] = ("REPLACE", static f => ((Replace)f).Arguments),
        [typeof(Substitute)] = ("SUBSTITUTE", static f => ((Substitute)f).Arguments),
        [typeof(Rept)] = ("REPT", static f => ((Rept)f).Arguments),
        [typeof(Proper)] = ("PROPER", static f => ((Proper)f).Arguments),
        [typeof(Exact)] = ("EXACT", static f => ((Exact)f).Arguments),
        [typeof(CharFunction)] = ("CHAR", static f => ((CharFunction)f).Arguments),
        [typeof(Code)] = ("CODE", static f => ((Code)f).Arguments),
        [typeof(UniChar)] = ("UNICHAR", static f => ((UniChar)f).Arguments),
        [typeof(Unicode)] = ("UNICODE", static f => ((Unicode)f).Arguments),
        [typeof(Clean)] = ("CLEAN", static f => ((Clean)f).Arguments),
        [typeof(Fixed)] = ("FIXED", static f => ((Fixed)f).Arguments),
        [typeof(Dollar)] = ("DOLLAR", static f => ((Dollar)f).Arguments),
        [typeof(NumberValueFunction)] = (
            "NUMBERVALUE",
            static f => ((NumberValueFunction)f).Arguments
        ),
        [typeof(TextBefore)] = ("TEXTBEFORE", static f => ((TextBefore)f).Arguments),
        [typeof(TextAfter)] = ("TEXTAFTER", static f => ((TextAfter)f).Arguments),
        [typeof(ValueToText)] = ("VALUETOTEXT", static f => ((ValueToText)f).Arguments),
        [typeof(RegexTest)] = ("REGEXTEST", static f => ((RegexTest)f).Arguments),
        [typeof(RegexExtract)] = ("REGEXEXTRACT", static f => ((RegexExtract)f).Arguments),
        [typeof(RegexReplace)] = ("REGEXREPLACE", static f => ((RegexReplace)f).Arguments),
        [typeof(Choose)] = ("CHOOSE", static f => ((Choose)f).Arguments),
        [typeof(HLookup)] = ("HLOOKUP", static f => ((HLookup)f).Arguments),
        [typeof(Lookup)] = ("LOOKUP", static f => ((Lookup)f).Arguments),
        [typeof(Column)] = ("COLUMN", static f => ((Column)f).Arguments),
        [typeof(Columns)] = ("COLUMNS", static f => ((Columns)f).Arguments),
        [typeof(XMatch)] = ("XMATCH", static f => ((XMatch)f).Arguments),
        [typeof(Address)] = ("ADDRESS", static f => ((Address)f).Arguments),
        [typeof(Areas)] = ("AREAS", static f => ((Areas)f).Arguments),
        [typeof(FormulaText)] = ("FORMULATEXT", static f => ((FormulaText)f).Arguments),
        [typeof(AverageIf)] = ("AVERAGEIF", static f => ((AverageIf)f).Arguments),
        [typeof(AverageIfs)] = ("AVERAGEIFS", static f => ((AverageIfs)f).Arguments),
        [typeof(MaxIfs)] = ("MAXIFS", static f => ((MaxIfs)f).Arguments),
        [typeof(MinIfs)] = ("MINIFS", static f => ((MinIfs)f).Arguments),
        [typeof(AverageA)] = ("AVERAGEA", static f => ((AverageA)f).Arguments),
        [typeof(MaxA)] = ("MAXA", static f => ((MaxA)f).Arguments),
        [typeof(MinA)] = ("MINA", static f => ((MinA)f).Arguments),
        [typeof(SumProduct)] = ("SUMPRODUCT", static f => ((SumProduct)f).Arguments),
        [typeof(SumX2MY2)] = ("SUMX2MY2", static f => ((SumX2MY2)f).Arguments),
        [typeof(SumX2PY2)] = ("SUMX2PY2", static f => ((SumX2PY2)f).Arguments),
        [typeof(SumXMY2)] = ("SUMXMY2", static f => ((SumXMY2)f).Arguments),
        [typeof(Subtotal)] = ("SUBTOTAL", static f => ((Subtotal)f).Arguments),
        [typeof(Median)] = ("MEDIAN", static f => ((Median)f).Arguments),
        [typeof(ModeSngl)] = ("MODE.SNGL", static f => ((ModeSngl)f).Arguments),
        [typeof(Large)] = ("LARGE", static f => ((Large)f).Arguments),
        [typeof(Small)] = ("SMALL", static f => ((Small)f).Arguments),
        [typeof(RankEq)] = ("RANK.EQ", static f => ((RankEq)f).Arguments),
        [typeof(RankAvg)] = ("RANK.AVG", static f => ((RankAvg)f).Arguments),
        [typeof(PercentileInc)] = ("PERCENTILE.INC", static f => ((PercentileInc)f).Arguments),
        [typeof(PercentileExc)] = ("PERCENTILE.EXC", static f => ((PercentileExc)f).Arguments),
        [typeof(PercentRankInc)] = ("PERCENTRANK.INC", static f => ((PercentRankInc)f).Arguments),
        [typeof(PercentRankExc)] = ("PERCENTRANK.EXC", static f => ((PercentRankExc)f).Arguments),
        [typeof(QuartileInc)] = ("QUARTILE.INC", static f => ((QuartileInc)f).Arguments),
        [typeof(QuartileExc)] = ("QUARTILE.EXC", static f => ((QuartileExc)f).Arguments),
        [typeof(TrimMean)] = ("TRIMMEAN", static f => ((TrimMean)f).Arguments),
        [typeof(StDevS)] = ("STDEV.S", static f => ((StDevS)f).Arguments),
        [typeof(StDevP)] = ("STDEV.P", static f => ((StDevP)f).Arguments),
        [typeof(StDevA)] = ("STDEVA", static f => ((StDevA)f).Arguments),
        [typeof(StDevPA)] = ("STDEVPA", static f => ((StDevPA)f).Arguments),
        [typeof(VarS)] = ("VAR.S", static f => ((VarS)f).Arguments),
        [typeof(VarP)] = ("VAR.P", static f => ((VarP)f).Arguments),
        [typeof(VarA)] = ("VARA", static f => ((VarA)f).Arguments),
        [typeof(VarPA)] = ("VARPA", static f => ((VarPA)f).Arguments),
        [typeof(AveDev)] = ("AVEDEV", static f => ((AveDev)f).Arguments),
        [typeof(DevSq)] = ("DEVSQ", static f => ((DevSq)f).Arguments),
        [typeof(GeoMean)] = ("GEOMEAN", static f => ((GeoMean)f).Arguments),
        [typeof(HarMean)] = ("HARMEAN", static f => ((HarMean)f).Arguments),
        [typeof(Skew)] = ("SKEW", static f => ((Skew)f).Arguments),
        [typeof(SkewP)] = ("SKEW.P", static f => ((SkewP)f).Arguments),
        [typeof(Kurt)] = ("KURT", static f => ((Kurt)f).Arguments),
        [typeof(Standardize)] = ("STANDARDIZE", static f => ((Standardize)f).Arguments),
        [typeof(Correl)] = ("CORREL", static f => ((Correl)f).Arguments),
        [typeof(Pearson)] = ("PEARSON", static f => ((Pearson)f).Arguments),
        [typeof(CovarianceP)] = ("COVARIANCE.P", static f => ((CovarianceP)f).Arguments),
        [typeof(CovarianceS)] = ("COVARIANCE.S", static f => ((CovarianceS)f).Arguments),
        [typeof(Rsq)] = ("RSQ", static f => ((Rsq)f).Arguments),
        [typeof(Slope)] = ("SLOPE", static f => ((Slope)f).Arguments),
        [typeof(Intercept)] = ("INTERCEPT", static f => ((Intercept)f).Arguments),
        [typeof(Steyx)] = ("STEYX", static f => ((Steyx)f).Arguments),
        [typeof(ForecastLinear)] = ("FORECAST.LINEAR", static f => ((ForecastLinear)f).Arguments),
        [typeof(Fisher)] = ("FISHER", static f => ((Fisher)f).Arguments),
        [typeof(FisherInv)] = ("FISHERINV", static f => ((FisherInv)f).Arguments),
        [typeof(Phi)] = ("PHI", static f => ((Phi)f).Arguments),
        [typeof(Permut)] = ("PERMUT", static f => ((Permut)f).Arguments),
        [typeof(PermutationA)] = ("PERMUTATIONA", static f => ((PermutationA)f).Arguments),
        [typeof(Prob)] = ("PROB", static f => ((Prob)f).Arguments),
        // Compatibility aliases: the distinct node preserves the legacy spelling on un-parse.
        [typeof(Compat.Mode)] = ("MODE", static f => ((Compat.Mode)f).Arguments),
        [typeof(Compat.StDev)] = ("STDEV", static f => ((Compat.StDev)f).Arguments),
        [typeof(Compat.StDevP)] = ("STDEVP", static f => ((Compat.StDevP)f).Arguments),
        [typeof(Compat.Var)] = ("VAR", static f => ((Compat.Var)f).Arguments),
        [typeof(Compat.VarP)] = ("VARP", static f => ((Compat.VarP)f).Arguments),
        [typeof(Compat.Rank)] = ("RANK", static f => ((Compat.Rank)f).Arguments),
        [typeof(Compat.Percentile)] = ("PERCENTILE", static f => ((Compat.Percentile)f).Arguments),
        [typeof(Compat.PercentRank)] = (
            "PERCENTRANK",
            static f => ((Compat.PercentRank)f).Arguments
        ),
        [typeof(Compat.Quartile)] = ("QUARTILE", static f => ((Compat.Quartile)f).Arguments),
        [typeof(Compat.Covar)] = ("COVAR", static f => ((Compat.Covar)f).Arguments),
        [typeof(Compat.Forecast)] = ("FORECAST", static f => ((Compat.Forecast)f).Arguments),
        // Wave 5 — Date and time.
        [typeof(Date)] = ("DATE", static f => ((Date)f).Arguments),
        [typeof(Time)] = ("TIME", static f => ((Time)f).Arguments),
        [typeof(DateValue)] = ("DATEVALUE", static f => ((DateValue)f).Arguments),
        [typeof(TimeValue)] = ("TIMEVALUE", static f => ((TimeValue)f).Arguments),
        [typeof(Year)] = ("YEAR", static f => ((Year)f).Arguments),
        [typeof(Month)] = ("MONTH", static f => ((Month)f).Arguments),
        [typeof(Day)] = ("DAY", static f => ((Day)f).Arguments),
        [typeof(Hour)] = ("HOUR", static f => ((Hour)f).Arguments),
        [typeof(Minute)] = ("MINUTE", static f => ((Minute)f).Arguments),
        [typeof(Second)] = ("SECOND", static f => ((Second)f).Arguments),
        [typeof(Days)] = ("DAYS", static f => ((Days)f).Arguments),
        [typeof(Days360)] = ("DAYS360", static f => ((Days360)f).Arguments),
        [typeof(EDate)] = ("EDATE", static f => ((EDate)f).Arguments),
        [typeof(EoMonth)] = ("EOMONTH", static f => ((EoMonth)f).Arguments),
        [typeof(Weekday)] = ("WEEKDAY", static f => ((Weekday)f).Arguments),
        [typeof(WeekNum)] = ("WEEKNUM", static f => ((WeekNum)f).Arguments),
        [typeof(IsoWeekNum)] = ("ISOWEEKNUM", static f => ((IsoWeekNum)f).Arguments),
        [typeof(DateDif)] = ("DATEDIF", static f => ((DateDif)f).Arguments),
        [typeof(YearFrac)] = ("YEARFRAC", static f => ((YearFrac)f).Arguments),
        [typeof(NetworkDays)] = ("NETWORKDAYS", static f => ((NetworkDays)f).Arguments),
        [typeof(NetworkDaysIntl)] = (
            "NETWORKDAYS.INTL",
            static f => ((NetworkDaysIntl)f).Arguments
        ),
        [typeof(Workday)] = ("WORKDAY", static f => ((Workday)f).Arguments),
        [typeof(WorkdayIntl)] = ("WORKDAY.INTL", static f => ((WorkdayIntl)f).Arguments),
        // F1 — volatile clock functions.
        [typeof(Now)] = ("NOW", static f => ((Now)f).Arguments),
        [typeof(Today)] = ("TODAY", static f => ((Today)f).Arguments),
        // F1 — volatile RNG functions.
        [typeof(Rand)] = ("RAND", static f => ((Rand)f).Arguments),
        [typeof(RandBetween)] = ("RANDBETWEEN", static f => ((RandBetween)f).Arguments),
    }.ToFrozenDictionary();
}
