using System.Globalization;
using System.Text;
using Danfma.MySheet.Expressions;

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

    private static void Write(StringBuilder builder, Expression expression, string context, int minPrecedence)
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

            case UnionReference union:
                builder.Append('(');
                WriteList(builder, union.Areas, context);
                builder.Append(')');
                break;

            case NameReference name:
                builder.Append(name.Name);
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
    private static bool IsSimpleSheetName(string name)
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
    // FunctionCall keeps the name it was parsed/registered with.
    private static (string Name, Expression[] Arguments) Call(Function function) =>
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
            Expressions.Index f => ("INDEX", f.Arguments),
            VLookup f => ("VLOOKUP", f.Arguments),
            XLookup f => ("XLOOKUP", f.Arguments),
            Offset f => ("OFFSET", f.Arguments),
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
            _ => throw new NotSupportedException(
                $"No Excel function name registered for node '{function.GetType().Name}'."
            ),
        };
}
