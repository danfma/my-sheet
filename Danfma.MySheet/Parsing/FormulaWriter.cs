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

        Write(
            builder,
            expression,
            contextSheetName,
            minPrecedence: 0,
            depth: 0,
            deltaRow: 0,
            deltaColumn: 0
        );

        return builder.ToString();
    }

    // depth counts Write's own recursion; WriteBinary and WriteList are plain helpers invoked FROM here
    // (never independent recursion entry points), so they just pass the incremented depth through to their
    // own Write calls instead of tracking it themselves.
    //
    // deltaRow/deltaColumn (G3 spike, node-delta shared formulas): the ambient shift a SharedFormulaSlave
    // wrapper pushes for the duration of rendering its shared Master tree — zero everywhere else, so every
    // pre-spike call site (which always passes 0,0) renders byte-identical text. An AnchoredCellReference/
    // AnchoredRangeReference applies it (unless its own $-anchor flag is set) to render the SAME text the
    // legacy fully-expanded slave tree would have rendered — the writer intentionally never re-emits '$'
    // (neither did the pre-spike CellReference/RangeReference path: NormalizeCellId always strips it), so
    // parity here means "the same bare A1-style text", not "preserves anchors".
    private static void Write(
        StringBuilder builder,
        Expression expression,
        string context,
        int minPrecedence,
        int depth,
        int deltaRow,
        int deltaColumn
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
                WriteList(builder, union.Areas, context, depth + 1, deltaRow, deltaColumn);
                builder.Append(')');
                break;

            case NameReference name:
                builder.Append(name.Name);
                break;

            case DynamicRange dyn:
                Write(
                    builder,
                    dyn.Start,
                    context,
                    AtomPrecedence,
                    depth + 1,
                    deltaRow,
                    deltaColumn
                );
                builder.Append(':');
                Write(builder, dyn.End, context, AtomPrecedence, depth + 1, deltaRow, deltaColumn);
                break;

            case UnaryOperation { Operator: UnaryOperator.Percent } percent:
                Write(
                    builder,
                    percent.Operand,
                    context,
                    PercentPrecedence,
                    depth + 1,
                    deltaRow,
                    deltaColumn
                );
                builder.Append('%');
                break;

            case UnaryOperation prefix:
                builder.Append(prefix.Operator == UnaryOperator.Negate ? '-' : '+');
                Write(
                    builder,
                    prefix.Operand,
                    context,
                    PrefixPrecedence,
                    depth + 1,
                    deltaRow,
                    deltaColumn
                );
                break;

            case BinaryOperation binary:
                WriteBinary(builder, binary, context, depth + 1, deltaRow, deltaColumn);
                break;

            case Function function:
                var (functionName, arguments) = Call(function);
                builder.Append(functionName).Append('(');
                WriteList(builder, arguments, context, depth + 1, deltaRow, deltaColumn);
                builder.Append(')');
                break;

            // --- G3 spike (node-delta shared formulas) ---------------------------------------------------
            // These three only ever appear when a shared-formula slave was loaded through the anchored/
            // delta path (WorksheetStreamLoader.ExpandSlave). Rendering applies the AMBIENT delta (0 outside
            // a SharedFormulaSlave, that slave's own (DeltaRow, DeltaColumn) while rendering its Master) —
            // the same bare "A1"-style text the pre-spike fully-expanded slave tree would have rendered (no
            // '$' is ever re-emitted, matching the existing CellReference/RangeReference rule above).
            case AnchoredCellReference anchoredCell:
                WriteSheetQualifier(builder, anchoredCell.SheetName, context);
                builder.Append(
                    new CellAddress(
                        anchoredCell.ColumnAbsolute
                            ? anchoredCell.Column
                            : anchoredCell.Column + deltaColumn,
                        anchoredCell.RowAbsolute ? anchoredCell.Row : anchoredCell.Row + deltaRow
                    ).ToId()
                );
                break;

            case AnchoredRangeReference anchoredRange:
                WriteSheetQualifier(builder, anchoredRange.SheetName, context);
                builder
                    .Append(
                        new CellAddress(
                            anchoredRange.StartColumnAbsolute
                                ? anchoredRange.StartColumn
                                : anchoredRange.StartColumn + deltaColumn,
                            anchoredRange.StartRowAbsolute
                                ? anchoredRange.StartRow
                                : anchoredRange.StartRow + deltaRow
                        ).ToId()
                    )
                    .Append(':')
                    .Append(
                        new CellAddress(
                            anchoredRange.EndColumnAbsolute
                                ? anchoredRange.EndColumn
                                : anchoredRange.EndColumn + deltaColumn,
                            anchoredRange.EndRowAbsolute
                                ? anchoredRange.EndRow
                                : anchoredRange.EndRow + deltaRow
                        ).ToId()
                    );
                break;

            case SharedFormulaSlave slave:
                // A slave wrapper is only ever the WHOLE tree of a cell (never nested inside a larger
                // expression it was parsed as part of), so it always renders at the root — minPrecedence 0,
                // a fresh ambient delta from this slave's own (DeltaRow, DeltaColumn), replacing (not
                // composing with) whatever delta was ambient here (always 0 — nesting one slave inside
                // another never happens).
                Write(
                    builder,
                    slave.Master,
                    context,
                    0,
                    depth + 1,
                    slave.DeltaRow,
                    slave.DeltaColumn
                );
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
        int depth,
        int deltaRow,
        int deltaColumn
    )
    {
        var precedence = Precedence(binary);

        // '^' is right-associative (like the Parser); every other operator is left-associative, so the
        // tighter minimum goes on the opposite side to force parentheses only where re-parsing needs them.
        var rightAssociative = binary.Operator == BinaryOperator.Power;

        Write(
            builder,
            binary.Left,
            context,
            rightAssociative ? precedence + 1 : precedence,
            depth,
            deltaRow,
            deltaColumn
        );
        builder.Append(Token(binary.Operator));
        Write(
            builder,
            binary.Right,
            context,
            rightAssociative ? precedence : precedence + 1,
            depth,
            deltaRow,
            deltaColumn
        );
    }

    private static void WriteList(
        StringBuilder builder,
        Expression[] items,
        string context,
        int depth,
        int deltaRow,
        int deltaColumn
    )
    {
        for (var i = 0; i < items.Length; i++)
        {
            if (i > 0)
            {
                builder.Append(',');
            }

            Write(builder, items[i], context, minPrecedence: 0, depth, deltaRow, deltaColumn);
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

    // Node-type -> Excel-function-name resolution, backed by FunctionRegistry.ByType (the mirror of the
    // Parser's FunctionRegistry.ByName -- see that type for the single source-of-truth entry per built-in).
    // A custom FunctionCall keeps the name it was parsed/registered with; it has no registry entry (it is not
    // a built-in, it is the runtime fallback for an unrecognized name), so it stays a one-line special case
    // here. Every built-in -- Sum included, via its custom GetArguments accessor in the registry -- resolves
    // uniformly through ByType: one hash lookup + one delegate invocation, regardless of which function it
    // is (Call runs on EVERY cell of a Formulas-mode export and every FORMULATEXT call, so this intentionally
    // never falls back to an isinst chain). Internal so the defined-name validator can reuse it as the single
    // source of a function node's argument list when walking for unqualified refs.
    internal static (string Name, Expression[] Arguments) Call(Function function) =>
        function switch
        {
            FunctionCall f => (f.Name, f.Arguments),
            _ when FunctionRegistry.ByType.TryGetValue(function.GetType(), out var entry) => (
                entry.Name,
                entry.GetArguments(function)
            ),
            _ => throw new NotSupportedException(
                $"No Excel function name registered for node '{function.GetType().Name}'."
            ),
        };
}
