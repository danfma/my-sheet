using System.Globalization;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Parsing;

/// <summary>
/// A Pratt (top-down operator precedence) parser turning a token stream into an
/// <see cref="Expression"/> tree. Cell references are resolved against <c>sheetName</c>.
/// </summary>
internal sealed class Parser(
    List<Token> tokens,
    string sheetName,
    int deltaRow = 0,
    int deltaColumn = 0,
    bool anchored = false
)
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

    // Guards against StackOverflowException (uncatchable, kills the process) on a pathological formula —
    // e.g. 10,000 nested parentheses from a hostile or corrupted .xlsx. 256 is generous: Excel itself caps
    // function nesting at 64 levels, so any legitimate formula sits far below this. Recursion nests through
    // two independent entry points that do not call each other: ParseExpression (parens, unary chains,
    // nested function arguments, ranges) and ParseQualifiedReference (chained cross-sheet range endpoints,
    // e.g. Sheet1!A1:Sheet2!B1:Sheet3!C1:...). Both increment/decrement the same counter, so depth is
    // tracked across the whole parse regardless of which path it grows through.
    private const int MaxDepth = 256;

    private int _depth;

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
        // No try/finally: a ParseException aborts the whole ParseFormula call, and the Parser is a
        // one-shot instance (a fresh one is constructed per parse — see ExpressionParser), so a counter
        // left incremented past an exception never leaks into a later parse.
        if (++_depth > MaxDepth)
        {
            throw new ParseException("Formula nesting is too deep", Current.Position);
        }

        var left = ParsePrefix(Advance());

        while (rightBindingPower < LeftBindingPower(Current.Type))
        {
            left = ParseInfix(Advance(), left);
        }

        _depth--;

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
                    Advance();
                    var second = ParseExpression(0);

                    // Two areas is the overwhelmingly common shape: build the array directly instead of a
                    // scratch List that only gets thrown away right after ToArray() — see ParseFunctionCall's
                    // single-argument fast path for the same reasoning.
                    if (Current.Type != TokenType.Comma)
                    {
                        Expect(TokenType.RParen);
                        return new UnionReference([inner, second]);
                    }

                    var areas = new List<Expression> { inner, second };

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

        // G3 spike (node-delta shared formulas): a bounded range whose both endpoints are anchored cell
        // refs — Data!A1:A2 inside a shared-formula master — gets its own anchored node so the WHOLE range
        // shifts per slave without a per-slave re-parse. This must be checked BEFORE TryBuildOpenRange:
        // that helper's own CellReference-only endpoint check would otherwise never fire for anchored
        // endpoints, and letting it "resolve" both endpoints numerically would silently freeze the range at
        // the master's literal position (losing the per-slave shift) instead of anchoring it correctly.
        if (
            left is AnchoredCellReference anchoredStart
            && right is AnchoredCellReference anchoredEnd
        )
        {
            return new AnchoredRangeReference(
                anchoredStart.Column,
                anchoredStart.Row,
                anchoredStart.ColumnAbsolute,
                anchoredStart.RowAbsolute,
                anchoredEnd.Column,
                anchoredEnd.Row,
                anchoredEnd.ColumnAbsolute,
                anchoredEnd.RowAbsolute,
                anchoredStart.SheetName
            );
        }

        // The ':' operator forces reference semantics: a letters-only endpoint is a COLUMN and an
        // integer endpoint a ROW, even when a defined name of the same spelling exists. This yields the
        // whole-column/row and one-sided open references (A:A, 1:5, A2:A, A:A10, A1:C).
        if (TryBuildOpenRange(left, right, SheetOf(left, right), out var open))
        {
            return open;
        }

        // Endpoints that are not statically resolvable (a reference-returning function like INDEX/OFFSET,
        // a parenthesised reference) become a DynamicRange, resolved at evaluation time.
        return new DynamicRange(left, right);
    }

    // The sheet a range lives on: the left endpoint's sheet when it carries one, else the right's, else
    // the parser's context sheet (a column-/row-only endpoint carries no sheet).
    private string SheetOf(Expression left, Expression right) =>
        left is CellReference lc ? lc.SheetName
        : left is AnchoredCellReference la ? la.SheetName
        : right is CellReference rc ? rc.SheetName
        : right is AnchoredCellReference ra ? ra.SheetName
        : sheetName;

    // Combines two endpoints (each contributing what it knows) into an open range: the left gives the
    // lower bounds, the right the upper bounds. When all four limits are known it degrades to a plain
    // RangeReference so the existing bounded path is never regressed.
    private static bool TryBuildOpenRange(
        Expression left,
        Expression right,
        string sheet,
        out Expression result
    )
    {
        if (
            TryEndpoint(left, out var colMin, out var rowMin)
            && TryEndpoint(right, out var colMax, out var rowMax)
        )
        {
            if (
                colMin is { } cMin
                && colMax is { } cMax
                && rowMin is { } rMin
                && rowMax is { } rMax
            )
            {
                result = new RangeReference(
                    new CellAddress(cMin, rMin).ToId(),
                    new CellAddress(cMax, rMax).ToId(),
                    sheet
                );
                return true;
            }

            result = OpenRangeReference.Create(colMin, colMax, rowMin, rowMax, sheet);
            return true;
        }

        result = null!;
        return false;
    }

    // Reads what a range endpoint knows: a cell gives (column,row); a letters-only name gives a column
    // (row open); a positive-integer number gives a row (column open). Anything else is not an endpoint.
    private static bool TryEndpoint(Expression expression, out int? column, out int? row)
    {
        switch (expression)
        {
            case CellReference cell:
                var address = CellAddress.Parse(cell.Id);
                column = address.Column;
                row = address.Row;
                return true;

            case NameReference name
                when CellAddress.TryParseColumn(name.Name, out var parsedColumn):
                column = parsedColumn;
                row = null;
                return true;

            case NumberValue { Value: var value }
                when value >= 1 && value <= int.MaxValue && value == Math.Floor(value):
                column = null;
                row = (int)value;
                return true;

            default:
                column = null;
                row = null;
                return false;
        }
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
            return BuildCellReference(token.Text, sheetName);
        }

        // A bare name: a LET-bound name resolved at evaluation time (#NAME? if unbound).
        return new NameReference(token.Text);
    }

    private Expression ParseQualifiedReference(string sheet)
    {
        // This recurses into itself (below) for chained cross-sheet range endpoints
        // (Sheet1!A1:Sheet2!B1:Sheet3!C1:...) WITHOUT going through ParseExpression, so it needs its own
        // depth check against the same counter — see MaxDepth's doc comment.
        if (++_depth > MaxDepth)
        {
            throw new ParseException("Formula nesting is too deep", Current.Position);
        }

        // The qualifier is a fresh tokenizer substring per formula, so N cross-sheet references to the same
        // sheet would otherwise each hold their own copy of the name (~24MB of duplicate "Data" strings at
        // K1 scale). Intern it here — the single point where a qualified SheetName enters the AST — so every
        // reference to a sheet shares ONE instance. string.Intern is exact (Ordinal), so the token's casing
        // is preserved verbatim (FormulaWriter echoes it, resolution is OrdinalIgnoreCase either way), and it
        // is the SAME pool MemoryPack's InternStringFormatter uses on Load, so a parsed name and a loaded name
        // converge on one instance. Sheet names are a tiny, bounded set, so process-lifetime interning is cheap.
        sheet = string.Intern(sheet);

        Expect(TokenType.Bang);
        var first = Advance();

        // A qualified range/open-range: Data!A1:B2, Data!A:A, Data!1:5, Data!A1:C. Both endpoints live on
        // the qualified sheet; the ':' forces reference semantics on a column-/row-only endpoint.
        if (Current.Type == TokenType.Colon)
        {
            Advance();
            var second = Advance();

            // The right endpoint may itself be sheet-qualified: Sheet1!A1:Sheet2!B2 (or Sheet1!A1:Sheet1!B2).
            // `second` was then the sheet name and a '!' follows. Parse it as its own qualified reference,
            // then span: same sheet → a plain RangeReference; different sheets → a DynamicRange, which the
            // cross-sheet guard resolves to #REF! (Excel parity) instead of silently using one sheet.
            if (Current.Type == TokenType.Bang)
            {
                if (
                    !TryEndpointToken(first, sheet, out var leftEndpoint)
                    || leftEndpoint is not CellReference leftCell
                    || ParseQualifiedReference(second.Text) is not CellReference rightCell
                )
                {
                    // Report at the RIGHT endpoint — that is the malformed side in this branch.
                    throw new ParseException(
                        "Expected a cell reference after '!'",
                        second.Position
                    );
                }

                _depth--;

                return string.Equals(
                    leftCell.SheetName,
                    rightCell.SheetName,
                    StringComparison.OrdinalIgnoreCase
                )
                    ? new RangeReference(leftCell.Id, rightCell.Id, leftCell.SheetName)
                    : new DynamicRange(leftCell, rightCell);
            }

            if (
                TryEndpointToken(first, sheet, out var left)
                && TryEndpointToken(second, sheet, out var right)
                && TryBuildOpenRange(left, right, sheet, out var range)
            )
            {
                _depth--;

                return range;
            }

            throw new ParseException("Expected a cell reference after '!'", first.Position);
        }

        if (first.Type != TokenType.Identifier || !IsCellReference(first.Text))
        {
            throw new ParseException("Expected a cell reference after '!'", first.Position);
        }

        _depth--;

        return BuildCellReference(first.Text, sheet);
    }

    // Turns a range-endpoint token into the expression a qualified open range is built from: a cell id
    // becomes a sheet-qualified CellReference, a letters-only identifier a column NameReference, an
    // integer a row NumberValue. Validation of "letters-only column" / "integer row" happens in
    // TryEndpoint when the range is built.
    private bool TryEndpointToken(Token token, string sheet, out Expression endpoint)
    {
        if (token.Type == TokenType.Identifier)
        {
            endpoint = IsCellReference(token.Text)
                ? BuildCellReference(token.Text, sheet)
                : new NameReference(token.Text);
            return true;
        }

        if (token.Type == TokenType.Number)
        {
            endpoint = new NumberValue(double.Parse(token.Text, CultureInfo.InvariantCulture));
            return true;
        }

        endpoint = null!;
        return false;
    }

    // G3 spike (node-delta shared formulas): builds the reference node for a cell-shaped token, either the
    // ordinary normalized/shifted CellReference (delta==0 or the legacy per-slave shift mode) or, in the
    // Parser's ANCHORED mode (used ONLY to parse a shared-formula group's master once — see
    // ExpressionParser.ParseAnchoredMasterBody), an AnchoredCellReference that keeps its ($-anchor,
    // column, row) components so a SharedFormulaSlave can shift it per-cell at evaluation time instead of
    // this Parser re-parsing the token per slave.
    private Expression BuildCellReference(string text, string sheet)
    {
        if (!anchored)
        {
            return new CellReference(NormalizeReference(text), sheet);
        }

        var (column, row, columnAbsolute, rowAbsolute) = ParseAnchorComponents(text);

        return new AnchoredCellReference(column, row, columnAbsolute, rowAbsolute, sheet);
    }

    // Decomposes a cell-reference token into (column, row, $-column, $-row) WITHOUT shifting — the anchored
    // twin of ShiftCellId, reusing its exact token shape (letters, optional interior '$', digits). Assumes
    // the token already passed IsCellReference (every call site does), so the malformed-shape guard branches
    // ShiftCellId carries for parity with the legacy textual shifter (extra '$', 4+ letter columns — shapes
    // the Tokenizer never actually produces) are intentionally not duplicated here; see the spike report.
    private static (
        int Column,
        int Row,
        bool ColumnAbsolute,
        bool RowAbsolute
    ) ParseAnchorComponents(string text)
    {
        var index = 0;
        var columnAbsolute = text[0] == '$';

        if (columnAbsolute)
        {
            index = 1;
        }

        var lettersStart = index;

        while (index < text.Length && char.IsLetter(text[index]))
        {
            index++;
        }

        var letterCount = index - lettersStart;
        var rowAbsolute = index < text.Length && text[index] == '$';

        if (rowAbsolute)
        {
            index++;
        }

        var digitsStart = index;
        var column = 0;

        for (var i = lettersStart; i < lettersStart + letterCount; i++)
        {
            column = column * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
        }

        var row = int.Parse(text.AsSpan(digitsStart), CultureInfo.InvariantCulture);

        return (column, row, columnAbsolute, rowAbsolute);
    }

    // Strips absolute markers ('$') and upper-cases, e.g. $A$1 -> A1. The reference identifies the same
    // cell regardless of '$'; absolute/relative only matters for Excel copy/fill, which we do not do.
    private static string NormalizeCellId(string text) => StripDollars(text).ToUpperInvariant();

    // Normalizes a cell-reference token, applying the shared-formula delta (if any) to its RELATIVE
    // components — the token text still carries the '$' markers the AST drops, so this is the single
    // point where "shift the copy like Excel" can happen on the token stream.
    private string NormalizeReference(string text) =>
        deltaRow == 0 && deltaColumn == 0 ? NormalizeCellId(text) : ShiftCellId(text);

    // Exact-parity port of the textual shifter's token shape: ^($?)([A-Za-z]{1,3})($?)([0-9]+)$.
    // Tokens outside that shape (extra '$', 4+ letters) were copied verbatim by the text rewrite, so
    // here they normalize WITHOUT shifting. '$'-anchored components do not move.
    private string ShiftCellId(string text)
    {
        var index = 0;
        var columnAbsolute = text[0] == '$';

        if (columnAbsolute)
        {
            index = 1;
        }

        var lettersStart = index;

        while (index < text.Length && char.IsLetter(text[index]))
        {
            index++;
        }

        var letterCount = index - lettersStart;

        if (letterCount is < 1 or > 3)
        {
            return NormalizeCellId(text);
        }

        var rowAbsolute = index < text.Length && text[index] == '$';

        if (rowAbsolute)
        {
            index++;
        }

        var digitsStart = index;

        while (index < text.Length && char.IsDigit(text[index]))
        {
            index++;
        }

        if (index != text.Length || index == digitsStart)
        {
            return NormalizeCellId(text);
        }

        var column = 0;

        for (var i = lettersStart; i < lettersStart + letterCount; i++)
        {
            column = column * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
        }

        if (!columnAbsolute)
        {
            column += deltaColumn;
        }

        // An anchored row keeps its original digits (leading zeros included — the text rewrite echoed
        // them verbatim); a relative one is re-rendered from the shifted number.
        var rowText = rowAbsolute
            ? text[digitsStart..]
            : (
                int.Parse(text.AsSpan(digitsStart), CultureInfo.InvariantCulture) + deltaRow
            ).ToString(CultureInfo.InvariantCulture);

        var letters = string.Empty;

        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }

        return letters + rowText;
    }

    private static string StripDollars(string text) =>
        text.Contains('$') ? text.Replace("$", string.Empty) : text;

    private Expression ParseFunctionCall(Token name)
    {
        Expect(TokenType.LParen);

        var arguments = ParseArgumentList();

        Expect(TokenType.RParen);

        var functionName = NormalizeFunctionName(name.Text);

        // Built-in: typed record with parse-time arity validation (a wrong count throws, like Excel
        // rejecting it at entry). Otherwise a generic call resolved at runtime against the workbook's
        // custom-function registry (#NAME? if never registered).
        if (!FunctionRegistry.ByName.TryGetValue(functionName, out var spec))
        {
            return new FunctionCall(functionName, arguments);
        }

        if (arguments.Length < spec.MinArgs || arguments.Length > spec.MaxArgs)
        {
            throw new ParseException(
                $"Function '{functionName}' does not accept {arguments.Length} argument(s)",
                name.Position
            );
        }

        return spec.Create(arguments);
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

    // Reads the comma-separated argument list up to (not including) the closing ')'. Zero and one argument
    // are, by far, the most common shapes (PI()/NOW()/TRUE(), SQRT(x)/ABS(x)/LEN(x)/…) and are built directly
    // as an array — no scratch List, so no throwaway backing array behind the final ToArray() copy. Two or
    // more arguments still route through a List: the eventual count is not known ahead of the comma scan, so
    // there is no allocation-free way to size the array up front (see the M1 write-up for why presizing the
    // List does not help: its own first growth already lands on the same capacity the default gives it).
    private Expression[] ParseArgumentList()
    {
        if (Current.Type == TokenType.RParen)
        {
            return [];
        }

        var first = ParseArgument();

        if (Current.Type != TokenType.Comma)
        {
            return [first];
        }

        var arguments = new List<Expression> { first };

        while (Current.Type == TokenType.Comma)
        {
            Advance();
            arguments.Add(ParseArgument());
        }

        return arguments.ToArray();
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
