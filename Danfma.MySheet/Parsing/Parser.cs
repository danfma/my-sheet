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
            throw new ParseException(
                ParseErrorKind.UnexpectedToken,
                $"Unexpected token '{Current.Text}'",
                Current.Position,
                Current.Text
            );
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
            throw new ParseException(
                ParseErrorKind.NestingTooDeep,
                "Formula nesting is too deep",
                Current.Position,
                Current.Text
            );
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

            // Item 43 (sweep 31-35-43): an error literal (#REF!, #N/A, ...) is a primary expression —
            // the same ErrorValue node evaluation already produces at runtime (1/0, a ghost-sheet
            // reference), just reached from SYNTAX instead. Error.FromDisplay reuses the existing
            // singletons (ErrorValue.DivByZero, .Reference, ...) so this is the one bridge point, not a
            // parallel mapping.
            //
            // Round 2 (I-3): Excel writes a DELETED SHEET's own qualifier as #REF! (=Other!A1 becomes
            // =#REF!A1 once "Other" is deleted — measured on the oracle, Aspose.Cells 26.7.0/26.6.0,
            // PLAIN=CSE, 2026-09-14). So #REF! in THIS prefix position plays the same role a real
            // Identifier does before '!' in ParseQualifiedReference. A following cell-shaped endpoint,
            // row/column range, #REF!, parenthesized group, or spill marker belongs to that deleted
            // reference and is discarded. A boolean, name, or structured reference is instead parsed as
            // its own expression after the prefix; function calls and qualifiers remain invalid. See
            // ParseDeletedSheetQualifierContinuation for the bounded classification.
            case TokenType.Error:
                var errorValue = Error.FromDisplay(token.Text).ToErrorValue();

                if (token.Text == ErrorValue.Reference.ErrorCode)
                {
                    return ParseDeletedSheetQualifierContinuation(errorValue);
                }

                return errorValue;

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

            case TokenType.LBrace:
                return ParseArrayConstant(token);

            case TokenType.BracketedSpecifier:
                // A `[...]` with no table name before it — the three shapes S1 keeps permanently out of
                // scope, each with its own message so the error names what the user wrote instead of
                // "unexpected token". This arm is also why the tokenizer's reader is unconditional rather
                // than gated on the previous token: the reader makes one token out of every `[`, and the
                // parser is the single place that decides the shape is unsupported.
                throw new ParseException(
                    ParseErrorKind.UnsupportedStructuredReference,
                    UnsupportedPrefixBracket(token.Text),
                    token.Position,
                    token.Text
                );

            default:
                throw new ParseException(
                    ParseErrorKind.UnexpectedToken,
                    $"Unexpected token '{token.Text}'",
                    token.Position,
                    token.Text
                );
        }
    }

    private Expression ParseArrayConstant(Token opening)
    {
        var values = new List<Expression>();
        var columns = 0;
        var currentColumns = 0;

        while (true)
        {
            var sign = 1d;
            if (Current.Type == TokenType.Minus)
            {
                Advance();
                sign = -1d;
            }

            var token = Advance();
            Expression value = token.Type switch
            {
                TokenType.Number => new NumberValue(
                    sign * double.Parse(token.Text, CultureInfo.InvariantCulture)
                ),
                TokenType.String when sign > 0 => new StringValue(token.Text),
                TokenType.Identifier when sign > 0 && IsBoolean(token.Text, out var boolean) =>
                    new BooleanValue(boolean),
                TokenType.Error when sign > 0 => Error.FromDisplay(token.Text).ToErrorValue(),
                _ => throw InvalidArrayElement(token),
            };

            values.Add(value);
            currentColumns++;

            if (Current.Type == TokenType.Comma)
            {
                Advance();
                continue;
            }

            if (Current.Type == TokenType.Semicolon)
            {
                if (columns == 0)
                {
                    columns = currentColumns;
                }
                else if (columns != currentColumns)
                {
                    throw InvalidArrayElement(Current);
                }

                currentColumns = 0;
                Advance();
                continue;
            }

            if (Current.Type != TokenType.RBrace)
            {
                throw InvalidArrayElement(Current);
            }

            Advance();
            columns = columns == 0 ? currentColumns : columns;
            if (currentColumns != columns)
            {
                throw InvalidArrayElement(Current);
            }

            return new ArrayConstant(values.ToArray(), values.Count / columns, columns);
        }

        ParseException InvalidArrayElement(Token token) =>
            new(
                ParseErrorKind.UnexpectedToken,
                "Array constants accept only rectangular literal values",
                token.Position,
                token.Text
            );
    }

    // Classifies only the measured deleted-reference continuations. Reference-shaped endpoints are
    // discarded, while a boolean, name, or structured reference is parsed independently after the lost
    // qualifier. Operators and terminators remain outside the continuation.
    private Expression ParseDeletedSheetQualifierContinuation(ErrorValue errorValue)
    {
        if (Current.Type == TokenType.Bang)
        {
            Advance();
        }

        if (Current.Type == TokenType.Identifier && !IsDeletedReferenceEndpoint(Current))
        {
            if (tokens[_index + 1].Type is TokenType.LParen or TokenType.Bang)
            {
                return errorValue; // Leave the invalid function/qualified shape for ParseFormula to reject.
            }

            if (IsR1C1Reference(Current.Text))
            {
                return errorValue; // R1C1 syntax is invalid in this A1-only continuation.
            }

            var expression = ParseExpression(RangeBindingPower);

            // Preserve the independently parsed name instead of letting TryBuildOpenRange reinterpret its
            // letters as a column endpoint. A scalar defined name then produces #VALUE!, not #REF!.
            if (Current.Type == TokenType.Colon)
            {
                var colon = Advance();
                var right = ParseExpression(RangeBindingPower);
                RejectNonReferenceErrorEndpoint(expression, colon, "before");
                RejectNonReferenceErrorEndpoint(right, colon, "after");
                return new DynamicRange(expression, right);
            }

            return expression;
        }

        if (
            Current.Type == TokenType.Number
            && tokens[_index + 1].Type == TokenType.Colon
            && tokens[_index + 2].Type == TokenType.Number
        )
        {
            Advance();
            Advance();
            Advance();
            return errorValue;
        }

        if (!ConsumeReferenceEndpoint())
        {
            return errorValue; // bare #REF!, or an operator/terminator follows — nothing to absorb
        }

        if (Current.Type == TokenType.Colon)
        {
            Advance();
            if (!ConsumeReferenceEndpoint() && Current.Type == TokenType.Identifier)
            {
                ConsumeColumnEndpoint();
            }
        }

        if (Current.Type == TokenType.DeletedReferenceSpill)
        {
            Advance();
        }

        return errorValue;
    }

    private bool IsDeletedReferenceEndpoint(Token token) =>
        (
            IsExcelGridCellReference(token.Text)
            && tokens[_index + 1].Type != TokenType.BracketedSpecifier
        ) || (tokens[_index + 1].Type == TokenType.Colon && IsColumnEndpoint(token.Text));

    private static bool IsR1C1Reference(string text)
    {
        if (text.Length < 4 || text[0] is not ('R' or 'r'))
        {
            return false;
        }

        var index = 1;
        while (index < text.Length && char.IsAsciiDigit(text[index]))
        {
            index++;
        }

        if (index == 1 || index >= text.Length || text[index] is not ('C' or 'c'))
        {
            return false;
        }

        var columnStart = ++index;
        while (index < text.Length && char.IsAsciiDigit(text[index]))
        {
            index++;
        }

        return index == text.Length && index > columnStart;
    }

    private static bool IsColumnEndpoint(string text)
    {
        return CellAddress.TryParseColumn(text, out var column) && column <= ExcelMaxColumn;
    }

    // One endpoint of the absorbed run: a cell-shaped identifier, another #REF! token, or a
    // parenthesized expression (parsed — and so recursively resolved, e.g. a nested #REF!(...) — via
    // the ordinary grammar, then discarded). Reports whether it found one to consume.
    private bool ConsumeReferenceEndpoint()
    {
        if (Current.Type == TokenType.LParen)
        {
            ParsePrefix(Advance());
            return true;
        }

        if (
            (Current.Type == TokenType.Identifier && IsDeletedReferenceEndpoint(Current))
            || (Current.Type == TokenType.Error && Current.Text == ErrorValue.Reference.ErrorCode)
        )
        {
            Advance();
            return true;
        }

        return false;
    }

    private bool ConsumeColumnEndpoint()
    {
        if (IsColumnEndpoint(Current.Text))
        {
            Advance();
            return true;
        }

        return false;
    }

    // Names the out-of-scope prefix shape from its payload, which is all there is to go on: a leading '@' is
    // the current-row form written without a table name (`[@Valor]`, `[@]`), and an all-digits payload is the
    // external-workbook INDEX (`[1]Sheet1!A1`) — the only external spelling a saved file carries, since the
    // package stores the workbook name in an externalLink part and refers to it by number. A payload that is
    // neither cannot be pinned down: `[Valor]` is an implicit-table column reference and `[Book1.xlsx]` is a
    // by-name external reference, and nothing in the token tells them apart, so the third message names both
    // possibilities instead of asserting one. The implicit-table shapes — `[Valor]`, and equally the
    // current-row spellings `[@Valor]` and `[@]` — are the ones the oracle itself rejects outside a table,
    // all with the same message "Invalid table reference, formula should be in table when specifing no
    // table name" (Aspose.Cells 26.6.0, PLAIN entry, 2026-09-11, all three measured), because they are
    // valid only INSIDE their own table, which the parser has no cell context to check.
    private static string UnsupportedPrefixBracket(string text)
    {
        var payload = text[1..^1];

        if (payload.StartsWith('@'))
        {
            return $"This-row structured references are not supported: '{text}'";
        }

        if (payload.Length > 0 && IsAllDigits(payload))
        {
            return $"External-workbook references are not supported: '{text}'";
        }

        return $"A structured reference with no table name is not supported: '{text}' — qualify the column "
            + "with its table name; an external-workbook reference is out of scope either way";

        static bool IsAllDigits(string text)
        {
            foreach (var c in text)
            {
                if (!char.IsAsciiDigit(c))
                {
                    return false;
                }
            }

            return true;
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
        // Item 43 (sweep 31-35-43): measured (Aspose.Cells 26.7.0/26.6.0, both entry modes, 2026-09-14),
        // ONLY #REF! is accepted as a range endpoint — Excel's own broken-reference spelling. Every other
        // error literal there (A1:#N/A, #DIV/0!:A1, ...) is a PARSE error on the oracle ("Invalid data
        // after/before range sign ':'"), even though the same literal parses fine everywhere else. #REF!
        // itself needs no special case below: it matches none of TryEndpoint's arms, so it already falls
        // through to DynamicRange, which resolves to #REF! when an endpoint cannot be resolved — the
        // right answer for exactly this endpoint.
        RejectNonReferenceErrorEndpoint(left, colon, "before");

        var right = ParseExpression(RangeBindingPower);

        RejectNonReferenceErrorEndpoint(right, colon, "after");

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

    // Item 43: the asymmetric "before/after" message mirrors Aspose's own two distinct messages for the
    // two sides of ':'. #REF! passes through untouched (see ParseRange's comment on why it needs no arm).
    private static void RejectNonReferenceErrorEndpoint(
        Expression endpoint,
        Token colon,
        string side
    )
    {
        if (
            endpoint is ErrorValue { ErrorCode: var code }
            && code != ErrorValue.Reference.ErrorCode
        )
        {
            throw new ParseException(
                ParseErrorKind.ExpectedCellReference,
                $"Expected a cell reference {side} ':' but found '{code}'",
                colon.Position,
                code
            );
        }
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

            // An ABSOLUTE row endpoint ($1 in $1:$1, 1:$1, Data!$1:$1000). The Tokenizer lexes '$1' as an
            // identifier (a '$' starts a name so that $A$1 works), so it reaches here as a NameReference
            // rather than a NumberValue; without this arm the range degraded to a DynamicRange over two
            // unbound names (#REF!) and the sheet-qualified form threw (issue #8). Like the column arm
            // above, the '$' is a fill/copy marker only and is dropped.
            case NameReference name when CellAddress.TryParseRow(name.Name, out var parsedRow):
                column = null;
                row = parsedRow;
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

        // A structured (table) reference: this identifier is the table name and the `[...]` token carries the
        // whole specifier. Three ordering facts put the arm exactly here: the first two are forced by the
        // call shape itself, and only the third is measured.
        //  - AFTER the LParen check above, so `SUM(...)` still wins. Safe because FunctionRegistry.ByName is
        //    consulted at exactly one site, behind an Expect(LParen) in ParseFunctionCall, so `Name[` can
        //    never look like a call.
        //  - BEFORE the IsCellReference check below, which is MANDATORY: that check is unbounded (any
        //    letters-then-digits string) and answers TRUE for "Tabela1", "Table1", "Sales2024" and "ABC123" —
        //    Excel's own DEFAULT table names included — so an arm placed after it would build a CellReference
        //    and leave this token dangling. Jumping the IsBoolean check costs nothing and keeps `TRUE[Valor]`
        //    one coherent failure (a table Table.ValidateName forbids, so #NAME? at resolution) instead of a
        //    boolean with a bracket dangling after it.
        //  - The bracket must start exactly where the identifier's SOURCE SPAN ends. `Position + Text.Length`
        //    is that end for ReadIdentifier, whose Text is the raw slice; for ReadQuotedName it NEVER is,
        //    because Text is decoded while the span also carries the two quotes (and any doubled one) — so
        //    the test rejects a quoted table name unconditionally, not by luck, and rejects a whitespace gap
        //    for the ordinary reason. Both are rejected by the oracle too (Aspose.Cells 26.6.0, PLAIN entry,
        //    2026-09-11, over a Data!Tabela1 whose Valor column sums to 60): `SUM('Tabela1'[Valor])` is
        //    `Invalid "'"` and `SUM(Tabela1 [Valor])` is "Invalid table reference, formula should be in table
        //    when specifing no table name". Quoting is why this is a guard and not a nicety — ReadQuotedName
        //    delivers a DECODED identifier, so `'My Table'[Valor]` would build a node the writer renders back
        //    as the unparsable `My Table[Valor]`. Taking no arm leaves both with the error they already have.
        if (
            Current.Type == TokenType.BracketedSpecifier
            && Current.Position == token.Position + token.Text.Length
        )
        {
            return StructuredReferenceSyntax.Parse(token.Text, Advance());
        }

        if (IsBoolean(token.Text, out var boolean))
        {
            return new BooleanValue(boolean);
        }

        // A cell-shaped token becomes a cell ONLY when it names something in Excel's grid. IsCellReference
        // above is unbounded on purpose (see its doc), so without the grid check "Tabela1" — Excel's own
        // default pt-BR table name — would parse as the always-blank cell TABELA1 and every formula using a
        // bare table name would answer a silent 0 (measured on the oracle, Aspose.Cells 26.6.0, 2026-09-11,
        // PLAIN entry: SUM(Tabela1) 66, ROWS 3, COUNTA 9, ISREF TRUE — and 0/1/0/FALSE from the blank cell).
        // Phase 5 ruling R3, mechanism (a): route the identifier that is cell-SHAPED but NOT grid-bounded to
        // a NameReference instead — the same set both name validators (Table.ValidateName,
        // NamedReferences.IsValidName via IsExcelGridCellReference) reserve — and let the NAME path resolve
        // a registered table at EVALUATION time (NamedReferences.TryResolveRaw / NameReference.Evaluate).
        // The parser holds a sheet, no workbook, so it cannot know that "Tabela1" IS a table; what it can
        // know is that it is not a cell. The helper itself stays untouched, and so does its unbounded pin
        // (ExcelGridCellReferenceTests.IsCellReference_StaysUnbounded): the tightening lives here, at the
        // one classification that feeds BuildCellReference.
        if (IsCellReference(token.Text) && IsExcelGridCellReference(token.Text))
        {
            return BuildCellReference(token.Text, sheetName);
        }

        // A bare name: a LET-bound name resolved at evaluation time (#NAME? if unbound). Since R3 this is
        // also where a cell-shaped token outside Excel's grid lands — including a bare table name, which
        // name resolution answers with the table's data body when one is registered.
        return new NameReference(token.Text);
    }

    private Expression ParseQualifiedReference(string sheet)
    {
        // This recurses into itself (below) for chained cross-sheet range endpoints
        // (Sheet1!A1:Sheet2!B1:Sheet3!C1:...) WITHOUT going through ParseExpression, so it needs its own
        // depth check against the same counter — see MaxDepth's doc comment.
        if (++_depth > MaxDepth)
        {
            throw new ParseException(
                ParseErrorKind.NestingTooDeep,
                "Formula nesting is too deep",
                Current.Position,
                Current.Text
            );
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

        // Item 43 (sweep 31-35-43): `Sheet1!#REF!`, `'My Sheet'!#REF!`. Measured (Aspose.Cells 26.7.0/
        // 26.6.0, PLAIN=CSE, 2026-09-14): #REF! is the ONE error literal accepted after a sheet
        // qualifier — exactly the same exception a ':' range endpoint makes
        // (RejectNonReferenceErrorEndpoint below) — because the qualifier is irrelevant once the
        // reference is broken: `Sheet1!#REF!` evaluates to plain `#REF!`, same as the unqualified
        // literal. Every OTHER error literal there is "Invalid data before reference sign" on the
        // oracle (round 2, I-1: `Sheet1!#N/A`, `'My Sheet'!#DIV/0!`, `SUM(Sheet1!#VALUE!)` all throw).
        // Aspose's OWN formula-text round trip keeps the qualifier; MySheet's does not (ErrorValue
        // carries no sheet, and the value never depends on it) — a deliberate divergence, documented in
        // docs/workbook-and-expressions.md and docs/excel-interop.md (Scope and limitations), both
        // twins. Returning here immediately (rather than falling into the range/cell-reference checks
        // below) also means a trailing `:A1` is picked up by the ORDINARY Pratt loop once this call
        // returns, exactly like the unqualified `#REF!:A1` case.
        if (first.Type == TokenType.Error)
        {
            if (first.Text != ErrorValue.Reference.ErrorCode)
            {
                throw new ParseException(
                    ParseErrorKind.ExpectedCellReference,
                    "Expected a cell reference after '!'",
                    first.Position,
                    first.Text
                );
            }

            _depth--;

            return ErrorValue.Reference;
        }

        // `Data!Tabela1[Valor]`: a sheet qualifier on a table reference. Out of scope by S1 and a DELIBERATE
        // divergence, not parity — measured (Aspose.Cells 26.6.0, PLAIN entry, 2026-09-11) the oracle ACCEPTS
        // it, answers 60 and stores the formula back with the qualifier STRIPPED, from the table's own sheet,
        // from another sheet, and with the qualifier quoted ('Data'!Tabela1[Valor]) alike; table names are
        // workbook-scoped, so the qualifier carries no information. The guard's real argument is
        // diagnosability: without it the error depends on how the table happens to be SPELLED —
        // `Data!Tabela1[Valor]` would hit the !IsCellReference throw below (ExpectedCellReference), while
        // `Data!Sales2024[Valor]` would pass IsCellReference (unbounded, so TRUE for that spelling), build a
        // nonsense CellReference and only fail later as a dangling UnexpectedToken. One `if` gives both
        // spellings the same kind, at the table name's own position. It guards THIS endpoint only: a bracket
        // on the right-hand endpoint of a qualified range (`Data!A1:Tabela1[Valor]`) is consumed by the Colon
        // branch below and keeps the generic token error, since that shape is nothing a producer writes.
        if (Current.Type == TokenType.BracketedSpecifier)
        {
            throw new ParseException(
                ParseErrorKind.UnsupportedStructuredReference,
                "A sheet qualifier cannot be applied to a table reference",
                first.Position,
                first.Text
            );
        }

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
                        ParseErrorKind.ExpectedCellReference,
                        "Expected a cell reference after '!'",
                        second.Position,
                        second.Text
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

            // Item 43 round 2 (M-1): unlike the two `&&`-chained checks this replaced, building the
            // open range is no longer part of the success gate — an endpoint that resolves to
            // something recognisable but not a column/row/cell (ErrorValue.Reference, from the arm
            // added below) still counts as a VALID qualified range, just not an open one; it falls to
            // DynamicRange exactly like the unqualified ':' branch does (ParseRange, above) for the
            // same shape.
            if (
                TryEndpointToken(first, sheet, out var left)
                && TryEndpointToken(second, sheet, out var right)
            )
            {
                _depth--;

                return TryBuildOpenRange(left, right, sheet, out var range)
                    ? range
                    : new DynamicRange(left, right);
            }

            throw new ParseException(
                ParseErrorKind.ExpectedCellReference,
                "Expected a cell reference after '!'",
                first.Position,
                first.Text
            );
        }

        if (first.Type != TokenType.Identifier || !IsCellReference(first.Text))
        {
            throw new ParseException(
                ParseErrorKind.ExpectedCellReference,
                "Expected a cell reference after '!'",
                first.Position,
                first.Text
            );
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

        // Item 43 round 2 (M-1): #REF! is the one error literal accepted as a qualified range endpoint
        // too (Sheet1!A1:#REF!), mirroring the unqualified ':' rule (RejectNonReferenceErrorEndpoint)
        // and the '!' rule above (I-1) — only #REF! ever plays a reference role. TryBuildOpenRange never
        // resolves it to a column/row (TryEndpoint has no arm for ErrorValue), so this always falls
        // through to the DynamicRange fallback in the ':' branch above, which resolves to #REF! when an
        // endpoint cannot be resolved — the right answer for exactly this endpoint.
        if (token.Type == TokenType.Error && token.Text == ErrorValue.Reference.ErrorCode)
        {
            endpoint = ErrorValue.Reference;
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
                ParseErrorKind.InvalidArgumentCount,
                $"Function '{functionName}' does not accept {arguments.Length} argument(s)",
                name.Position,
                name.Text
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
                ParseErrorKind.ExpectedToken,
                $"Expected {type} but found '{Current.Text}'",
                Current.Position,
                Current.Text
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

    // Internal: the PARSER's own "looks like a cell reference" — the predicate that decides whether a bare
    // identifier becomes a CellReference (ParseIdentifier) or an endpoint of a qualified range. It is NO
    // LONGER any validator's rule: the defined-name validator (NamedReferences.IsValidName) was repointed at
    // IsExcelGridCellReference, which is the predicate the table-name validator (Table.ValidateName) already
    // used, so both name rules are bounded to Excel's grid while this one stays unbounded on purpose — see
    // IsExcelGridCellReference below for why the two must differ.
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

    // Excel's grid: XFD1048576 is the last cell.
    private const int ExcelMaxColumn = 16_384;
    private const int ExcelMaxRow = 1_048_576;

    // Internal: BOTH name validators' "looks like a cell reference" — Table.ValidateName's and, since the
    // defined-name repoint, NamedReferences.IsValidName's, because tables and defined names share one
    // namespace and must reserve the same names. It differs from IsCellReference
    // on purpose. IsCellReference is UNBOUNDED — any letters-then-digits string, because MySheet's grid has
    // no ceiling and ParseIdentifier depends on that — so it says Excel's own default table names, "Tabela1"
    // and "Table1", are cells. Excel's rule is that a table name may not be a reference INTO ITS GRID, so
    // this check is bounded twice: on the letter run (1-3 ASCII letters — four or more can never label a
    // column at or below XFD) and on the grid (column <= 16,384, row 1..1,048,576, at most 7 digits so the
    // row can never overflow). "T1" is a real cell and is rejected as a name; "Tabela1", "XFE1" and
    // "A1048577" are not cells and are accepted. Standalone: it must not reuse CellAddress.TryGetColumnRow,
    // which neither strips '$' nor guards its row accumulator.
    internal static bool IsExcelGridCellReference(string text)
    {
        text = StripDollars(text);

        var letters = 0;
        var column = 0;
        while (letters < text.Length && char.IsAsciiLetter(text[letters]))
        {
            column = column * 26 + (char.ToUpperInvariant(text[letters]) - 'A' + 1);
            letters++;
        }

        if (letters is 0 or > 3 || column > ExcelMaxColumn)
        {
            return false;
        }

        var digits = text.Length - letters;
        if (digits is 0 or > 7)
        {
            return false;
        }

        var row = 0;
        for (var i = letters; i < text.Length; i++)
        {
            if (!char.IsAsciiDigit(text[i]))
            {
                return false;
            }

            row = row * 10 + (text[i] - '0');
        }

        return row >= 1 && row <= ExcelMaxRow;
    }
}
