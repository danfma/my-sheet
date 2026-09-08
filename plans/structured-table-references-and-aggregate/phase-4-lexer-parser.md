# Phase 4: Bracket lexing and the structured-reference grammar

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `lexer-parser`. Design dependencies: `table-model-registry`, `reference-semantics`. Adversarial verifier verdict: **needs-revision** (2 blockers, 1 major, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

The tokenizer gets ONE new fat token (`TokenType.BracketedSpecifier`) carrying the whole `[...]` suffix
verbatim, read by an unconditional, context-free `ReadBracketedSpecifier` that mirrors `ReadQuotedName`
(Tokenizer.cs:159-191) but does a balanced-bracket scan honouring Excel's `'`-escape instead of decoding. Thin
tokens (LBracket/Hash/Comma) are not merely worse, they are IMPOSSIBLE: whitespace is discarded at
Tokenizer.cs:32 and `Parser` receives only `List<Token>` with no source text (Parser.cs:10-16), so
`Tabela1[Total (USD)]` could never be reassembled from `Identifier LParen Identifier RParen`. All grammar
(item split, `'`-decode, specifier map) and its inverse (escape, canonical rendering) live in ONE new internal
helper, `Parsing/StructuredReferenceSyntax.cs`, so encode and decode cannot drift — the same reason
`IsSimpleSheetName` sits next to `WriteSheetQualifier` (FormulaWriter.cs:344-382). The parser emits ONE node,
`TableReference(string TableName, string? ColumnName, TableArea Area) : Reference` (one union tag 322, one
writer arm, one future DependencyExtractor arm); the arm goes into `ParseIdentifier` BETWEEN the LParen check
(:316-319) and the IsBoolean check (:321), because I measured that `Parser.IsCellReference("Tabela1")`,
`("Table1")` and `("Sales2024")` all return TRUE — so any arm placed after :326 would turn the DEFAULT Excel
table name into a cell reference. `IsFullyAnchored` ACCEPTS the node (same reasoning as NameReference at
AnchoredFormulaSupport.cs:37: no position component once `[@Col]` is out of scope), which keeps the commonest
real shape — one table formula shared down thousands of rows — on the fast shared-master path. Every grammar
rule below was validated by running a prototype over 44 inputs in /tmp/probe-lexer-parser; that run is what
found two writer bugs (`Table[#Data]` rendering as `Table[]`, and a space-padded column name losing its spaces
on re-parse) before any repo code was written.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Item 6's item-list grammar ('inner content taken verbatim then DecodeName'd', THEN 'collect 0-2 leading specifiers then at most one column') combined with item 12's writer, and item 15's canonical row `Tabela1[['#OfItems]]`.
      *Measured evidence:* I implemented item 6 + item 12 literally (/tmp/verify-lexer-
      parser/proto/Program.cs) and ran it. Result row: `Tabela1['#OfItems]` -> column "#OfItems" -> written
      `Tabela1[['#OfItems]]` -> **rewrite REJECTED as Invalid: unknown specifier '#OfItems'**, and the direct
      input `Tabela1[['#OfItems]]` -> `!! Invalid: unknown specifier '#OfItems'`. Cause: after DecodeName,
      `'#OfItems` becomes `#OfItems`, whose first char is `#`, so the item is classified as a specifier and
      ParseSpecifier rejects it. The spec author's OWN prototype does the opposite and is correct: /tmp/probe-
      lexer-parser/Program.cs:100 tests `inner.Length > 0 && inner[0] == '#'` on the RAW inner text, and only
      calls `Decode(inner, ...)` at line 111 for the column branch — its run prints `Tabela1[['#OfItems]] ->
      Column = #OfItems ... stable? yes`. So the written spec and the prototype it claims to codify disagree,
      and the written one breaks the advertised round-trip invariant for every column name beginning with `#`.
      *Correction:* State the ordering rule explicitly in item 6: an item is a SPECIFIER iff its RAW
      (undecoded) inner text starts with an unescaped `#`; DecodeName runs only after classification, on the
      column branch. Verification #6 (FormulaWriterTests total 71, failed 0) fails as written until this is
      fixed. Alternative, Excel-truer fix: have item 12's writer emit the single-bracket form
      `Tabela1['#OfItems]` for a `#`-leading column — that is literally the MS doc's own example
      (`=DeptSalesFYSummary['#OfItems]`, verified on the cited page), and it re-parses through the payload-
      first-char-`'` branch with no ordering rule needed.
- [ ] **B2.** Item 17 ('Leave the 5 genuinely table-specific tests (:129, :162, :193, :213, :243) failing for the excel-loader phase to rewrite') and verification #8 ('total: 14. EXPECT failed: 5 while this phase stands alone — exactly [those 5 names]').
      *Measured evidence:* tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs:128-297: every one of those 5
      tests asserts only `warnings.Count`, `warnings[0].Kind == UnparsableFormula`, `.Subject == "B4"`,
      `.Detail` non-empty, and the cached-value fallback (`999.0`, `Blank`, `42.0`/`43.0`). None asserts the
      formula text or the failure reason. Danfma.MySheet.Excel/WorksheetStreamLoader.cs:479-528 is a SINGLE
      `try { ParseFormulaBody ... } catch (ParseException)` -> `DegradeToCachedLiteral`, so
      `SUM(Tabela1[@Valor])` (item 6: leading `@` -> UnsupportedStructuredReference, still a ParseException)
      degrades identically. All 5 therefore PASS after the swap; the run reports failed: 0, not failed: 5.
      Worse: after the swap NOTHING in either suite asserts anything about `SUM(Tabela1[Valor])` on load —
      which is exactly the phase's own risk #1 (it stops degrading and, with no registry, yields an error
      where Excel's cached 999 used to appear). The verification suite goes fully green while the named
      regression ships. Also the item's arithmetic is wrong: 11, not 12, of the 14 tests use the constant
      (:134, :168, :196, :217, :218, :255, :308, :340, :375, :398, :428, :538 across 11 test methods; :100,
      :460, :589 do not).
      *Correction:* Correct verification #8 to `failed: 0`, and add a test that actually detects the change:
      keep one `SUM(Tabela1[Valor])` fixture and assert the NEW observable behaviour of this phase in
      isolation (parses, no UnparsableFormula warning, cell evaluates to #REF!) so the loss of the degradation
      path is pinned rather than hidden. Note in the commit body that this assertion is expected to be
      rewritten once table-model-registry lands.

## Major corrections

- [ ] **M1.** The risks section enumerates the downstream default arms that need a TableReference case (DependencyExtractor.cs:216, ReferenceGuard.cs:96-98) but omits ArrayEvaluation entirely; ArrayEvaluation is never mentioned anywhere in the spec.
      *Evidence:* Danfma.MySheet/Expressions/ArrayEvaluation.cs:123-201 (`Probe`) and :420-481
      (`TryBuildOperand`) admit exactly `RangeReference`, `AnchoredRangeReference`, `Row{[RangeReference]}`,
      `Row{[AnchoredRangeReference]}`, `BinaryOperation`, `If`, refuse `OpenRangeReference`, and end with
      `default: operand = new ScalarOperand(expression.Evaluate(context)); return true;` (:478-480). A
      `TableReference` therefore broadcasts a single `ComputedValue.Reference` as a 1x1 scalar, so
      `SUM(Tabela1[Valor]*2)`, `SUM(IF(Tabela1[Col]>0,1,0))` and `SMALL(IF(Tabela1[Col]<>"",...),1)` silently
      return the wrong answer instead of Excel's elementwise result. The file's own AnchoredRangeReference arm
      comment (:131-136) names this exact failure: 'it fell through to the `default` case below (an opaque
      scalar ... always yields #VALUE!), silently breaking the mini-CSE idiom'. So the precedent for needing
      an arm is documented in the file the spec did not consult.
      *Correction:* Add ArrayEvaluation (`Probe` + `TryBuildOperand`, resolving via `TryResolveReference` to a
      concrete RangeReference and delegating to `BuildRange`, mirroring the AnchoredRangeReference arms) to
      the risks/handoff list alongside DependencyExtractor and ReferenceGuard, and require a
      MiniCseConsumerTests case. Also state that `NamedReferences.CaptureValue` (NamedReferences.cs:59-69)
      needs NO arm — I verified its `_ => expression.Evaluate(context)` default yields
      `ComputedValue.Reference` for TableReference precisely because item 3's Evaluate mirrors DynamicRange —
      so a future contributor does not 'fix' it.

## Implementation items

- [ ] **1.** Add `BracketedSpecifier` to `TokenType` in Danfma.MySheet/Parsing/Token.cs, after `Bang` (:25) and before `EndOfInput` (:26), with a one-line comment: "The whole `[...]` suffix of a structured (table) reference, Text = the raw bracketed text INCLUDING the brackets, undecoded."
      *Files:* `Danfma.MySheet/Parsing/Token.cs`
      *Why:* TokenType is internal (Token.cs:3), never serialized, and I grepped for numeric use of it: the
      only `(int)`-cast on an enum in the engine is WorkdayFunctions.cs:19 on DayOfWeek, so inserting a member
      cannot shift anything meaningful. TokenizerTests.cs:7-11 golden strings join `t.Type` names, not
      ordinals, so no existing golden string changes.
- [ ] **2.** Append three members to the PUBLIC `ParseErrorKind` enum in Danfma.MySheet/Parsing/ParseException.cs, after `NestingTooDeep` (:38): `UnterminatedBracketedReference` ("a structured reference whose closing ']' is missing: =Tabela1[Valor"), `InvalidStructuredReference` ("balanced brackets, content Excel's own grammar rejects: Tabela1[], Tabela1[#Bogus], a column followed by a specifier"), `UnsupportedStructuredReference` ("valid Excel that MySheet does not model: Tabela1[@Valor], Tabela1[#This Row], implicit-table [Valor], column span Tabela1[[A]:[C]], multi-column list, external-workbook [1]Sheet1!A1, sheet-qualified Data!Tabela1[Valor]").
      *Files:* `Danfma.MySheet/Parsing/ParseException.cs`
      *Why:* Three kinds, not one: the enum already distinguishes UnterminatedString (:20) from
      UnterminatedQuotedName (:22-23) rather than lumping them, and the Invalid/Unsupported split is the
      difference an integrator needs — "your file is malformed" vs "MySheet has not implemented this yet",
      which is also what the docs' limitation list has to enumerate. Append at the end because the enum is
      public (ParseException.cs:7-39). Tie-break rule for the implementation: when it is uncertain whether
      Excel accepts a shape, classify Unsupported, never Invalid — we do not accuse the user's file of being
      wrong.
- [ ] **3.** Create Danfma.MySheet/Expressions/TableReference.cs with `public enum TableArea { Data, All, Headers, Totals, HeadersAndData, DataAndTotals }` and `[MemoryPackable] public sealed partial record TableReference(string TableName, string? ColumnName, TableArea Area) : Reference`, overriding `Evaluate` and `TryResolveReference` by delegating to the registry seam owned by the table-model-registry/reference-semantics phase (`TableReferences.TryResolve(this, context, out var range)`; `Evaluate` => `ComputedValue.Reference(range)` when it resolves else `ComputedValue.Error(Error.Ref)`, exactly DynamicRange.cs:42-56).
      *Files:* `Danfma.MySheet/Expressions/TableReference.cs`
      *Why:* One node type, not two: `Table[#All]` and `Table[Col]` differ only in whether ColumnName is null,
      so a nullable column + an area enum buys one union tag, one FormulaWriter arm, one DependencyExtractor
      arm and one ReferenceGuard arm instead of two of each. `TableArea` is declared in the same file as its
      record, matching BinaryOperation.cs:5 and UnaryOperation.cs:5. `Data` is member 0 so `default` is the
      Excel default area. NO SheetName member: table names are workbook-scoped in Excel, exactly like
      Workbook.DefinedNames (Workbook.cs:123-124) and NameReference (NameReference.cs:12), and the sheet comes
      from the registry. Deriving from `Reference` (Reference.cs:3-10) is what lets the resolved concrete
      RangeReference light up the six range-only consumers (VLookup.cs:20-21, LookupFunctions.cs:71-72,
      Index.cs:33-34/:178-179, Offset.cs:84) for free.
- [ ] **4.** Add `[MemoryPackUnion(322, typeof(TableReference))]` immediately after the `[MemoryPackUnion(321, typeof(SharedFormulaSlave))]` line in Danfma.MySheet/Expressions/Expression.cs:352, with a comment naming the feature, and fix the now-stale policy comment at Expression.cs:14-16 ("Add new tags at 319+") to say 323+.
      *Files:* `Danfma.MySheet/Expressions/Expression.cs`
      *Why:* I verified the union tags are contiguous 0..321 with no gaps (Expression.cs:18 through :352), so
      322 is the next free tag. The comment at :14-16 already warns it drifts; leaving it saying 319 after
      this lands would mis-direct the next contributor into reusing a live tag.
- [ ] **5.** Create Danfma.MySheet/Parsing/StructuredReferenceSyntax.cs — one `internal static class` holding BOTH halves of the syntax: (a) `internal static bool TryFindClosingBracket(string text, int openIndex, out int closeIndex)` — scan from `text[openIndex]=='['`, `depth=1`, on `'` skip TWO chars (the escape and its escapee), on `[` depth++, on `]` depth-- and return when 0, run off the end => false; (b) `internal static Expression Parse(string tableName, Token token)` implementing the grammar in the next item; (c) `internal static void Write(StringBuilder builder, TableReference reference)` implementing the rendering two items down; (d) privates `SplitItems`, `DecodeName`, `ParseSpecifier`, `IsSimpleColumnName`, `EscapeName`.
      *Files:* `Danfma.MySheet/Parsing/StructuredReferenceSyntax.cs`
      *Why:* A ~120-line grammar does not belong inside the already-813-line Parser.cs, and splitting decode
      (parser) from encode (writer) across two files is exactly how a round-trip bug is born — the repo keeps
      the sheet-name pair together for the same reason (FormulaWriter.IsSimpleSheetName :366-382 next to
      WriteSheetQualifier :344-361). Precedent for a small dedicated Parsing/ helper:
      AnchoredFormulaSupport.cs. ONE scanner used by three callers (the tokenizer's reader, the item splitter,
      and the per-item bracket check) guarantees the tokenizer can never accept a payload the splitter mis-
      splits.
- [ ] **6.** Implement `StructuredReferenceSyntax.Parse(tableName, token)`: payload = `token.Text[1..^1].Trim()`; empty => InvalidStructuredReference("empty specifier"); first char `@` => UnsupportedStructuredReference; first char `#` => single specifier via ParseSpecifier; first char `[` => item-list form (SplitItems at top-level `,`, rejecting a top-level `:` as UnsupportedStructuredReference "column span"; each item trimmed must satisfy `TryFindClosingBracket(item, 0, out var c) && c == item.Length - 1`, else Invalid; inner content taken verbatim then DecodeName'd); anything else => the whole payload is one column name, DecodeName'd, Area = Data. Collect 0-2 leading specifiers then at most one column, in that order: 0 specifiers => Data; 1 => that area; `[#Headers],[#Data]` => HeadersAndData; `[#Data],[#Totals]` => DataAndTotals; any other pair or 3+ specifiers => Invalid; a specifier after a column => Invalid; a second column => Unsupported. ParseSpecifier maps `#All/#Data/#Headers/#Totals` OrdinalIgnoreCase, `#This Row` => Unsupported, anything else => Invalid. DecodeName: `'` escapes the next char literally (so `''`->`'`, `']`->`]`), a trailing `'` with nothing after => Invalid("dangling escape"). Every throw uses `token.Position + offset` so the reported position lands inside the brackets.
      *Files:* `Danfma.MySheet/Parsing/StructuredReferenceSyntax.cs`
      *Why:* Escape set and specifier set are Microsoft's documented ones, fetched today from the "Using
      structured references with Excel tables" page (f5ed2452-2337-4f71-bed3-c8ae6d2b276e): escape characters
      are exactly "Left bracket ([), Right bracket (]), Pound sign (#), Single quotation mark ('), At sign
      (@)", specifiers are exactly #All/#Data/#Headers/#Totals/#This Row/@. Up to two specifiers plus a column
      is not an invention: that page's own example is `=DeptSales[[#Headers], [#Data], [% Commission]]`, which
      also proves decorative spaces after the separator are legal — hence the Trim of each item OUTSIDE its
      brackets, and why interior content is taken verbatim (`Table[[ Col ]]` is the escape hatch for a
      genuinely space-padded header while `Table[ Col ]` trims). Decoding cannot happen in the tokenizer the
      way ReadQuotedName does it (Tokenizer.cs:170-178): decoding `'[` before splitting would corrupt the item
      split. I ran this exact algorithm over 44 inputs; all 39 accepted forms are round-trip stable and all 5
      rejected classes report the intended kind.
- [ ] **7.** Add `ReadBracketedSpecifier(int start)` to Danfma.MySheet/Parsing/Tokenizer.cs next to ReadQuotedName (:159-191), and dispatch to it from `NextToken` with `if (c == '[') return ReadBracketedSpecifier(start);` inserted after the `'` case (:57-60) and before `return ReadOperator(start)` (:62). Body: `if (!StructuredReferenceSyntax.TryFindClosingBracket(text, _position, out var close)) throw new ParseException(ParseErrorKind.UnterminatedBracketedReference, "Unterminated bracketed reference", start, text[start..]);` then `var token = new Token(TokenType.BracketedSpecifier, text[start..(close + 1)], start); _position = close + 1; return token;`. Comment it as the deliberate divergence from ReadQuotedName: Text is RAW (brackets included, undecoded) because decoding is per-item and happens after the split.
      *Files:* `Danfma.MySheet/Parsing/Tokenizer.cs`
      *Why:* Unconditional (no `private TokenType _previous` gate) keeps the class's documented context-free
      property (Tokenizer.cs:9-11, one `_position` field) AND gives better diagnostics: with the reader firing
      anywhere, `=SUM([Valor])` and `=[1]Sheet1!A1` arrive at the parser as one token that ParsePrefix can
      reject with a message naming the shape, whereas a `_previous`-gated reader would leave them as
      UnexpectedCharacter '[' — and S1 requires out-of-scope shapes to fail CLEARLY. I measured the only cost:
      `=[1]Sheet1!A1` moves from UnexpectedCharacter@0 to UnsupportedStructuredReference@0 (both
      ParseException, both degrade identically in the loader), while the quoted external form `='[1]Data'!A1`
      is untouched because `'` is dispatched first at :57 (probe confirms it still parses to a CellReference
      on a sheet named "[1]Data"). Text includes the brackets so ParseException.Token keeps its documented
      meaning — "the offending token's text as it appeared in the formula" (ParseException.cs:71-73), the same
      convention ReadString/ReadQuotedName use for their unterminated payloads — and so generic messages from
      Expect (:783-795) and ParseFormula (:47-57) print `[Valor]`, not a bare `Valor`. A single
      `text[start..(close+1)]` slice also means no StringBuilder at all, unlike ReadQuotedName.
- [ ] **8.** In Danfma.MySheet/Parsing/Parser.cs `ParseIdentifier` (:308), insert the new arm between the LParen check (ends :319) and the IsBoolean check (:321): `if (Current.Type == TokenType.BracketedSpecifier) { return StructuredReferenceSyntax.Parse(token.Text, Advance()); }` with a comment naming the two ordering hazards.
      *Files:* `Danfma.MySheet/Parsing/Parser.cs`
      *Why:* Before :326 is MANDATORY and measured, not hypothetical: `Parser.IsCellReference` returns TRUE
      for "Tabela1", "Table1", "Sales2024", "Vendas2024" and "ABC123" (probe over the real internal method),
      because it only requires letters-then-digits (Parser.cs:788-812) — so with the arm after :326, Excel's
      DEFAULT table name would build a CellReference and leave the bracket token dangling. Before :321 costs
      nothing and covers a table named TRUE/FALSE. AFTER :316 (LParen) is required so `SUM(...)` still wins,
      and it is safe because FunctionRegistry.ByName is consulted at exactly one site behind an Expect(LParen)
      (Parser.cs:629), so `Name[` can never look like a call. Consequence accepted deliberately: `A1[Col]`
      also parses to a TableReference on a table Excel could never name; the parser is context-free (it holds
      only a sheet name, Parser.cs:10-16) so this must be a resolution-time #NAME?/#REF!, not a parse error.
- [ ] **9.** In Danfma.MySheet/Parsing/Parser.cs `ParsePrefix` (:88), add `case TokenType.BracketedSpecifier:` immediately before the `default:` at :143, throwing `ParseErrorKind.UnsupportedStructuredReference` with a message that discriminates the three prefix-position shapes by their first payload char — `[@...]`/`[@]` "this-row structured references are not supported", `[<digits>]` "external-workbook references are not supported", otherwise "implicit-table structured references ([Column]) are not supported: qualify the reference with the table name" — passing `token.Position` and `token.Text`.
      *Files:* `Danfma.MySheet/Parsing/Parser.cs`
      *Why:* This single arm is where S1's three explicitly-out-of-scope shapes land with a clear error:
      implicit-table `[Valor]`, `[@Valor]` when written without a table name, and the external-workbook
      `[1]Sheet1!A1` form. Without it they fall to the `default:` at :143-149 as a generic UnexpectedToken,
      which S1 forbids. It is also the reason the unconditional tokenizer reader beats a `_previous`-gated
      one.
- [ ] **10.** In Danfma.MySheet/Parsing/Parser.cs `ParseQualifiedReference`, insert immediately after `var first = Advance();` (:360) and before the Colon branch (:364): `if (Current.Type == TokenType.BracketedSpecifier) throw new ParseException(ParseErrorKind.UnsupportedStructuredReference, "A sheet qualifier cannot be applied to a table reference", first.Position, first.Text);`
      *Files:* `Danfma.MySheet/Parsing/Parser.cs`
      *Why:* `Data!Tabela1[Valor]` is explicitly out of scope and Excel agrees: table names are workbook-
      scoped, so Excel rejects a sheet qualifier on one (it accepts only a workbook qualifier,
      `[Book1.xlsx]Table1[Col]`). Without this guard the error depends on the table name's spelling —
      `Data!Tabela1[Valor]` hits the `!IsCellReference` throw at :420-428 with ExpectedCellReference, but
      `Data!Sales2024[Valor]` passes IsCellReference (measured TRUE), builds a nonsense CellReference and only
      fails later as UnexpectedToken. One `if` makes both spellings report the same clear kind at the table
      name's position.
- [ ] **11.** In Danfma.MySheet/Parsing/AnchoredFormulaSupport.cs `IsFullyAnchored`, add `TableReference` to the NameReference arm at :37 (`NameReference or TableReference => true`) and extend that arm's comment: a structured reference resolves by table name against the workbook-scoped registry and carries no position component, since `[@Col]` — the only position-dependent structured form — is out of scope.
      *Files:* `Danfma.MySheet/Parsing/AnchoredFormulaSupport.cs`
      *Why:* Accept, not reject. Correctness is the same argument the file already makes for NameReference at
      :34-37 ("identical for every slave, exactly as it is identical for every cell of an ordinary formula
      referencing the same name"). Rejecting would be safe but would put the WORST case on the slow path: a
      table formula column shared down thousands of rows is precisely the shape that would then fall back to
      per-slave token re-parsing via ExpressionParser.ParseSharedFormulaBody (ExpressionParser.cs:57-66)
      instead of one shared master tree. Note in the comment that both Parser modes (anchored and legacy-
      delta) build the identical node because a BracketedSpecifier token carries no shiftable component.
- [ ] **12.** In Danfma.MySheet/Parsing/FormulaWriter.cs `Write`, add `case TableReference table: StructuredReferenceSyntax.Write(builder, table); break;` next to the NameReference arm (:136-138) and before the `default:` NotSupportedException (:252-255). Implement `StructuredReferenceSyntax.Write`: table name bare when `FormulaWriter.IsSimpleSheetName`-shaped (letters/digits/_/. and no leading digit) else `'`-quoted with `''` doubling; then column-only => `Table[col]` when `IsSimpleColumnName(col)` (every char `char.IsLetterOrDigit`) else `Table[[escaped]]`; exactly one specifier and no column => the single-bracket form `Table[#All]`; otherwise the item-list form `Table[[#Headers],[#Data],[col]]` with each item bracketed and the column escaped. `EscapeName` prefixes `'` to every `[ ] # ' @`. `#Data` is implicit ONLY when a column is present, so `TableReference(T, null, Data)` renders `T[#Data]`. Ignore the ambient deltaRow/deltaColumn (a structured reference has no position component).
      *Files:* `Danfma.MySheet/Parsing/FormulaWriter.cs`, `Danfma.MySheet/Parsing/StructuredReferenceSyntax.cs`
      *Why:* The arm is mandatory, not optional: without it FORMULATEXT and every Formulas-mode export throw
      NotSupportedException at FormulaWriter.cs:252-255. `Precedence` (:384-404) already returns
      AtomPrecedence via `_ =>`, so the atom is never parenthesized — no change needed there. Two rules came
      out of the prototype run as BUGS I had reasoned wrong: rendering Area=Data with no column produced `T[]`
      (fixed by making #Data implicit only alongside a column), and rendering a non-simple column bare
      produced `T[ Col ]`, which re-parsed to "Col" and broke round-trip (fixed by always using the double-
      bracket form for a non-alphanumeric name). The IsSimpleColumnName predicate is deliberately STRICTER
      than IsSimpleSheetName (:366-382): `_` and `.` are in Microsoft's 33-character "requires the extra
      brackets" list, and `char.IsLetterOrDigit` is a one-line superset of that list that can never under-
      bracket, which is the safe direction since Excel always accepts extra brackets
      (`=DeptSalesFYSummary[[Total $ Amount]]` is the doc's own canonical form). ROUND-TRIP INVARIANT: for
      every canonical spelling, parse->write is byte-identity; for every non-canonical spelling, write is
      idempotent from the second pass and the re-parsed tree is MemoryPack-identical.
- [ ] **13.** Extend tests/Danfma.MySheet.Tests/Parsing/TokenizerTests.cs with: `StructuredReference_IsOneToken` (`Shape("Tabela1[Valor]")` == "Identifier BracketedSpecifier EndOfInput"; also the composite `Tabela1[[#Data],[% Comissao]]` and `Tabela1[Total (USD)]` — same three-token shape, proving `(`/`)`/`,`/space inside brackets are not lexed as operators), `BracketedSpecifier_KeepsTheRawTextIncludingBrackets` (`Single(...).Text` == "[Total (USD)]", Position == 7), `UnterminatedBracket_Throws` (`Tokenizer.Tokenize("Tabela1[Valor")` throws ParseException), and `StructuredReference_TerminatesAtTheMatchingBracket` (`Shape("T[[#Data],[a]]:T2[b]")` == "Identifier BracketedSpecifier Colon Identifier BracketedSpecifier EndOfInput").
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/TokenizerTests.cs`
      *Why:* No existing test in either suite contains a `[` inside a formula string (grepped `"=[^"]*\[`
      across tests/ — zero hits), so nothing in this file breaks; `InvalidCharacter_Throws` at :74-78 uses
      `#`, which stays UnexpectedCharacter because no top-level `#` case is added. The last case is the
      measured evidence (probe: `T[[#Data],[a]]:T2[b]` closes at index 13) that `Table[a]:Table2[b]` reaches
      ParseRange (:173) and becomes a DynamicRange via the :217 fallback — a range spanning two table columns,
      working for free.
- [ ] **14.** Add tests/Danfma.MySheet.Tests/Parsing/StructuredReferenceTests.cs with the repo's local `Calc`-less parse harness (`ExpressionParser.Parse("=" + f, new Sheet { Name = "Sheet1" })`) asserting the NODE for each accepted S1 form: `Tabela1[Valor]` => (Tabela1, Valor, Data); `[#All]/[#Data]/[#Headers]/[#Totals]` and lowercase `[#totals]` => (null column, the area); `[[#Data],[Valor]]`, `[[#All],[Valor]]`, `[[#Headers],[#Data]]`, `[[#Data],[#Totals]]`, `[[#Headers],[#Data],[% Comissao]]`; `[[Valor]]` => (Valor, Data); `[Sales Amount]`, `[[Total $ Amount]]`, `[Total (USD)]`, `[a,b]`, `[a:b]`, `[Unit_Price]`, `['#OfItems]` => "#OfItems", `['[bracket']]` => "[bracket]", `[a'']` => "a'", `[ Col ]` => "Col", `[[ Col ]]` => " Col ", `[[a@b]]` => "a@b"; plus `'My Table'[Valor]` (quoted table name) and `SUM(Tabela1[Valor])` inside a call.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/StructuredReferenceTests.cs`
      *Why:* These 24 cases are exactly the accepted rows of the prototype run, so the expected values are
      measured, not guessed. `'My Table'[Valor]` works because ReadQuotedName emits a decoded Identifier token
      (Tokenizer.cs:178) that the new ParseIdentifier arm then sees. The naming and location follow the
      convention that parse/behaviour tests live in tests/Danfma.MySheet.Tests/Parsing/<Family>Tests.cs.
- [ ] **15.** Extend tests/Danfma.MySheet.Tests/Parsing/ParseExceptionTests.cs (currently 14 tests, all green) with one test per new kind using its existing `Throws(formula)` helper: `UnterminatedBracketedReference` ("=Tabela1[Valor" => Token "[Valor", Position 7), `InvalidStructuredReference` x3 ("=Tabela1[]", "=Tabela1[#Bogus]", "=Tabela1[[Valor],[#Data]]"), `UnsupportedStructuredReference` x6 ("=Tabela1[@Valor]", "=Tabela1[#This Row]", "=SUM([Valor])", "=Tabela1[[Col1]:[Col3]]", "=Tabela1[[A],[B]]", "=Data!Tabela1[Valor]") and one for the external form ("=[1]Sheet1!A1" => UnsupportedStructuredReference, Position 0).
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/ParseExceptionTests.cs`
      *Why:* This file is the repo's contract test for structured diagnosability (its class comment ties it to
      issue #8) and is the only place `UnexpectedCharacter` is asserted (:34) — pinning the external-workbook
      form here documents the deliberate kind change from UnexpectedCharacter to
      UnsupportedStructuredReference. Position 7 for the unterminated case and the Token payloads are taken
      from the probe's measured offsets.
- [ ] **16.** Extend tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs (currently 49 tests): add to the `RoundTrips_CanonicalText` [Arguments] list (:18-61) the canonical spellings `Tabela1[Valor]`, `Tabela1[#All]`, `Tabela1[#Data]`, `Tabela1[#Headers]`, `Tabela1[#Totals]`, `Tabela1[[#All],[Valor]]`, `Tabela1[[#Headers],[#Data]]`, `Tabela1[[#Data],[#Totals]]`, `Tabela1[[#Headers],[#Data],[% Comissao]]`, `Tabela1[[Sales Amount]]`, `Tabela1[[Total (USD)]]`, `Tabela1[['#OfItems]]`, `Tabela1[[ Col ]]`, `SUM(Tabela1[Valor])`, `'My Table'[Valor]`; and to `NormalizesEquivalentText` (:69-81) the pairs ("Tabela1[[Valor]]","Tabela1[Valor]"), ("Tabela1[[#Data],[Valor]]","Tabela1[Valor]"), ("Tabela1[[#Data], [Valor]]","Tabela1[Valor]"), ("Tabela1[ Col ]","Tabela1[Col]"), ("Tabela1[#totals]","Tabela1[#Totals]"), ("Tabela1[Sales Amount]","Tabela1[[Sales Amount]]"), ("Tabela1[Rev#1]","Tabela1[[Rev'#1]]").
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`
      *Why:* The two existing harnesses already express exactly the invariant this phase needs and nothing new
      has to be invented: `RoundTrips_CanonicalText` (:62-66) is byte-identity, and `NormalizesEquivalentText`
      (:73-82) additionally compares MemoryPack bytes of the re-parsed tree (`Structure(Parse(written))`),
      which simultaneously proves the new union tag 322 serializes. Every expected value here is a measured
      row of the prototype run.
- [ ] **17.** Change the `StructuredFormula` constant in tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs:21 from "SUM(Tabela1[Valor])" to "SUM(Tabela1[@Valor])" so the ~6 warning-machinery tests that merely need SOME unparsable formula (`Load_UnparsableFormula_WithEmptyCachedValue_LeavesTheCellBlank` :303, `_WithNonNumericCachedValue` :335, `_WithCachedStringResult` :367, `_WithCachedErrorResult` :392, `_WithOutOfRangeSharedStringIndex` :418, `Load_RejectedMasterReusingAnotherGroupsIndex` :511) keep working, and update the class comment at :8-18 which currently states "MySheet has no table model and its tokenizer has no `[`". Leave the 5 genuinely table-specific tests (:129, :162, :193, :213, :243) failing for the excel-loader phase to rewrite, and say so in the commit body.
      *Files:* `tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* 12 of this class's 14 tests use `SUM(Tabela1[Valor])` as their unparsable-formula vehicle, so
      this phase breaks them all; most are not about tables at all. `[@Valor]` is the right replacement
      because S1 keeps the this-row form out of scope permanently (so the vehicle cannot rot the way `{1;2}`
      could once array literals or the dynamic-array phase land) and it exercises the new
      UnsupportedStructuredReference path through the loader's existing negative parse cache
      (WorksheetStreamLoader.cs:94, :460) and warning sites (:500-515, :555-566).
- [ ] **18.** OPTIONAL, NEEDS A GO/NO-GO DECISION (see openQuestions): tighten `Parser.IsCellReference` (Parser.cs:788-812) to Excel's actual token shape — optional `$`, 1..3 ASCII letters (`is >= 'A' and <= 'Z' or >= 'a' and <= 'z'`, not `char.IsLetter`), optional `$`, then 1+ digits — and add tests: `IsCellReference` false for "Tabela1"/"Table1"/"Sales2024"/"表1" and still true for "Q1"/"ABC123"/"$A$1"/"A01"; `=Tabela1` parses to NameReference; `Workbook.DefineName("Table1", new NumberValue(1))` no longer throws.
      *Files:* `Danfma.MySheet/Parsing/Parser.cs`, `tests/Danfma.MySheet.Tests/Parsing/ExpressionParserTests.cs`, `tests/Danfma.MySheet.Tests/Parsing/NamedRangeTests.cs`
      *Why:* Measured defect, not speculation: `Workbook.DefineName("Tabela1", …)` and `DefineName("Table1",
      …)` THROW today ("must not look like a cell reference") because NamedReferences.cs:192 delegates to this
      method — Excel's own default table name is unusable as a name — and `=Tabela1` silently evaluates to 0
      as a cell reference to column "TABELA". Excel's rule is that a name is reserved only when it IS a real
      grid address (columns are at most 3 ASCII letters), and the codebase ALREADY encodes that shape in one
      place: ShiftCellId rejects `letterCount is < 1 or > 3` at Parser.cs:552-556, so IsCellReference
      contradicts its neighbour. `char.IsLetter` also makes the CJK default table name `表1` a cell reference.
      Blast radius measured: the only 4+letter-plus-digit tokens in any test formula string are sheet names
      (Sheet1/Sheet8/Sparse2 — the qualifier is never run through IsCellReference) and function names
      (ATAN2/DAYS360/SUMXMY2 — consumed by the LParen branch at :316 first), and
      `DefineName_CellShapedName_Throws` (NamedRangeTests.cs:224-232) uses "A1", which stays reserved. It is
      OPTIONAL for S1 (the new arm sits before :326, so `Table1[Col]` parses either way) but REQUIRED for
      Excel-parity on bare `=Table1` and for S2's table-name validation; ship it as its own `fix(parser)`
      commit with the doc note, or explicitly document that bare table names are unsupported.

## Verification Plan

- [ ] `dotnet csharpier check .`
      → expected: Exit code 0, no "Formatted" or diff output. CSharpier 1.3.0 is pinned in dotnet-tools.json
      and this is the pre-commit gate.
- [ ] `dotnet build Danfma.MySheet.slnx -c Release`
      → expected: "Build succeeded" with 0 errors and 0 warnings. Fails with CS0117/CS8509 if TableReference
      or the TableArea members are missing an arm somewhere.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TokenizerTests/*"`
      → expected: "Test run summary: Passed!" with total: 13, failed: 0 (9 existing + 4 new). In particular
      InvalidCharacter_Throws must still pass, proving a bare `#` outside brackets is still
      ParseErrorKind.UnexpectedCharacter.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/StructuredReferenceTests/*"`
      → expected: "Passed!", failed: 0, total: 24 or more — one assertion per accepted S1 form.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/ParseExceptionTests/*"`
      → expected: "Passed!" with total: 25, failed: 0 (14 existing + 11 new). Baseline measured today: total
      14, failed 0.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/FormulaWriterTests/*"`
      → expected: "Passed!" with total: 71, failed: 0 (49 existing + 15 canonical + 7 normalization
      [Arguments] rows). Baseline measured today: total 49, failed 0. A failure in NormalizesEquivalentText
      means the writer is not idempotent or tag 322 does not serialize.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build`
      → expected: "Test run summary: Passed!", failed: 0 for the whole core suite. No pre-existing test
      contains a `[` inside a formula string (verified by grep), so nothing else should move.
- [ ] `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release --no-build -- --treenode-filter "/*/*/TableInteropTests/*"`
      → expected: total: 14. EXPECT failed: 5 while this phase stands alone — exactly
      Load_StructuredReferenceFormula_ReportsWarning_AndCellFallsBackToCachedValue,
      _WithNoCachedValue_LeavesTheCellBlank, _WithoutOptions_DoesNotThrow,
      Load_RepeatedStructuredReferenceFormula_WarnsOncePerAffectedCell and
      Load_SharedStructuredReferenceMaster_DoesNotAbort..., which the excel-loader phase rewrites. failed: 0
      is only expected once that phase lands. Any OTHER failing name in this class means the vehicle swap in
      item 17 was incomplete.
- [ ] `cd /tmp/probe-lexer-parser && dotnet run --project probe.csproj`
      → expected: The 44-row grammar table reprints with "stable? yes" on every accepted row and the five
      rejected classes (Unsupported/Invalid) unchanged — the standalone oracle the repo implementation must
      match row for row. (The single `'My Table'[Valor]` NO! row is a harness artefact: the probe passes the
      quotes through as part of the table name, which the real tokenizer strips at Tokenizer.cs:178.)

## Risks carried by this phase

- SEQUENCING: this phase alone silently REMOVES the current degradation path. Once `Tabela1[Valor]` parses, WorksheetStreamLoader.cs:500-515/:555-566 no longer raises ExcelLoadWarningKind.UnparsableFormula and no longer falls back to the cached <v>, so with no table registry yet the cell evaluates to an error instead of Excel's number. It must NOT be merged before table-model-registry + resolution-and-graph, or must be merged in the same PR.
- EXPORT: FormulaWriter will now emit `Tabela1[Valor]` into SaveAsExcel output while S3 says no <table> part is written (ExcelExport writes formula text; the part-writing precedents are ExcelExport.cs:107-127/:317-330). The result is an xlsx containing a structured reference to a table that does not exist in the file, which Excel reports as an error or repairs. The export phase must choose: pre-resolve TableReference nodes to their A1 range at export time, emit an ExcelSaveWarning, or document it. MergeIntoExcel is unaffected (it copies the TableDefinitionPart verbatim, ExcelMerge.cs:325/:438).
- DOWNSTREAM DEFAULT ARMS: DependencyExtractor.Visit's `default: return;` (DependencyExtractor.cs:216-217) contributes nothing AND does not set AlwaysDirty, and ReferenceGuard's `default: return null` (ReferenceGuard.cs:96-98) turns a table on a deleted sheet into a silently-empty range instead of #REF!. Those arms belong to resolution-and-graph but are a correctness prerequisite for every node this parser now emits; without them the failure is silent, not loud.
- BEHAVIOUR CHANGE (item 18 only): tightening IsCellReference moves `=Vendas2024` from a silent 0 to #NAME? and starts ACCEPTING DefineName("Table1"). No existing test covers either (grep evidence in the item), but it is a public behaviour change needing a docs paragraph and a `fix(parser)` commit, and it also relaxes the defined-name reservation enforced at NamedReferences.cs:190-192.
- WHITESPACE LENIENCE: `Tabela1 [Valor]` is accepted because SkipWhitespace runs before the reader (Tokenizer.cs:32, measured: `SUM (A1)` already lexes to Identifier LParen), whereas in Excel a space there is the intersection operator. Pre-existing class of lenience (TokenizerTests.cs:63-69 pins it for SUM); document, do not fix.
- RENDERING DIVERGENCE: for a column name containing a special character MySheet writes `Tabela1[['#OfItems]]` where Microsoft's example is `Tabela1['#OfItems]`. Excel accepts both (extra brackets are always legal, `[[Total $ Amount]]` is the doc's own canonical form), but FORMULATEXT text will differ from Excel's for such columns. The MS doc is self-contradictory here — `#` appears both in the 33-character double-bracket list and in the single-bracket example — so the never-under-bracket rule was chosen for safety.
- TOKEN.TEXT IS RAW: the BracketedSpecifier payload is undecoded, the one deliberate divergence from ReadQuotedName (Tokenizer.cs:170-178). If a future contributor "fixes" the reader to decode, the composite form silently breaks: decoding `'[`/`']` before the item split corrupts the split. This must be stated in the reader's comment, not left implicit.
- CONTEXT-FREE CONSEQUENCE: `A1[Col]`, `TRUE[Col]` and `'[1]Sheet1'[Col]` all parse to a TableReference on a table Excel could never name, so the failure is a resolution-time #NAME?/#REF! rather than a syntax error. Accepted: the Parser holds only a sheet name (Parser.cs:10-16) and cannot see the table registry.

## Open questions owned by this phase

- Does Excel ACCEPT the single-bracket form for a column name containing a space (`Tabela1[Sales Amount]`), or does it require `Tabela1[[Sales Amount]]`? We accept both on read and always WRITE the double-bracket form (the safe direction, since extra brackets are documented-canonical). Someone with Excel should confirm the writer's canonical form before we advertise FORMULATEXT parity.
- What is Excel's real rule for which characters trigger the double-bracket form versus only the `'` escape? The MS doc lists `#` among the 33 double-bracket characters AND shows `=DeptSalesFYSummary['#OfItems]` with single brackets. We chose "any non-alphanumeric => double brackets + escape the five", which never under-brackets. Tightening it later is a writer-only change; the node model does not move.
- Does Excel trim decorative whitespace inside a simple column specifier (`Tabela1[ Valor ]`)? We trim, and `Tabela1[[ Valor ]]` is the escape hatch for a genuinely space-padded header. If the answer is "Excel does not trim", the fix is one line in DecodeName plus the registry doing a trimmed-retry lookup.
- Is `Tabela1[[ColA],[ColB]]` legal Excel (a two-column multi-area)? We classify it UnsupportedStructuredReference rather than Invalid precisely because we are unsure; if it is legal it becomes a natural follow-up (a UnionReference of two resolved column ranges), and no parse-level change would be needed beyond the classifier.
- GO/NO-GO on item 18 (tightening Parser.IsCellReference to 1-3 ASCII letters). It is what makes `Workbook.DefineName("Table1", …)` stop throwing — measured today, it throws — and what would let a bare `=Tabela1` resolve to the table's data body like Excel. It is not required for S1. Without it, the docs must state that bare table names are unsupported AND that `=Tabela1` yields 0 rather than an error, which is a silent wrong answer for any real xlsx that uses the bare form.
- Should `Tabela1[[#Headers],[#Data]]` / `[[#Data],[#Totals]]` (the two legal Excel specifier pairs, TableArea.HeadersAndData / DataAndTotals) really be in scope? S1 did not list them either way; the grammar produces them for one extra loop iteration, and the MS doc's own example (`[[#Headers], [#Data], [% Commission]]`) proves Excel emits them. If the resolution phase would rather not implement two extra area arms, its switch must reject them EXHAUSTIVELY (no default fall-through) or a wrong range ships silently — I recommend keeping them and requiring an exhaustive switch.

## Phase Summary

_(write when phase completes)_
