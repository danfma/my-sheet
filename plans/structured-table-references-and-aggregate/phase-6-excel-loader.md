# Phase 6: The .xlsx loader reads <table> parts into Workbook.Tables

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `excel-loader`. Design dependencies: `table-model-registry`, `lexer-parser`, `reference-semantics`, `resolution-and-graph`. Adversarial verifier verdict: **needs-revision** (1 blocker, 4 majors, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

Read the tables inside the existing per-sheet loop of `ExcelFile.Load` (right after
`WorksheetStreamLoader.Load` at ExcelFile.cs:141), not in a separate pass beside `LoadDefinedNames` at :145 —
the loop already holds both the MySheet sheet name (`sheet.Name`, :138) and the `WorksheetPart` (:139), so a
second pass would have to re-walk `workbook.Sheets` and duplicate the skip rule at :130-136, which is a bug
source for zero benefit (resolution is evaluation-time, so ordering is free — `LoadDefinedNames` running after
every formula is parsed already proves it). The mapping is `displayName` (MEASURED: the schema requires
`displayName` and `name` is optional — the OpenXmlValidator reports "The required attribute 'displayName' is
missing" only for the former), `ref` parsed into header/first-data/last-data/totals rows with `headerRowCount
?? 1` and `totalsRowCount ?? 0` (MEASURED: ClosedXML omits both at their defaults, and `totalsRowShown` is
useless for geometry — a table WITH a totals row wrote `totalsRowCount="1"` and no `totalsRowShown`, while
tables WITHOUT one wrote `totalsRowShown="0"` and no count), and the ordered `<tableColumn>/@name` values
taken RAW: the OpenXML SDK already decodes XML entities (`&amp;`, `&#39;`, `&quot;` all came back as the
literal characters), so the loader stores exactly the text Phase 1's lexer produces after undoing Excel's
`'`-escapes, and the only extra decode needed is OOXML's `_xHHHH_` form for control characters, which the SDK
does NOT touch (MEASURED: `_x000a_` survived verbatim). The whole read costs 62 KB and 0.23 ms on a 5000x10
table versus 29.8 MB and 340 ms to materialize the worksheet DOM (MEASURED), so the streaming invariant holds.
Two measured failure modes force a wider catch than `LoadDefinedNames`' `ParseException or ArgumentException`
(ExcelFile.cs:194): enumerating `TableDefinitionParts` over a dangling relationship throws
`InvalidOperationException`, and `tdp.Table` on a corrupt part throws `XmlException`/`InvalidDataException` —
none of which can happen today because nothing ever touches the part, so this phase introduces the failure
mode and must contain it behind a new `ExcelLoadWarningKind.InvalidTableDefinition`.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Item `Load_TableWithATotalsRow_ExcludesItFromTheDataBody`: fixture built with ClosedXML `t.ShowTotalsRow = true; t.Field("Valor").TotalsRowFunction = XLTotalsRowFunction.Sum;` then assert `warnings.Count == 0`, `D1 == 42`, `D2 == 4`, `D3 (SUM(Tabela1[#Totals])) == 42`.
      *Measured evidence:* MEASURED. ClosedXML 0.105.0 writes the totals cell as an IMPLICIT-TABLE formula
      with no cached value: xl/worksheets/sheet1.xml contains `<x:row r="4"><x:c
      r="B4"><x:f>SUBTOTAL(109,[Valor])</x:f></x:c></x:row>` — no `<v>`. Loading that exact fixture through
      the CURRENT ExcelFile.Load gives: `LOAD p3.xlsx: warnings=1 UnparsableFormula/B4/Unexpected character
      '[' (at position 13)` and B4 is Blank. `[Valor]` is the implicit-table form that S1 puts EXPLICITLY OUT
      OF SCOPE, so it will still be a ParseException after Phase 1. Therefore `warnings.Count == 0` fails (it
      is 1, UnparsableFormula/B4) and `SUM(Tabela1[#Totals])` sums A4:B4 = blank + blank = 0, not 42. The
      geometry claims the item rests on ARE correct (`ref="A1:B4" totalsRowCount="1"`, `autoFilter
      ref="A1:B3"`, `totalsRowShown` absent — all measured), so only the fixture/assertions are broken.
      *Correction:* After `wb.SaveAs`, overwrite B4 through the existing `inject` mechanism with a literal
      cell (`<c r="B4"><v>42</v></c>`, no `<f>`), keeping `ref="A1:B4" totalsRowCount="1"`. Then all four
      assertions hold. If the SUBTOTAL cell is kept deliberately, the item must instead assert exactly one
      `UnparsableFormula` for B4 and drop `SUM(Tabela1[#Totals]) == 42` (it would be 0).

## Major corrections

- [ ] **M1.** Item `Load_TableWithNoHeaderRow_UsesGeneratedColumnNames`: "ClosedXML's `SetShowHeaderRow(false)` rewrites the ref to drop row 1 and names the columns after the discarded data (MEASURED: `ref="A2:B2" headerRowCount="0"` with names "10" and "20") — a misleading fixture", therefore hand-inject `ref="A2:B3"` with `<tableColumn name="Column1"/><tableColumn name="Column2"/>` and assert `SUM(T[Column1])` covers both rows.
      *Evidence:* MEASURED, and the cited measurement is wrong. ClosedXML 0.105.0 `t.SetShowHeaderRow(false)`
      on the A1:B3 fixture produced `<x:table ... ref="A2:B3" headerRowCount="0">` with `<x:tableColumn id="1"
      name="Item"/><x:tableColumn id="2" name="Valor"/>` — the REAL header names, not "10"/"20" — and it
      DELETED row 1 from the sheet (sheetData now starts `<x:row r="2">` with A2="a", B2=10, A3="b", B3=32).
      Separately, the specified assertion is vacuous: in the base fixture column A holds the text "a"/"b", so
      `SUM(T[Column1])` = 0 under a correct geometry AND under every wrong one.
      *Correction:* Use ClosedXML's real headerless output as the fixture (`ref="A2:B3"`, headerRowCount=0,
      names Item/Valor) — no InjectTablePart needed here — and assert `SUM(Tabela1[Valor]) == 42`, which
      distinguishes correct geometry (B2:B3 = 42) from mistaking row 2 for a header (B3 = 32). Keep
      InjectTablePart for the malformed-part test, which genuinely needs it.
- [ ] **M2.** The mapping the loader records: TryMap produces `TableDefinition` from name + geometry (headerRow/firstDataRow/lastDataRow/totalsRow, firstColumn/lastColumn) + ordered column names; the docs row says the entry holds "the table's name (`displayName`), its geometry (…) and its ordered `<tableColumn>` names".
      *Evidence:* The sheet name is never recorded anywhere. `TryMap(Table table, string sheetName, out
      TableDefinition definition, out string reason)` takes `sheetName` but the item's enumerated assignments
      never use it, and `TryParseReference` deliberately "reject[s] anything containing '$' or '!'" so the
      `ref` cannot carry it either — confirmed MEASURED: every fixture wrote a bare `ref="A1:B3"` / `"A1:B4"`
      / `"A2:B3"`. `ExcelFile.cs:138` is the only place holding the sheet. Without it, a `TableDefinition`
      resolving to a RangeReference has no `SheetName` (Reference.cs / RangeReference(StartId, EndId,
      SheetName)), so `Tabela1[Valor]` written on Sheet2 would resolve against Sheet2's grid.
      *Correction:* State explicitly that `TableDefinition` carries the owning sheet name and that TryMap
      assigns it from the `sheetName` parameter; add it to the docs/excel-interop.md mapping-table row ("…the
      owning sheet, the table's name, its geometry…"). Add an assertion to
      `Load_Table_RegistersItInWorkbookTables_WithoutWarnings` pinning the recorded sheet name.
- [ ] **M3.** `InvalidTableDefinition`'s enum doc (and the docs/excel-interop.md row) advertise "a name Excel's own name rules reject" as a reject reason, and Read's structure is "(2) per-table try wrapping `tablePart.Table` and the mapping … (3) on a successful map, `if (workbook.Tables.ContainsKey(...)) {…} continue;` then `workbook.DefineTable(definition)`".
      *Evidence:* No specified code path produces that reject reason: TryMap's five rejects are headerRows>1,
      geometry overflow, column-count mismatch, empty column name, duplicate column names — none is a table-
      NAME check. The only place name validation can live is `DefineTable`, which S2 says mirrors
      `DefineName`; `Workbook.DefineName` calls `NamedReferences.ValidateName(name)` (Workbook.cs:549) which
      throws `ArgumentException` unless the name "start[s] with a letter or underscore, contain[s] only
      letters, digits, '.' or '_'" (NamedReferences.cs:164-172). Excel permits a backslash in a table name, so
      `displayName="My\Table"` is a legal Excel table that ValidateName rejects. Step (3) reads as OUTSIDE the
      try ("wrapping `tablePart.Table` and the mapping"), so that ArgumentException would escape
      `ExcelFile.Load` — turning a file that loads today into a hard failure, the exact regression risk #1
      exists to prevent. The item also never states the `if (!TryMap(...)) { Warn(...); continue; }` branch at
      all, so the malformed-table test's `Subject == "T"` has no specified producer.
      *Correction:* State that step (3) — the duplicate check AND `workbook.DefineTable(definition)` — is
      INSIDE the per-table try (that is what the `ArgumentException` in the filter is for), and add the
      missing `if (!TryMap(table, sheetName, out var definition, out var reason)) { Warn(options, name ??
      sheetName, reason); continue; }` step. Add a test with `displayName="My\Table"` (InjectTablePart)
      asserting one `InvalidTableDefinition` and no throw.
- [ ] **M4.** Verification: `grep -rn "No Excel Tables\|Sem Tabelas do Excel\|models the first and not the second\|modela o primeiro e não o segundo" docs/ README.md` → "No matches (grep exits 1). Each of these strings is a sentence this phase must have replaced; a hit means a doc edit was skipped."
      *Evidence:* The pt-BR pattern matches nothing even TODAY. docs/pt-BR/workbook-and-expressions.md:396
      actually reads `> O MySheet modela o primeiro, e não a segunda — veja` (comma; feminine "a segunda").
      Ran the grep as written: it returns 3 hits, all English/pt-BR-excel-interop, none from pt-BR/workbook-
      and-expressions.md. So after the three English edits plus the pt-BR excel-interop edit, this
      verification passes while the pt-BR workbook-and-expressions mirror edit — the one the spec specifies
      with no line cite ("Mirror into …") — can be silently skipped.
      *Correction:* Change the pt-BR pattern to `modela o primeiro` (or `e não a segunda`). Verified it
      matches docs/pt-BR/workbook-and-expressions.md:396 today.

## Implementation items

- [ ] **1.** Add `internal static bool TryParse(string? id, out int row, out int column)` to /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/CellId.cs and reimplement the existing `Parse` (:8-20) as `TryParse(id, out r, out c) ? (r, c) : throw new FormatException(...)`, keeping `Parse`'s signature intact for its six existing call sites. TryParse must require at least one letter and at least one digit, reject anything after the digits, and use `int.TryParse` so no exception escapes.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/CellId.cs`
      *Why:* A table's `ref` is untrusted producer text and the SDK does NOT validate it — MEASURED:
      `ref="NOT-A-RANGE"` came back from `Table.Reference` as that literal string, and a missing `ref` came
      back null. `CellId.Parse` (CellId.cs:19) ends in `int.Parse(id.AsSpan(index))`, which on "NOT-A-RANGE"
      throws FormatException from inside Load with nothing catching it. TryParse is the smallest reusable
      component; the existing callers (ExcelExport.cs:166, ExcelMerge.cs:128,:399,
      SharedFormulaShifter.cs:21-22, WorksheetStreamLoader.cs:151,:180-181,:376) all read ids the SDK already
      validated, so they keep the throwing overload.
- [ ] **2.** Append `InvalidTableDefinition` as the FIFTH member of `ExcelLoadWarningKind` in /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs, after `UnparsableCellLiteral` (:77), with this doc comment in the house style of :45-77: "An Excel <b>Table</b> (<c>&lt;table&gt;</c> part) that could not be registered in <see cref=\"Workbook.Tables\"/>: a missing or malformed <c>ref</c>, a <c>headerRowCount</c> other than 0 or 1, a column count that disagrees with the <c>ref</c>'s width, an empty or duplicated column name, a name Excel's own name rules reject, a name another table already claimed, or a table part the package cannot produce at all. The table is skipped and the rest of the workbook loads normally — its cells are still ordinary cells — but a structured reference into it then evaluates to <c>#NAME?</c>, which is exactly what Excel shows for a table that does not exist. <see cref=\"ExcelLoadWarning.Subject\"/> is the table's <c>displayName</c>, or the SHEET name when the part could not be read far enough to have one." Also extend the `Subject` param doc at :29-33 to mention the table/sheet subject.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs`
      *Why:* The enum is public; appending keeps every existing numeric value stable. A new kind rather than
      reusing `InvalidDefinedName` because the Subject is a table (or a sheet) rather than a defined name, and
      because a host filtering on kind must be able to tell "one table was dropped, its formulas now say
      #NAME?" from "one name was dropped". The #NAME? sentence is the Excel behaviour being matched (P0):
      Excel resolves a structured reference to an unknown table as #NAME?, so skipping a corrupt table is more
      faithful than today's cached-value fallback — but it is a visible change and the warning is what makes
      it observable.
- [ ] **3.** Reword the `UnparsableFormula` doc comment at /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs:59-67 — it currently names a structured reference as the cause "which the tokenizer has no `[` for", which becomes false. Replacement body: "A cell whose formula text failed to parse — a syntax MySheet's parser does not accept. Now that structured references are supported, the remaining common causes are the structured-reference shapes still out of scope (the this-row form <c>Tabela1[@Valor]</c>, a column span <c>Tabela1[[Q1]:[Q3]]</c>, the implicit-table form <c>[Valor]</c>), array literals (<c>{1;2;3}</c>), and genuinely malformed or otherwise unsupported formula text." Keep the existing sentences about the cached-value fallback and the shared-formula master verbatim (:62-66), then add: "A structured reference whose TABLE is missing or was skipped is NOT this warning — it parses fine and evaluates to <c>#NAME?</c>; see <see cref=\"InvalidTableDefinition\"/>."
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs`
      *Why:* The warning kind must survive (it still fires for `[@Col]`, spans, implicit-table refs and
      `{1;2;3}` — MEASURED today: `{3;1;2}` throws ParseException "Unexpected character '{'"), but its
      documented cause is now wrong and would send a host looking for a table problem that does not exist. The
      final cross-reference sentence closes the one gap a reader will otherwise hit: the two table-related
      warnings are mutually exclusive.
- [ ] **4.** Create /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs holding `internal static class TableDefinitionReader` with entry point `public static void Read(WorksheetPart part, Workbook workbook, string sheetName, ExcelLoadOptions? options)`. Body: (1) materialize the parts inside a try — `List<TableDefinitionPart> parts; try { parts = [.. part.TableDefinitionParts]; } catch (Exception exception) when (exception is InvalidOperationException or OpenXmlPackageException) { Warn(options, sheetName, exception.Message); return; }`; (2) `foreach (var tablePart in parts)` with a per-table try wrapping `tablePart.Table` and the mapping, catching `XmlException or InvalidDataException or InvalidOperationException or ArgumentException or FormatException` and warning with the table's displayName when known, else `sheetName`; (3) on a successful map, `if (workbook.Tables.ContainsKey(definition.Name)) { Warn(...); continue; }` then `workbook.DefineTable(definition)`. Add a class doc comment stating that `TableDefinitionParts` is a package-relationship enumeration, so this reads the tiny `<table>` part without materializing the worksheet DOM (62 KB / 0.23 ms vs 29.8 MB / 340 ms measured on a 5000x10 table), which is why it does not break WorksheetStreamLoader's streaming contract.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs`
      *Why:* A new file rather than growing ExcelFile.cs (209 lines, currently Load + LoadDefinedNames): the
      project's pattern is one focused reader per concern — SharedStringsStreamReader.Read,
      WorksheetStreamLoader.Load, SharedFormulaShifter, XlsxNumbers — and the mapping plus validation is ~130
      lines. The two-stage try is forced by measurement, not caution: with the table part physically deleted
      from the package but the relationship kept, `SpreadsheetDocument.Open` and `WorksheetParts.First()` both
      SUCCEED and the throw lands on enumeration — `InvalidOperationException: Part: /xl/tables/table1.xml
      doesn't exist in the package` — so a single try around the loop would silently drop a healthy third
      table when the second is corrupt, while no try at all would turn a file that loads today into a hard
      failure. `tdp.Table` on garbage XML threw `XmlException` ("Data at the root level is invalid") and on a
      wrong root element `InvalidDataException` ("Cannot load the root element from the part") — both outside
      the `ParseException or ArgumentException` filter used at ExcelFile.cs:194. First-registration-wins on a
      duplicate name is checked HERE rather than by making `DefineTable` throw, so the public API stays
      symmetric with `DefineName`'s last-one-wins overwrite (Workbook.cs:551, :584) while the loader still
      refuses to repoint a name that formulas on an earlier sheet already resolve through.
- [ ] **5.** Add `private static bool TryMap(Table table, string sheetName, out TableDefinition definition, out string reason)` to TableDefinitionReader.cs. Name: `table.DisplayName?.Value` when non-blank, else `table.Name?.Value` when non-blank, else fail with reason "the <table> part has no displayName". Geometry: `TryParseReference(table.Reference?.Value, out firstRow, out firstColumn, out lastRow, out lastColumn)`; `var headerRows = table.HeaderRowCount?.Value ?? 1;` `var totalsRows = table.TotalsRowCount?.Value ?? 0;` then `headerRow = headerRows == 0 ? 0 : firstRow`, `firstDataRow = firstRow + (int)headerRows`, `lastDataRow = lastRow - (int)totalsRows`, `totalsRow = totalsRows == 0 ? 0 : lastRow`. Columns: `table.TableColumns?.Elements<TableColumn>()` in document order, each `DecodeColumnName(column.Name?.Value)`. Reject (return false with a specific reason string) when: headerRows > 1; `firstRow + headerRows + totalsRows > lastRow + 1`; the column count differs from `lastColumn - firstColumn + 1`; any column name is null/empty/whitespace; two column names are equal under OrdinalIgnoreCase. ACCEPT `lastDataRow == firstDataRow - 1` (a header-only table with an empty data body) and pass it through unchanged.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs`
      *Why:* Every default and every gate here is measured, not assumed. `HeaderRowCount` came back null for a
      normal ClosedXML table whose header row plainly exists, so `?? 1` is required (ECMA's default).
      `TotalsRowCount=1` with `TotalsRowShown=null` on a totals table versus `TotalsRowShown=False` with
      `TotalsRowCount=null` on three non-totals tables proves `totalsRowShown` is a UI flag and only
      `totalsRowCount` is geometry; the `ref` INCLUDES the totals row (measured `ref="A1:B4"` with `autoFilter
      ref="A1:B3"`), which is exactly why `lastDataRow` must subtract it — getting this wrong double-counts
      the totals cell in `SUM(Tabela1[Valor])`. The rejects are all reachable and all schema-VALID (measured:
      `headerRowCount="2"`, `count="3"` over a 2-column ref, `name=""`, a missing `name` attribute, and
      `c`/`C` duplicates each round-tripped through the SDK and the OpenXmlValidator without complaint), so
      the schema cannot be relied on and the loader is the only gate. Excel writes only 0 or 1 header rows and
      its ListObject header is a single row, so >1 is skipped rather than guessed at. `headerRowCount=0` means
      the ref starts at the first DATA row and the `<tableColumn>` names are Excel's generated ones (measured:
      ClosedXML wrote `ref="A2:B2" headerRowCount="0"` with names taken from the dropped row), so `headerRow =
      0` is the sentinel for "no header row" and `[#Headers]` becomes the reference-semantics phase's error
      case.
- [ ] **6.** Add `private static bool TryParseReference(string? reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)` to TableDefinitionReader.cs: reject null/whitespace; split on the first ':' (no ':' means a single-cell ref, start == end); `CellId.TryParse` each half; then require `firstRow <= lastRow && firstColumn <= lastColumn`, normalizing nothing. Do NOT strip '$' or a sheet qualifier — a table `ref` never carries either (MEASURED: every fixture wrote a bare `A1:B3` / `C3:C4` form); reject anything containing '$' or '!' as malformed so a producer surprise surfaces as a warning instead of a wrong range.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs`
      *Why:* Deliberately NOT routed through `ExpressionParser.ParseFormulaBody` the way `LoadDefinedNames`
      handles `refersTo` (ExcelFile.cs:187-190): a table `ref` is a plain bounding box, and parsing it into an
      Expression only to destructure it back into four integers would add a dependency on how the parser types
      unqualified references and would accept nonsense like `A1:B3+1`. `CellId.TryParse` is the same primitive
      `WorksheetStreamLoader.ApplyDimensionHint` already uses for the identically-shaped `dimension/@ref`
      (WorksheetStreamLoader.cs:180-181), so this is the established pattern in this assembly.
- [ ] **7.** Add `private static string DecodeColumnName(string raw)` to TableDefinitionReader.cs, undoing OOXML's `_xHHHH_` escaping for CONTROL characters only: scan for "_x"; accept exactly four hex digits followed by '_'; decode `_x005f_` to '_' and `_xHHHH_` to that char when HHHH < 0x0020; leave every other `_x…_` sequence literal. Fast-path `return raw` when `raw.IndexOf("_x", StringComparison.Ordinal) < 0`. Document that XML entity decoding is already done by the SDK and must NOT be repeated, and that Excel's formula-side `'`-escapes (`''`, `'[`, `']`, `'#`, `'@`) live on the LEXER side (Phase 1) — both halves meet at this raw text, compared OrdinalIgnoreCase.
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs`
      *Why:* MEASURED, and this is the whole column-name-fidelity answer. The SDK decodes entities: injected
      `A&amp;B &lt; C &gt; D &quot;q&quot;` came back as `A&B < C > D "q"`, and `Owner&#39;s Share` came back
      as `Owner's Share` — so a second decode pass would corrupt them. Spaces and parentheses need nothing at
      all: `AMOUNT IN USD For Line 1` and `(A) NAME OF PFIC` round-tripped byte-for-byte, and so did `Net
      [USD] #1 @Rate`, which is the proof that the `'`-escaping is a FORMULA-text convention and never appears
      in the part. The one thing the SDK does not touch is `_xHHHH_`: injected `Line _x000a_ Break` came back
      with the escape literal, and a raw newline cannot be used instead because XML attribute-value
      normalization would turn it into a space — so a header containing a line break is unreachable from a
      formula without this decode. Restricting the decode to code points < 0x20 (plus `_x005f_`) is what makes
      it safe in both directions: those are the only characters a producer is FORCED to escape, so decoding
      them cannot misfire on a producer that escapes nothing, while a user-typed literal `_x0020_` is left
      alone.
- [ ] **8.** Add `private static void Warn(ExcelLoadOptions? options, string subject, string detail)` to TableDefinitionReader.cs — one line: `options?.OnWarning?.Invoke(new ExcelLoadWarning(ExcelLoadWarningKind.InvalidTableDefinition, subject, detail));`
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/TableDefinitionReader.cs`
      *Why:* Mirrors the null-tolerant single-call shape already used at ExcelFile.cs:199-205 and
      WorksheetStreamLoader.cs:574-576, so a null `options` costs nothing (ExcelLoadOptions doc at
      ExcelFile.cs:15-22 makes the pays-nothing-when-null contract explicit) and the five reject sites stay
      one-liners.
- [ ] **9.** Insert `TableDefinitionReader.Read(worksheetPart, workbook, sheet.Name, options);` in /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs immediately after the `WorksheetStreamLoader.Load(...)` call at :141, inside the per-sheet loop, with the comment: "Tables come from a package RELATIONSHIP (TableDefinitionParts), not from the sheet XML — the streaming loader breaks out at </sheetData> and never sees <tableParts>. Read after the cells so this sheet's warnings stay in document order; registration order does not matter because a structured reference resolves at evaluation time, exactly like a defined name."
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs`
      *Why:* The per-sheet loop is the only place that already holds BOTH halves of what the mapping needs:
      the MySheet sheet name (`sheet` from :138) and the `WorksheetPart` (:139). A separate pass beside
      `LoadDefinedNames` (:145) would have to re-iterate `workbookPart.Workbook.Sheets.Elements<XlsxSheet>()`,
      re-apply the null-name/null-relId skip at :130-136 and re-call `GetPartById` — duplicated logic whose
      only failure mode is drifting apart from the loop it copies, and a sheet the loop skipped would silently
      get its tables read against a sheet name that no `Workbook.Sheets` entry has. Ordering is genuinely
      free: `LoadDefinedNames` runs at :145 AFTER every formula on every sheet is already parsed (comment at
      :144), which proves parse-time never consults the registry. MEASURED that this placement is safe for
      streaming: enumerating `TableDefinitionParts` and reading the table DOM allocated 62 KB in 0.23 ms while
      `part.Worksheet` cost 29.8 MB and 340 ms, and the worksheet `GetStream` still worked after touching a
      child part.
- [ ] **10.** Update the `ExcelFile` class doc comment at /Volumes/Web/../Danfma.MySheet.Excel/ExcelFile.cs:80-87 (real path /Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs) — after the sentence about shared-formula slaves, add: "Excel <b>Tables</b> are read from each worksheet's <c>&lt;table&gt;</c> parts into <see cref=\"Workbook.Tables\"/> (name, geometry and column names), which is what makes a structured reference such as <c>Tabela1[Valor]</c> resolve; the parts are reached through package relationships, so this does not materialize any worksheet DOM."
      *Files:* `/Volumes/Work/Develop/MySheet/Danfma.MySheet.Excel/ExcelFile.cs`
      *Why:* The class comment is the one place that currently enumerates what Load does and does not model;
      leaving it silent about tables while the mapping table in docs/excel-interop.md changes would put the
      two out of sync. The streaming clause is there because that comment's central claim (":82-84 the OpenXML
      DOM is never materialized") is exactly what a reader will suspect this feature broke.
- [ ] **11.** In /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs replace the single `private const string StructuredFormula = "SUM(Tabela1[Valor])";` (:21) with TWO constants: `private const string UnsupportedStructuredFormula = "SUM(Tabela1[@Valor])";` (rejected at PARSE — the this-row form is out of scope) and `private const string UntokenizableFormula = "SUM({1;2;3})";` (rejected by the TOKENIZER — array literals are out of scope). Point every current `StructuredFormula` use at `UnsupportedStructuredFormula` EXCEPT the two tests whose comments depend on a tokenizer-stage failure: `Load_SharedStructuredReferenceMaster_DoesNotAbort_AndTheGroupFallsBackToCachedValues` (:243, comment at :245-247 "The master's TOKENIZATION is what fails here (before any parse)") and `Load_RejectedMasterReusingAnotherGroupsIndex_LeavesTheLegitimateGroupIntact` (:511, comment at :532 "whose text cannot be tokenized at all"), which must use `UntokenizableFormula`.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* Ten of the fourteen tests in this file use `SUM(Tabela1[Valor])` purely as "some text the parser
      rejects" — the negative-cache and <v>-decode tests at :162, :193, :213, :303, :335, :367, :392, :418 and
      the two shared-master tests. That text now PARSES, so all ten break at once; the const swap repairs them
      in one edit and, as a bonus, pins S1's out-of-scope boundary from the loader side. The split into two
      constants is required because a `[@Valor]` payload may well tokenize cleanly (Phase 1 could emit one
      bracketed-specifier token and reject `@` in the parser), which would silently move those two tests off
      the tokenizer-stage path their comments claim to cover; `{1;2;3}` is guaranteed to fail in the tokenizer
      (MEASURED today: "Unexpected character '{'") and stays out of scope per S1.
- [ ] **12.** REWRITE `Load_TableWithOrdinaryFormulas_LoadsAsAPlainRange_WithoutWarnings` (TableInteropTests.cs:99-114) as `Load_Table_RegistersItInWorkbookTables_WithoutWarnings`: keep the fixture and the `warnings.Count == 0` / `B2 == 10` / `B4 == 42` (re-evaluated, cache says 999) assertions verbatim, and REPLACE the `DefinedNames.ContainsKey("Tabela1")` assertion at :108 with three: `workbook.Tables.ContainsKey("Tabela1")` is true; `workbook.Tables["Tabela1"]`'s column names are `["Item", "Valor"]` in order; and `workbook.DefinedNames.ContainsKey("Tabela1")` is still FALSE — tables and names are separate maps.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* Its comment at :107 ("MySheet has no table model: the table's name is not a resolvable
      reference") is the assertion this whole phase inverts, but the DefinedNames half stays TRUE and is worth
      keeping: it pins that the loader does not smuggle tables into `Workbook.DefinedNames` (Workbook.cs:123),
      which is the shortcut a future implementer would reach for. The 999-vs-42 cached-value lie already in
      the fixture (:94) keeps proving re-evaluation.
- [ ] **13.** DELETE `Load_StructuredReferenceFormula_ReportsWarning_AndCellFallsBackToCachedValue` (TableInteropTests.cs:128-147) and replace it with `Load_StructuredReference_EvaluatesAgainstTheRegisteredTable`: same fixture, `FormulaCell("B4", new CellFormula("SUM(Tabela1[Valor])"), "999")`, then assert `warnings.Count == 0` and `workbook.GetCellValue("Data", "B4").ToDouble() == 42.0` — the true sum, NOT the cached 999. Add a second assertion that the formula is live: set `workbook["Data"]["B2"] = new NumberValue(8)`, `workbook.InvalidateCache()`, then `B4 == 40.0`.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* This is the one test asserting the exact opposite of the feature (:131-137: one UnparsableFormula
      warning, cell reads 999), so it is deleted rather than adapted. The 999 cache is what makes the
      replacement airtight — a loader that still degraded would read 999 and only a real resolution reads 42.
      The edit-and-recalculate half is the cheapest end-to-end proof that the structured reference reached the
      dirty graph rather than being frozen at load time, and it copies the technique already used at :575-580.
- [ ] **14.** ADD `Load_StructuredReference_WithOutOfScopeThisRowForm_StillDegradesWithAWarning` to TableInteropTests.cs: fixture injects `FormulaCell("B4", new CellFormula("Tabela1[@Valor]*2"), "999")`; assert exactly one warning, `Kind == ExcelLoadWarningKind.UnparsableFormula`, `Subject == "B4"`, and `B4 == 999.0`; assert no `InvalidTableDefinition` warning was raised (the table itself is fine).
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* S1 puts `[@Column]` explicitly out of scope and requires a clear ParseException, and
      ExcelFile.cs's reworded `UnparsableFormula` comment now names this shape as the primary remaining cause
      — the doc claim needs a test. Asserting the ABSENCE of InvalidTableDefinition is what separates "the
      formula shape is unsupported" from "the table failed to register", the distinction the two warning kinds
      exist to make.
- [ ] **15.** ADD `Load_TableWithATotalsRow_ExcludesItFromTheDataBody` to TableInteropTests.cs with a new fixture writer `WriteTotalsTableFixture(Action<SheetData>? inject)`: ClosedXML writes Item/Valor over A1:B3 with 10 and 32, then `var t = data.Range("A1:B3").CreateTable("Tabela1"); t.ShowTotalsRow = true; t.Field("Valor").TotalsRowFunction = XLTotalsRowFunction.Sum;` (producing `ref="A1:B4" totalsRowCount="1"` with the totals in row 4). Inject `SUM(Tabela1[Valor])` at D1 and `ROWS(Tabela1[#All])` at D2 and `SUM(Tabela1[#Totals])` at D3. Assert D1 == 42 (NOT 84), D2 == 4, D3 == 42, and `warnings.Count == 0`.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* This is the highest-value new test because it is the one geometry mistake that produces a
      plausible-looking wrong number rather than an error: the `ref` INCLUDES the totals row (MEASURED:
      `ref="A1:B4" totalsRowCount="1"` while `autoFilter ref="A1:B3"` stops before it), so a loader that
      forgets `lastDataRow = lastRow - totalsRows` returns 84 for a column whose Excel value is 42. Excel's
      rule being matched: `Table[Column]` is the data body only, `[#All]` is header + data + totals,
      `[#Totals]` is the totals row. `ShowTotalsRow` + `TotalsRowFunction` is confirmed working in ClosedXML
      0.105.0 by probe.
- [ ] **16.** ADD `Load_TableColumnNamesWithSpacesParenthesesAndAnApostrophe_ResolveFromFormulas` to TableInteropTests.cs with a fixture whose headers are `AMOUNT IN USD For Line 1` (A1), `(A) NAME OF PFIC` (B1) and `Owner's Share` (C1) over A1:C3, and injected formulas `SUM(Tabela1[AMOUNT IN USD For Line 1])`, `COUNTA(Tabela1[(A) NAME OF PFIC])` and `SUM(Tabela1[Owner''s Share])` (the apostrophe DOUBLED, Excel's escape). Assert each evaluates to the expected number with no warnings, and assert `workbook.Tables["Tabela1"]`'s column names are the three RAW strings with a single apostrophe in the third.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* MEASURED that these are exactly the shapes that round-trip through the part untouched — ClosedXML
      wrote `name="AMOUNT IN USD For Line 1"`, `name="(A) NAME OF PFIC"` and `name="Owner's Share"`, and the
      SDK handed all three back verbatim — so the ONLY way this test fails is if the two halves disagree about
      escaping: the loader adding a decode it should not, or the lexer failing to collapse `''` to `'`. That
      is the silent-mismatch risk this dimension was told to close, and asserting both the registry contents
      AND the formula result localizes a failure to one side or the other.
- [ ] **17.** ADD `Load_ItemSpecifiers_ResolveAgainstTheLoadedGeometry` to TableInteropTests.cs: on the plain A1:B3 fixture, inject `ROWS(Tabela1[#All])`, `ROWS(Tabela1[#Data])`, `COUNTA(Tabela1[#Headers])` and `SUM(Tabela1[[#Data],[Valor]])` into D1:D4 and assert 3, 2, 2 and 42 with no warnings. ALSO add `Load_TableWithNoHeaderRow_UsesGeneratedColumnNames` using a hand-injected table part (`headerRowCount="0"` over `ref="A2:B3"` with `<tableColumn name="Column1"/><tableColumn name="Column2"/>`, via a new `InjectTablePart(string path, string tableXml)` helper that overwrites the first `TableDefinitionPart`'s stream), asserting `SUM(T[Column1])` covers BOTH rows of the ref and no warning is raised.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* S1 puts `[#All]`/`[#Data]`/`[#Headers]`/`[#Totals]` and the composite `[[#Data],[Column]]` in
      scope, and each one reads a DIFFERENT field of the geometry the loader derived, so this is the cheapest
      way to pin all four fields from the .xlsx side. The headerless case needs hand-injected XML because
      ClosedXML's `SetShowHeaderRow(false)` rewrites the ref to drop row 1 and names the columns after the
      discarded data (MEASURED: `ref="A2:B2" headerRowCount="0"` with names "10" and "20") — a misleading
      fixture. `InjectTablePart` is also the helper the malformed-table test below needs.
- [ ] **18.** ADD `Load_DynamicStructuredReferenceThroughIndirect_Resolves` to TableInteropTests.cs: inject `SUM(INDIRECT("Tabela1["&A1&"]"))` at D1 with A1 already holding the text "Valor"… — since A1 is the table's own header, use a cell OUTSIDE the table: put `Valor` at D5 and inject `SUM(INDIRECT("Tabela1["&D5&"]"))` at D1. Assert D1 == 42 and no warnings.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* Indirect.cs:58-66 parses its text with `ExpressionParser.ParseFormulaBody` and then calls
      `parsed.TryResolveReference`, so this shape works for free the moment the lexer accepts `[` and the node
      overrides the virtual — but "for free" is exactly the kind of claim that quietly stops being true, and
      this is the only test that exercises the runtime-constructed path. It also guards the `IsVolatile =>
      true` interaction (Indirect.cs:16): a table reference reached through INDIRECT must still recompute.
- [ ] **19.** ADD `Load_MalformedTablePart_WarnsAndSkipsTheTable_WithoutFailingTheLoad` to TableInteropTests.cs using `InjectTablePart` to plant `<table … name="T" displayName="T" ref="NOT-A-RANGE"><tableColumns count="1"><tableColumn id="1" name="C"/></tableColumns></table>`: assert exactly one warning with `Kind == InvalidTableDefinition`, `Subject == "T"`, non-empty Detail; `workbook.Tables.Count == 0`; and that the sheet's ordinary cells are intact (`B3 == 32`). ADD a sibling `Load_TablePartMissingFromThePackage_WarnsWithTheSheetAsSubject` that deletes the `xl/tables/table1.xml` zip entry (System.IO.Compression, ZipArchiveMode.Update) while leaving the relationship, asserting one `InvalidTableDefinition` warning with `Subject == "Data"` and `B3 == 32`.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* Both are MEASURED failure modes that this phase INTRODUCES, because nothing touches the table
      part today: `ref="NOT-A-RANGE"` sails through the SDK and the validator and would reach `int.Parse`
      (CellId.cs:19) as a FormatException, and a dangling part makes the `TableDefinitionParts` enumeration
      throw `InvalidOperationException: Part: /xl/tables/table1.xml doesn't exist in the package` AFTER
      `SpreadsheetDocument.Open` has already succeeded. The second test is the only thing that pins the outer
      try, and the differing Subject (table name vs sheet name) is what documents which stage failed.
- [ ] **20.** ADD `SaveAsExcel_WorkbookWithTables_WritesNoTablePart` to /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/ExcelExportTests.cs: load a table fixture, `SaveAsExcel` to a temp path in both `FormulaMode.ValuesOnly` and `FormulaMode.Formulas`, reopen with `SpreadsheetDocument.Open(path, false)` and assert `WorksheetParts.First().TableDefinitionParts.Count() == 0` in both, plus (Formulas mode) that the `<f>` text round-trips as `SUM(Tabela1[Valor])`. Do NOT touch ExcelExport.cs or ExcelMerge.cs.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/ExcelExportTests.cs`
      *Why:* S3 puts table-part WRITING out of scope, and the honest way to keep it out of scope is to pin it
      as a deliberate contract rather than leave it as an absence nobody checked. This test is also the one
      that makes the export trap visible in CI: in Formulas mode the file now carries a structured reference
      with no table to resolve it, so Excel opens it showing #NAME? — see risks. `MergeIntoExcel` needs no
      test change at all: it opens `isEditable: true` and copies non-owned nodes verbatim (ExcelMerge.cs:325,
      :438), so the `<tableParts>` element and the `TableDefinitionPart` survive untouched, already documented
      and already covered.
- [ ] **21.** Rewrite the TableInteropTests.cs class doc comment (:8-18): it currently states "MySheet has no table model and its tokenizer has no `[`, so a structured reference cannot parse". Replacement: the file covers (a) Excel Tables loading into `Workbook.Tables` and their structured references EVALUATING, and (b) the degrade path that survives for shapes MySheet still rejects — the this-row form, column spans, implicit-table refs, array literals, and malformed text. Keep the last paragraph about ClosedXML writing the fixture and formulas being injected through the OpenXML SDK verbatim, and add that a malformed `<table>` part is injected by overwriting the `TableDefinitionPart`'s stream because ClosedXML validates what it writes.
      *Files:* `/Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      *Why:* The comment is the file's contract statement and every sentence of it is now false. Consider (but
      do NOT do in this phase) splitting the eight non-table degrade tests into their own file — they use
      `WriteTableFixture` only as a scratch sheet — because moving eight tests and duplicating three helpers
      is more transcription risk than the tidiness is worth right now.
- [ ] **22.** Rewrite the Excel Table row of the mapping table in /Volumes/Work/Develop/MySheet/docs/excel-interop.md (:66, currently "Nothing — its cells load as an ordinary range"): "| Excel **Table** (a `<table>` part, a.k.a. a ListObject) | An entry in [`Workbook.Tables`](workbook-and-expressions.md#excel-tables): the table's name (`displayName`), its geometry (header row, data body, totals row — derived from `ref`, `headerRowCount` and `totalsRowCount`) and its ordered `<tableColumn>` names. Its cells stay ordinary cells too. **Structured references** (`Tabela1[Valor]`, `Tabela1[#All]`, `Tabela1[[#Data],[Valor]]`) resolve against it. A table MySheet cannot register — malformed `ref`, `headerRowCount` above 1, a column count disagreeing with `ref`, an empty or duplicated column name, a name Excel's own rules reject, or a name already taken — is skipped and reported as `InvalidTableDefinition`; its structured references then evaluate to `#NAME?`, as in Excel. `SaveAsExcel` still writes no `<table>` part. |"
      *Files:* `/Volumes/Work/Develop/MySheet/docs/excel-interop.md`
      *Why:* This row is the mapping table's only statement about tables and it currently promises the
      opposite of the shipped behaviour. Naming the six reject reasons here (rather than only in the XML doc
      comment) is what lets a host reading the docs predict which of their files will warn. Note this file is
      being edited concurrently — the row's line number has already moved once this session, so anchor on the
      row's leading cell text, not on :66.
- [ ] **23.** Rewrite the limitations bullet at /Volumes/Work/Develop/MySheet/docs/excel-interop.md:235-241 ("**No Excel Tables and no structured references**") as "**Excel Tables load, but MySheet never writes one**": tables are read into `Workbook.Tables` and `Tabela1[Valor]` / `[#All]` / `[#Data]` / `[#Headers]` / `[#Totals]` / `[[#Data],[Valor]]` resolve; still unsupported and still degrading via `UnparsableFormula` are `[@Valor]`, column spans `[[Q1]:[Q3]]` and the implicit-table form `[Valor]`; the table's geometry is fixed at load time and does not grow when rows are appended; `SaveAsExcel` writes no `<table>` part, so exporting in `FormulaMode.Formulas` a workbook whose formulas use structured references produces a file Excel opens with `#NAME?` — export values (the default) or use `MergeIntoExcel` into a template that already owns the table. Keep the existing final sentences about `MergeIntoExcel` copying the `<table>` part and `<tableParts>` through untouched and NOT resizing `ref` (:239-241) verbatim. Also update the `UnparsableFormula` guidance paragraph at :124-127 ("A structured reference into an Excel Table is the common cause") to name the out-of-scope shapes instead.
      *Files:* `/Volumes/Work/Develop/MySheet/docs/excel-interop.md`
      *Why:* Three separate claims in this bullet are now false (not modeled / name and columns dropped /
      `Tabela1[Valor]` does not parse) and one is newly true and dangerous (the Formulas-mode export produces
      a file Excel cannot resolve). The export sentence is the single most important doc addition in this
      phase: it is the only user-visible way S3's scope boundary bites, and without it a host will file it as
      a bug.
- [ ] **24.** Apply the pt-BR mirror edits: /Volumes/Work/Develop/MySheet/docs/pt-BR/excel-interop.md line 68 (the `| **Tabela** do Excel …` row) and the `- **Sem Tabelas do Excel e sem referências estruturadas**` bullet at :244-251, plus the `UnparsableFormula` guidance paragraph at :131. Translate the two English replacements above; the bullet heading becomes "**Tabelas do Excel são carregadas, mas o MySheet nunca escreve uma**". Keep the existing `MergeIntoExcel` sentences at :249-251 verbatim.
      *Files:* `/Volumes/Work/Develop/MySheet/docs/pt-BR/excel-interop.md`
      *Why:* docs/pt-BR is a confirmed full 11-file mirror and docs/pt-BR/README.md:3 states the English
      prevails on divergence — which makes a stale mirror tolerable but sloppy, and these are the two
      paragraphs a Portuguese-speaking host would act on. Anchor on the heading text: this file's line numbers
      have also shifted this session.
- [ ] **25.** Rewrite the blockquote at /Volumes/Work/Develop/MySheet/docs/workbook-and-expressions.md:379-383, whose closing claim "MySheet models the first and not the second" is now false. Replacement: a named range is still not a Table, and MySheet now models both — a name is a static alias for one expression, a table is a named region with named columns, a header row and an optional totals row addressed by its own syntax; they live in separate maps (`Workbook.DefinedNames` / `Workbook.Tables`) and separate grammars (`Sales` resolves as a name, `Sales[Valor]` as a table), so MySheet tolerates one identifier being both even though Excel forbids it; what MySheet does NOT model is a table whose range GROWS as rows are appended — the geometry is whatever `DefineTable` or the `.xlsx` loader recorded. Keep the cross-reference to excel-interop.md#scope-and-limitations and add one to the new `#excel-tables` section. Mirror into /Volumes/Work/Develop/MySheet/docs/pt-BR/workbook-and-expressions.md.
      *Files:* `/Volumes/Work/Develop/MySheet/docs/workbook-and-expressions.md`, `/Volumes/Work/Develop/MySheet/docs/pt-BR/workbook-and-expressions.md`
      *Why:* This blockquote was written in the commit that shipped the UnparsableFormula degradation and is
      the most load-bearing wrong sentence left in the docs. The name/table-collision clause is worth stating
      explicitly because it is a real deviation from Excel in the permissive direction: the loader
      deliberately does not check `DefinedNames` for a colliding table name (Workbook.cs:123), and it cannot
      cause an ambiguity because the two are distinguished by the following `[` at Parser.ParseIdentifier. The
      no-growth clause depends on the table-model-registry phase's decision — if it implements a growing
      range, drop that sentence.

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded with 0 errors. Danfma.MySheet.Excel compiles against the new
      Workbook.DefineTable / Workbook.Tables surface, so a failure here means the table-model-registry phase's
      API does not match TableDefinitionReader.TryMap's construction of TableDefinition.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release`
      → expected: "Test run summary: Passed!" with failed: 0 and total >= 97. Baseline MEASURED before this
      phase: total 88, failed 0, ~774 ms. The delta is the 9 new tests specified above (one of which replaces
      a deleted test), so total should be 88 - 1 + 9 = 96 at minimum, 97 with the export pin.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release -- --treenode-filter "/*/*/TableInteropTests/*"`
      → expected: "Passed!", failed: 0, total >= 21. Baseline MEASURED: total 14. Filter syntax verified
      working in this repo. Use this while iterating on the fixtures; the full-suite run above is the gate.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release`
      → expected: "Passed!", failed: 0. The engine suite must stay green: this phase edits nothing under
      Danfma.MySheet/ except nothing at all, so any failure here belongs to another phase (the registry, the
      lexer or the reference semantics) and must not be attributed to the loader.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier check .`
      → expected: Exit code 0 and no "would be reformatted" lines. The pre-commit hook runs this, so an
      unformatted TableDefinitionReader.cs blocks the commit.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -rn "No Excel Tables\|Sem Tabelas do Excel\|models the first and not the second\|modela o primeiro e não o segundo" docs/ README.md`
      → expected: No matches (grep exits 1). Each of these strings is a sentence this phase must have
      replaced; a hit means a doc edit was skipped.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -n "tokenizer has no" Danfma.MySheet.Excel/ExcelFile.cs tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs`
      → expected: No matches (grep exits 1). The phrase appears today in the UnparsableFormula doc comment
      (ExcelFile.cs:61) and the TableInteropTests class comment (:10-11); both must be reworded.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -c "StructuredFormula" tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs && grep -n "UnsupportedStructuredFormula\|UntokenizableFormula" tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs | wc -l`
      → expected: The first grep must not report any bare `StructuredFormula` identifier surviving on its own
      (only the two new names, which contain the substring — so read the second number instead: at least 12
      lines referencing the two new constants, with `UntokenizableFormula` used exactly twice, at the two
      shared-master tests). If any test still feeds SUM(Tabela1[Valor]) to the degrade path it will fail the
      run above with "Expected 1 but got 0" warnings.

## Risks carried by this phase

- NEW HARD-FAILURE MODE, MEASURED: `part.TableDefinitionParts` throws `InvalidOperationException: Part: /xl/tables/table1.xml doesn't exist in the package` when the relationship survives but the part does not — and `SpreadsheetDocument.Open` (ExcelFile.cs:118) plus `WorksheetParts.First()` both succeed first, so today such a file loads fine. Without the outer try in TableDefinitionReader.Read, this phase turns a loading file into a throwing one. `tdp.Table` adds two more: `XmlException` on non-XML content and `InvalidDataException` ("Cannot load the root element from the part") on a wrong root element. None are covered by the `ParseException or ArgumentException` filter this codebase uses at ExcelFile.cs:194, so copying that filter is the mistake to avoid.
- EXPORT ROUND-TRIP TRAP, and the sharpest edge in S3: after this phase, load(.xlsx with a table) -> `SaveAsExcel(new ExcelExportOptions { FormulaMode = FormulaMode.Formulas })` writes `<f>SUM(Tabela1[Valor])</f>` into a package with NO `<table>` part, so Excel opens it showing #NAME?. Today this is impossible because the formula never parsed. `ExcelExportOptions` has no warning channel (nothing analogous to ExcelLoadOptions.OnWarning), so there is no way to tell the host at export time — documentation and the SaveAsExcel_WritesNoTablePart test are the only mitigations available inside S3.
- A table this loader SKIPS silently changes its formulas from "Excel's cached number" to #NAME?, because a structured reference to an unregistered table parses fine and only fails at evaluation (the DefinedNames model, ExcelFile.cs:144-145). That is more Excel-faithful than the current cached-value fallback — Excel shows #NAME? for a missing table — but on a corrupt file the host sees numbers turn into errors, and ONLY the `InvalidTableDefinition` warning explains why. Hosts that pass no `ExcelLoadOptions` (the `Load(string)` overload at ExcelFile.cs:91) get no signal at all.
- TEN of the fourteen tests in tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs feed `SUM(Tabela1[Valor])` (:21) to the degrade path and will fail the moment the lexer accepts `[`. Swapping the constant to `SUM(Tabela1[@Valor])` fixes all ten but COUPLES them to Phase 1 rejecting `@` with a ParseException; if Phase 1 ever supports the this-row form, those ten fail again. The two shared-master tests are worse: their comments (:245-247, :532) assert a TOKENIZER-stage failure, and `[@Valor]` may well tokenize cleanly, which would silently move them onto the parse-stage path they were written to avoid — hence the separate `SUM({1;2;3})` constant.
- `_xHHHH_` decoding is a judgement call that can misfire. Restricting it to code points < 0x20 plus `_x005f_` makes it safe for any producer that escapes correctly and a no-op for one that escapes nothing, but a column whose literal header text contains `_x000a_` written by a producer that does NOT escape underscores will be silently renamed. The alternative — no decoding at all — leaves a header containing a line break permanently unreachable from a formula, because a raw newline in an XML attribute is normalized to a space by every conformant parser (which is why the escape exists). MEASURED: the SDK does not decode it; injected `Line _x000a_ Break` came back verbatim.
- Column names are stored UNTRIMMED and compared OrdinalIgnoreCase. If Excel trims trailing whitespace when it writes `tableColumn/@name` but the formula carries the untrimmed text (or the reverse), a header with a trailing space becomes unreachable. Trimming defensively is worse: it would break the match for a producer that does preserve the space. No real Excel-produced file was available to settle it.
- docs/excel-interop.md and tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs BOTH changed on disk during this analysis (a fourth warning kind `UnparsableCellLiteral` was added to ExcelFile.cs, five decode tests were added, and two fixtures switched their cached `<v>` from "42" to a deliberate "999"). Every line number cited for those three files will have drifted; anchor edits on member names, constant names and heading text, and re-read before editing.

## Open questions owned by this phase

- Does Excel escape a control character in `tableColumn/@name` as `_x000a_`, as `&#10;`, or not at all? The specified decoder is correct for the first two (it is a no-op for the second, since the SDK decodes numeric entities — MEASURED). If Excel writes a raw newline, the attribute-value normalization makes the header and the formula text genuinely unmatchable and the column is unreachable by design. Settle it by opening one real Excel-authored file whose table header contains an Alt+Enter line break and reading `xl/tables/table1.xml`.
- What does Excel return for `Table[#Headers]` on a `headerRowCount="0"` table, and for `Table[#Totals]` when `totalsRowCount` is 0? Assumed #REF! for both. The loader's job is only to record `headerRow = 0` / `totalsRow = 0`; the reference-semantics phase owns the error. Verify against Excel or Aspose.Cells before that phase hard-codes it.
- `headerRowCount="2"` is schema-valid (MEASURED: the OpenXmlValidator accepted it) but Excel's ListObject header is a single row and Excel writes only 0 or 1. This design SKIPS the table with a warning. The alternative — treat the LAST of the header rows as the header — is a guess nobody can check without a producer that emits it. Is skipping acceptable, or should it degrade to headerRowCount 1?
- A header-only table (`ref="A1:B1"`, `headerRowCount=1`, zero data rows) is schema-valid (MEASURED). This design passes the empty geometry through, which means `TableDefinition` must accept `LastDataRow == FirstDataRow - 1` rather than throwing — a hard requirement on the table-model-registry phase. What does Excel return for `Tabela1[Valor]` there: #REF!, or a one-cell range at the phantom first data row?
- When a producer writes `name` and `displayName` DIFFERENT (MEASURED: legal, since only `displayName` is schema-required, and the SDK reads both independently), which does Excel use to resolve `Foo[Col]`? This design uses `displayName` and falls back to `name` only when `displayName` is absent, justified by the schema requiring `displayName`. No file exists to test the divergent case; ECMA's prose describes BOTH attributes as formula-relevant, so the prose alone does not settle it.
- Should a follow-up phase teach `SaveAsExcel` to emit the `<table>` part, closing the Formulas-mode export trap? S3 says no, and the cost is a documented file that Excel opens with #NAME?. If the answer is yes, the write side is small — `AddNewPart<TableDefinitionPart>()` plus a `<tableParts>` child on the worksheet, following the SharedStringRegistry.WriteTo precedent at ExcelExport.cs:327-330 — but it lands after the streaming `WriteWorksheet` has already closed `</worksheet>`.
- Should the eight non-table degrade tests (TableInteropTests.cs:303-457, :460-508, :589-623) move to their own `UnparsableFormulaInteropTests.cs`? They use the table fixture only as a scratch sheet, and after this phase the file's name misdescribes most of its contents. Deferred as pure tidying with real transcription risk.

## Phase Summary

_(write when phase completes)_

## CONTROLLER RULINGS after the re-verification (2026-09-11) — binding

Read the re-verification below first; these settle its T0 questions.

1. **File ownership between Phase 4 T6a and this phase.** Phase 4 T6a keeps ONLY item 17's `StructuredFormula` constant swap, correction B2's arithmetic, and the `WorksheetStreamLoader.cs:521` comment (all become false the moment `[` lexes, which is Phase 4's own doing). **Every doc site is this phase's** — `docs/excel-interop.md` and pt-BR, both `workbook-and-expressions.md` blockquotes (`:934-939`, `:993-999` and pt-BR twins), `docs/serialization.md:337` / pt-BR `:367` — because this phase is what makes their sentences true. Phase 4 T6a's docs items are struck from its brief. `TableInteropTests.cs` is edited by three tasks in TIME order, never concurrently: Phase 6 T1 flips `:123` NOW; Phase 4 T6a swaps the constant later and rebases onto it; Phase 6 T2 rewrites the evaluation tests last.
2. **Two constants** (`StructuredFormula` for `Tabela1[Valor]`, `ImplicitTableFormula` for `[[#This Row],[Valor]]`) — the design's split wins, and the fixture text is the `#This Row` item-list Aspose actually stores, not the hand-typed `[@Valor]`.
3. **Item 19's dangling-part test is DELETED.** `ExcelFile.Load` throws on that file today (OpenXml 3.5.1, `ExcelFile.cs:128`) before any reader runs; changing that is out of scope. The two failure modes this phase introduces are garbage XML and a wrong root, both measured.
4. **Items 1 and 6 are DELETED; items 4-5 are rewritten as the audit's §5 first bullet**: `TryMap` calls the shipped `Workbook.DefineTable(name, sheetName, ref, columns, hasHeader, hasTotals)` inside the per-table try and catches `ArgumentException` (which covers `Table.Validate`'s every reject) plus `XmlException`, `InvalidDataException`, `InvalidOperationException`. The `$A$1:$B$3` form is ACCEPTED, as the overload accepts it.
5. **M2 is moot** (`Table.SheetName` exists). **M3's mechanism is corrected**: the validator is `Table.ValidateName`, the Excel-legal names it rejects are `My\Table` and `TRUE`, and the outcome is two warnings per such table with the cached number surviving — a structural limit under P0, documented as "a name MySheet's tokenizer could never read".
6. **Header-only table:** the oracle answers `SUM(Tabela1[Valor])` = 0 and `ROWS(Tabela1[#Data])` = 0, NOT `#REF!`. `Table.TryGetColumnRange` returning false for `DataRowCount == 0` hands Phase 5 a decision; Phase 5's re-verification carries it. This phase only registers the table.
7. **The bracket scanner must accept a raw newline inside `[…]`** — Aspose stores `SUM(Tabela1[Line\nBreak])` with a real `\n` in `<f>`. Added to Phase 4's T1/T2/T3 briefs; without it item 7's `_x000a_` decode is unreachable from a real file.
8. **Every committed Aspose fixture has its `Evaluation Warning` sheet removed** through the SDK after generation, and no pin asserts the sheet count.
9. **This phase's T1 (the reader) and T3's export/merge half start NOW**, in a worktree off `main`, concurrently with Phase 11c and before Phase 4 — they need nothing from either. The branch is NOT merged before the 3.20.0 release; it rebases onto whatever `main` is when Phase 4's T1 lands. T2 waits for Phases 4 and 5; T4 is last.

## Re-verification against main @ a99ab43 (2026-09-11)


Tree: `/Volumes/Work/Develop/MySheet`, branch `main`, clean. Read-only: nothing in the repo was modified; every experiment ran under `/private/tmp/claude-501/-Volumes-Work-Develop-MySheet/e0408f78-ba55-43f2-ba34-329382660371/scratchpad/p6-reverify/` (fixtures in `fixtures/`, Aspose builder in `aspose/`, loader probe in `loadprobe/`, logs `aspose-out.log`, `aspose-out2.log`, `loadprobe-out.log`). Oracle: Aspose.Cells 26.6.0 via a copy of `/tmp/aspose-probe-fable`, PLAIN entry, `ListObjects.Add` + `DisplayName`. Loader side: DocumentFormat.OpenXml **3.5.1** (the version `Danfma.MySheet.Excel` resolves today), ClosedXML 0.105.0 (the test project's version). Baselines measured on this tree: Release build 0 warnings; Excel suite **93 / 0** (design says 88); `TableInteropTests` **14** methods; `ExcelFile.cs` 209 lines.

Tags: [Certain] = measured or read in the tree this session; [Likely] = strong inference from measured facts; [Guessing] = marked as such.

## Ranked findings — what does the most damage unnoticed

1. **[Certain] The design targets a type that does not exist and a validator that is not the one in the path.** Items 4, 5, M2, M3 and the verification plan construct a `TableDefinition` with `headerRow/firstDataRow/lastDataRow/totalsRow/firstColumn/lastColumn`; Phase 3 shipped `Danfma.MySheet/Table.cs:26-35`: `Table(Name, SheetName, FirstRow, LastRow, FirstColumn, HasHeaderRow, HasTotalsRow, ColumnNames)` with the data band DERIVED (`:39-68`). M3 argues name validation from `NamedReferences.ValidateName` (`Workbook.cs:549` → real `:604`, and it is `DefineName`'s validator); `DefineTable` (`Workbook.cs:659-680`) calls `Table.Validate()` → `Table.ValidateName` (`Table.cs:149-176`), a DIFFERENT rule that accepts `Tabela1`/`Table1` and rejects `C`/`R`, R1C1, `TRUE`/`FALSE`, a backslash and anything `Parser.IsExcelGridCellReference` accepts. M2 (the sheet name is never recorded) is moot: the record has `SheetName` and `Validate` rejects a blank one (`Table.cs:248-256`).
2. **[Certain] Items 1 and 6 are redundant with what Phase 3 shipped, and item 6's `$` rule contradicts it.** `Workbook.DefineTable(string name, string sheetName, string reference, IReadOnlyList<string> columnNames, bool hasHeaderRow = true, bool hasTotalsRow = false)` (`Workbook.cs:690-740`) already parses the xlsx `ref` string through `TryParseTableReference` (`:744-772`, over `CellAddress.TryParseA1`), rejects a width/column-count mismatch (`:718-726`), and throws `ArgumentException` for everything the design's `TryParseReference` was written to catch. Measured on the shipped overload: `NOT-A-RANGE` → ArgumentException "is not an A1 range"; `Data!A1:B3` → rejected; `A1:B3+1` → rejected; **`$A$1:$B$3` → ACCEPTED** (item 6 says reject `$`); `A1:B1` header-only → accepted (the open question 4 "hard requirement on Phase 3" is met); `A1:B1` with header+totals → rejected "spans 1 row(s) but … need 2"; empty column name → rejected; `c`/`C` duplicates → rejected. `CellId.TryParse` (item 1, plus reimplementing `Parse` for six call sites) buys nothing the overload does not already do; drop both items and have `TryMap` call the overload inside the per-table try.
3. **[Certain] The "NEW HARD-FAILURE MODE" is not new, and the outer try in item 4 cannot contain it.** With `xl/tables/table1.xml` deleted and the relationship kept, `SpreadsheetDocument.Open` and `WorksheetParts.First()` do succeed (the design's measurement reproduces), but **`ExcelFile.Load` throws TODAY** on that file: `InvalidOperationException: Part: /xl/tables/table1.xml doesn't exist in the package`, raised from `workbookPart.Workbook` at `ExcelFile.cs:128` (stack: `StrictNamespaceFeature.get_Found` → `OpenXmlPackage.LoadAllParts` → `OpenXmlPart.Load`), before any sheet or table reader runs. So item 4's two-stage try is justified only by the corrupt-XML cases (still real: garbage → `XmlException`, wrong root → `InvalidDataException`, both measured on 3.5.1), the risk bullet's "today such a file loads fine" is false, and item 19's second test `Load_TablePartMissingFromThePackage_WarnsWithTheSheetAsSubject` (one warning, `Subject == "Data"`, `B3 == 32`) **cannot pass** without changing `ExcelFile.Load` itself — out of this phase's scope. Delete that test or rescope it to "the load throws the same exception before and after this phase".
4. **[Certain] Phase 4's T6a and Phase 6 own the same files with conflicting edits.** Phase 4 item 17 / T6a claims `TableInteropTests.cs` (ONE constant, `SUM(Tabela1[@Valor])`), the `WorksheetStreamLoader.cs:520-522` comment, `docs/excel-interop.md:66, :125, :243-246` (+ pt-BR) and the `workbook-and-expressions.md` Tables blockquote; Phase 6 items 11 (TWO constants), 21-25 rewrite the same sites. Phase 4's B2 also adds a `SUM(Tabela1[Valor])` test asserting `#REF!`/`#NAME?` that Phase 6 item 13 deletes. Whoever merges second rebases onto rewritten text; the controller must assign each site to one phase before either brief is dispatched.
5. **[Certain] Real files carry `[[#This Row],[Valor]]`, not `[@Valor]`.** Aspose stores `=SUM(Tabela1[@Valor])` in the sheet XML as `<f>SUM(Tabela1[[#This Row],[Valor]])</f>` and `=Tabela1[@Valor]*2` as `<f>Tabela1[[#This Row],[Valor]]*2</f>` (f8, `Data!D2` = 10, `I2` = 20). Item 14's fixture (`Tabela1[@Valor]*2`) and item 3's doc text exercise the hand-typed form only; the shape the loader meets from Excel/Aspose is the `#This Row` item-list. Phase 4 item 6 lists `#This Row` among the specifiers but its classification is stated only for a leading `@`; item 14 must inject BOTH spellings and item 3 must name `[#This Row]`.
6. **[Certain] The totals-row blocker B1 is still real, and Aspose changes which correction applies.** ClosedXML 0.105.0 re-measured: `<x:c r="B4" s="0"><x:f>SUBTOTAL(109,[Valor])</x:f></x:c>` — no `<v>`; today's load gives exactly one `UnparsableFormula/B4`, B4 Blank. Aspose writes the same formula WITH a cache: `<c r="B4"><f>SUBTOTAL(109,[Valor])</f><v>42</v></c>` (part: `ref="A1:B4" totalsRowCount="1"`, no `totalsRowShown`, `autoFilter ref="A1:B3"`). Either way `[Valor]` is S1 out-of-scope, so `warnings.Count == 0` fails; with an Aspose-shaped fixture the honest assertions are ONE `UnparsableFormula` at B4 plus `D1` = 42, `ROWS([#All])` = 4, `SUM([#Totals])` = 42 (the cached 42 makes it hold), `SUM([[#Data],[#Totals]])` = 84.
7. **[Certain] M4's corrected grep pattern matches nothing either.** Phase 3 rewrote both blockquotes: `docs/workbook-and-expressions.md:934-939` now reads "MySheet models names *and* the table model … but neither the structured-reference syntax nor the self-growing range" and pt-BR `:986-992` "modela os nomes *e* o modelo da tabela … mas nem a sintaxe de referência estruturada". `grep -rn "No Excel Tables\|Sem Tabelas do Excel\|models the first and not the second\|modela o primeiro e não o segundo\|modela o primeiro\|e não a segunda" docs/ README.md` → exit 1 TODAY, so verification step 6 passes before any edit is made. The sentences this phase must retire are now different (list under §5).
8. **[Certain] Two doc sites the design does not know exist assert the opposite of this phase.** `docs/workbook-and-expressions.md:993-999` (the "## Tables" blockquote, Phase 3): "`ExcelFile.Load` does not populate the registry from an xlsx `<table>` part either"; pt-BR `:1045-1052`. `docs/serialization.md:337` "(`ExcelFile.Load` does not populate the …" and pt-BR `:367` "ainda não preenche o registro". None is in items 22-25.
9. **[Certain] Open question 1 is settled: Excel-style producers write `_x000a_`.** Aspose wrote `<tableColumn id="4" name="Line_x000a_Break"/>` for a header typed with `\n`, the SDK hands it back verbatim (`cols=[…|Line_x000a_Break|…]`), and the FORMULA is stored with a raw newline: `<f>SUM(Tabela1[Line\nBreak])</f>` = 300. Item 7's decode is therefore necessary (not a judgement call) — and it creates a Phase 4 requirement nobody has stated: the bracket scanner must accept a raw newline inside `[…]`.
10. **[Certain] Aspose accepts two table names `Table.ValidateName` rejects, and one of them is also unlexable.** `DisplayName = "My\Table"` → saved as `name="My\Table" displayName="My\Table"`, `=SUM(My\Table[Valor])` = 42; `DisplayName = "TRUE"` → `=SUM(TRUE[Valor])` = 42. `DefineTable("My\\Table", …)` and `("TRUE", …)` both throw `ArgumentException` "'…' is not a valid table name" (measured). Today's loader on those files: `My\Table[Valor]` → `UnparsableFormula` "Unexpected character '\' (at position 6)" (the tokenizer, not the bracket), `TRUE[Valor]` → "Unexpected character '[' (at position 8)". So M3's scenario is real and its outcome is two warnings per such table (one `InvalidTableDefinition`, one `UnparsableFormula`) with Excel's cached number surviving — acceptable under P0 as a structural limit (the tokenizer never reads `\` or `TRUE` into an identifier, which is exactly why `Table.ValidateName` rejects them), but item 2's doc text and item 22's row must say so. Aspose rejects `T1`, `A1`, `R1C1` ("Invalid text(cell reference) for the defined name") and `My Table` — consistent with `ValidateName`.
11. **[Certain] The export trap and the merge claim, measured.** `SaveAsExcel` of a loaded table workbook (ValuesOnly and Formulas) writes **no** `TableDefinitionPart` and no `<tableParts>` for `Data`; a hand-registered `DefineTable` is also dropped (and no defined name is written for it). `MergeIntoExcel` into a copy of the same file: `TableDefinitionParts = 1`, `ref="A1:B3"`, `<tableParts>` PRESENT — the table survives, for the reason Phase 4's audit gave (merge rewrites only the worksheet stream, `ExcelMerge.cs:63` + `MergeWorksheet` `:209` copies every non-`sheetData` element with `WriteNode`), NOT because of `:325`/`:438` (row/cell copies). Item 20's "already covered" is false: no test in `ExcelMergeTests.cs` touches a table part.
12. **[Likely] Item 11's constant swap collides with Phase 4's, and its arithmetic is off.** 12 usages in 11 of 14 methods (`:138, :172, :200, :221, :222, :259, :312, :344, :379, :402, :432, :542`), not "ten of the fourteen". The two-constant split is the better design (Phase 4's re-verification notes `[@Valor]` may tokenize cleanly, which moves the two shared-master tests off the tokenizer path their comments claim) — but only one phase can own the constant.
13. **[Certain] Every Aspose-produced fixture carries an extra sheet.** Unlicensed Aspose adds a worksheet named `Evaluation Warning` on save; `ExcelFile.Load` loads it as a third sheet (`sheets: Data, Evaluation Warning`). A committed end-to-end fixture must either have that sheet removed through the SDK after generation or the pin must not assert the sheet count. The design's in-test ClosedXML fixtures (`WriteTableFixture`) do not have this problem.

## 1. Anchors, exhaustive

Exact and unmoved (re-read): `ExcelFile.cs` — `:15-22` options doc, `:29-33` Subject doc, `:45-77` enum, `:59-67` UnparsableFormula doc (`:61-62` "tokenizer has no"), `:62-66` fallback sentences, `:77` `UnparsableCellLiteral`, `:80-87` class doc (`:82-84` streaming claim), `:91` `Load(string)`, `:118` `Open`, `:128` sheet loop, `:130-136` skip rule, `:138` `sheet`, `:139` `worksheetPart`, `:141` `WorksheetStreamLoader.Load`, `:144-145` comment + `LoadDefinedNames`, `:187-190` (call is `:188-191`), `:194` `ParseException or ArgumentException`, `:199-205` warning invoke, "209 lines" (`ExcelFile.cs` last changed 507477b, 2026-09-08, the commit the design was written against). `CellId.cs:8-20`, `:19`. Call sites `ExcelExport.cs:166`, `ExcelMerge.cs:128`, `:399`, `SharedFormulaShifter.cs:21-22`, `WorksheetStreamLoader.cs:151`, `:180-181`, `:376`. `WorksheetStreamLoader.cs:574-576`. `Indirect.cs:16`, `:58-66`. `TableInteropTests.cs` `inject` mechanism `:46-53`. `docs/excel-interop.md:66` (row) and `:124-127` (guidance paragraph); pt-BR `:68`, `:131`.

Moved or wrong:

| Design cites | Now | Note |
| --- | --- | --- |
| `TableDefinition` (items 4, 5, M2, M3, verification 1, open q. 4) | `Danfma.MySheet/Table.cs:26-35` record `Table` | eight positional members, geometry derived `:39-68`, `TryGetColumnRange` `:114-133` (`[#Data]` only, false when `DataRowCount == 0`) |
| `Workbook.cs:549` `NamedReferences.ValidateName` | `:604` (in `DefineName`); `DefineTable` `:659-680` calls `table.Validate()` `:664` → `Table.ValidateName` `Table.cs:149` | M3 cites the wrong validator |
| `Workbook.cs:551`, `:584` (last-one-wins) | `:607` (`DefinedNames[name] = reference`), second overload `:620-646` (`:641`) | `DefineTable` also replaces: `_tables[registered.Name] = registered` `:675`, doc `:649-650` |
| `Workbook.cs:123` `DefinedNames` | `:131` | `Tables` property `:153`, `ThrowIfNameTaken` `:777-793` (symmetric name/table collision, NOT in the design) |
| `NamedReferences.cs:164-172` | `Danfma.MySheet/Expressions/NamedReferences.cs:166-185`, message `:177-182`, `IsValidName` `:186-206` | file path was never right (Phase 3 re-verification says the same) |
| `ExcelExport.cs:327-330` `SharedStringRegistry.WriteTo` | `:324-335` (`AddNewPart` `:331`) | |
| `ExcelMerge.cs:325`, `:438` "copies non-owned nodes verbatim" | lines exist and are row/cell `WriteNode`s; the element-level pass-through that keeps `<tableParts>` is `MergeWorksheet` `:209`; the part survives because only the sheet stream is rewritten | conclusion true, citation wrong |
| `TableInteropTests.cs:21` const | `:23` | 12 usages / 11 methods |
| `:99-114` baseline test | `:101-130` (`Load_TableWithOrdinaryFormulas_LoadsAsAPlainRange_WithoutWarnings`) | Phase 3 (9b2189c) already changed `:108` → `:123-124`: asserts `workbook.Tables.Count == 0` AND `DefinedNames.ContainsKey("Tabela1")` false; comment `:107` → `:121-122` now says "The loader does not populate the table registry" |
| `:94` 999 lie | `:105-106` | |
| `:128-147` delete target | `:132-163`; `:131-137` → `:147-153` | |
| `:162, :193, :213, :243, :303, :335, :367, :392, :418, :511` | `:166, :197, :217, :247, :307, :339, :371, :396, :422, :515` (+4) | |
| `:245-247` tokenization comment | `:250-251`; `:532` → `:536` | |
| `:575-580` edit-and-recalc technique | `:579-583` | |
| `:8-18` class comment | `:8-20`; Phase 3 already rewrote it — "MySheet has no table model" is GONE, only "the tokenizer has no `[`" (`:12`) remains false-after-Phase-4 | item 21's "every sentence of it is now false" is stale |
| `:303-457, :460-508, :589-623` (open q. 7) | `:307-461, :464-512, :593-627` | |
| `docs/excel-interop.md:235-241` limitations bullet | `:243-250`; `MergeIntoExcel` sentences `:239-241` → `:248-250` | bullet heading now "**The loader ignores Excel Tables, and structured references do not parse**" (Phase 3), not "No Excel Tables …" |
| pt-BR `excel-interop.md:244-251`, `:249-251` | `:260-268`, `:266-268` | heading "**O carregador ignora Tabelas do Excel, e referências estruturadas não são reconhecidas**" |
| `docs/workbook-and-expressions.md:379-383` blockquote | `:934-939`, rewritten by Phase 3 | plus NEW `:993-999` (Tables blockquote) not in the design |
| pt-BR `workbook-and-expressions.md:396` | `:986-992`, rewritten | plus NEW `:1045-1052` |
| item 10 path `/Volumes/Web/../Danfma.MySheet.Excel/ExcelFile.cs` | typo; `:80-87` itself is exact | |
| verification 7 `grep -n "tokenizer has no" ExcelFile.cs TableInteropTests.cs` | misses the third site `WorksheetStreamLoader.cs:521` (which Phase 4 T6a claims) | |
| counts: 88 Excel, "total >= 97", "88 - 1 + 9 = 96" | 93 today; 93 − 1 + 9 = 101, 102 with the export pin | `TableInteropTests` 14 still exact |
| M4 "pt-BR pattern `modela o primeiro` … verified it matches `:396` today" | matches nothing today | see finding 7 |
| "SDK version" (implicit) | DocumentFormat.OpenXml 3.5.1 | the dangling-part behaviour differs from what the design measured |

## 2. Behavioural claims, checked

| Claim | Result |
| --- | --- |
| Loader with a `<table>` part today: warns, ignores, or crashes? | [Certain] Ignores the part; every structured-reference cell degrades with `UnparsableFormula` "Unexpected character '[' (at position N)" and reads the cached `<v>`. Measured on all nine Aspose fixtures: `Tables.Count = 0`, `DefinedNames = 0`, cached numbers surface (f1 `D1` = 42 … `D9` = 42; `D6` INDIRECT = `#REF!` because the inner parse fails; `D7` = `#REF!` from the cached `t="e"`). No crash. |
| Dangling table relationship loads today | [Certain] FALSE on OpenXml 3.5.1 — `ExcelFile.Load` throws `InvalidOperationException` from `ExcelFile.cs:128` (see finding 3). |
| `tdp.Table` on garbage → `XmlException`; wrong root → `InvalidDataException` | [Certain] Both reproduce: "Data at the root level is invalid. Line 1, position 1." / "Cannot load the root element from the part. The part contains invalid data." Today's loader loads both files fine (it never touches the part), so THESE two are the failure modes this phase introduces. |
| `ref="NOT-A-RANGE"` comes back verbatim; missing `ref` → null; `headerRowCount="2"` accepted; missing `displayName` → `DisplayName` null with `Name` set | [Certain] All reproduce. `$A$1:$B$3` also comes back verbatim. |
| SDK decodes entities, leaves `_xHHHH_` alone | [Certain] `A&amp;B &lt; C &quot;q&quot; &#39;s` → `A&B < C "q" 's`; `Line _x000a_ Break _x005f_x` verbatim. |
| `HeaderRowCount` null on a normal table; `TotalsRowCount=1`/`TotalsRowShown=null` with totals; `TotalsRowShown=False`/`TotalsRowCount=null` without | [Certain] Aspose writes exactly the ClosedXML pattern the design measured: `totalsRowShown="0"` and no count without totals; `totalsRowCount="1"` and no `totalsRowShown` with. `ref` INCLUDES the totals row (`A1:B4`, `autoFilter ref="A1:B3"`). |
| Headerless table: ClosedXML rewrites to `ref="A2:B3" headerRowCount="0"` with the REAL names and deletes row 1 (M1) | [Certain] Aspose does the same: `ref="A2:B3" headerRowCount="0"`, names `Item`/`Valor`, sheet `A1`/`B1` Blank on load. M1's correction stands; the "Column1/Column2" fixture in item 17 is wrong for both producers (Aspose rejects `Tabela1[Column1]`: "Invalid table column"). |
| Aspose writes `name` and `displayName` | [Certain] Both, identical, on every fixture. No divergent-case file (open question 5 stays open). |
| `SaveAsExcel` drops the table part | [Certain] Both modes; hand-registered table also dropped. `<f>` round-trip in Formulas mode cannot be measured yet (the formula degrades to a literal today: `<x:c r="D1"><x:v>42</x:v></x:c>`). |
| `MergeIntoExcel` keeps the part | [Certain] Part + `<tableParts>` survive; `B2` rewritten to 8 and `D1` to the literal 42 (the documented stale-number limitation). |
| Excel test baseline 88 / TableInteropTests 14 | [Certain] 93 / 14. |
| Streaming cost (62 KB / 0.23 ms vs 29.8 MB / 340 ms) | not re-measured; the placement argument does not depend on the exact numbers. |
| `Indirect.cs` parses via `ParseFormulaBody` then `TryResolveReference` | [Certain] `:58-68`; `IsVolatile => true` `:16`. |
| "ExcelLoadOptions has no analogue on export" | [Certain] `ExcelExportOptions` `:26` has only `FormulaMode`. |
| No test pins the `ExcelLoadWarningKind` member count | [Certain] grep: none. Appending is safe. |

## 3. Dependencies

- **Phase 3 (Complete): everything items 2, 4, 5, 7, 8, 9, 10 need is in the tree** — `Workbook.Tables` (`:153`, `IReadOnlyDictionary`, case-insensitive), `DefineTable(Table)` `:659`, the A1 overload `:690`, `Table` `Table.cs:26`. [Certain] What Phase 3 changed relative to the design's assumptions: (a) `DefineTable` REPLACES on a duplicate name (design assumed it might throw — its `ContainsKey` first-wins check still works and is unreachable from a valid file, Excel forbids duplicates); (b) `DefineTable` throws for a name a DEFINED NAME already holds (`ThrowIfNameTaken`), and `LoadDefinedNames` runs AFTER the sheets, so a corrupt file with both gets `InvalidTableDefinition`-free registration and then an `InvalidDefinedName` warning from `DefineName` — the design never states this ordering consequence; (c) `Table.Validate` throws `ArgumentOutOfRangeException` for rows/columns < 1 — a subclass of `ArgumentException`, so item 4's filter covers it; (d) `DefineTable` snapshots `ColumnNames` (`:670-673`), so the loader may pass any list.
- **Phase 4 T1 (`TableReference` node) + T5 (parser arms) + Phase 5 (resolver): needed by items 13-18 (evaluation assertions), item 11 (the constant swap breaks only when `[` lexes), item 3/21 (doc text that says the shapes parse), items 22-25 (docs that say formulas resolve), and item 20's Formulas-mode `<f>` half.** [Certain] Nothing in items 2, 4-10 references the node.
- **Phase 6 CAN start before Phase 4 merges** — the reader (`TableDefinitionReader.Read`, `ExcelFile.cs:141` call, enum member, `_xHHHH_` decode) is testable at the REGISTRY level: after `Load`, assert `workbook.Tables["Tabela1"]` equals `Table("Tabela1","Data",1,3,1,true,false,["Item","Valor"])`, the geometry via `FirstDataRow/LastDataRow/HeaderRow/TotalsRow`, the warnings for malformed parts, and that `DefinedNames` stays empty. Every existing `TableInteropTests` test keeps passing (they assert `Tables.Count == 0` only at `:123` — that one line flips). What MUST wait: items 13-18 and the docs claims that a structured reference evaluates. [Certain] Phase 4's re-verification already reached the same conclusion ("Phases 4, 5 and 6 can be developed concurrently and merged in any order, provided no release ships before Phase 6").
- **Ordering hazard**: `TableInteropTests.cs:123` (`Tables.Count == 0`) turns red the moment the reader lands, so Phase 6's first task rewrites it as item 12 says; Phase 4 T6a must not touch that assertion.

## 4. Blocker and majors, re-judged

- **B1 — still true, correction updated** (finding 6). ClosedXML's no-`<v>` totals cell reproduces byte for byte. With Aspose's `<v>42</v>` the "inject a literal" workaround is unnecessary; assert one `UnparsableFormula/B4` instead of zero warnings, and `D3` = 42 holds through the cache. Note the coincidence trap the master plan warns about: 42 in the cache equals the true `SUM([#Totals])`; make the cached totals value a lie (Aspose cannot — it computes it — so for the in-test ClosedXML fixture inject `<c r="B4"><v>999</v></c>` and assert 999, which is what the cache-vs-evaluate distinction needs).
- **M1 — still true, and now measured on the oracle too** (Aspose = ClosedXML behaviour; `SUM(Tabela1[Valor])` = 42, `ROWS([#All])` = 2, `ROWS([#Data])` = 2, `COUNTA([#Headers])` = 1 — an error counted as one element, per Phase 4's note; `SUM(Tabela1[Item])` = 0). Correction stands; `InjectTablePart` is still needed only for malformed parts.
- **M2 — MOOT.** `Table.SheetName` exists and `Validate` rejects a blank one. Replace with: "`TryMap` passes `sheetName` (`ExcelFile.cs:138`) as `Table.SheetName`; pin it in `Load_Table_RegistersItInWorkbookTables_WithoutWarnings` and with the cross-sheet fixture (f6: a formula on `Report` sums `Data`'s table — oracle 42; `Data!Tabela1[Valor]` is stored back by Aspose WITHOUT the qualifier)".
- **M3 — still true in conclusion, wrong in mechanism** (finding 10). The validator is `Table.ValidateName`, the rejected-but-Excel-legal names are `My\Table` and `TRUE` (both measured accepted by Aspose), and the `ArgumentException` escapes only if step (3) sits outside the try. Keep the correction; replace the cited rule and add `TRUE` as a second row. `InvalidTableDefinition`'s doc must say "a name MySheet's tokenizer could never read (a backslash, `TRUE`/`FALSE`)" rather than "a name Excel's own name rules reject" — Excel accepts both.
- **M4 — stale in both directions** (finding 7). Neither the original nor the corrected pattern matches anything on main. Replace the whole verification grep with the sentences that are false today (see §5).

## 5. What the design does not cover and now must

- **The `Table` record and `DefineTable`'s validation** (findings 1, 2). Rewrite items 4-6 as: `TryMap` = name (`displayName` ?? `name`), `headerRows = HeaderRowCount ?? 1`, `totalsRows = TotalsRowCount ?? 0`, reject `headerRows > 1` and `totalsRows > 1` with a reason, decode column names (item 7), then `workbook.DefineTable(name, sheetName, table.Reference?.Value ?? "", columns, headerRows == 1, totalsRows == 1)` INSIDE the per-table try, catching `ArgumentException` (covers `ArgumentOutOfRangeException` and every `Table.Validate` reject: name, blank sheet, geometry, width mismatch, empty/duplicate column) plus `XmlException`, `InvalidDataException`, `InvalidOperationException`. The warning `Detail` is then `DefineTable`'s message, which already names the table and the offending member. Items 1 and 6 disappear; item 19's `NOT-A-RANGE` test keeps its assertions (the reject now comes from `Workbook.cs:712-716`).
- **A table name `ValidateName` rejects** → one `InvalidTableDefinition` (Subject = displayName), table skipped, its formulas ALSO fail to lex (`\`) or parse (`TRUE[`) → `UnparsableFormula` + cached value. Pin both warnings on the `My\Table` fixture (`f5-name-My_Table.xlsx`: oracle `D1` = 42, cached 42).
- **Totals/header flags versus `Table`'s geometry**: `headerRowCount` ∈ {absent→1, 0, 1}, `totalsRowCount` ∈ {absent→0, 1}; `ref` spans header through totals — exactly `Table.FirstRow..LastRow` with `HasHeaderRow/HasTotalsRow`. Measured mapping rows: f1 `A1:B3` → (1,3,header,no totals) FirstDataRow 2, LastDataRow 3; f2 `A1:B4 totalsRowCount=1` → (1,4,header,totals) data 2..3, TotalsRow 4; f3 `A2:B3 headerRowCount=0` → (2,3,no header) data 2..3, HeaderRow null; f7 `A1:B1` → (1,1,header) DataRowCount 0.
- **Header-only table** (open question 4): Aspose `SUM(Tabela1[Valor])` = **0** and `ROWS(Tabela1[#Data])` = **0** on `ref="A1:B1"`. `DefineTable` accepts it (measured). `Table.TryGetColumnRange` returns false for `DataRowCount == 0` and its doc hands the caller "the `#REF!` decision" — that is a Phase 5 divergence to record there (P0: the oracle says 0, not `#REF!`); Phase 6 only registers.
- **`[#Totals]` with no totals row** = `#REF!` (f1 `D7`, cached `t="e"`); `[[#Data],[#Totals]]` shrinks to the data body (42 without totals, 84 with) — confirms Phase 4 Ruling 2. `SUM(Tabela1)` is STORED as `SUM(Tabela1[])` in the sheet XML (f1 `D9`), so the loader will meet the empty-specifier form Phase 4 measured as accepted.
- **The serialized third member and the two wire goldens**: [Certain] this phase adds no `Workbook` member and no union tag, so `CellStoreTests.Wire_IsByteIdentical_AfterNumericKeys` and `SheetNameInterningTests.Wire_IsByteIdentical_AfterInterning` are untouched; a table registered by the loader round-trips through `Save`/`Load` on Phase 3's member (Phase 3 measured `Tabela1` round-tripping). Worth one pin: `ExcelFile.Load(fixture).Save()` → `Load` → `Tables["Tabela1"]` equal — the loader is the first REAL producer of that member. `docs/serialization.md:337` and pt-BR `:367` must stop saying `ExcelFile.Load` does not populate it (`:345` is Phase 4's sentence).
- **Docs, both twins**: replace verification step 6 with a grep for the sentences that are false after this phase — English: `does not populate` (`workbook-and-expressions.md:997`, `serialization.md:337`), `The loader ignores Excel Tables` (`excel-interop.md:243`), `the LOADER does not populate it yet` (`:66`), `cannot be parsed and degrades` (`:66`), `A structured reference into an Excel\nTable is the common cause` (`:125-126`); pt-BR: `não preenche` (`workbook-and-expressions.md:1049`, `serialization.md:367`), `O carregador ignora Tabelas` (`excel-interop.md:260`), `o CARREGADOR ainda não o preenche` (`:68`), `é a causa mais comum` (`:133`). `docs/pt-BR/README.md:3` still says English prevails. Diff the twins afterwards (lessons: a bounded script replacement swallowed four sentences once).
- **The `[#This Row]` shape** (finding 5) in item 3's doc text, item 14's fixture, item 22/23's out-of-scope list.
- **`_x000a_` and the raw newline in `<f>`** (finding 9): item 7 stays; add to Phase 4's brief that the bracket scanner must accept `\n` inside a specifier, or `Line\nBreak` degrades with `UnparsableFormula` and the decode is unreachable from a real file.
- **The end-to-end release pin**: one Aspose-authored `.xlsx` committed under the Excel test project (f1 shape plus the totals and names variants), loaded through `ExcelFile.Load` with `OnWarning`, asserting the oracle's numbers below and `warnings` limited to the expected `UnparsableFormula` rows. Strip or tolerate the `Evaluation Warning` sheet (finding 13). This is the only test that proves "a real xlsx with `SUM(Tabela1[Valor])` no longer says `#NAME?`" — the master plan's release gate.

## 6. Task decomposition (disjoint files) and the rows each task pins

Expected values are Aspose.Cells 26.6.0, PLAIN entry, measured this session; fixture shapes as built by `aspose/Program.cs`. `Tabela1` = `Data!A1:B3`, header `Item`/`Valor`, rows `a`/10 and `b`/32.

**T0 — controller, before dispatch.** Assign ownership of `TableInteropTests.cs`, `WorksheetStreamLoader.cs:521`, `docs/excel-interop.md` (+pt-BR) and the `workbook-and-expressions.md` blockquotes between Phase 4 T6a and Phase 6 T4 (finding 4). Decide the constant policy (one vs two). Decide whether item 19's dangling-part test is deleted (recommended) or rescoped.

**T1 — Reader (no Phase 4 dependency).** Files: `Danfma.MySheet.Excel/TableDefinitionReader.cs` (new), `Danfma.MySheet.Excel/ExcelFile.cs` (items 2, 9, 10; item 3's doc text only after T0 rules on wording), new `tests/Danfma.MySheet.Excel.Tests/TableDefinitionReaderTests.cs` (registry-level pins; leaves `TableInteropTests.cs` to T2 except the one-line flip at `:123`). Pins (all `[Certain]` from the SDK view):
| fixture | `Tables["Tabela1"]` | warnings |
| --- | --- | --- |
| plain (`ref="A1:B3"`) | `("Tabela1","Data",1,3,1,true,false,["Item","Valor"])`; `FirstDataRow` 2, `LastDataRow` 3 | 0 |
| totals (`A1:B4 totalsRowCount=1`) | `(…,1,4,1,true,true,…)`; `TotalsRow` 4, `LastDataRow` 3 | 0 (ClosedXML fixture: 1 `UnparsableFormula/B4`) |
| headerless (`A2:B3 headerRowCount=0`) | `(…,2,3,1,false,false,["Item","Valor"])`; `HeaderRow` null | 0 |
| header-only (`A1:B1`) | `(…,1,1,1,true,false,…)`; `DataRowCount` 0 | 0 |
| names (`A1:E3`) | columns `["AMOUNT IN USD For Line 1","(A) NAME OF PFIC","Owner's Share","Line\nBreak"," Padded "]` (4th DECODED from `Line_x000a_Break`, 5th untrimmed) | 0 |
| two tables + second sheet | `Tabela1`, `Tabela2` (`Data`, `F1:G3`), `Tabela3` (`Other`, `A1:A3`); `Tables.Count` 3 | 0 |
| `My\Table` / `TRUE` displayName | not registered | 1 `InvalidTableDefinition`, Subject = displayName |
| `NOT-A-RANGE`, missing `ref`, `headerRowCount="2"`, `count` mismatch, empty/duplicate column name, garbage XML, wrong root | not registered; `B3` = 32 intact | 1 `InvalidTableDefinition` each; Subject = displayName when readable else sheet name |
| `DefinedNames.ContainsKey("Tabela1")` | false on every fixture | |
| `Save` → `Load` | `Tables["Tabela1"]` equal to the loaded one | |

**T2 — Evaluation pins (after Phase 4 T5 and Phase 5 merge).** Files: `tests/Danfma.MySheet.Excel.Tests/TableInteropTests.cs` (items 11-19 minus the dangling test, plus the committed Aspose fixture under `tests/Danfma.MySheet.Excel.Tests/Fixtures/`). Oracle rows:
| cell / formula | plain | totals | headerless | header-only |
| --- | --- | --- | --- | --- |
| `SUM(Tabela1[Valor])` | 42 | 42 | 42 | 0 |
| `ROWS(Tabela1[#All])` | 3 | 4 | 2 | — |
| `ROWS(Tabela1[#Data])` | 2 | 2 | 2 | 0 |
| `COUNTA(Tabela1[#Headers])` | 2 | — | 1 (`#REF!` counted) | — |
| `SUM(Tabela1[[#Data],[Valor]])` | 42 | — | — | — |
| `SUM(INDIRECT("Tabela1["&D10&"]"))`, D10 = `Valor` | 42 | — | — | — |
| `SUM(Tabela1[#Totals])` | `#REF!` | 42 | — | — |
| `SUM(Tabela1[[#Data],[#Totals]])` | 42 | 84 | — | — |
| `SUM(Tabela1)` (stored `Tabela1[]`) | 42 | — | 42 | — |
| `SUM(Tabela1[Item])` | — | — | 0 | — |
| `SUM(Tabela1[@Valor])` in D2 (stored `[[#This Row],[Valor]]`) | 10 → S1: `UnparsableFormula`, cached 10 | | | |
| `Tabela1[@Valor]*2` in I2 (stored `[[#This Row],[Valor]]*2`) | 20 → `UnparsableFormula`, cached 20 | | | |
| names fixture: `SUM([AMOUNT IN USD For Line 1])` / `COUNTA([(A) NAME OF PFIC])` / `SUM([Owner''s Share])` / `SUM([Line\nBreak])` / `SUM([[ Padded ]])` | 42 / 2 / 3 / 300 / 15 (Aspose stores `Owner''s Share` doubled even when typed single) | | | |
| cross-sheet: `Report!D1 =SUM(Tabela1[Valor])`, `D2 =SUM(Data!Tabela1[Valor])` (stored unqualified) | 42 / 42 | | | |
| `My\Table[Valor]`, `TRUE[Valor]` | 42 on the oracle; MySheet: `UnparsableFormula` + cached 42, plus T1's `InvalidTableDefinition` | | | |
Edit-and-recalculate half of item 13 (`B2` = 8 → 40) has no oracle row; it is a MySheet liveness pin. The `SUM(Tabela1[Valor]*2)` / `IF` shapes from Phase 5 are plain-entry `#VALUE!` on the oracle — do not add them here.

**T3 — Export and merge pins (no Phase 4 dependency for the first half).** Files: `tests/Danfma.MySheet.Excel.Tests/ExcelExportTests.cs`, `tests/Danfma.MySheet.Excel.Tests/ExcelMergeTests.cs`. Pins: `SaveAsExcel` in both modes → `Data`'s `TableDefinitionParts.Count() == 0`, no `<tableParts>` (measured); `MergeIntoExcel` into a copy of the table fixture → `TableDefinitionParts.Count() == 1`, `ref == "A1:B3"`, `<tableParts>` present (measured); Formulas-mode `<f>` text `SUM(Tabela1[Valor])` — after Phase 4 only (the writer's canonical form is single-bracket per Ruling 2). Locate the worksheet by sheet name, never `WorksheetParts.First()` (part order is not document order — measured `rId4` before `rId3`).

**T4 — Docs (after T1 and T2, one owner for both twins).** Files: `docs/excel-interop.md`, `docs/pt-BR/excel-interop.md`, `docs/workbook-and-expressions.md` (`:934-939` and `:993-999`), `docs/pt-BR/workbook-and-expressions.md` (`:986-992`, `:1045-1052`), `docs/serialization.md:337`, `docs/pt-BR/serialization.md:367`. Verification: the §5 grep list returns exit 1; `diff` of each twin's edited paragraphs against the English; the out-of-scope list names `[@Col]`, `[[#This Row],[Col]]`, `[[A]:[B]]`, implicit `[Col]`, and the two structurally unlexable table names.

Concurrency: T1 and T3's first half run now, in parallel, with disjoint files; T2 waits for Phases 4 and 5; T4 last. No two tasks touch the same file.
