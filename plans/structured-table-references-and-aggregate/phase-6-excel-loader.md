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
