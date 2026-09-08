# Phase 3: Table model, Workbook.DefineTable/Tables, and the third serialized Workbook member

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `table-model-registry`. Design dependencies: none. Adversarial verifier verdict: **needs-revision** (3 blockers, 2 majors, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

The table registry becomes `Workbook`'s third serialized member — a `[MemoryPackInclude] private
Dictionary<string, Table> _tables` declared after `DefinedNames`, exposed as `public
IReadOnlyDictionary<string, Table> Tables` (the `Sheet._cells`/`Sheet.Cells` pattern, Sheet.cs:25-35), written
only by `DefineTable`. `Table` is a public MemoryPackable positional record whose geometry mirrors the xlsx
`<table>` element exactly — `FirstRow`/`LastRow`/`FirstColumn` plus `HasHeaderRow`/`HasTotalsRow` — with
header/data/totals rows as derived `[MemoryPackIgnore]` properties, which makes contradictory geometry
unrepresentable and makes the zero-data-row table fall out arithmetically instead of needing sentinels; column
lookup is a lazily built OrdinalIgnoreCase `Dictionary` (measured 6.3x faster than a linear scan at 30
columns, 21x at 80, break-even ~12). Two decisions in the assigned scope are wrong and I am overriding them
with measurements: **`NamedReferences.ValidateName` cannot be reused for table names** — it rejects `Tabela1`
and `Table1`, Excel's own default table names, because `Parser.IsCellReference` (Parser.cs:788-812) treats any
letters-then-digits string as a cell reference with no column/row bound — so a dedicated `Table.ValidateName`
using a new Excel-grid-bounded check is required; and the version counter should **fold** into the existing
counter, renamed `Workbook.DefinitionsVersion`, because `RecalculationEngine.EnsureFresh` only ever rebuilds
the whole graph. I measured that a version bump alone leaves stale VALUES behind (`SUM(Rng)` kept returning 3
after `Rng` widened from A1:A2 to A1:A4, even through `engine.Recalculate([])` which reported `rebuilt=True,
dirty=0`) — a pre-existing `DefineName` bug that table redefinition would inherit, so the engine must treat a
definitions change as a full-invalidation event. The wire consequences are measured on the exact `Workbook`
shape: header `0x02` → `0x03`, +4 bytes for the empty map, old readers throw
`MemoryPackSerializationException: ... property count is 2 but binary's header maked as 3`, and old files
(header `0x01` and `0x02`, both verified in the frozen fixtures) still load with the member arriving `null`.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Item 16 treats CellStoreTests.PreChangeCellsWireGolden as the only frozen wire golden, and item 15's rationale states: "Verified there is exactly one hardcoded 0x02 assumption in the whole repo, this comment (grep over Danfma.MySheet and tests found no other)." Verification #5 expects "Passed! failed: 0 for the whole core suite".
      *Measured evidence:* There is a SECOND frozen full-Workbook wire golden.
      tests/Danfma.MySheet.Tests/Parsing/SheetNameInterningTests.cs:18 `private const string
      PreInterningWireGolden = "AgIAAAD7////BAAAAERhdGED+////wQAAABEYXRhAAAAAAQAAAD9////AgAAAEEx..."` (header
      byte `Ag` = 0x02), asserted at :51 `await
      Assert.That(Convert.ToBase64String(bytes)).IsEqualTo(PreInterningWireGolden);` inside
      `Wire_IsByteIdentical_AfterInterning` (:47). MEASURED: I implemented items 1/2/3/4/9/10/12/13/14
      verbatim in /tmp/verify-table-model-registry/engine and ran the copied core suite. Unpatched baseline:
      `total: 1203, failed: 0`. Patched: `total: 1203, failed: 2` —
      CellStoreTests.Wire_IsByteIdentical_AfterNumericKeys AND
      SheetNameInterningTests.Wire_IsByteIdentical_AfterInterning, both "differs at index 1: AwIA... vs
      AgIA...". Their grep missed it because the 0x02 is encoded as base64 `Ag`, not the literal "0x02".
      *Correction:* Add an item mirroring 16/17/18 for
      tests/Danfma.MySheet.Tests/Parsing/SheetNameInterningTests.cs: rename PreInterningWireGolden ->
      PreTablesInterningWireGolden, add the regenerated constant (same mechanical delta: [0x03] + old[1..] +
      four 0x00), repoint Wire_IsByteIdentical_AfterInterning, and add the byte-delta guard test. Delete the
      "exactly one hardcoded 0x02 assumption" claim from item 15's rationale and restate verification #5's
      expectation as 1203 passing (baseline measured today).
- [ ] **B2.** Item 15: splitting IsStale into HasDefinitionChange()/HasSheetStructureChange() and having Recalculate call InvalidateCache when DefinitionsChanged fixes the measured stale-value trap.
      *Measured evidence:* MEASURED with item 15 implemented verbatim (HasDefinitionChange() =>
      _workbook.NamesVersion != _namesSnapshot; EnsureFresh reporting FreshnessResult; Recalculate
      invalidating; EstimateImpact returning ImpactEstimate.Full). Sequence A (Recalculate first) works:
      `mode=FullFallback dirty=-1 rebuilt=True -> B1=15`. Sequence B (the documented plan-then-act pair):
      DefineName("Rng","Data!A1:A4") -> `EstimateImpact([])` = `ImpactEstimate { RecommendFull = True,
      ConeSize = -1, Reason = definição (nome/tabela) alterada }` -> `Recalculate([])` = `mode=Partial dirty=0
      rebuilt=False` and **Main!B1 stays 3** (want 15). Cause: RecalculationEngine.cs:167-178 EnsureFresh
      refreshes the snapshot as a side effect (`_namesSnapshot = _workbook.NamesVersion;` :176), so whichever
      public method runs first consumes the signal — and EstimateImpact is documented at
      RecalculationEngine.cs:93-96 as analysing "WITHOUT recomputing or evicting ... surfaced so the host can
      plan". Item 21's test (Recalculate first) passes and would never catch this.
      *Correction:* Make the invalidation sticky, not derived from the snapshot comparison: set a `private
      bool _pendingDefinitionInvalidation` when EnsureFresh observes DefinitionsVersion != snapshot, have
      EstimateImpact READ it without clearing, and clear it only in Recalculate's InvalidateCache arm. Add the
      EstimateImpact-then-Recalculate ordering to item 21's test (and its DefineName twin).
- [ ] **B3.** Item 3/12: `Table` memoizes column lookup in a `[MemoryPackIgnore] private Dictionary<string,int>? _columnIndex` field on the record, and item 12 stores a defensive snapshot via `var registered = table with { ColumnNames = table.ColumnNames.ToArray() };`.
      *Measured evidence:* C# copies ALL instance fields (including manually declared private ones) in a
      record's generated copy constructor, so `with` carries the stale memo. MEASURED: `var q0 = new
      Table("Tabela","Data",1,3,1,true,false,["a","b"]); q0.TryGetColumnIndex("a", out _); var q1 = q0 with {
      ColumnNames = ["a","b","c"] };` -> `q1.TryGetColumnIndex("c")` = **False** (truth: true/2); after
      `wb.DefineTable(q1)`, `wb.Tables["Tabela"].TryGetColumnIndex("c")` = **False** too. With the names
      reordered instead (`with { ColumnNames = ["c","b","a"] }`) it returns index 2 for "c" (truth 0).
      Resizing a table by appending a column — `wb.Tables["T"] with { ColumnNames = [...] }` after any
      evaluation has warmed the memo — is the natural host flow, and item 4 makes TryGetColumnRange return
      false, so the reference-semantics phase renders `T[NewCol]` as a permanent silent #REF!/#NAME? with a
      wrong sheet column in the reorder case.
      *Correction:* Two options. (a) Move the memo off the record (registry-side index keyed by table name, or
      a ConditionalWeakTable) — this also fixes the equality/hash defect below. (b) Keep it but make
      DefineTable build a fresh `new Table(...)` rather than `with`, AND make TryGetColumnIndex self-
      invalidate (`if (map is null || map.Count != ColumnNames.Count) rebuild`). Option (a) is the only one
      that makes the public `with` safe for hosts.

## Major corrections

- [ ] **M1.** Risk #7 states the only Table-equality hazard is that ColumnNames compares by reference, and prescribes "Tests and hosts must assert on individual members."
      *Evidence:* The private memo field also participates in the synthesized Equals/GetHashCode. MEASURED:
      `var cols = new[]{"a","b"}; var x = new Table("Tabela","Data",1,3,1,true,false,cols); var y = new
      Table(...same, same cols instance...);` -> `x == y` True, hashes equal. Then `x.TryGetColumnIndex("a",
      out _);` -> `x == y` **False**, hashes **differ**. So a public record's GetHashCode mutates on first
      column lookup: a Table used as a Dictionary/HashSet key escapes its bucket after the resolver touches
      it, and two structurally identical tables compare unequal for a reason invisible in the record's printed
      members.
      *Correction:* Move the memo off the record (preferred — fixes the blocker above too), or override
      Equals/GetHashCode on Table over the eight serialized members plus a sequence comparison of ColumnNames.
      Record the real hazard in risks instead of the narrower ColumnNames one.
- [ ] **M2.** Item 8: "TryParseColumn and TryParseRow already strip '$', reject row 0 and guard the row accumulator against overflow (:98-102) — composing them is the small-reusable-component move and adds no new parsing rules." Item 5 explicitly rejects reusing CellAddress.TryGetColumnRow for exactly this reason ("no ... overflow guard").
      *Evidence:* CellAddress.cs:123-146 TryParseColumn has NO overflow guard: `column = column * 26 +
      (char.ToUpperInvariant(raw) - 'A' + 1);` with only a `char.IsLetter` check. MEASURED via the composed
      TryParseA1 as item 8 specifies it: `TryParseA1("AAAAAAAAAAAAAAAAAAAAAAAA1")` -> `ok=True col=-965696553
      row=1`. So DefineTable's A1 overload accepts a wrapped column and then fails with a misleading "spans N
      columns but M names" / ArgumentOutOfRangeException(FirstColumn) message instead of "'...' is not a valid
      A1 range"; a letter run that wraps to a positive value with a matching width passes both checks and
      stores wrong geometry silently.
      *Correction:* Bound the letter run in TryParseA1 (reject > 3 letters, matching item 5's Excel bound) or
      add `if (column > (int.MaxValue - d) / 26) return false;` inside TryParseColumn, mirroring
      TryParseRow:100-102. Correct item 8's rationale — the row accumulator is guarded, the column accumulator
      is not.

## Implementation items

- [ ] **1.** Create Danfma.MySheet/Table.cs with `[MemoryPackable] public sealed partial record Table(string Name, string SheetName, int FirstRow, int LastRow, int FirstColumn, bool HasHeaderRow, bool HasTotalsRow, IReadOnlyList<string> ColumnNames)`. FirstRow/LastRow span the WHOLE table range (header row through totals row, exactly the xlsx `ref`); FirstColumn is the 1-based leftmost sheet column. Do NOT add a LastColumn member — it is derived from ColumnNames.Count.
      *Files:* `Danfma.MySheet/Table.cs`
      *Why:* This is the xlsx model verbatim, which removes an entire class of invalid input. Measured against
      a real ClosedXML-written file: a table with a totals row reports ref="A1:B4" totalsRowCount="1" (header
      row 1, data 2..3, totals 4), so ref INCLUDES the totals row. The requester's suggested shape (headerRow,
      firstDataRow, lastDataRow) (a) omits the column origin entirely, so no column ordinal can be mapped to a
      sheet column, and (b) lets the three row numbers contradict each other (headerRow=5 with
      firstDataRow=2). Positional record + IReadOnlyList<string> measured byte-stable across the array-backed
      (fresh) and List-backed (deserialized) shapes, which CellStoreTests.RoundTrip_NewToNew_IsByteStable:71
      requires. Named `Table` (not TableDefinition) because Workbook.DefinedNames/DefineName's symmetric pair
      is Tables/DefineTable; the collision with DocumentFormat.OpenXml.Spreadsheet.Table inside the Excel
      assembly is resolved by the alias already used there for the identical Sheet collision, ExcelFile.cs:4
      `using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;`. Verified no `Table` type exists in
      Danfma.MySheet today.
- [ ] **2.** In Table.cs add the derived geometry as `[MemoryPackIgnore]` get-only properties: `ColumnCount => ColumnNames.Count`, `LastColumn => FirstColumn + ColumnNames.Count - 1`, `HeaderRow => HasHeaderRow ? FirstRow : (int?)null`, `TotalsRow => HasTotalsRow ? LastRow : (int?)null`, `FirstDataRow => FirstRow + (HasHeaderRow ? 1 : 0)`, `LastDataRow => LastRow - (HasTotalsRow ? 1 : 0)`, `DataRowCount => Math.Max(0, LastDataRow - FirstDataRow + 1)`.
      *Files:* `Danfma.MySheet/Table.cs`
      *Why:* Probed: get-only expression-bodied properties marked [MemoryPackIgnore] compile inside a
      [MemoryPackable] record and do not appear on the wire (an 8-member record with 4 extra derived
      properties serialized as 8 members, 81 bytes). [MemoryPackIgnore] on non-serialized public members is
      the codebase convention — Sheet.Cells:34-35 and Sheet.Count:83-84. DataRowCount clamped at 0 is what
      makes a header-only table (`ref="A1:A1"` with headerRowCount=1) a legal, non-throwing model state rather
      than a negative row count.
- [ ] **3.** In Table.cs add `[MemoryPackIgnore] private Dictionary<string, int>? _columnIndex;` and `public bool TryGetColumnIndex(string columnName, out int index)` that lazily builds the OrdinalIgnoreCase map from ColumnNames via `Interlocked.CompareExchange(ref _columnIndex, created, null) ?? created`, then does one TryGetValue. Never use a field initializer on `_columnIndex`.
      *Files:* `Danfma.MySheet/Table.cs`
      *Why:* Dictionary over linear scan, measured (best-of-3, warmed, 1e6 lookups of the last column with
      different casing): n=5 1.4x, n=12 1.1x, n=30 6.3x, n=80 21.4x, n=256 46.4x; the dictionary is flat at
      ~17ns/lookup while the scan is linear. Excel tables are typically 10-50 columns and may legally reach
      16,384, so the dictionary is at worst a wash and usually a large win, and it is built once per table.
      Lazy + Interlocked is the codebase's exact race-free memo pattern: Workbook.ValueStore:68-80 and
      Sheet.GetStructuralIndex:63-72. The 'never `= new()` on the field' rule is mandatory here for the
      documented reason at Workbook.cs:36-38 and Sheet.cs:39-42 — MemoryPack bypasses field initializers,
      which I confirmed: a `[MemoryPackInclude]` member on a Workbook-shaped type came back null (not the
      initializer's value) when reading an older payload.
- [ ] **4.** In Table.cs add `public int SheetColumnAt(int columnIndex) => FirstColumn + columnIndex;` (1-based sheet column for a 0-based table column ordinal), and `public bool TryGetColumnRange(string columnName, out int sheetColumn, out int firstRow, out int lastRow)` returning the [#Data] rows for that column (false when the name is unknown OR DataRowCount == 0).
      *Files:* `Danfma.MySheet/Table.cs`
      *Why:* These two are the whole surface the reference-semantics phase needs to turn `Tabela1[Valor]` into
      a concrete RangeReference: it must never re-derive geometry itself. Returning false for DataRowCount ==
      0 hands that phase the #REF! decision at the one place that knows the table is empty, mirroring how
      DynamicRange.TryResolveReference:42-47 either produces a concrete RangeReference or nothing
      (DynamicRange.cs:53-56 then yields #REF!).
- [ ] **5.** Add `internal static bool IsExcelGridCellReference(string text)` to Danfma.MySheet/Parsing/Parser.cs immediately after IsCellReference (:788-812), with `private const int ExcelMaxColumn = 16_384; private const int ExcelMaxRow = 1_048_576;`. Implement it standalone and overflow-free: strip '$', require 1-3 letters followed by 1-7 digits, compute column and row, return column <= ExcelMaxColumn && row >= 1 && row <= ExcelMaxRow. Document in the comment WHY it differs from IsCellReference.
      *Files:* `Danfma.MySheet/Parsing/Parser.cs`
      *Why:* MEASURED BLOCKER for the assigned scope: `Parser.IsCellReference` (:788-812) has no column or row
      bound — it returns true for any letters-then-digits string — so `NamedReferences.ValidateName`
      (NamedReferences.cs:155-175, which calls it at :191) REJECTS 'Tabela1' and 'Table1', Excel's own default
      table names. Probed via Workbook.DefineName: 'Tabela1' REJECT, 'Table1' REJECT, 'Sales.Data' OK,
      'Vendas.2024' OK, 'Q1' REJECT, 'A1' REJECT. IsCellReference must NOT be bounded (MySheet's grid is
      deliberately unbounded — no MaxColumn/MaxRow constant exists anywhere in the engine, and
      Parser.ParseIdentifier:326 depends on the current behaviour), so the name rule needs its own bounded
      predicate. Keeping it in Parser.cs preserves the single-source-of-truth arrangement the existing comment
      at NamedReferences.cs:190-192 asserts. Do NOT reuse CellAddress.TryGetColumnRow:37-69 — it neither
      strips '$' nor guards the row accumulator against overflow, so a 255-char digit run could wrap into an
      in-grid row.
- [ ] **6.** Add `internal static void ValidateName(string name)` to Table.cs implementing Excel's documented table-name rule: reject null/empty/whitespace; reject length > 255; require the first char to be a letter or '_'; require every char to be a letter, digit, '.' or '_'; reject the single-character names "C", "c", "R", "r"; reject `Parser.IsExcelGridCellReference(name)`; reject an R1C1 shape (`R` digits `C` digits); reject "TRUE"/"FALSE" (OrdinalIgnoreCase). Throw ArgumentException with the offending value quoted and `nameof(name)` as the paramName, in the message style of NamedReferences.ValidateName:167-173.
      *Files:* `Danfma.MySheet/Table.cs`, `Danfma.MySheet/Parsing/Parser.cs`
      *Why:* Excel's table-name rules per the 'Rename an Excel table' documentation: start with a letter,
      underscore or backslash; letters/numbers/periods/underscores thereafter; cannot be "C"/"c"/"R"/"r";
      cannot be a cell reference (A1 or R1C1 form); no spaces; 255 characters max. TRUE/FALSE is a MySheet-
      specific addition, not an Excel rule: the tokenizer reads those as BooleanValue (Parser.ParseIdentifier
      IsBoolean at :321), so `TRUE[Col]` could never reach a table branch — rejecting it at registration is
      honest fail-fast rather than a silent dead name. Backslash is deliberately rejected even though Excel
      allows it, because Tokenizer.ReadIdentifier:108-121 consumes only letters/digits/_/./$ so a backslash-
      prefixed table could never be written in a formula. The 255-char limit is enforced here even though
      ValidateName does not enforce it for defined names — a pre-existing gap recorded in risks, not a reason
      to ship a second one.
- [ ] **7.** Add `internal void Validate()` to Table.cs (mirroring ValueStoreOptions.Validate, WorkbookOptions.cs:106-130): call Table.ValidateName(Name); `ArgumentNullException.ThrowIfNull(SheetName)` and reject an empty/whitespace SheetName with ArgumentException; `FirstRow < 1` or `FirstColumn < 1` -> `ArgumentOutOfRangeException(nameof(FirstRow), FirstRow, "Rows are 1-based.")` / the column twin; `LastRow < FirstRow` -> ArgumentException; row span < (HasHeaderRow?1:0)+(HasTotalsRow?1:0) -> ArgumentException naming both flags; `ColumnNames.Count == 0` -> ArgumentException; any null/empty/whitespace column name -> ArgumentException naming its index; any duplicate column name under StringComparer.OrdinalIgnoreCase -> ArgumentException naming both indices and the repeated text. Row span == header+totals (zero data rows) is EXPLICITLY LEGAL.
      *Files:* `Danfma.MySheet/Table.cs`
      *Why:* Validation on the value type, invoked by the consumer, is the established pattern:
      Workbook(WorkbookOptions):64 calls `options.ValueStore.Validate()`. Keeping it OFF the constructor is
      deliberate and matches the DefineName/ValidateName split (Workbook.cs:546-548) — MemoryPack materializes
      records through the generated constructor, so a validating ctor would turn a hand-corrupted file into a
      throw during deserialization instead of at the API boundary. ArgumentOutOfRangeException(paramName,
      value, message) for 1-based geometry is the exact CellRef.Format precedent (CellRef.cs:67-76);
      ArgumentException for structural/name violations matches ValueStoreOptions (WorkbookOptions.cs:126-131)
      and NamedReferences.ValidateName. Case-insensitive column uniqueness is corroborated by an independent
      implementation: ClosedXML 0.105.0 refused to save a table whose headers were 'Col' and 'col' with
      `ArgumentException: The header row contains more than one field name 'col'` (measured), and MySheet must
      compare case-insensitively anyway since Excel resolves Table1[col] to Table1[Col]. Column names are
      human-authored ('Valor Total (R$)' in the measured fixture) and must never see
      Table.ValidateName/NamedReferences.ValidateName.
- [ ] **8.** Add `public static bool TryParseA1(string text, out int column, out int row)` to Danfma.MySheet/Expressions/CellAddress.cs: split at the first digit (skipping '$'), delegate the letter run to the existing TryParseColumn (:122-145) and the digit run to TryParseRow (:79-116), return false when either half is empty or fails.
      *Files:* `Danfma.MySheet/Expressions/CellAddress.cs`
      *Why:* The A1-range overload of DefineTable needs to parse a corner like "A1" or "$B$500".
      CellAddress.Parse:11-30 throws FormatException and does not strip '$'; TryGetColumnRow:37-69 does not
      strip '$' and has no row-0 or overflow guard. TryParseColumn and TryParseRow already strip '$', reject
      row 0 and guard the row accumulator against overflow (:98-102) — composing them is the small-reusable-
      component move and adds no new parsing rules. CellAddress is internal and Danfma.MySheet.csproj:28-30
      grants InternalsVisibleTo to Danfma.MySheet.Excel, so the loader could use it directly, but the parse
      belongs in DefineTable so there is exactly one implementation.
- [ ] **9.** In Danfma.MySheet/Workbook.cs, add AFTER the DefinedNames property (:123-124): `[MemoryPackInclude] private Dictionary<string, Table> _tables = new(StringComparer.OrdinalIgnoreCase);` followed by `[MemoryPackIgnore] public IReadOnlyDictionary<string, Table> Tables => _tables;`. MOVE the 'MemoryPack serializes members in declaration order; this MUST stay the LAST serialized member' comment from :119-122 onto `_tables`, and rewrite the DefinedNames comment to say it is now the SECOND of three.
      *Files:* `Danfma.MySheet/Workbook.cs`
      *Why:* Private serialized field + public read-only projection is the in-repo pattern for exactly this
      shape: Sheet.cs:25-27 (`[MemoryPackInclude] [CellStoreFormatter] private CellStore _cells`) with
      Sheet.cs:34-35 (`[MemoryPackIgnore] public IReadOnlyDictionary<string, Expression> Cells => _cells`),
      and Sheet.cs:16-24 documents the same 'must keep member position' constraint. It delivers the read-only
      view S2 asks for and is strictly better than DefinedNames, whose doc comment at :115-118 has to warn
      that direct mutation escapes NamesVersion — with no public setter that hole cannot exist for tables.
      Dictionary<string, Table> (key duplicating Table.Name on the wire) rather than Table[] because
      Sheets:106-107 and DefinedNames:123-124 already duplicate their keys; ~10 bytes per table is not worth
      diverging.
- [ ] **10.** In Workbook.RestoreComparers (:143-152) add the third branch verbatim in the DefinedNames style: `_tables = _tables is null ? new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase) : new Dictionary<string, Table>(_tables, StringComparer.OrdinalIgnoreCase);`
      *Files:* `Danfma.MySheet/Workbook.cs`
      *Why:* Both halves are measured on the exact Workbook shape (parameterless [MemoryPackConstructor],
      property initializers, a [MemoryPackInclude] private Dictionary field). (1) Reading an older 2-member
      payload into the 3-member type leaves the private field NULL — the field initializer does NOT survive,
      so the null branch is mandatory, not defensive. (2) A round-tripped NEW payload came back with the
      comparer LOST: `ContainsKey("tabela1")` was false for a key stored as "Tabela1", so the rebuild is
      mandatory too. This is precisely why the DefinedNames branch at :148-151 exists.
- [ ] **11.** Rename `Workbook._namesVersion` (:133-134) to `_definitionsVersion` and `internal long NamesVersion` (:139) to `internal long DefinitionsVersion`, updating its XML doc to 'the count of DefineName + DefineTable calls'. Update the 6 code/comment sites: RecalculationEngine.cs:71, :90, :176, :184, :187 (rename `_namesSnapshot` to `_definitionsSnapshot`), Workbook.cs:117, SheetStructuralIndex.cs:14, DirtyGraph/ReverseDependencyGraph.cs:6, and the Portuguese comment at tests/Danfma.MySheet.Tests/DirtyGraph/RecalculationEngineTests.cs:243.
      *Files:* `Danfma.MySheet/Workbook.cs`, `Danfma.MySheet/RecalculationEngine.cs`, `Danfma.MySheet/SheetStructuralIndex.cs`, `Danfma.MySheet/DirtyGraph/ReverseDependencyGraph.cs`, `tests/Danfma.MySheet.Tests/DirtyGraph/RecalculationEngineTests.cs`
      *Why:* FOLD, do not add a second counter — I agree with the framing that folding is strictly simpler and
      strictly equivalent, with one caveat worth recording: it is equivalent only because IsStale's sole
      consequence is a WHOLESALE rebuild (EnsureFresh:167-178 calls `DirtyEngine.Build(_workbook)`
      unconditionally), so no consumer can ever benefit from distinguishing a name change from a table change;
      a future finer-grained rebuild would have to split them again. The rename is what makes the fold honest
      rather than a lie in IsStale:187, and it is free: the member is `internal`, there is no PublicAPI.*.txt
      / approval test / TreatWarningsAsErrors in this repo, and the compiler finds every site. All 9
      references are internal code or comments — verified by grep.
- [ ] **12.** Add `public void DefineTable(Table table)` to Workbook.cs directly after DefineName(string, string) (:564-588): `ArgumentNullException.ThrowIfNull(table)`; take a defensive snapshot `var registered = table with { ColumnNames = table.ColumnNames.ToArray() };`; call `registered.Validate()`; call the shared `ThrowIfNameTaken` guard; `_tables[registered.Name] = registered;` then bump `_definitionsVersion` inside `unchecked { }`. XML doc must state that redefining an existing table REPLACES it, that the sheet need not exist, and that already-memoized values are not evicted (pointing at InvalidateCache).
      *Files:* `Danfma.MySheet/Workbook.cs`
      *Why:* Mirror of DefineName(string, Expression):546-556 — validate, assign, bump — including its
      replace-on-redefine semantics (`DefinedNames[name] = reference` at :551), which is also the right table
      behaviour: Excel forbids two tables sharing a name, so the only meaning of a second DefineTable with the
      same name is 'resize/redefine this table'. The defensive ColumnNames copy exists because Table memoizes
      a derived lookup dictionary keyed on those names (item 3); a caller mutating their List<string>
      afterwards would leave the memo silently lying, which is strictly worse than a stale reference. The
      sheet is deliberately NOT required to exist: DefineName does not check either, and the missing-sheet ->
      #REF! rule already lives at exactly one place, Workbook.EvaluateCell:339-342, which is also Excel's
      behaviour for a table on a deleted sheet. That also frees the loader from any ordering constraint
      relative to Sheets.Add (ExcelFile.cs:126).
- [ ] **13.** Add the A1-range convenience overload `public void DefineTable(string name, string sheetName, string reference, IReadOnlyList<string> columnNames, bool hasHeaderRow = true, bool hasTotalsRow = false)` to Workbook.cs, plus `private static bool TryParseTableReference(string reference, out int firstColumn, out int lastColumn, out int firstRow, out int lastRow)` which splits on a single ':' (accepting a colon-free single-cell ref), parses each corner with CellAddress.TryParseA1, and normalizes with Math.Min/Math.Max. Throw ArgumentException (paramName `nameof(reference)`) for an unparsable ref, and a second ArgumentException when `lastColumn - firstColumn + 1 != columnNames.Count`. Then delegate to DefineTable(new Table(...)).
      *Files:* `Danfma.MySheet/Workbook.cs`
      *Why:* Yes, the A1-range overload is needed, and it is the primary form: the xlsx loader has
      `Table.Reference` as the A1 string "A1:B4" (measured on a real ClosedXML file) and nothing else, so
      without this overload the range parse gets duplicated inside Danfma.MySheet.Excel. It is also the direct
      analogue of DefineName's string overload (:564-588) — the Expression/geometry form plus a text-parsing
      convenience form. A separate geometry-explicit overload is NOT needed: `DefineTable(new Table("T",
      "Data", firstRow: 1, lastRow: 500, firstColumn: 1, hasHeaderRow: true, hasTotalsRow: false, columnNames:
      [...]))` already IS it, with named arguments doing better than six positional ints. sheetName stays a
      SEPARATE parameter rather than a qualified "Data!A1:B500": sheet names are plain unescaped strings
      everywhere in the model (CellReference/RangeReference SheetName), and only
      FormulaWriter.WriteSheetQualifier:344-361 quotes them, so requiring a caller to pre-quote a sheet name
      would be the one place in the public API demanding formula syntax. The width cross-check is the one
      genuine redundancy in the inputs (Excel guarantees ref width == tableColumns count), so it is checked
      exactly here and nowhere else.
- [ ] **14.** Add `private void ThrowIfNameTaken(string name, bool definingTable)` to Workbook.cs and call it from DefineTable(Table) and from both DefineName overloads (:546, :564) after their existing ValidateName call: a table name that collides with an existing DefinedNames key, or a defined name that collides with an existing Tables key, throws ArgumentException naming the conflict and the other namespace.
      *Files:* `Danfma.MySheet/Workbook.cs`
      *Why:* Excel's rule, and S2's 'table names are workbook-unique': Excel's Name Manager shares one
      namespace for tables and defined names and refuses a duplicate. Enforcing it symmetrically in one helper
      keeps a single invariant instead of an order-dependent one, and it is safe for the loader:
      ExcelFile.LoadDefinedNames already catches ArgumentException and degrades to an InvalidDefinedName
      warning (ExcelFile.cs:194-206), so a pathological file that violates Excel's own rule produces a
      warning, never a failed load. No existing host can break, since no workbook can contain a table before
      this feature exists.
- [ ] **15.** In RecalculationEngine.cs, split IsStale (:185-210) into `private bool HasDefinitionChange() => _workbook.DefinitionsVersion != _definitionsSnapshot;` and `private bool HasSheetStructureChange()` (the sheet count/identity/version loop at :192-208), and change EnsureFresh (:167-178) to report both, e.g. returning a `private readonly record struct FreshnessResult(bool Rebuilt, bool DefinitionsChanged)`. In Recalculate (:118-160), when DefinitionsChanged is true, call `_workbook.InvalidateCache()` and return `new RecalculationResult(RecalculationMode.FullFallback, -1, [], rebuilt, "definição (nome/tabela) alterada — recompute completo")`. In EstimateImpact (:97-110) return `ImpactEstimate.Full("definição (nome/tabela) alterada")` in the same case.
      *Files:* `Danfma.MySheet/RecalculationEngine.cs`
      *Why:* MEASURED, and this is the concrete staleness trace the scope asks for. Fixture Data!A1..A4 =
      1,2,4,8; Main!B1 = =SUM(Rng) with Rng = Data!A1:A2. (1) Plain workbook: first read 3; redefine Rng to
      Data!A1:A4; read again -> STILL 3; after InvalidateCache -> 15. (2) With an engine and ComputeAll:
      redefine, then `engine.Recalculate([])` reported `mode=Partial rebuilt=True dirty=0` and B1 STAYED 3.
      (3) `EstimateImpact` after a redefinition also served a cone of 1 without noticing. So a version bump
      rebuilds the GRAPH (layer 2) but never evicts layer-3 values, which Workbook.InvalidateCache:372-388
      alone clears (_valueStore.Clear + _rangeCache.Clear); the structural index is correctly untouched
      (:376-378). DefineTable must NOT call InvalidateCache itself — that would nuke the whole cache on every
      one of N registrations during a load and defeat the engine's incremental contract for a host that
      redefines a table and then reports its edits. The right owner is the engine, the component whose job is
      keeping values consistent with structure, and FullFallback with DirtyCellCount -1 is the existing
      vocabulary for 'the whole workbook is stale' (:128-135, :141-147). This also fixes the identical pre-
      existing DefineName bug; land it as its own `fix(recalc):` commit so versionize attributes it correctly.
- [ ] **16.** Update the format comment in Workbook.Serialization.cs:41-43: '(Workbook = 0x02)' becomes '(Workbook = 0x03 since the table registry; 0x01/0x02 for older files)'. Leave SniffContainerHeader (:625-627) and Deserialize (:630-636) unchanged.
      *Files:* `Danfma.MySheet/Workbook.Serialization.cs`
      *Why:* The sniff compares four bytes against the "MSWM" magic; only the claim 'never M (0x4D)' matters
      and 0x03 still satisfies it. No code change is needed and none should be made: the length guard
      `bytes.Length >= ContainerHeaderLength` (9) still holds — an empty 3-member workbook is 13 bytes
      (measured), up from 9. Verified there is exactly one hardcoded 0x02 assumption in the whole repo, this
      comment (grep over Danfma.MySheet and tests found no other).
- [ ] **17.** In tests/Danfma.MySheet.Tests/CellStoreTests.cs: rename the frozen constant PreChangeCellsWireGolden (:20-33) to PreTablesWireGolden and KEEP its bytes; add a new `CellsWireGolden` constant holding the regenerated base64; repoint Wire_IsByteIdentical_AfterNumericKeys (:60-66) at CellsWireGolden. Regenerate mechanically, not by capture: newBytes = [0x03] + oldBytes[1..] + [0x00,0x00,0x00,0x00]. Verified values — old golden 726 bytes, first byte 0x02, last four bytes 00 00 00 00; new golden 730 bytes, first byte 0x03, base64 begins `AwIAAAD7////BAAAAERhdGE` and ends `...AQEAAAAAAAAcQAAAAAAAAAAA==`.
      *Files:* `tests/Danfma.MySheet.Tests/CellStoreTests.cs`
      *Why:* I decoded the constant and computed the transformation: base64 chunks 11, decoded 726 bytes, head
      `02 02 00 00 00 fb ff ff`, tail `... 1c 40 00 00 00 00` where the trailing four zeros are the empty
      DefinedNames map. Because the new member is an empty Dictionary serialized as a zero length (measured:
      `03 ...` + `00 00 00 00`), the delta is provably confined to byte 0 and four appended zeros. Keeping the
      old constant costs nothing and turns the next two items into permanent proofs; capturing a fresh base64
      from a debug print would discard the only evidence that nothing else moved.
- [ ] **18.** Add `[Test] public async Task Wire_NewGolden_IsPreTablesGoldenPlusEmptyTablesMember()` to CellStoreTests.cs asserting, over `old = Convert.FromBase64String(PreTablesWireGolden)` and `actual = MemoryPackSerializer.Serialize(BuildWireFixture())`: old[0] == 0x02; actual[0] == 0x03; actual.Length == old.Length + 4; `actual.AsSpan(1, old.Length - 1).SequenceEqual(old.AsSpan(1))`; and the last four bytes of actual are all zero.
      *Files:* `tests/Danfma.MySheet.Tests/CellStoreTests.cs`
      *Why:* This is the answer to 'prove the regeneration is not masking a real break'. Every byte of the
      historical golden except the object-header count must reappear at the same offset, so a formatter
      change, a member reorder, a CellStoreFormatter regression or an accidental extra member all still fail —
      only the one intended schema addition passes. It converts a one-time trust exercise into a permanent
      guard, and it costs no new fixture file.
- [ ] **19.** Add `[Test] public async Task PreTablesGolden_StillLoads_WithEmptyTables()` to CellStoreTests.cs: deserialize `Convert.FromBase64String(PreTablesWireGolden)` into a Workbook, assert `Tables.Count == 0`, `DefinedNames.Count == 0`, and repeat the cell/id assertions of RoundTrip_PreservesCellsAndEvaluatesIdentically (:82+).
      *Files:* `tests/Danfma.MySheet.Tests/CellStoreTests.cs`
      *Why:* The backward-compat leg of the one-way boundary, derived from bytes the repo already froze — no
      new binary fixture. Measured on the exact Workbook shape: a 2-member payload read by a 3-member type
      leaves the private field null, so this test is what pins RestoreComparers' null branch (item 10) rather
      than leaving it as untested defensiveness.
- [ ] **20.** Create tests/Danfma.MySheet.Tests/TableRegistryTests.cs covering: DefineTable(Table) then Tables lookup by a differently-cased name; redefinition replaces and Tables.Count stays 1; the A1 overload parsing "A1:B4" with hasTotalsRow:true yields FirstDataRow 2 / LastDataRow 3 / DataRowCount 2 / LastColumn 2; a header-only "A1:A1" table gives DataRowCount 0 and TryGetColumnRange false; TryGetColumnIndex is case-insensitive and false for an unknown name; ArgumentException for a duplicate-by-case column name, an empty column name, zero columns, an empty sheet name, LastRow < FirstRow, ref width != columnNames.Count, and a name colliding with a DefinedNames key; ArgumentOutOfRangeException for FirstRow 0 and FirstColumn 0; ACCEPTS "Tabela1", "Table1", "Vendas.2024", "_x" and REJECTS "A1", "XFD1048576", "R1C1", "C", "My Table", "Table-1", a 256-char name; a Save/Load round-trip preserving every member; and DefinitionsVersion advancing on each DefineTable.
      *Files:* `tests/Danfma.MySheet.Tests/TableRegistryTests.cs`
      *Why:* TUnit convention `[Test] public async Task X()` + `await
      Assert.That(actual).IsEqualTo(expected)`, function-behaviour style per tests/Danfma.MySheet.Tests.
      'Tabela1'/'Table1' MUST be asserted accepted — they are Excel's default table names (measured in a real
      ClosedXML-authored file: name="Tabela1" displayName="Tabela1") and they are exactly what
      NamedReferences.ValidateName rejects today, so this assertion is the regression guard for item 5. Assert
      on individual MEMBERS, never `Assert.That(wb.Tables["T"]).IsEqualTo(theTable)`: measured, record
      equality on the IReadOnlyList<string> member is reference equality, so the defensive copy in DefineTable
      makes the stored instance unequal to the passed one.
- [ ] **21.** Add to tests/Danfma.MySheet.Tests/MemoryPackCompatibilityTests.PreNamespaceFixture_LoadsAndReevaluates (:23-79) one assertion `await Assert.That(workbook.Tables.Count).IsEqualTo(0);`, and the same assertion to the fixture-loading test in tests/Danfma.MySheet.Tests/ContainerVersionCompatibilityTests.cs (:13+).
      *Files:* `tests/Danfma.MySheet.Tests/MemoryPackCompatibilityTests.cs`, `tests/Danfma.MySheet.Tests/ContainerVersionCompatibilityTests.cs`
      *Why:* Both fixtures keep passing, and I verified WHY rather than assuming. Fixtures/workbook-pre-
      namespaces.msgpack.bin begins `01 02 00 00 00 fb ff ff ff ...` — a ONE-member Workbook from before
      DefinedNames existed, which is the proof that this exact append-only move has already been made once and
      survived. Fixtures/container-v2-brotli-warm.bin: I decompressed its body with BrotliStream (magic MSWM,
      version 2, modelLength 154) and the MODEL's first byte is 0x02, with its last four bytes `00 00 00 00`
      (empty DefinedNames), and the container header layout is entirely independent of the model's member
      count. Both counts are below the new declared 3, and MemoryPack's count-below-declared tolerance is
      measured on the Workbook shape. The assertions convert 'still passes' into 'pinned'.
- [ ] **22.** Add `[Test] public async Task RedefiningATable_ForcesAFullRecompute()` to tests/Danfma.MySheet.Tests/DirtyGraph/RecalculationEngineTests.cs, plus a DefineName twin: build Data!A1..A4 = 1,2,4,8 with Main!B1 = =SUM(Rng) over Data!A1:A2, ComputeAll, create the engine, read 3, redefine to Data!A1:A4, call Recalculate with an EMPTY edited set, and assert Mode == RecalculationMode.FullFallback, DirtyCellCount == -1, StructureRebuilt == true, and that B1 now reads 15.
      *Files:* `tests/Danfma.MySheet.Tests/DirtyGraph/RecalculationEngineTests.cs`
      *Why:* This is the measured failure written down: today that sequence returns `mode=Partial rebuilt=True
      dirty=0` and B1 stays 3. The DefineName twin documents that item 15 fixes a pre-existing bug and not
      just a new-feature gap, and it guards both namespaces against a future regression once the counters are
      folded.
- [ ] **23.** In docs/serialization.md, rewrite the byte-identity bullet at :85-86 so the permanent contract is scoped to the WRITE SHAPE (a cold save is the raw model with no container header, deterministic for a given model) rather than to the byte length across versions, and note that a schema ADDITION shifts the bytes and has already done so once (header 0x01 -> 0x02 when DefinedNames was added). Then add a new subsection after the container-v3 subsection (:209-221) titled 'Forward-compatibility: the table registry (a third `Workbook` member)', following the :199-207 template structure exactly: one paragraph naming Workbook.Tables/DefineTable/ExcelFile.Load, then a 'This is a one-way compatibility boundary' bullet list covering (a) the structured-reference union tag(s), (b) the object-header change 0x02 -> 0x03 plus the four bytes an empty registry adds, (c) files written by this or a later version cannot be opened by older ones, quoting `MemoryPackSerializationException: Workbook property count is 2 but binary's header maked as 3, can't deserialize about versioning.`, and (d) older files (0x01 or 0x02) keep loading because MemoryPack tolerates a header count below the declared member count and leaves absent members null, which RestoreComparers turns into an empty registry — citing the frozen 0x01 pre-namespaces fixture as the precedent.
      *Files:* `docs/serialization.md`
      *Why:* The exception text is verbatim from a probe on the Workbook shape, including MemoryPack's own
      typo ('maked'), so a user hitting it can search for it. The :199-207 shared-formula subsection is the
      template the scope requires and plans/shared-formula-delta-production.md:223,:262 is the write-up
      precedent. The correction at :85-86 is necessary and honest: 'byte-for-byte identical to every prior
      version' was already imprecise before this change, since the pre-namespaces fixture's 0x01 header proves
      the member count grew once. If the structured-reference node's union tag is not yet assigned when this
      item is executed, write bullets (b)-(d) now and leave (a) to be filled in by the reference-semantics
      phase; the two halves are independent.
- [ ] **24.** Mirror the previous item into docs/pt-BR/serialization.md: correct the byte-identity bullet at :90-92 and add the translated subsection after the pt-BR container-v3 subsection, keeping the section order and heading levels identical to the English file.
      *Files:* `docs/pt-BR/serialization.md`
      *Why:* docs/pt-BR is a confirmed full mirror (11 files each, identical names) and docs/pt-BR/README.md:3
      states 'Em caso de divergência, o inglês prevalece.' — the mirror must not drift on a compatibility
      boundary, which is the one kind of documentation a user consults after a failed Load.
- [ ] **25.** In docs/workbook-and-expressions.md, add a 'Tables' subsection next to §Named ranges (:373) documenting Workbook.Tables and both DefineTable overloads: the geometry convention (FirstRow/LastRow span the whole range including header and totals rows, as the xlsx ref does), that table names follow Excel's table-name rule and share the namespace with defined names, that column names are compared case-insensitively and must be unique, that redefining replaces, that the sheet need not exist (a table on a missing sheet is #REF! at evaluation), that a table with zero data rows is legal, and that a runtime redefinition does not evict memoized values — call InvalidateCache, or let RecalculationEngine.Recalculate detect it and fall back to a full recompute. Mirror into docs/pt-BR/workbook-and-expressions.md.
      *Files:* `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`
      *Why:* §Named ranges is where a host looks for the DefinedNames/DefineName pair, so the symmetric
      Tables/DefineTable pair belongs beside it. The staleness paragraph is the documented form of the
      measured behaviour in item 15 and prevents the exact trap I reproduced (SUM(Rng) returning 3 after the
      range widened).
- [ ] **26.** Do NOT hand-edit CHANGELOG.md or the <Version> element in Danfma.MySheet/Danfma.MySheet.csproj:10 (3.16.0). Land the registry as a `feat(serialization):`- or `feat(eval):`-scoped Conventional Commit (with the RecalculationEngine change as a separate `fix(recalc):`) so versionize derives 3.17.0 during .github/workflows/release.yml.
      *Files:* `CHANGELOG.md`, `Danfma.MySheet/Danfma.MySheet.csproj`
      *Why:* CHANGELOG.md is versionize-generated and the release workflow runs a bare `versionize`
      (release.yml:39-42), which bumps the csproj Version itself; there is no .versionize config file
      overriding that. Scopes already in use are parser, eval, serialization, recalc and dirty-graph. Hand-
      editing either file would be overwritten or would double-bump.

## Verification Plan

- [ ] `dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded. The MemoryPack source generator must emit no MEMPACK diagnostics for
      Table.cs (the [MemoryPackIgnore] derived properties and the ignored private _columnIndex field are both
      known-good shapes — probed) and none for the new [MemoryPackInclude] private _tables field on Workbook.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build --treenode-filter "/*/*/CellStoreTests/*"`
      → expected: total: 12, failed: 0 (baseline is 9 passing — measured today; the three added tests are
      Wire_NewGolden_IsPreTablesGoldenPlusEmptyTablesMember, PreTablesGolden_StillLoads_WithEmptyTables and
      the repointed Wire_IsByteIdentical). A failure in Wire_NewGolden_... means the wire delta is NOT
      confined to byte 0 plus four appended zeros and something other than the intended member changed — do
      not regenerate the golden to make it pass.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build --treenode-filter "/*/*/*CompatibilityTests/*"`
      → expected: total: 2, failed: 0 (same count as the measured baseline). These load Fixtures/workbook-pre-
      namespaces.msgpack.bin (verified model header 0x01) and Fixtures/container-v2-brotli-warm.bin (verified
      inner model header 0x02) and now also assert Tables.Count == 0.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build --treenode-filter "/*/*/TableRegistryTests/*"`
      → expected: failed: 0, and the name-validation cases in particular must show 'Tabela1', 'Table1' and
      'Vendas.2024' ACCEPTED while 'A1', 'XFD1048576', 'R1C1', 'C', 'My Table' and 'Table-1' are rejected with
      ArgumentException. If Tabela1/Table1 are rejected, Table.ValidateName is still routing through
      Parser.IsCellReference instead of Parser.IsExcelGridCellReference.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build`
      → expected: Passed! failed: 0, skipped: 0 for the whole core suite. RoundTrip_NewToNew_IsByteStable and
      WarmStartSaveLoadTests.Save_Default_IsByteIdenticalToRawMemoryPack must both still pass — they compute
      both sides at runtime, and an IReadOnlyList<string> member was measured byte-stable across its array-
      backed and List-backed shapes.
- [ ] `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release --no-build`
      → expected: Passed! failed: 0. This phase adds no Excel-side code, so the six existing TableInteropTests
      must be untouched and still green; NamedRangeInteropTests must stay green, proving the DefineName cross-
      namespace guard did not start rejecting real defined names.
- [ ] `dotnet csharpier check .`
      → expected: No files reported as needing formatting (exit 0). Run `dotnet csharpier format .` first if
      it reports any; the pre-commit and pre-push hooks both gate on this.
- [ ] `grep -rn "NamesVersion" Danfma.MySheet Danfma.MySheet.Excel tests | grep -v obj/`
      → expected: No matches. Baseline is 9 matches (Workbook.cs:117/:139,
      RecalculationEngine.cs:71/:90/:176/:184/:187, SheetStructuralIndex.cs:14, ReverseDependencyGraph.cs:6,
      RecalculationEngineTests.cs:243); every one must have become DefinitionsVersion.
- [ ] `python3 -c "import base64,re;s=open('tests/Danfma.MySheet.Tests/CellStoreTests.cs').read();g=lambda n:base64.b64decode(''.join(re.findall(r'\"([^\"]*)\"',re.search(n+r'\s*=\s*(.*?);',s,re.S).group(1))));o,nw=g('PreTablesWireGolden'),g('CellsWireGolden');print(len(o),len(nw),o[0],nw[0],nw[1:len(o)]==o[1:],nw[len(o):]==b'\x00'*4)"`
      → expected: 726 730 2 3 True True — the two frozen constants in the test file satisfy the mechanical
      derivation independently of the test runner.

## Risks carried by this phase

- MEASURED, and it invalidates part of the assigned scope: `NamedReferences.ValidateName` (NamedReferences.cs:155-175) CANNOT be reused for table names. It rejects 'Tabela1' and 'Table1' — Excel's own default table names, and the exact name the existing tests/Danfma.MySheet.Excel.Tests/TableInteropTests fixture uses — because Parser.IsCellReference (Parser.cs:788-812) accepts any letters-then-digits string as a cell reference with no column/row bound. Reusing it would make the xlsx loader unable to register almost every real-world table. Items 5 and 6 exist solely because of this; if they are skipped the whole feature silently fails on real files.
- MEASURED pre-existing bug that table redefinition inherits: bumping the version counter rebuilds the dependency graph but never evicts layer-3 values. With Rng = Data!A1:A2, Main!B1 = =SUM(Rng) reading 3, redefining Rng to Data!A1:A4 and calling engine.Recalculate([]) reported mode=Partial rebuilt=True dirty=0 and B1 STAYED 3; EstimateImpact likewise served a stale cone. Only Workbook.InvalidateCache:372-388 clears it. If item 15 is dropped, `DefineTable` ships a silent stale-value trap and the existing `DefineName` one stays unfixed.
- docs/serialization.md:85-86 claims the cold format is 'byte-for-byte identical to every prior version. This is a permanent contract, guarded by a regression test.' That claim was ALREADY imprecise: Fixtures/workbook-pre-namespaces.msgpack.bin starts with header 0x01 (one member), so the member count already grew once when DefinedNames was added. This phase makes it 0x03 and adds 4 bytes to EVERY cold file, tables or not. Item 24 must correct the claim rather than quietly append an exception to it, and the pt-BR mirror (:90-92) must move with it.
- Naming the public type `Table` collides with DocumentFormat.OpenXml.Spreadsheet.Table inside Danfma.MySheet.Excel, which does `using DocumentFormat.OpenXml.Spreadsheet;` (ExcelFile.cs:3). The excel-loader phase must add `using XlsxTable = DocumentFormat.OpenXml.Spreadsheet.Table;` — a verbatim mirror of the alias already in that file at ExcelFile.cs:4 for the identical `Sheet` collision. If that phase instead renames MySheet's type, the public API loses the DefinedNames/Tables symmetry S2 asks for.
- The symmetric name guard (item 14) adds a NEW failure mode to the shipped `DefineName` API: a defined name colliding with a registered table now throws. Excel forbids the collision so no real file should hit it, and ExcelFile.LoadDefinedNames:194-206 already catches ArgumentException and degrades to an InvalidDefinedName warning, but a third-party writer that violates Excel's rule will now lose that defined name with a warning where it previously loaded. Confirm by running the Excel suite (NamedRangeInteropTests in particular).
- Table.Validate deliberately does NOT reject two tables OVERLAPPING on the same sheet, which Excel does forbid. Nothing in the approved scope ever asks 'which table contains this cell' (`[@Column]` and implicit-table `[Column]` are explicitly out per S1), so overlap is unobservable — but if a later phase adds `[@Column]`, overlap detection becomes mandatory and the check will have to be retrofitted into DefineTable.
- Record equality on `Table` is reference equality for the IReadOnlyList<string> ColumnNames member (measured: two Tables with identical contents but different list instances compare unequal), and DefineTable stores a defensive copy, so `workbook.Tables["T"]` never equals the instance the caller passed. Tests and hosts must assert on individual members. If a future consumer needs value equality, ColumnNames has to move behind a custom Equals — do not discover this via a confusing test failure.
- Table.ValidateName enforces Excel's 255-character limit while NamedReferences.ValidateName still enforces no length limit for defined names. That asymmetry is deliberate (the new validator is written to Excel's documented rule) but a reviewer will read it as an inconsistency; the defined-name gap is pre-existing and out of scope for this phase.
- The `Tables` member duplicates each table's name on the wire (dictionary key plus Table.Name), and nothing validates that they agree after deserializing a hand-crafted file. This exactly matches Sheets (Workbook.cs:106-107, where Sheet.Name duplicates the key and is likewise unvalidated), so it is consistency rather than a new hazard — but it is a real, accepted redundancy.
- Item 24's bullet (a) needs the union tag number of the structured-reference node, which the reference-semantics phase assigns (next free tag is 322). If that number is not yet fixed, write the Workbook-member half of the subsection and leave the tag bullet for that phase; do NOT block items 1-23 on it, and do NOT guess a tag.

## Open questions owned by this phase

- Does Excel return #REF! for `Tabela1[Coluna]` when the table's data body is empty (ref covers only the header row), and for `Tabela1[#Totals]` when the table has no totals row? [Likely] yes — the structured-reference documentation states a specifier for an area the table does not have resolves to #REF! — but I could not confirm it from a primary source in this environment. The MODEL decision does not depend on it (item 4 returns false and hands the choice to the reference-semantics phase); the RESOLUTION decision does. Settle it against the Microsoft 'Using structured references with Excel tables' page or a real Excel run before that phase codes the empty-area arm.
- Is `TRUE`/`FALSE` a legal Excel TABLE name? MySheet must reject it regardless (Tokenizer/Parser read it as a boolean literal at Parser.ParseIdentifier:321, so `TRUE[Col]` can never reach a table branch), but the user-facing doc should say whether that is Excel's rule or a MySheet limitation. Same question for a leading backslash, which Excel's defined-name rule allows and which item 6 rejects because Tokenizer.ReadIdentifier:108-121 has no '\\' in its character set — confirm with the lexer-parser phase that it is not adding one.
- Is `<table>/@name` or `@displayName` authoritative for the name a structured reference uses? Measured on a ClosedXML-authored file the two are always identical ('Tabela1'/'Tabela1'), and Excel's formula engine uses displayName. Recommend the loader read DisplayName with a fallback to Name; confirm against a genuinely Excel-authored .xlsx (this repo has no such fixture — all Excel fixtures are generated by ClosedXML at test time), since a file where they differ would silently register the wrong key.
- Should Workbook.DefineName really cross-reject names that a table already uses (item 14)? It is Excel's rule and it keeps one invariant, but it changes the behaviour of a shipped API and the two syntaxes are actually disjoint in MySheet, so the collision is harmless to resolution. The alternative — enforce uniqueness only in the new DefineTable and document the asymmetry — is lower risk. This needs a product decision, not more analysis.
- How should Table.Validate treat a table whose declared column count exceeds Excel's 16,384-column grid, or whose LastRow exceeds 1,048,576? MySheet's grid is deliberately unbounded (no MaxColumn/MaxRow constant exists in the engine), so rejecting them would be the first place the engine imposes Excel's limits. Item 7 currently does not bound them; confirm that is the intent.

## Phase Summary

_(write when phase completes)_
