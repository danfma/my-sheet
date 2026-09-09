# Structured table references, AGGREGATE, and the blocking reference-semantics gaps

Make `Table[Column]` a first-class reference in the MySheet formula engine, implement `AGGREGATE`, and close
the reference-semantics gaps that stop the affected real-world workbooks from evaluating even once the first
two exist. Target version 3.17.0 (minor).

## Provenance and status of this plan

This plan is the output of seven parallel design agents, each followed by an adversarial verifier that re-read
the cited source and PROVED claims by running probes. Every phase came back `needs-revision`: 12 blockers and
19 majors were found and are folded in below as explicit correction items. Nothing here is unreviewed design.

**Do not trust a line number in this document.** They were accurate when written; several of the cited files changed on disk during the analysis itself. Anchor every edit on member names, constant names and heading text, and re-read the file before editing it.

### The governing principle (P0)

Match Excel's (or Aspose.Cells') observable behaviour whenever possible. Never adopt a simpler-but-different
semantics because it is easier to implement. Where MySheet genuinely cannot match Excel — it has no spill
model and no hidden-row model — the deviation must be explicit in the code, justified by the specific missing
capability, and documented for users. For every semantic decision, state which Excel behaviour is being
reproduced. Where Excel's behaviour is unknown, it goes in this plan's open decisions, never into a silent
guess.

### Settled scope

- **S1 — syntax.** `Table[Column]` (spaces, punctuation and Excel's `'`-escaping inside the brackets), plus `Table[#All]`, `Table[#Data]`, `Table[#Headers]`, `Table[#Totals]` and the composite `Table[[#Data],[Column]]`. Explicitly out: `[@Column]`, column spans `[[Col1]:[Col3]]`, implicit-table `[Column]`. Out-of-scope shapes fail with a clear `ParseException` and therefore keep degrading to the cached value plus an `UnparsableFormula` warning on load.
- **S2 — public API.** `Workbook.DefineTable(...)` plus a read-only `Workbook.Tables`, mirroring the `DefinedNames`/`DefineName` pair. Table names are workbook-unique, as in Excel.
- **S3 — interop.** The loader reads `<table>` parts and auto-registers them. `SaveAsExcel` does not write them. `MergeIntoExcel` is unchanged.
- **S4 — cell boundary.** Excel's implicit intersection, in two halves, both verified against Microsoft's "Implicit intersection operator: @" article. For a RANGE: intersect with the formula cell's own row and column — a single-column range spanning the formula's row yields that row's cell, a single-row range spanning its column yields that column's cell, a 1x1 range yields itself, anything else `#VALUE!`. For a computed ARRAY (`FILTER`, `SEQUENCE`, …): the TOP-LEFT value. No `ComputedValueKind.Reference` may escape as a cell value. Consumed inside a function is unchanged. No spill.
- **S5 — dynamic arrays.** `FILTER`/`SORT`/`UNIQUE`/`SEQUENCE` in the consumed position only. Bare in a cell follows S4's array half. The no-spill limitation is documented prominently.
- **S6 — AGGREGATE.** Full `function_num` 1-19 x options 0-7. Documented caveat: MySheet has no hidden-row model, so options 1/3/5/7 behave as 0/2/4/6, consistent with `SUBTOTAL` already mapping 101-111 to 1-11.
- **S7 — versioning.** 3.17.0 minor, with a new one-way compatibility subsection in `docs/serialization.md` covering both the new union tags and the new third `Workbook` member.
- **S8 — blocking gaps.** `ROW`/`COLUMN` must accept any reference-producing expression. `SUMPRODUCT` must accept a computed-array argument.

### What the verification measured that changes the shape of the work

- **`SUBTOTAL` silently returns wrong numbers for array arguments — a pre-existing bug, not part of the request.** Measured: `SUBTOTAL(9,ROW(A1:A3))` = 1 where Excel gives 6; `SUBTOTAL(9,(A1:A3<>0)*1)` = `#VALUE!` where Excel gives 2; `SUBTOTAL(2,…)` = 0 and `SUBTOTAL(3,…)` = 1 where Excel gives 3. The range form (`SUBTOTAL(9,A1:A3)` = 14) is correct. The feature request lists SUBTOTAL as "confirmed working". Phase 2 fixes it, because AGGREGATE shares the same scan.
- **There are TWO frozen full-`Workbook` wire goldens, not one.** `CellStoreTests.PreChangeCellsWireGolden` and `SheetNameInterningTests.PreInterningWireGolden`. The second was missed by name-based grep because the `0x02` member-count byte appears as base64 `Ag`. Measured: adding the third `Workbook` member fails BOTH.
- **A C# record's `with` copies private fields, so a memoized column index goes stale.** Measured: after warming the memo, `table with { ColumnNames = [...] }` reports a missing column as absent and a reordered column at the wrong index. The memo must live off the record.
- **`RecalculationEngine.EnsureFresh` consumes the staleness signal as a side effect.** Measured: `EstimateImpact` followed by `Recalculate` leaves a stale value behind, because whichever public method runs first refreshes the snapshot. The invalidation must be sticky.
- **`NamedReferences.ValidateName` cannot be reused for table names.** Measured: it rejects `Tabela1` and `Table1` — Excel's own default table names — because `Parser.IsCellReference` accepts any letters-then-digits string with no column/row bound. Reusing it would make the loader unable to register almost every real table. The same fact forces the parser's bracket arm to sit BEFORE the `IsCellReference` check.
- **A version bump alone leaves stale VALUES behind.** Measured on the shipped `DefineName`: after widening a name's range, `SUM(name)` kept its old answer even through `Recalculate` reporting `rebuilt=True`. Only `InvalidateCache` clears layer 3. `DefineTable` would inherit this; the fix repairs `DefineName` too.
- **The array/mini-CSE gap exists for defined names too and is left unfixed.** Measured: `COUNT((Rng<>"")*1)` = 1 where the literal range gives 3. After this work, tables will be right and defined names still wrong — a visible inconsistency. See the open decisions.

## For Future Agents

As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to
`Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with
zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases
are done, fill in **Final Recap** and **Deployment Plan**.

Rules specific to this repo, which override habit:

- **TDD is not optional here.** Write the failing test, run it, confirm it fails for the stated reason, then implement. Several items below exist only because a verifier ran the design and watched it fail.
- **A fixture value that coincides with the correct answer is an assertion that cannot fail.** When a test distinguishes "read from cache" from "computed", the cached value must be a deliberate lie.
- Tests run with `dotnet run --project tests/<proj>/<proj>.csproj`, NOT `dotnet test` (TUnit on Microsoft.Testing.Platform). Filter with `-- --treenode-filter "/*/*/<Class>/*"`.
- Gates before any commit: `dotnet csharpier format .` then `check .`, `dotnet build Danfma.MySheet.slnx -c Release` with zero warnings, and both suites green. Baseline measured for this plan: **1203** core + **88** Excel.
- `CHANGELOG.md` is generated by versionize — never hand-write it. Use Conventional Commits; scopes in use are `parser`, `eval`, `serialization`, `recalc`, `dirty-graph`, `excel`.
- **Union tags must be assigned in ONE coordinated edit.** Two phases append `[MemoryPackUnion]` tags. The next free tag is 322 as of writing. If two phases each write `322`, MemoryPack fails at type initialization with a duplicate-tag error, not at compile time. Re-count `Expression.cs` before trusting any number, and fix the stale "Add new tags at 319+" comment while you are there.
- `docs/pt-BR/` is a full mirror (11 files, identical names). Every doc edit needs its twin. `README.md` has no mirror.
- Golden-value tests cite the Microsoft docs page plus its GUID and the fetch date. Fetch it at implementation time; do not invent a GUID.

## Phases

Each phase lives in its own file so a future agent loads only the one it is working on. Order below is a
topological sort of the design dependencies, then earliest user value. Phases 1 and 2 have no dependencies and
each delivers something verifiable on its own; the corpus formulas need Phase 1 even after Phases 3-5 land,
which is why it goes first.

| # | Phase | File | Items | Blockers | Majors | Status |
| --: | --- | --- | --: | --: | --: | --- |
| 1 | ROW/COLUMN over any reference, Excel's implicit intersection, array-native SUMPRODUCT | [`phase-1-reference-semantics.md`](structured-table-references-and-aggregate/phase-1-reference-semantics.md) | 27 | 2 | 2 | **Complete** — `feat/reference-semantics`, `32a67f6..c04e78b`, pending merge |
| 2 | AGGREGATE(function_num, options, ref1, [k]) on a shared SUBTOTAL/order-statistics core | [`phase-2-aggregate.md`](structured-table-references-and-aggregate/phase-2-aggregate.md) | 18 | 1 | 1 | Not started |
| 3 | Table model, Workbook.DefineTable/Tables, and the third serialized Workbook member | [`phase-3-table-model-registry.md`](structured-table-references-and-aggregate/phase-3-table-model-registry.md) | 26 | 3 | 2 | Not started |
| 4 | Bracket lexing and the structured-reference grammar | [`phase-4-lexer-parser.md`](structured-table-references-and-aggregate/phase-4-lexer-parser.md) | 18 | 2 | 1 | Not started |
| 5 | Structured-reference resolution and cross-cutting graph integration | [`phase-5-resolution-and-graph.md`](structured-table-references-and-aggregate/phase-5-resolution-and-graph.md) | 24 | 1 | 5 | Not started |
| 6 | The .xlsx loader reads <table> parts into Workbook.Tables | [`phase-6-excel-loader.md`](structured-table-references-and-aggregate/phase-6-excel-loader.md) | 25 | 1 | 4 | Not started |
| 7 | FILTER / SORT / UNIQUE / SEQUENCE as mini-CSE producers | [`phase-7-dynamic-arrays.md`](structured-table-references-and-aggregate/phase-7-dynamic-arrays.md) | 21 | 2 | 4 | Not started |

Keep this table's Status column in step with each phase file's own `Status:` line — it is the first thing a
resuming agent reads.

## Open decisions that need a human, not more analysis

These block no phase from starting, but each one changes a shipped behaviour or an advertised guarantee. They
are collected here so they are answered deliberately rather than by omission.

- **Should the array/mini-CSE gap be fixed for defined names too, or shipped inconsistent?** It is a measured, pre-existing silent-wrong-answer bug (`COUNT((Rng<>"")*1)` = 1, truth 3). Fixing it means threading an `EvaluationContext` overload through `ArrayEvaluation.IsArrayEligible`/`Probe` and deciding scalar-vs-range at four call sites — roughly the size of three items again. Fixing it makes tables and names consistent; deferring ships a documented inconsistency.
- **Should `DefineName` start rejecting a name a table already uses?** It is Excel's rule and keeps one invariant, but it adds a failure mode to a shipped API, and in MySheet the two syntaxes are disjoint so the collision is harmless to resolution. Alternative: enforce uniqueness only in `DefineTable` and document the asymmetry.
- **Should a follow-up teach `SaveAsExcel` to emit the `<table>` part?** S3 says no. The measured cost of no: load a table workbook, export in `FormulaMode.Formulas`, and Excel opens the result showing `#NAME?` because the formula now parses but the part is absent. Today that is impossible because the formula never parsed. `ExcelExportOptions` has no warning channel, so documentation is the only mitigation inside S3.
- **Should `MergeIntoExcel` stop writing a degraded cell's stale cached number?** Already shipped and documented as a limitation, but under P0 the Excel-faithful answer is arguably to leave the target's formula alone so Excel recomputes, rather than freezing a number no engine derived from current inputs.
- ~~**Should `ROWS`/`COLUMNS`/`AREAS` be brought in line with `ROW`'s new error propagation?**~~ **DECIDED 2026-09-08: yes, in Phase 1, under P0.** `ROWS(NoSuchName)` returns 1 today where Excel gives `#NAME?`. It is a behaviour change to four functions and ships as `feat(eval):`, not `fix`. See Phase 1's "Scope addition" section.
- **Should bare `=SEQUENCE(5)` really return 1?** That is Excel's verified `@`-on-array behaviour and what Excel writes when upgrading a legacy formula, but it is not what modern Excel does with the un-prefixed formula. A user who types `=FILTER(A:A,B:B>0)` and sees one value will read it as a bug. `#VALUE!` is the defensible alternative.

### Questions settleable without a human, by experiment

Several phases need an Excel oracle for a semantic this plan could not verify (what `[#Totals]` returns when
the table has no totals row; whether implicit intersection of a 2-D range uses both axes; whether SUBTOTAL
skips a nested AGGREGATE; how SORT orders mixed types). Neither machine in this session has Excel or
LibreOffice. Two oracles are available and should be used before the affected arm is coded:

- **The self-verifying fixture.** Author a `.xlsx` in real Excel containing the formula in question, then read the cached `<v>`/`t="e"` Excel stored next to it — the loader already reads that cache, so the fixture answers the question by itself and becomes a regression test.
- **Aspose.Cells 26.6.0**, already referenced by `benchmarks/Danfma.MySheet.Benchmark`. It resolves every formula in the original report correctly, so it can serve as the oracle for anything Excel is not on hand for.

## Final Recap

_(write when all phases complete)_

## Deployment Plan

_(write when all phases complete)_
