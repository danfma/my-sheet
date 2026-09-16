# Sweep: robustness and parser gaps (items 19, 27, 52, 53, 54) plus the re-measurement of items 14-30 on Aspose.Cells 26.7.0

Close the only confirmed evaluation crash (item 19 `WORKDAY`/`WORKDAY.INTL` overflow), confirm item 27 (huge column references), and add the three missing formula syntaxes that real spreadsheets use:
- array constants `{1,2;3,4}` (item 53);
- `INDEX`'s `area_num` over union references (item 52);
- structured-reference column spans `Tabela1[[A]:[B]]` (item 54).

In parallel, re-measure the older open items 14-30 on the 26.7.0 oracle to prune the backlog. Out of scope: fixing items 14-18, 20-26 and 28-30 themselves, which only get classified here, and items 44-51 and 55-56. Source of truth for the items: `plans/structured-table-references-and-aggregate/phase-11-excel-compatibility-sweep.md`.

**Decisions (user, 2026-09-15):**
- Delegation: all implementation and review runs on opencode agents (GPT-5.6 sol for hard work and reviews, terra for routine work), bound by `.superpowers/sdd/AGENT-RULES.md`. Each phase runs an implement/review loop of at most 5 rounds.
- MemoryPack: array constants may add ONE new AST node, appended to the union as tag 329 (append-only, with a Save/Load round-trip test).
- Final gate: a Claude Fable 5.1 code-quality review of the whole sweep range, then its fix wave.
- Release: local ff-only merge into `main` with gates, then notify the user. The user owns push and release.

**Controller ledger:** `.superpowers/sdd/sweep-robustness-parser/progress.md`. **Briefs and reports:** `.superpowers/sdd/sweep-robustness-parser/`.

## For Future Agents

As work proceeds:
- mark checkboxes `- [x]` as items complete;
- when a phase is done, set its status to `Complete` and write its **Phase Summary**: what was done, key decisions, and anything needed to continue with zero context;
- run the phase's **Verification Plan** and record the result before moving on.

When all phases are done, fill in **Final Recap** and **Deployment Plan**.

Every oracle number must be measured on Aspose.Cells 26.7.0, one formula per workbook, in both PLAIN and CSE modes. Never pin an unmeasured value (AGENT-RULES section 3). Time performance claims only with uncached evaluations (`--open-range-lookup` pattern).

## Phase 0: Re-measure items 14-30 on 26.7.0 (no code)
Status: Complete

- [x] Rebuild every fixture in items 14-30. Measure Aspose 26.7.0 in PLAIN and CSE, and MySheet at `856a6c5`.
- [x] Classify each item: STILL-DIVERGENT / NOW-MATCHES / ORACLE-CHANGED / ROBUSTNESS / NOT-MEASURABLE.
- [x] Confirm the robustness rows: item 19 (still throws) and item 27 (value returned, expected `#REF!`).
- [x] Update the sweep file: close NOW-MATCHES items with evidence, re-record ORACLE-CHANGED numbers, and order the remaining items by impact.

### Verification Plan
- `reports/phase-0-remeasure.md` has a row for every formula in items 14-30, with fixture, 26.6 recorded, 26.7 PLAIN/CSE and MySheet values.
- `git status --short` in `MySheet-oracle` is empty.

### Phase Summary
Terra re-measured every formula in items 14-30 on Aspose.Cells 26.7.0 (report `.superpowers/sdd/sweep-robustness-parser/reports/phase-0-remeasure.md`). No cited 26.6.0 oracle value moved.
- **STILL-DIVERGENT:** 14, 15, 16, 17, 18, 20, 21, 23 (partly), 24, 28 and 29.
- **NOW-MATCHES:** 22, which Phase 11a had already closed; this re-measurement confirms it.
- **ROBUSTNESS:** 19 still throws; the oracle gives 45362, but MySheet's contract is `#NUM!`. For 27 nothing throws, so the ruling is: an out-of-grid range is `#REF!`, and a bare invalid token is `#NAME?`, which matches the oracle. 26 and 30 depend on inconsistent oracle behaviour.
- **NOT-MEASURABLE:** 25, whose original fixture was never recorded.
- **Follow-up order** is recorded in the sweep file: 17, 18, 29, 28, then 14/16/20/21/15, then 23-24.

## Phase 1: Robustness: item 19 (WORKDAY overflow) and item 27 (huge column references)
Status: Complete (branch `feat/rp-robustness` @ `9a24eb9`, reviewed Clean at round 1)

- [x] Red pins: `WORKDAY(45362,-3000000000)` and `WORKDAY.INTL(45362,-3000000000)` must not throw. Cover the interval ends (`int.MinValue`, `int.MinValue+1`, `±3e9`).
- [x] Contract: no evaluation throws, and a result outside Excel's date range is `#NUM!`. Fix the saturating `(int)Math.Truncate` / `Math.Abs` path in `WorkdayFunctions.cs`.
- [x] Audit sibling date functions for the same truncation/overflow pattern. Pin and fix any that throw.
- [x] Item 27: confirm and pin `#REF!` through real formulas (`=COUNTA(FXSHRXW1:FXSHRXW1)`, a 24-letter column), on both sides of the boundary.

### Verification Plan
- Every new pin is green, and a mutation restoring the saturating cast turns them red.
- Gates: `dotnet csharpier check .`; `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` with 0 warnings; both suites `failed: 0`; union count 328.

### Phase Summary
- **Item 19** (`c8ffcfa`): `Workday.Advance`, shared by WORKDAY and WORKDAY.INTL, rejects non-finite and out-of-`Int32` day counts before converting them, and widens `Math.Abs` to `long`. Every extreme now returns `#NUM!` instead of throwing `OverflowException`. MySheet's contract is deliberately not oracle parity: Aspose returns 45362, an out-of-range serial, or throws.
- **Audit and review fuzz**: no other date function throws. The fuzz ran 450 evaluations (numeric-slot and holiday extremes) with 0 exceptions.
- **Item 27** (`9a24eb9`, tests only):
  - `COUNTA(FXSHRXW1:FXSHRXW1)` → `#REF!`
  - a 24-letter bare token → `#NAME?` (oracle agrees)
  - `XFD1` → 0 (valid)
  - `XFE1` → `#REF!`

  The `CellAddress` guard still protects the public A1 parsing API.
- **New divergence, registered for Phase 5**: `NETWORKDAYS(45362,45363,3000000000)` returns 2 against the oracle's `#NUM!`.

## Phase 2: Item 53, array constants
Status: **Complete @ `8234e45`**. It closed after review round 9 came back Clean and the controller's final perf gate passed. The history below is kept for the record.
- Review round 1: I1 `{-0}` must be rejected; I2 `ROW({1})` must give `#REF!`; M1 per-consumer arms. All three closed.
- Review round 2 found three new issues:
  - scalar `ROW(1)` gives `#VALUE!` where the oracle gives `#REF!`;
  - VLOOKUP wildcard matching and LET-bound tables fail on the array route;
  - open HLOOKUP is +23% slower.
- Fix round 2 (`c7853db`, `33364c8`, `bc4ffa0`, `aa84676`) closed the ROW/COLUMN issue, the table-source decision and the reference probe order.
- The controller perf gate (`reports/phase-2-controller-perf-r2.md`) passed XLOOKUP, MATCH, XMATCH and VLOOKUP. The +83% was an artifact of the replaced harness. Two HLOOKUP problems remain: open approximate is bimodal, and closed exact is +14%.
- Review round 3 found:
  - unpinned tilde moves in COUNTIF and XMATCH;
  - wildcard classification and pattern building run per candidate inside the loop;
  - the tilde rows `MATCH("a~*")` and `VLOOKUP("~")` still diverge.
- Fix round 3 (`c85a8f5`, `507dd3b`):
  - Closed: the harness restore (with text and wildcard scenarios added), the tilde pins, the wildcard strategy hoisted out of the loop, and `MATCH("a~*")`.
  - Registered as sweep item 65: the tilde rule differs by family (lookup vs criteria vs SEARCH).
  - Controller perf gate r3 closed P1 and P2 (the HLOOKUP issues) and found P3: closed-range approximate VLOOKUP is +80% at steady state (0.00118 → 0.00212 ms, 4/4 runs).
  - It also showed that the committed one-warm-up gate mostly measures the snapshot build.
  - Fix round 4 (`8f7fec1`, `83e3e67`) is Clean.
    - Cause, measured by counters: a `UsesWildcards` probe ran on every numeric lookup, ahead of the snapshot index.
    - Fix: guard the probe with the key kind. Closed approximate VLOOKUP is now -2.7% against base.
    - The harness gained a `--steady` mode.
  - Review round 4 (fix rounds 3 and 4) found one Important issue: `XMATCH("a*",{"x","ab"},2)` returns `#N/A` where the oracle gives 2. Mode 2 over an array constant is broken; the range route works. Everything else is closed.
  - Controller perf gate r4 was inconclusive: load 10-13 from a stray process at 100% CPU plus agents. It reruns with no agent active.
  - Fix round 5 (`b4d2b88`) is Clean. XMATCH mode 2 now materialises computed arrays into the shared matcher.
    - Registered an oracle defect: reverse wildcard search over an array returns an impossible position, 1 where the range answer is 2. MySheet keeps 2.
    - Review round 5 returned Fixes required.
      - The mode-2 branch materialises closed references (ranges, names, table refs, INDEX/OFFSET) on every evaluation and bypasses the snapshot.
      - 2D arrays return 2 where the oracle gives `#VALUE!`.
      - Evaluation order and draw count are unchanged.
    - Fix round 6 (`ee7b81c`): references stay on `ExpandCached`, backed by a route pin. One 2D shape rule returns `#VALUE!`, and modes are validated before the array error. Blocked only on oracle and probe coverage, because the table registration in the probe failed.
    - Review round 6 returned Fixes required.
      - The route pin was vacuous (Critical): the assertion itself built the snapshot.
      - The shape rule was restricted to `ArrayConstant`, so SEQUENCE, IF and LET in 2D return 1 where the oracle gives `#VALUE!`.
      - Open ranges and unions are not checked.
    - Fix round 7 (`6792f6d`) is Clean.
      - A route seam proves references never materialise; every mutated form goes red.
      - Admission is general.
      - Open rectangles have a shape; unions return `#N/A`, as measured.
      - Draw counts are unchanged.
    - Review round 7 returned Fixes required: `XMATCH(1,SEQUENCE(1000),0)` allocates 70x more than before (345 → 24,177 B per evaluation) because every computed array is materialised.
      - The diagnostics hook, a null check, was accepted by controller ruling.
      - The `TryBuildIf` reuse is safe and the draw count went down.
    - Fix round 8 (`28bac64`) is Clean: a single `TryEvaluateStream` for shape and every scan, and all modes stream.
      - Allocations are now below the original baseline: mode 0 153 B vs 345, mode 2 32k vs 56k.
      - The route pin went red before the fix; the probe output is identical.
    - Review round 8 (terra) is running.
    - Controller final perf gate (`856a6c5` vs `28bac64`):
      - Tables (TABLE_ONLY, 10 pairs) pass within +8.4%.
      - Gap-free `XMATCH(A:A)` is consistently +65%, with later scenarios in the same process bimodal. Cause, confirmed in the code: `TryGetShape` calls `ToBoundedRange` on every evaluation, allocating a `RangeReference` plus two id strings.
    - Review round 8 is Clean, with 0 findings: allocations reproduced, O(1) `ElementAt`, a single stream, draws unchanged, attack rows equal CSE.
    - Fix round 9 (`8234e45`) is Clean. `XMATCH(A:A)` drops from 440 to 32 B per evaluation (base 37) through the structural short-circuit, and the probe output is identical.
    - The controller reran final-gate part 1 with no agent active (6 pairs). The bimodal slowdown and the table regressions are gone.
      - Residual: gap-free XMATCH at +12.7%, with half the runs overlapping the base. Allocation is back at base level.
      - Controller ruling: the residual is accepted. **Phase 2 perf gate PASS @ `8234e45`.**
    - Review round 9 is running on terra. Phase 2 closes when it comes back Clean.
    - Controller perf gate r4 then runs against `b4d2b88`.

- [ ] Measure the oracle first:
  - shapes: row `{1,2,3}`, column `{1;2;3}`, 2D `{1,2;3,4}`, mixed types, negatives, errors;
  - invalid forms: ragged arrays, expressions inside braces, empty elements;
  - consumers: SUM/AVERAGE/SUMPRODUCT, INDEX, MATCH, XLOOKUP, VLOOKUP, ROWS/COLUMNS, IF/AND, LET, names;
  - bare-cell reading, and use in a criteria slot.
- [ ] Tokenizer and parser: `{ … }` with `,` as column separator and `;` as row separator. New AST node `ArrayConstant`, MemoryPack union tag 329.
- [ ] Evaluation as a computed array through the existing array machinery. The Phase 11c fence holds: array constants are refused in criteria range slots, as SEQUENCE is.
- [ ] Formula-text write-back (`FORMULATEXT`, expression text), a MemoryPack Save/Load round trip, and xlsx load/export.
- [ ] Docs twins.

### Verification Plan
- Every measured oracle row is pinned. A parser mutation dropping the row separator turns the 2D pins red.
- A Save/Load round-trip test passes; union count is 329; `DefinedNameArrayEligibilityTests` and `ArrayBindingTests` are unchanged.
- Gates as in Phase 1.

### Phase Summary
- **Scope delivered:**
  - Array constants: tokenizer, parser, the `ArrayConstant` node (MemoryPack union count 329), write-back, and evaluation through `IArrayProducer`.
  - Along the way:
    - one table-source decision for VLOOKUP and HLOOKUP (`LookupTable`/`LookupGrid`);
    - shared wildcard matching resolved once per evaluation, with `~` escape pins;
    - ROW and COLUMN over scalars and computed arrays return `#REF!`;
    - XMATCH: mode validated first, references stay on the snapshot or expansion route, computed arrays stream, one 2D shape rule (open 2D gives `#VALUE!`, unions give `#N/A`), and an allocation-free open-range shape check.
- **Oracle rulings:**
  - Registered defect: reverse wildcard search over an array; MySheet returns 2.
  - Sweep items 65 (lookup tilde family), 69 (TRANSPOSE), 70 (intersection operator) and 71 (lifted XMATCH array).
- **Performance:** controller gates are interleaved, run on an identical harness, and measure steady state.
  - Tables are within +8.4%.
  - XMATCH allocation is at base level: 32 B, and 153 B for a computed array.
  - The gap-free XMATCH residual of about +12% was accepted by controller ruling.
- **Loops:** 9 fix and 9 review rounds.
  - Defects found by review: a broad branch, a vacuous route pin, a concrete-type check, 70x allocation, per-evaluation `ToBoundedRange`.
  - Defects found by the controller: harness replacement, a range 38x too small, per-candidate wildcard work, a numeric wildcard probe.
- **Suites at `8234e45`:** core 3473/0, Excel 138/0.
- **CLOSED @ `8234e45`.**

## Phase 3: Item 52, INDEX `area_num`
Status: Complete (review round 2 Clean). Branch `feat/rp-index-areas` @ `8e3dc12`.

- [x] Measure the oracle first:
  - `INDEX((A1:A3,C1:C3),r,c,n)` for n = 1, 2, 3 (out of range), 0, negative, fractional, text, and error;
  - zero row or column with `area_num`;
  - `SUM`/`AREAS`/`ROWS` over the result;
  - single-area references with n = 1 and n = 2;
  - unions that include table references. Aspose threw `Invalid table column` because the probe did not register the table as a ListObject. The row is re-measured in Phase 4 with the table registered on both engines.
- [x] Registry arity 4. `Index` selects the area from a union reference before applying the shared `ValidateAxes` rule.
- [x] Docs twins: the INDEX row.

### Verification Plan
- Every measured row is pinned. A mutation ignoring `area_num` turns the pins red. Existing `IndexZeroAxisTests` stay green.
- Gates as in Phase 1.

### Phase Summary
- **Commits:**
  - `afc3986` feat: `area_num` selects a union area.
  - `c68856f` docs.
  - `380a7fe` fix: validate the missing sheet only on the selected area.
  - `6650eed` refactor: one `TrySelectArea` helper for both the value and reference paths.
  - `8e3dc12` test: adversarial and single-draw pins.
- **Behaviour** (Aspose 26.7.0, PLAIN and CSE agree):
  - `area_num` is 1-based and truncated;
  - 0 or a negative value gives `#VALUE!`; a value beyond the area count gives `#REF!`;
  - an unselected missing-sheet area does not poison the result (`SUM(INDEX((Ghost!A1:A3,C1:C3),0,1,2))` = 60).
- **Registered oracle defect:** `INDEX(U,2,1,2)` over a defined-name union gives Aspose `#VALUE!` against its own `AREAS(U)` = 2. MySheet keeps 20.
- **Review:** round 1 found I1, I2 and M1; round 2 was Clean.
  - Mutations: 2/36, 4/36, 25/36 red.
  - Suites: core 3428/0, Excel 137/0.
  - Union count 328; the divergence probe changed 0 rows.
- **Deferred to Phase 6:** the array-constant `area_num` row `INDEX((A1:A3,C1:C3),2,1,{2})`, once the Phase 2 parser is integrated.

## Phase 4: Item 54, structured-reference column spans
Status: **Complete @ `bc31695`** (integrated onto `feat/rp-integrate` as `e6c6796`/`b8e6cf9`). The reversed-span divergence is registered as an oracle defect; the history below is kept for the record.
Commit `35bd5ef` on `feat/rp-column-spans`: a column span is modelled as a first/last column pair and resolves through `Table.GetRegion`, so INDIRECT is covered too.
- Oracle: 30 rows, table registered in both engines, PLAIN = CSE. Pins 42/42, mutations red.
- Four-argument INDEX over a table union is recorded for Phase 6.
- Review round 1 returned Fixes required. Two findings were ruled pre-existing and moved to sweep items: the space intersection operator (item 70) and xlsx export without tables (item 72).
- Open finding: `INDEX` over a reversed span returns 10 in MySheet and `#REF!` in Aspose, although SUM over the same span agrees at 660. The divergence probe also lacked span rows.
- Fix round 1 ended Blocked, as it should: neither condition held.
  - For the reversed span, Aspose gives COLUMNS 0, INDEX `#REF!`, COUNTIF 0 and OFFSET 6600, yet SUM 660 and a two-column XLOOKUP row (220). Those results contradict each other.
  - Controller ruling: registered as an oracle defect. MySheet keeps the normalised forward region, which is its current behaviour.
- Fix round 2 (`bc31695`, test-only): the reversed-span rows are pinned to the forward region, with a comment citing the defect. The spelling pin keeps the user's text. The real divergence probe shows spans moving from parse error to the CSE value.
  - The missing-column contract and the current-row exclusion (now sweep item 73) were not treated as blockers.
- **Phase 4 CLOSED @ `bc31695`.** It is integrated after Phase 2.

- [ ] Measure the oracle first:
  - `Tabela1[[Valor]:[Qtd]]`;
  - with special items (`[#Data]`, `[#Headers]`, `[#Totals]`, `[#All]`, `[@]`);
  - reversed column order, non-adjacent columns, missing column;
  - consumers: SUM, COUNTIF, XLOOKUP return, INDEX;
  - a header-only table.
- [ ] Parser plus `TableReference` resolution of a contiguous column span. Formula-text round trip.
- [ ] Docs twins.

### Verification Plan
- Every measured row is pinned, and the parser round-trip test passes.
- Gates as in Phase 1.

### Phase Summary
- **Commits:** `35bd5ef fix(parsing): support structured column spans` and `bc31695 test(parsing): pin reversed structured spans`, on `feat/rp-column-spans` (base `83e3e67`).
- **Design:** `TableReference.LastColumnName` is an appended MemoryPack member, so the union count stays at 329. The syntax takes a top-level colon between two decoded endpoints, with the spacing canonicalised. Resolution goes through `Table.GetRegion`, one normalised rectangle, so `INDIRECT` works through the same path.
- **Oracle:** 30 rows, with the table registered in both engines and PLAIN = CSE. They cover data, totals, header-only tables, escaped headers, every row specifier, reversed spans, single-column spans and `INDIRECT`.
- **Rulings:**
  - Reversed spans are a registered ORACLE DEFECT: Aspose gives COLUMNS 0, INDEX `#REF!`, COUNTIF 0 and OFFSET 6600, but SUM 660 and a two-column XLOOKUP result. MySheet keeps the forward region and the user's spelling.
  - A missing column gives `#REF!`, under the existing contract.
  - The current-row form `[@...]` stays excluded; sweep item 73.
- **Moved to sweep items as pre-existing:**
  - item 70: the space intersection operator;
  - item 72: xlsx export writes no tables.
- **Review:** round 1 found the reversed-INDEX divergence. Fix round 1 was Blocked, correctly, because neither pre-set condition held. Fix round 2 added test-only pins. The controller waived review round 2 because the commit is test-only and its mutation was verified.
- **Suites:** core 3488/0, Excel 138/0. The divergence probe with spans moved rows only toward CSE, apart from the registered defect.
- **CLOSED @ `bc31695`.** It is integrated after Phase 2.

## Phase 5: Follow-ups from the consumer's calc-divergences document
Status: **Complete** (2026-09-16).
- 5a is Complete (`reports/phase-5a-divergences.md`).
- 5c is CLOSED @ `5196852`. The user decided the contract on 2026-09-15: normalise to `0`.
- 5b is CLOSED @ `aa683fb` after review round 8 came back Clean. It was integrated by fast-forward onto `feat/rp-integrate`, where the gates are green.

Source: `~/MYSHEET-CALC-DIVERGENCES.md` (user request 2026-09-15). As of 3.21.0 the document records these open items on the MySheet side:
- (A) structured references and `INDIRECT("Table[Col]")` (`#NAME?`/`#REF!` without registered table metadata);
- (B) blank lookup keys in `MATCH`/`COUNTIF` (70 cells where Aspose errors and MySheet returns a value);
- (C) IEEE negative zero (`=-SUM(zeros)` gives `-0` where Aspose gives `0`).

The document's other items are downstream harness work, or were already fixed in 3.20.0/3.21.0: Bugs 6, 7, 8 and 9, through sweep 31 item 31.

- [x] 5a: measure A, B and C on the oracle with fair fixtures (a table registered on both engines; blank-key variants: no cell, `=""`, empty string). Classify each as ENGINE-GAP, HARNESS-ONLY, ORACLE-INCONSISTENT or CONTRACT-CHOICE.
  - (A) is HARNESS-ONLY: with the table registered on both engines, all 10 forms agree, so the consumer must register table metadata.
  - (B) is an ENGINE-GAP: absent and empty-text keys in approximate MATCH, XMATCH, XLOOKUP, COUNTIF/COUNTIFS/SUMIF and the `"="` criterion.
  - (C) is a CONTRACT-CHOICE.
  - (D) the controls agree.
- [x] 5b: fix the (B) ENGINE-GAP. Brief: `phase-5b-blank-keys-brief.md`, finalised with addenda 1-3. Dispatched on sol at 16:00 on base `b8e6cf9`.
  - The first attempt ended its turn after announcing a step; the controller resumed it.
  - Implemented as `c31b0cf` (Clean):
    - rules per family (MATCH exact and approximate, XMATCH/XLOOKUP, VLOOKUP/HLOOKUP, criteria with `"="` and `"<>"`);
    - the absent key is classified once, via `LookupMatching.IsAbsentKey`;
    - 43 pins; every changed row matches CSE;
    - core 3646/0.
  - Controller 3-way perf gate: **PASS**.
    - TABLE_ONLY: VLOOKUP and COUNTIF +2-3%, SUMIF +6%.
    - The part-A flags on open-range VLOOKUP/HLOOKUP come from bimodal slow runs that also appear at base `856a6c5`, so they are environmental.
  - Review round 1 returned Fixes required:
    - I1: derived keys (INDEX, OFFSET, IF, defined names) are not classified as absent. `XMATCH` returns 1 where the oracle gives 3.
    - I2: approximate VLOOKUP with an empty-text key and no empty-text candidate returns 30 where the oracle gives `#N/A`.
    - I3: the "last empty text" preference has no pin.
    - The scope audit and the allocation check passed.
  - Fix round 1 (`e9ddbda`) is Clean.
    - One classification point, `LookupMatching.EvaluateKey`, over the resolved reference, with volatile functions evaluated once.
    - One shared V/H rule, `LookupGrid.FindLastExactText`.
    - Core 3673/0; allocations did not increase.
  - Controller perf gate at `e9ddbda`: **PASS**.
    - TABLE_ONLY put VLOOKUP, COUNTIF and SUMIF within +7%.
    - A numeric HLOOKUP flag was ruled out by a focused 12-round rerun: -0.2% to +0.8% against `b8e6cf9`. All commits alternate between the same two modes.
  - Review round 2 returned Fixes required:
    - I1: a derived key wrapped in LET is not classified.
    - I2: SUMIF with a derived key returning 0 was **wrong**. On the discriminating fixture `C1:C4 = 1,2,4,8` the oracle gives 1, so the SUMIF-only arm must go.
    - I3: approximate MATCH with empty text over a range returns 2; the oracle gives `#N/A`.
  - Fix round 2 (`f5fde75`) closed I1-I3.
  - Review round 3 returned Fixes required:
    - The exact empty-text family (MATCH/XMATCH/XLOOKUP/VLOOKUP/HLOOKUP/LOOKUP) accepts an absent candidate. This is item 59 core, pre-existing at `856a6c5`.
    - A multi-cell OFFSET is still evaluated twice in `Capture`/`EvaluateKey`, and the singleton classifiers are duplicated.
    - `LET(r,D1:D2,XMATCH(r,B1:B4))` returns `#N/A` where the oracle gives 3 (conditional ruling: fix only if the direct form already matches).
    - `SUM(r:B4)` is registered as item 74, with its 5b attribution rejected.
  - Fix round 3 (`7e1262d`), core 3722/0:
    - `ExactMatcher` routes empty-text keys through `IsExactText` (shared, no per-function arm).
    - LOOKUP takes the last exact text or `#N/A`.
    - Resolution happens once, through `ResolvedReferenceValue`. The stack diagnosis showed `f5fde75` had regressed the LET draw count (1 at `856a6c5`); it is 1 again.
    - F3 was registered as item 75, because the direct form diverges too.
  - Controller perf gate: PASS. The focused reruns resolved the flags; a +4-6% residual on `COUNTIF(XLOOKUP)` over a gapped column sits below the 10% threshold.
  - Review round 4 returned Fixes required:
    - Approximate XMATCH/XLOOKUP with an empty-text key and no candidate does not return `#N/A`.
    - The multi-cell key `#VALUE!` branch is unpinned. The controller widened this: the approximate-MATCH matrix regressed away from a consistent oracle (union) and from CSE (OFFSET/INDEX/LET/name).
    - LOOKUP over array constants: `LOOKUP(3,{1,2,3,4},{10,20,30,40})` gives 10 vs oracle 30.
    - Layering and dead code.
    - Docs twins incomplete.
  - Fix round 4 (`fe2c7bc`), core 3741/0:
    - An empty-text key in approximate modes stops at the exact scan (`#N/A` with no candidate).
    - Multi-cell keys follow one rule: a bare reference gives `#VALUE!`, a derived reference keeps the reference. Every row is back to its `b8e6cf9` value.
    - Both LOOKUP vectors go through `ArgumentFlattening.MaterializeVector`: `LOOKUP(3,{1,2,3,4},{10,20,30,40})` 10 → 30, computed vector `#VALUE!` → 20.
    - Classification moved to `ResolvedReferenceValue`.
    - Docs twins updated.
  - Perf gate for round 4: PASS.
  - Review round 5 (terra) returned Fixes required: `LOOKUP(2,{1,2,3}*TICK()^0)` draws the volatile function twice, and the `LOOKUP(1,{1,2}/0)` row changed without an oracle measurement.
  - Fix round 5 (`9f252ae`), core 3754/0:
    - LOOKUP's volatile computed vector draws once, through `TryUnresolvedError(preserveComputedArray)`.
    - Error elements in the vector are skipped per the oracle: `LOOKUP(2,1/(B1:B4=2),C1:C4)` = 20, `LOOKUP(1,{1,2}/0)` = `#N/A`.
    - MATCH's double draw, pre-existing since `b8e6cf9`, is registered as item 76.
  - Review round 6 (sol, independent) returned Fixes required:
    - The multi-cell matrix is confirmed: 126/126 cells equal `b8e6cf9`. LOOKUP draws once.
    - `LOOKUP(2,IF(TRUE,NoSuch,B1:B4))` went `#NAME?` → `#N/A` (regression from round 5).
    - `LOOKUP(1,NoSheet!A1:A3*1)` gives `#N/A` vs oracle `#REF!`.
    - The approximate-MATCH absent-cell mutation turns no pin red.
  - Fix round 6 (`5092106`), core 3773/0:
    - Scalar structural errors in the vector propagate: IF/CHOOSE selecting `NoSuch` gives `#NAME?`; a missing sheet gives `#REF!`.
    - Array elements with errors are skipped.
    - The approximate MATCH absent pin now kills its mutation.
    - `LOOKUP(1,1/0)`, a PLAIN/CSE split, is registered as item 77.
  - The controller found duplicated IF/CHOOSE selection logic across `ReferenceGuard`/`ReferencePosition`, plus a result-slot heuristic taken from one row. Fix round 7 (`b466562`, `2bbda0d`) consolidated them into one `LookupVector` component:
    - the shared helpers are restored to their pre-`9f252ae` shape;
    - existing tests only gained pins, with no expectation changed;
    - core 3784/0;
    - a CHOOSE result-slot split and the short result-vector quirks were registered as items 78 and 79.
  - Integration gates at `2bbda0d`: green (core 3784/0, Excel 138/0).
  - Review round 7 (sol) returned Fixes required, having confirmed the 126-cell matrix, 15 mutations and every draw route:
    - `LOOKUP(2,B1:B4,CHOOSE(9,C1:C4))` moved from `#VALUE!` to `#N/A`, while CSE gives `#VALUE!`.
    - `LOOKUP(1,(A1:A3)*NoSheet!A1)` moved from `#N/A` to `#REF!`, while both oracle modes give `#N/A`.
    - `LOOKUP(2,INDEX(B1:B4*1,0))` gives `#N/A` vs oracle 2; pre-existing, handled by a conditional ruling.
  - Fix round 8 was dispatched and then PAUSED at the user's request on 2026-09-15 20:07, aborted mid-run. Its partial test edit was saved to `scratchpad/phase-5b-fix-r8-partial.diff` and reverted, so `feat/rp-blank-keys` is clean at `2bbda0d`.
  - **Resume:** re-dispatch `phase-5b-fix-r8-brief.md` unchanged on sol, then review round 8 (`phase-5b-round-8-review-brief.md`, fill the head), then integrate and finish Phase 6.
  - **User scope, 2026-09-16:** first close what is on the consumer's calc-divergences list. Of the open sweep items only item 59 (family B, this phase) qualifies: A is HARNESS-ONLY, C is closed in 5c, and the Bug 9 residual, `IF("1")`, Bug 6 and Bug 8 already agree (5a section D).
    - Fix round 8 keeps T1 and T2 (phase regressions) and drops T3, registered as item 80.
    - Then Phase 6 and a local merge; the user pushes and releases.
    - The other open items (14-30, 44-56, 58, 60-80) follow after the release.
  - Review round 2's brief added two checks:
    - whether SUMIF with a derived key returning 0 while COUNTIF returns 1 is an oracle defect;
    - approximate MATCH with empty text over `0, blank, 5`, which diverges.
- [x] 5c: the user chose to normalise `-0` to `0` in computed values, so TEXT and the value read by consumers match Excel/Aspose. Brief: `phase-5c-negative-zero-brief.md`. Its files are disjoint from Phases 2-3, so it runs in parallel.
  - Implemented in `3b4cfe8` on `feat/rp-negative-zero`.
    - `ComputedValue.Number` maps both zero signs to `+0`.
    - `NumberFormatting.Format` retries with the positive magnitude when the rendered text is zero. Measured: `TEXT(-0.0001,"0.0")` is `"0.0"` on the oracle.
    - `RoundToDigits` canonicalises a rounded zero, which covers FIXED.
  - Review round 1 returned Fixes required.
    - The TEXT retry parsed the display text, so it missed `0%`, `$0.00`, literal text and literal minus signs, and it broke the `"0;(0)"` section.
    - The unconditional mutation went undetected.
    - TEXT/FIXED showed a plausible +7% slowdown.
    - Fix round 1 (`cb01536`, `3cb6917`, `bcef1df`) returned Clean. Measured rule over 105 rows:
      - the section is chosen by the ORIGINAL sign, before rounding;
      - a single-section format drops the automatic minus only when the scaled numeric pattern rounds to zero;
      - the decision is made before rendering, and the value renders once.
    - Review round 2 returned Fixes required, with one finding.
      - Closed: every round-1 finding, 10 of 11 rule-attack rows, and the `bcef1df` guards (accepted).
      - Open: an empty selected section passes through to .NET, so `TEXT(-0.4,"0;")` returns `"0.4"` where the oracle gives `""`. Fix round 2 (`57b0904`) is Clean: a selected empty section renders `""`, 11 rows pinned. Review round 3 is Clean: 0 findings; the 6-row attack matches CSE. Semantics closed at `57b0904`. The controller TEXT/FIXED gate is inconclusive: load was 12-13 because of a stray process at 100% CPU plus agents. The slowdown held in 5 of 6 runs: TEXT +24-36%, FIXED +7%, arithmetic control -0.6%. [Likely] cause: the format analysis runs on every evaluation. Perf round 1 (`94cae69`) is Clean: a bounded 256-entry cache per format string, following the `RegexCache` pattern. A counter confirmed the analysis dropped from 2,000 runs to 1, and the probe output is byte-identical. Its review found three issues: a non-atomic capacity check, no cap on key length, and duplicate factory runs under concurrency. Controller ruling: accepted, because all three match the existing `RegexCache` pattern. The hardening of both caches is registered as sweep item 68. The TEXT/FIXED gate at `94cae69` still shows +11-18% on TEXT, with FIXED +2.7% and the control +1.4%, at load around 8. Perf round 2 measured allocations and found them identical to base, which refutes the allocation hypothesis; no commit was made. Controller gate r2 confirmed a real CPU cost (+11-22% on TEXT; FIXED and the control flat). Reading the code found `ConcurrentDictionary.Count` on every cache hit (it takes all locks) and `Math.Pow` on every call. Perf round 3 (`5196852`) removed both, with IL proof and a byte-identical corpus. Controller gate r3 shows TEXT at +1% to +11%, with the control at +1.9%. The residual is the cost of the measured sign and section rule, and the controller accepted it. **5c CLOSED @ `5196852`.**
      - Sweep items registered: 66 (`@`), 67 (conditional sections).
    - The controller runs the TEXT/FIXED performance gate after fix round 2.
  - Pre-existing divergences it found were registered as sweep items 60-64.

### Verification Plan
- `reports/phase-5a-divergences.md` covers every formula family in A-C with both engines and every fixture variant.
- Each fix passes the Phase 1 gates, pins the measured rows, and runs the divergence probe with no row moving away from CSE.

### Phase Summary
- **5a:** measured the consumer's calc-divergences document and classified each item (`reports/phase-5a-divergences.md`):
  - A, structured references: HARNESS-ONLY. The consumer must register table metadata.
  - B, blank lookup keys: ENGINE-GAP.
  - C, negative zero: CONTRACT-CHOICE.
  - D: controls that already agree.
- **5c (`5196852`):** computed `-0` normalises to `0`, by user contract. TEXT section and sign rules: the section is selected by the original sign, and a single-section format drops the automatic minus only when the scaled pattern rounds to zero. `TextFormatCache` has no `Count`/`Math.Pow` on the hit path.
- **5b (item 59; `c31b0cf`..`aa683fb`, 10 commits, review round 8 Clean):** an absent key is distinct from an `=""` key across the lookup family and the criteria family. Key decisions:
  - **Classification:** `ResolvedReferenceValue.TryClassify` is the single classifier. Derived keys (INDEX, OFFSET, IF, CHOOSE, defined names, LET) resolve once and carry the singleton reference and absent bit, so a volatile key draws once.
  - **Exact matching:** in the shared `LookupMatching.ExactMatcher`, an empty-text key matches only exact text (`IsExactText`). An absent cell never matches `""`.
  - **Approximate matching:** an empty-text key needs an exact text candidate, otherwise `#N/A`. MATCH type 1 takes the last candidate and -1 the first; VLOOKUP/HLOOKUP take the last (`LookupGrid.FindLastExactText`); LOOKUP takes the last. XMATCH/XLOOKUP modes ±1 stop after the exact scan.
  - **Criteria family:** shared `Criteria.Parse` with the absent bit, and no per-function arm. The SUMIF-only arm was removed after a discriminating 1/2/4/8 fixture showed it was wrong.
  - **Multi-cell keys:** unchanged from `b8e6cf9`. A bare reference gives `#VALUE!`; a derived reference keeps the reference value. The 126-cell matrix is byte-identical to `b8e6cf9`. Lifting is item 75.
  - **LOOKUP vectors:** `ArgumentFlattening.MaterializeVector` serves both slots (fixing `LOOKUP(3,{1,2,3,4},{10,20,30,40})` 10 → 30, and the `LOOKUP(2,1/(B1:B4=2),C1:C4)` idiom → 20). The `Lookup/LookupVector.cs` companion owns:
    - selected-path structural errors: a missing sheet, and IF/CHOOSE through the memoised condition;
    - single materialisation;
    - slot error shaping: lookup-slot scalar errors propagate, array element errors are skipped, a selector's own error propagates, a selected singleton result error gives `#N/A`, and a binary missing sheet propagates in the lookup slot only against a scalar operand.
  - **Registered, not fixed:**
    - 74: a LET reference as a range endpoint;
    - 75: multi-cell key lifting;
    - 76: MATCH's computed-array double draw;
    - 77: `LOOKUP(1,1/0)`, a PLAIN/CSE split;
    - 78: the CHOOSE result-slot split, a controller ruling for the IF-consistent `#N/A`;
    - 79: short result vectors;
    - 80: a computed INDEX vector.
- **Process:** 8 fix rounds and 8 reviews. The lessons are in `tasks/lessons.md` (2026-09-15 section): a full family matrix up front, discriminating fixtures, stacks for draw counts, independent closure evidence, and consolidating rather than piling special cases into shared helpers.
- **Status for the consumer list:** A needs harness work (table metadata); B is closed by 5b; C is closed by 5c; Bugs 6/8/9 and `IF("1")` already agree.

## Phase 6: Integration, sweep file, lessons, Fable 5.1 final gate, local merge
Status: In progress (local merge and final gate).
- 2026-09-16: `feat/rp-integrate` fast-forwarded to `aa683fb` (45 commits since `856a6c5`). Gates are green: csharpier 420, Release build 0/0, core 3803/0, Excel 138/0, union count 329, attribution 0. Since the last perf gate (`fe2c7bc`, PASS), only LOOKUP code changed, so no new gate is needed.
- 2026-09-15 19:53: `feat/rp-integrate` was fast-forwarded to `2bbda0d` (Phase 5b through fix round 7; 44 commits since `856a6c5`). Gates are green: csharpier 420, Release build 0/0, core 3784/0, Excel 138/0, union count 329, attribution 0, `plans/`/`tasks/` untouched, fences unchanged. Phase 5b review round 7 is pending; later fix commits can fast-forward again.
- Earlier: `feat/rp-integrate` in worktree `MySheet-rpint` was @ `48e01c0`, carrying Phase 1, Phase 3 and Phase 5c, all cherry-picked with clean merge-tree forecasts.
- Gates are green: csharpier, Release build 0/0, core 3472/0, Excel 137/0, union count 328, attribution 0.
- Phase 2 (CLOSED @ `8234e45`) and Phase 4 (CLOSED @ `bc31695`) are now integrated too, so `feat/rp-integrate` is @ `b8e6cf9` with 35 commits. Gates are green: csharpier 417, build 0/0, core 3603/0, Excel 138/0, union count 329, attribution 0.
- Next: 5b (running), integrate 5b, then gates, sweep file, lessons, Fable 5.1 gate, fix wave, local ff merge, notify the user.
- The Fable brief draft is at `.superpowers/sdd/sweep-robustness-parser/final-gate/fable-brief.draft.md`.

- [x] Integrate the phase branches onto a single branch off `main`: forecast conflicts with `git merge-tree` and resolve them. Run semantic checks and the divergence probe.
- [x] Sweep file: items 19, 27, 52, 53, 54 and 59 CLOSED-BY with hashes; the Phase 0 classifications recorded; follow-ups registered as items 58-81.
- [ ] Housekeeping of `.superpowers/sdd/pending-verifications.md`. Lessons appended to `tasks/lessons.md`.
- [ ] Claude Fable 5.1 code-quality review of the sweep range, then the fix wave on opencode, then closure review.
- [ ] Local ff-only merge into `main` with gates on `main`, then notify the user. No push.

### Verification Plan
- Both suites `failed: 0` on `main`; csharpier clean; Release build with 0 warnings; union 329; attribution grep 0.
- The divergence probe shows no row moving away from CSE compared with `856a6c5`.

### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
