# Phase 11: Excel-compatibility sweep of the recorded divergences

Status: Not started   <!-- Not started | In progress | Complete -->

Adversarial verifier verdict: **needs-revision** (2 blockers, 7 majors, 14 minors — folded in below; every one was MEASURED on both sides with probes against the `1b1e2d3` Release build and Aspose.Cells 26.6.0, not argued). The file's own numbers all reproduce; the corrections are shapes it did not measure, plus two rulings (union, `INDEX(…,0,…)`) the addendum requires.

Executes AFTER Phase 10 and BEFORE Phase 3. Every number below was MEASURED on 2026-09-09, twice: MySheet at commit `1b1e2d3` (detached worktree, `dotnet build Danfma.MySheet.slnx -c Release`, console probe against the built `Danfma.MySheet.dll`) and **Aspose.Cells 26.6.0** on the same fixture. No value in this file is inferred; where a value is a MySheet-only observation it says so.

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0 **and its 2026-09-09 addendum** ("when I say Excel, I mean Aspose" — Aspose.Cells 26.6.0 as measured is THE authority; a measurement beats a Microsoft page; a divergence is a work item), the settled scope S1-S8, the repo rules (TDD, `dotnet run --project tests/…` not `dotnet test`, the gates, the union-tag hazard, the `docs/pt-BR/` mirror) and the open decisions. Each item below is one of those decisions, all of them marked **DECIDED → Phase 11**. This phase adds **no** node type, **no** MemoryPack union tag and **no** public API.

## Design decision

Seven recorded divergences, one sweep. Each subsection states the fixture, both measured columns, the code path at `1b1e2d3`, and the rule to adopt. Where MySheet and Aspose already AGREE the item is a docs relabel, not a code change — those are named too, because a "not measured" label is itself a defect under the addendum.

### (a) The nested-AGGREGATE skip — REVERSAL of Phase 2's ruling

Fixture (the one `docs/function-reference.md` already describes): `A1:A3` = 1, 2, 0 (sum 3). Three parallel triples, each summing to 11 with nothing skipped:
`C1` = `=SUBTOTAL(9,A1:A3)` (3), `C2` = `=AGGREGATE(9,4,A1:A3)` (3), `C3` = 5 — **both** kinds nested.
`D1` = `=SUBTOTAL(9,A1:A3)` (3), `D2` = 3 (plain), `D3` = 5 — **nested SUBTOTAL only**.
`F1` = 3 (plain), `F2` = `=AGGREGATE(9,4,A1:A3)` (3), `F3` = 5 — **nested AGGREGATE only**.

Full matrix, options 0-7 × three ranges × function 9 (SUM) and function 3 (COUNTA). Options 1/3/5/7 answered exactly as 0/2/4/6 on **both** sides (the hidden-row bit is a no-op for Aspose too, on a sheet with no hidden rows), so the table collapses to two columns per side; the divergent cells are in **bold**.

| formula | opts 0-3, MySheet | opts 0-3, **Aspose** | opts 4-7, both |
| --- | --: | --: | --: |
| `AGGREGATE(9,o,C1:C3)` (both) | **5** | **8** | 11 |
| `AGGREGATE(9,o,D1:D3)` (SUBTOTAL only) | 8 | 8 | 11 |
| `AGGREGATE(9,o,F1:F3)` (AGGREGATE only) | **8** | **11** | 11 |
| `AGGREGATE(3,o,C1:C3)` | **1** | **2** | 3 |
| `AGGREGATE(3,o,D1:D3)` | 2 | 2 | 3 |
| `AGGREGATE(3,o,F1:F3)` | **2** | **3** | 3 |

`SUBTOTAL` over the same three ranges, measured identical on both sides: `SUBTOTAL(9,C1:C3)` = 8, `SUBTOTAL(3,C1:C3)` = 2, `SUBTOTAL(9,D1:D3)` = 8, `SUBTOTAL(3,D1:D3)` = 2, `SUBTOTAL(9,F1:F3)` = 11, `SUBTOTAL(3,F1:F3)` = 3.

**The rule Aspose implements:** options 0-3 drop a referenced cell whose own formula is a `SUBTOTAL`, **and only a `SUBTOTAL`**; a nested `AGGREGATE` is never dropped, at any option, by either function. Options 4-7 drop nothing. AGGREGATE's nested-skip predicate is therefore **identical to SUBTOTAL's** — Microsoft's options table wording "Ignore nested SUBTOTAL and AGGREGATE functions" is half wrong as measured, and under the addendum the measurement wins. `NestedSkip` collapses from three members to two.

### (b) A scalar in a `ref` slot, and an error scalar in the `array` slot

Fixture `A1:A3` = 5, 0, 9; `B1:B3` = 1, 2, 3; a name `Sete` = `=7`; `Rng` = `Sheet1!$A$1:$A$3`.

| formula | MySheet | Aspose |
| --- | --- | --- |
| `SUBTOTAL(9,7)` | 7 | `#VALUE!` |
| `SUBTOTAL(9,A1:A3,7)` | 21 | `#VALUE!` |
| `SUBTOTAL(2,7)` / `SUBTOTAL(3,7)` | 1 / 1 | `#VALUE!` / `#VALUE!` |
| `SUBTOTAL(9,"7")` / `SUBTOTAL(9,TRUE)` | 0 / 0 | `#VALUE!` / `#VALUE!` |
| `SUBTOTAL(9,A1:A3,"")` | 14 (MySheet-only) | `#VALUE!` |
| `AGGREGATE(9,4,7)` / `AGGREGATE(9,6,7)` | 7 / 7 | `#VALUE!` / `#VALUE!` |
| `AGGREGATE(9,4,A1:A3,7)` / `AGGREGATE(9,0,A1:A3,7)` | 21 / 21 | `#VALUE!` / `#VALUE!` |
| `SUBTOTAL(9,Sete)` / `AGGREGATE(9,4,Sete)` | 7 / 7 (MySheet-only) | `#VALUE!` / `#VALUE!` |
| `AGGREGATE(9,4,A1:A3,B1)` — a single-CELL ref | 15 | 15 |
| `SUBTOTAL(9,A1)` / `SUBTOTAL(9,Rng)` | 5 / 14 | 5 / 14 |
| **array slot** `AGGREGATE(15,6,7,1)` | 7 | 7 |
| **array slot** `AGGREGATE(15,4,7,1)` / `AGGREGATE(14,6,7,1)` / `AGGREGATE(16,6,7,0.5)` | 7 / 7 / 7 | 7 / 7 / 7 |
| **array slot** `AGGREGATE(15,6,Sete,1)` | 7 (MySheet-only) | 7 |
| **array slot** `AGGREGATE(15,6,1/0,1)` | `#DIV/0!` | **`#VALUE!`** |

Rule: a `ref` slot takes a REFERENCE and nothing else — a cell, a range, a union, a name bound to one, or a function that returns one; anything that evaluates to a bare value is `#VALUE!`, whatever its type. The `array` slot is genuinely different: it takes a scalar (7 stays 7 on both sides) but rejects an ERROR scalar with `#VALUE!` — not with the error itself. Phase 2's summary recorded that last cell as `#DIV/0!`; **re-measured, it is `#VALUE!`** — that line of the Phase 2 summary is wrong and this phase supersedes it.

### (c) `MODE.SNGL` tie-break

| population | MySheet | Aspose |
| --- | --: | --: |
| `A1:A4` = 2, 1, 1, 2 | 1 | **2** |
| `A1:A4` = 1, 2, 2, 1 | 2 | **1** |
| `A1:A6` = 3, 1, 2, 1, 2, 3 | 1 | **3** |

`MODE` (the legacy alias) and `AGGREGATE(13,4,…)` answer identically to `MODE.SNGL` on both sides, on all three. Rule: among the values that tie on the winning count, Excel returns the one that appears FIRST in scan order — not the first to REACH the count. Every one of the three fixtures discriminates, which is why all three are pinned.

### (d) A 1×1 reference in the array slot

Fixture `E1:E3` = 5, `=1/0`, 9; `G1` = `"x"`, `G2` = `=1/0`.

| formula | MySheet | Aspose |
| --- | --- | --- |
| `AGGREGATE(15,6,E2,1)` / `AGGREGATE(15,6,E2:E2,1)` | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(14,6,E2,1)` / `AGGREGATE(14,6,E2:E2,1)` | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(16,6,E2:E2,0.5)` / `(17,6,…,2)` / `(18,6,…,0.5)` / `(19,6,…,2)` | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(15,2,E2:E2,1)` (option 2) | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(15,0,E2:E2,1)` (option 0) | `#DIV/0!` | `#DIV/0!` |
| `AGGREGATE(15,6,E2:E2,2)` (k past the population) | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(15,6,E2:E2,0)` (invalid k) | `#NUM!` | **`#DIV/0!`** |
| `AGGREGATE(15,6,G1:G1,1)` (1×1 TEXT) | `#NUM!` | `#NUM!` |
| `AGGREGATE(15,6,E1:E1,1)` / `AGGREGATE(15,6,A1,1)` | 5 / 5 | 5 / 5 |
| `AGGREGATE(15,6,E1:E2,1)` (2 cells) | 5 | 5 |
| `AGGREGATE(15,6,E2:E3,1)` / `AGGREGATE(15,6,E1:E3,2)` | 9 / 9 | 9 / 9 |

Rule: in the ARRAY slot a reference that denotes exactly ONE cell collapses to that cell's scalar value BEFORE the options and BEFORE the `k` bound, so an error there propagates as itself — at every option, for every code 14-19. Widen the range by one cell and the ordinary population rules resume (`E2:E3` → 9 on both). Only an ERROR value collapses visibly; 1×1 text is `#NUM!` on both sides already, and the reference form is unaffected (`AGGREGATE(9,6,E2)` = 0, `AGGREGATE(3,6,E2:E2)` = 0, `AGGREGATE(2,6,E2:E2)` = 0 — all four cells agree today).

### (e) The defined-name mini-CSE gap

Fixture `Rng` = `Sheet1!$A$1:$A$3`, `A1:A3` = 5, 0, 9 — the master plan's "What the verification measured" fixture.

Aspose emulates legacy (non-dynamic-array) Excel, so a mini-CSE formula must be entered as CSE to get array semantics; the non-CSE column is implicit intersection and carries no signal here. **The measurement that matters is that Aspose gives the name and the literal range the SAME answer, which is what MySheet does not.**

| formula | MySheet, name | MySheet, literal range | Aspose CSE (name) | Aspose CSE (range) |
| --- | --: | --: | --: | --: |
| `COUNT((X<>"")*1)` | **1** | 3 | 3 | 3 |
| `SUM((X<>0)*1)` | **1** | 2 | 2 | 2 |
| `COUNT(X*1)` | — | — | 3 | 3 |
| `SMALL(IF(X>0,X),1)` | **0** | — | 5 | — |
| `COUNT(ROW(X))` | 3 | 3 | 3 | 3 |
| `SUM(X)` / `ROWS(X)` | 14 / 3 | 14 / 3 | 14 / 3 | 14 / 3 |

Cause at `1b1e2d3`: `ArrayEvaluation.Probe` (`Danfma.MySheet/Expressions/ArrayEvaluation.cs:199-280`) has arms for `RangeReference`, `AnchoredRangeReference`, `OpenRangeReference`, `Row`/`Column` over `[NameReference or Reference]`, `BinaryOperation` and `If` — and **no arm for a bare `NameReference`**, so a name falls to `default: (true, false)` at `:277-279` (opaque scalar) and `TryBuildOperand`'s twin `default:` at `:358-361` broadcasts one `ScalarOperand`. The `ROW(Rng)` row above is the proof that the resolution machinery is already present and already threaded with a context: `ProbePosition` → `ResolvePositionRange` (`:494-…`) calls `NamedReferences.TryResolveReference(argument, context, out var reference, boundOpenRanges: false)`. **The master plan's "it needs an `EvaluationContext` in the probe" precondition is already satisfied** — Phase 1 shipped `IsArrayEligible(Expression, EvaluationContext)` (`:135`) and both switches take a context. The gap at `1b1e2d3` is one missing arm on each switch, not a signature change.

### (f) An `IF` whose branches are REFERENCES

Fixture `A1:A3` = 5, 0, 9; `B1:B3` = 1, 2, 3.

| formula | MySheet | Aspose |
| --- | --- | --- |
| `SUBTOTAL(9,IF(TRUE,A1:A3,B1:B3))` | `#VALUE!` | **14** |
| `SUM(IF(TRUE,A1:A3,B1:B3))` | `#VALUE!` | **14** |
| `ROWS(…)` / `COLUMNS(…)` | `#VALUE!` / `#VALUE!` | **3** / **1** |
| `SUBTOTAL(9,IF(A1>0,A1:A3,B1:B3))` (computed scalar condition) | `#VALUE!` | **14** |
| `SUBTOTAL(9,IF(1,A1:A3,B1:B3))` (numeric condition) | `#VALUE!` | **14** |
| `SUBTOTAL(9,IF(FALSE,A1:A3,B1:B3))` (the else branch) | `#VALUE!` | **6** |
| `AGGREGATE(9,4,IF(TRUE,A1:A3,B1:B3))` | `#VALUE!` | **14** |
| `COUNT(IF(TRUE,A1:A3,B1:B3))` | **0** — silent | **3** |
| `MATCH(9,IF(TRUE,A1:A3,B1:B3),0)` | `#N/A` | **3** |
| `INDEX(IF(TRUE,A1:A3,B1:B3),3)` | `#REF!` | **9** |
| `AREAS(…)` / `ISREF(…)` | `#VALUE!` / `FALSE` | **1** / **TRUE** |
| `ROW(IF(TRUE,A2:A3,B1:B3))` | `#VALUE!` | **2** |
| `SUMPRODUCT(IF(TRUE,A1:A3,B1:B3))` | `#VALUE!` | **14** |
| `SUBTOTAL(9,CHOOSE(1,A1:A3,B1:B3))` — the working twin | 14 | 14 |
| `ROWS(CHOOSE(…))` / `AREAS(CHOOSE(…))` / `ISREF(CHOOSE(…))` / `MATCH(9,CHOOSE(…),0)` | 3 / 1 / TRUE / 3 | 3 / 1 / TRUE / 3 |
| `SUBTOTAL(9,OFFSET(A1,0,0,3,1))` | 14 | 14 |

`CHOOSE` and `OFFSET` already do the right thing, and `CHOOSE` is the exact shape template: `Choose.Evaluate` (`Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs:22-40`) ends `return NamedReferences.CaptureValue(chosen, context);` and it also implements `TryResolveReference`. `If.Evaluate` (`Danfma.MySheet/Expressions/Logical/If.cs:8-25`, the whole file is 25 lines) instead ends `return Arguments[1].Evaluate(context);`, and a range node's `Evaluate` is `#VALUE!` — the same bug class as issue #8, which `CaptureValue` exists to fix. `If` has no `TryResolveReference` override, which is what makes `ROWS`/`AREAS`/`ISREF`/`INDEX`/`MATCH` fail.

Two things measured so the fix cannot break them. **The cell boundary is already correct and must stay so**: today a cell holding `=CHOOSE(1,A1:A3,B1:B3)` yields 0 at `D2` (implicit intersection onto `A2`) and `#VALUE!` at `D9` (no intersection); Aspose gives `#VALUE!` for a bare `=IF(TRUE,A1:A3,B1:B3)` written at a non-intersecting cell, matching. After the fix `E2` must become 0 and `E9` stay `#VALUE!` — `Workbook` already routes a top-level node through `NamedReferences.CaptureValue` and then `ImplicitIntersection` (`Danfma.MySheet/Workbook.cs:349-362`). **The mini-CSE array-condition path must not move**: `SUM(IF(A1:A3>4,A1:A3,B1:B3))` = 16 in MySheet today via `IfOperand`, and Aspose's non-CSE answer (`#VALUE!`, implicit intersection) gives no contrary signal — leave it alone.

### (g) The remaining "inferred" / "not measured" labels

Measured on both sides; **every one of these AGREES**, so they are relabels, not behaviour changes — except that two facts the docs never stated are now measured and should be written down.

| claim in `docs/function-reference.md` | fixture | MySheet | Aspose | verdict |
| --- | --- | --- | --- | --- |
| COUNTA under option 6 excludes error cells | `E1:E3` = 5/`#DIV/0!`/9 | `(3,4)`=3, `(3,6)`=2 | 3, 2 | confirmed |
| …and blanks are never counted | `E1:E4` | `(3,4)`=3, `(3,6)`=2 | 3, 2 | confirmed |
| COUNT is unaffected by the bit | same | `(2,4)`=2, `(2,6)`=2 | 2, 2 | confirmed |
| an all-error population under option 6 → `#NUM!` | `A1:A3` all `=1/0` | 14/15/16 → `#NUM!` | `#NUM!` | confirmed — but **array form only** |
| (unstated) the REFERENCE form on the same population | same | `(9,6)`=0, `(3,6)`=0, `(2,6)`=0, `(1,6)`=`#DIV/0!` | 0, 0, 0, `#DIV/0!` | new, measured, must be documented |
| SUBTOTAL does not skip a nested AGGREGATE | (a)'s `F1:F3` | 11 | 11 | confirmed |
| `SUBTOTAL(9,…)` over (a)'s `C1:C3` = 8 | (a) | 8 | 8 | confirmed |
| `k` past the surviving population → `#NUM!` | `E1:E3`, k=3 | `#NUM!` | `#NUM!` | confirmed |
| `SUBTOTAL(3,…)` counts error cells | `E1:E3` / all-error | 3 / 3 | 3 / 3 | confirmed |

## Blocking corrections — the design as written was WRONG here. Apply these first.

Every number below was re-measured by the verifier on 2026-09-09: MySheet at `1b1e2d3` (detached worktree, Release build, console probe against `Danfma.MySheet.dll`) and Aspose.Cells 26.6.0 (a copy of the designer's probe, extended). The file's own numbers all REPRODUCE — the full options 0-7 × {C, D, F} × {9, 3} matrix, every (b)/(c)/(d)/(e)/(f)/(g) row, the `CHOOSE` cell-boundary pair (E2 → 0, E9 → `#VALUE!`), and `AGGREGATE(15,6,1/0,1)` = `#VALUE!` on Aspose (Phase 2's summary line "only an error scalar is rejected → `#DIV/0!`" is wrong, as the designer says). The corrections are about shapes the file did not measure.

- [ ] **B1. Item 6's tie-break is underspecified and, on a 2-D range, wrong: MySheet's population scan is COLUMN-major; Aspose's "first encountered" is ROW-major.** `AggregateCodes.Gather`'s rectangle loop (`AggregateCodes.cs:76-78`), `RangeReference.Expand` (`RangeReference.cs:38-40`) and therefore `StatisticsMath.Collect` all walk column-outer/row-inner (documented at `RangeReference.cs:189` and pinned by `RangeSnapshotBuildTests`).
      *Measured evidence:* fixture `C1`=9, `D1`=1, `C2`=2, `D2`=1, `C3`=2, `D3`=5 (row-major scan 9,1,2,1,2,5 → first tied value 1; column-major scan 9,2,2,1,1,5 → 2). Aspose: `MODE.SNGL(C1:D3)` = **1**, `MODE(C1:D3)` = 1, `AGGREGATE(13,4,C1:D3)` = 1. MySheet today: 2 / 2 / 2 — and item 6's two-pass over the same column-major list ALSO answers 2, so the item cannot go green on any 2-D fixture and item 1 has no 2-D pin to catch it. The 1-D fixtures do not discriminate the order.
      *Correction:* state the rule as "the value that appears first in ROW-major scan order". Feed `StatisticsMath.Mode` a row-major population from the two callers only — `StatisticsMath.Collect`'s range arm (used by `ModeSngl.Compute`, `OrderStatistics.cs:58-67`, and by the `MODE` alias) and the code-13 path of `AggregateCodes.Gather` (`:450`) — by walking the rectangle row-outer/column-inner for THIS code (a `RangeBounds` loop, or a row-major twin of `Expand`). Do NOT flip the engine-wide column-major scan: `RangeSnapshot`, `SheetValueStore`'s fast paths and `RangeReference.cs:189` all state it and tests pin it; the master plan's "2-D scan order for the first error" question is unaffected (Aspose did not discriminate it — Phase 2). Add to item 1(c): `MODE.SNGL(C1:D3)` = `MODE(C1:D3)` = `AGGREGATE(13,4,C1:D3)` = 1 [today 2]; `MODE.SNGL(A1:A4)` on 5,3,3,5 = 5 [today 3]; and on 1,1,2,2 = 1 [today 1 — the no-regression pin that rules out the one-pass `>=` shortcut, which answers 2 there]. Aspose also answers 1 for `MODE.SNGL(D1:D3,C1:C3)` and `MODE.SNGL(B1:B4,C1:D3)` (arguments are scanned in order, each row-major) — pin one of them.

- [x] **B2.** **MOVED TO PHASE 11a — DELIVERED 2026-09-10** ([`phase-11a-unblocking-slice.md`](phase-11a-unblocking-slice.md)): see that file for the UNION outcome, which B2 predicted wrong (a union name resolves to the Scalar outcome, so the whole expression is scalar-only and B2's `resolved.Evaluate` arm never runs — `SUM((UnN<>0)*1)` is still 1, not the literal's `#VALUE!`), and for the criteria gate's EXACT predicate (`!ArrayEvaluation.IsBareReferenceNode(argument) && ArrayEvaluation.IsArrayEligible(argument, context)` — `TryStream`'s first two conditions; a looser gate leaves two rows silently failing). **Item 8's MissingSheet outcome does not broadcast a `#REF!` — a ghost-bound name evaluates to a REFERENCE value, and the "silent 1" survives the fix.** The item says `MissingSheet → (true, false)` "so the `#REF!` is what gets broadcast"; the broadcast is `ScalarOperand(expression.Evaluate(context))`, and `NameReference.Evaluate` of a name bound to `Ghost!$A$1:$A$3` is not an error.
      *Measured evidence (MySheet, `GhostName` = `Ghost!$A$1:$A$3`):* `=GhostName<>0` = **TRUE**, `=ISERROR(GhostName)` = FALSE, `SUM((GhostName<>0)*1)` = **1**, `COUNT((GhostName<>"")*1)` = **1** — and both stay 1 under the item as written. The literal twin: `SUM((Ghost!A1:A3<>0)*1)` = `#REF!`, `COUNT((Ghost!A1:A3<>"")*1)` = 0 (per-element `#REF!` through `RangeOperand.At` → `GetCellValueDense`, the path Phase 8's item 2 measured for `SUM(LEN(Ghost!A1:A3))`). Aspose: `=GhostName` = `#REF!`, `SUM((GhostName<>0)*1)` = `#REF!` (CSE and non-CSE), `COUNT(…)` = 0. Same defect on the union outcome: `SUM((UnN<>0)*1)` = **1** today (UnN = `(Sheet1!$A$1:$A$2,Sheet1!$A$3:$A$3)`) where the literal `SUM(((A1:A2,A3:A3)<>0)*1)` = `#VALUE!` — again because the NAME evaluates to a reference value while the literal `UnionReference.Evaluate` is `#VALUE!`. (`MyCell` is fine either way: 0 on both paths and on Aspose.)
      *Correction:* the helper's outcomes become: unresolved → opaque scalar (unchanged); resolved `RangeReference`, **including one whose sheet is missing** → `(true, true)` / `BuildRange(range, context)` — exactly the literal's arm, so the per-element `#REF!` flows and `SUM` → `#REF!`, `COUNT` → 0 like the literal and like Aspose; `OpenRangeReference` → refused (unchanged); a single cell or a union → `(true, false)` / `ScalarOperand(resolved.Evaluate(context))` — the RESOLVED node's value, never the `NameReference`'s. Item 1(e) then pins the three pairs name = literal: `SUM((GhostName<>0)*1)` = `SUM((Ghost!A1:A3<>0)*1)` = `#REF!` [today 1 vs `#REF!`], `COUNT((GhostName<>"")*1)` = 0 [today 1], `SUM((UnN<>0)*1)` = `#VALUE!` [today 1]. Drop the "five outcomes are not negotiable" sentence; the `MissingSheet` degrade was designed for `ROW`/`COLUMN` (a position vector has no cell to read) and is wrong for a value vector.

## Major corrections

- [ ] **M1. Item 2's predicate is root-node-only; Aspose (and Excel) skip a cell whose formula CONTAINS a `SUBTOTAL` call anywhere in its tree — evaluated or not, but never through a defined name.** `IsNested` (`AggregateCodes.cs:258-266`) matches `Subtotal => true` on the cell's ROOT node, so `=SUBTOTAL(9,A1:A3)+0` is a `BinaryOperation` and counts.
      *Measured evidence* (fixture `A1:A3` = 1,2,0; `X1` = the formula below worth 3 unless stated, `X2` = 3, `X3` = 5; Aspose / MySheet for `SUBTOTAL(9,X1:X3)`, and `AGGREGATE(9,0,X1:X3)` answers identically on every row): `=SUBTOTAL(9,A1:A3)+0` **8 / 11**; `=1+SUBTOTAL(9,A1:A3)` (4) **8 / 12**; `=IF(TRUE,SUBTOTAL(9,A1:A3),0)` **8 / 11**; `=SUM(SUBTOTAL(9,A1:A3))` **8 / 11**; `=AGGREGATE(9,4,A1:A3)+SUBTOTAL(9,A1:A3)` (6) **8 / 14**; `=SUBTOTAL(9,A1:A3)*1` **8 / 11**; `=SUBTOTAL(9,A1:A3)&""` (text "3") skipped by `SUBTOTAL(3,…)` → 2. The predicate is SYNTACTIC: `=IF(FALSE,SUBTOTAL(9,A1:A3),3)` (the call never runs) → 8, skipped. It does NOT see through names: `=SubName` with `SubName` = `=SUBTOTAL(9,Sheet1!$A$1:$A$3)` → 11, counted; nor through a plain reference (`=Q1` where `Q1` is a `SUBTOTAL` → counted); a string is not a call (`="x"&"SUBTOTAL"`, `=LEN("SUBTOTAL")` → counted). A nested `AGGREGATE` is never skipped in any of these shapes (`=AGGREGATE(1,4,A1:A3)` → 9 / 8 today; `=MyAgg` → 11). Shared-formula slaves: a `SUBTOTAL` master+slave pair is skipped on both sides (5 / 5); an `AGGREGATE` master+slave pair is 11 on Aspose and **5** here — item 2 fixes that one.
      *Correction:* add item **2b** (after 2, before 3): rewrite `IsNested` as a tree walk — `SharedFormulaSlave { Master }` → recurse; `Subtotal` → true; `NameReference` → false (never resolved); any other node → recurse into its children (a `Function`'s `Arguments`, a `BinaryOperation`'s `Left`/`Right`, a unary's operand, `DynamicRange` endpoints — the same shape `DependencyExtractor.Visit`/`VisitArguments` (`DirtyGraph/DependencyExtractor.cs:75-248`) already walks; reuse or mirror it, do not add a visitor API). Cost: one walk per referenced cell under options 0-3 and for every `SUBTOTAL`, on formula cells only (a literal cell short-circuits at the dense probe, as today). Add to item 1(a) the six arithmetic/`IF`/`SUM` shapes = 8 [today 11/12/11/11/14/11], the unevaluated-branch shape = 8 [today 11], and the two NON-skips as anti-vacuity pins: `=SubName` = 11 and `=LEN("SUBTOTAL")` (8) = 16. Items 11/12: "a referenced cell whose own formula is a `SUBTOTAL`" → "whose own formula contains a `SUBTOTAL` call (anywhere in it, taken or not — but not through a name)".

- [ ] **M2. Item 9 must cover `IFS` and `SWITCH`, which Aspose also resolves to their taken branch's reference.** Both live in `Danfma.MySheet/Expressions/Logical/LogicalFunctions.cs` (`Ifs` `:51-77`, `Switch` `:80-115`) and end in `Arguments[…].Evaluate(context)` exactly like `If`.
      *Measured evidence (Aspose / MySheet):* `SUBTOTAL(9,IFS(TRUE,A1:A3))` **14 / `#VALUE!`**; `SUBTOTAL(9,IFS(FALSE,A1:A3,TRUE,B1:B3))` **6 / `#VALUE!`**; `SUBTOTAL(9,SWITCH(1,1,A1:A3))` **14 / `#VALUE!`**; `SUBTOTAL(9,SWITCH(2,1,A1:A3,2,B1:B3))` **6 / `#VALUE!`**; `SUBTOTAL(9,SWITCH(3,1,A1:A3,B1:B3))` (the default branch) **6 / `#VALUE!`**; `ROWS(IFS(TRUE,A1:A3))` **3 / `#VALUE!`**; `ROWS(SWITCH(1,1,A1:A3))` **3 / `#VALUE!`**; `ISREF(IFS(TRUE,A1:A3))` / `ISREF(SWITCH(1,1,A1:A3))` **TRUE / FALSE**; `SUBTOTAL(9,IFS(FALSE,A1:A3))` `#N/A` / `#N/A` (no-regression). Also measured, free after the override: `COUNTIF(IF(TRUE,A1:A3,B1:B3),">0")` = 2 [today 0], `SUMIF(…,">0")` = 14 [today 0], `VLOOKUP(9,IF(TRUE,A1:B3,B1:B3),2,0)` = 3 [today `#REF!`] — the `CHOOSE` twins already give 2 / 14 / 3, so `CriteriaScan` and `VLOOKUP` resolve through `NamedReferences.TryResolveReference` and need no change; pin them in item 1(f). Do NOT extend to `IFERROR`: Aspose's `SUBTOTAL(9,IFERROR(A1:A3,B1:B3))` = 6 is legacy implicit intersection of the first argument at the formula's row, not a reference-returning `IFERROR`, and Phase 8 lifts `IFERROR` elementwise.
      *Correction:* item 9 edits three nodes, same recipe each: the taken branch returns `NamedReferences.CaptureValue(Arguments[i], context)`, and a `TryResolveReference` override re-runs the selection (conditions in order / `ValueCoercion.AreEqual` against the candidates, default branch included) and delegates to the chosen node; `false` on an error or on no match. Item 10(a) gets the bare `=IFS(TRUE,A1:A3)` and `=SWITCH(1,1,A1:A3)` at `E2`-row → 0 pins too (today `#VALUE!`).

- [ ] **M3. Item 5's non-reference scalar arm must reject every non-NUMBER scalar with `#VALUE!` — text and boolean too, not only an error.**
      *Measured evidence (Aspose / MySheet today):* `AGGREGATE(15,6,"7",1)` **`#VALUE!` / `#NUM!`**; `(15,6,TRUE,1)` **`#VALUE!` / `#NUM!`**; `(15,6,"x",1)` `#VALUE!` / `#NUM!`; `(15,6,E1&"",1)` `#VALUE!` / `#NUM!`; `(15,6,E1>0,1)` `#VALUE!` / `#NUM!`; `(15,6,TEXT(7,"0"),1)` `#VALUE!` / `#NUM!`; `(15,6,Txt,1)` with `Txt` = `="7"` `#VALUE!` / `#NUM!`. A NUMBER is accepted however it was produced: `"7"+0` 7 / 7, `VALUE("7")` 7 / 7, `H1+0` (`H1` = text "7") 7 / 7, `Sete` 7 / 7. A text or boolean CELL is different — `AGGREGATE(15,6,H1,1)`, `(15,6,H1:H1,1)`, `(15,6,H2,1)` (`H2` = TRUE) are `#NUM!` on BOTH sides — which confirms item 5's 1×1 collapse must short-circuit only on an ERROR value and let a non-error 1×1 cell fall into the ordinary population (where text is not counted → `#NUM!`).
      *Correction:* the scalar arm reads: number → feed the accumulator; anything else (text, boolean, error) → `ComputedValue.Error(Error.Value)`. Item 1(b) adds `AGGREGATE(15,6,"7",1)` = `AGGREGATE(15,6,TRUE,1)` = `AGGREGATE(15,6,E1&"",1)` → `ErrorValue.NotValue` [today `#NUM!`] and `AGGREGATE(15,6,"7"+0,1)` = 7 as the no-regression pin; item 11's array-slot sentence says "a non-numeric scalar (text, boolean, error) is `#VALUE!`".

- [ ] **M4. The "an error scalar is `#VALUE!`" half is provenance-dependent on Aspose; item 5 as written flips six shapes MySheet gets RIGHT today, and item 1 must not pin them the wrong way.**
      *Measured evidence (Aspose / MySheet today), `E2` = `=1/0`, `CellTwo` = `=Sheet1!$E$2`, `ErrName` = `=1/0`:* `#VALUE!` on Aspose for `1/0` (`#DIV/0!`), `NA()` (`#N/A`), `SQRT(-1)`, `IF(TRUE,1/0,1)`, `E1/0` (`#DIV/0!`), `1/0+E1`, `E1/E4`, `IF(E1>0,1/0,1)`, `SUM(E2)` (`#DIV/0!`), `N(E2)`, `--E2` (`#DIV/0!`), `Sete/0`, `ROW(A1)/0`; but **`#DIV/0!`** on Aspose for `E2+0` (`#DIV/0!` ✓), `E2*1` (✓), `CellTwo+0` (✓), `ErrName` (✓), `INDEX(E1:E3,2)+0` (✓), `IF(TRUE,E2,1)` (✓), and `E2:E2+0` / `E2:E2*1` (`#NUM!` here; `#DIV/0!` there, CSE and non-CSE) — at every option (`(15,0,E2+0,1)`, `(14,4,…)`, `(16,6,…,0.5)` all `#DIV/0!`). An error cell passed through `+0`/`*1`/a name/`IF`/`INDEX` keeps its identity; the same cell through `--`, `N()` or `SUM()` does not — no value-level rule reproduces that, and the master plan's addendum reserves "cannot match" for structural limits, which this is not; it is an oracle artifact.
      *Ruling:* keep item 5's arm (it is the best simple rule: 13 of the 20 measured shapes, against 6 today), but (i) item 1(b) pins ONLY `AGGREGATE(15,6,1/0,1)` and `AGGREGATE(15,6,NA(),1)` → `ErrorValue.NotValue` (the second proves the arm is not `#DIV/0!`-specific) and pins `AGGREGATE(15,6,E2+0,1)` at its post-item-5 value `#VALUE!` with a comment naming Aspose's `#DIV/0!` as a recorded divergence; (ii) the `E2+0` family goes into this phase's "Open questions" with the table above, for real Excel; (iii) item 11 must not present the sentence as Excel's rule without "for a computed error; an error cell passed through arithmetic answers `#DIV/0!` on Aspose"; (iv) add the 1×1 COMPUTED-array collapse to item 5: when `TryStream` succeeds and the stream has exactly one element that is an error, return that error (`AGGREGATE(15,6,E2:E2*1,1)` = `#DIV/0!` on Aspose, `#NUM!` here; `E1:E1*1` — a 1×1 number — falls through unchanged).

- [ ] **M5. The union in a `ref` slot: EXCEPTION REQUEST to the user — do not match Aspose; the measurement is an artifact of Aspose modelling a union as a flattened ARRAY, not a reference, and matching it in `SUBTOTAL` alone is incoherent while matching the model wholesale flips `ISREF`, `ROWS`, `SUMPRODUCT`, `COUNTIF` and the mini-CSE.** The standing rule says match; this is the explicit request not to, with the evidence.
      *Measured evidence (Aspose):* `ISREF((A1:A2,A3:A3))` = **FALSE** while `ISREF((A1:A3))` = TRUE; `ROWS((A1:A2,A3:A3))` = `#VALUE!` yet `ROWS(Un)` (the same union through a name) = 2 and `AREAS(Un)` = 2; `SUMPRODUCT((A1:A2,A3:A3))` = `#VALUE!`; `COUNTIF((A1:A2,A3:A3),">0")` = `#REF!`; `OFFSET((A1:A2,A3:A3),0,0)` = `#REF!`; `INDEX(Un,1,1,2)` = `#VALUE!` yet the literal `INDEX((A1:A2,A3:A3),1,1,2)` = 9; CSE `SUM(((A1:A2,A3:A3)<>0)*1)` = 2 and `SUM((Un<>0)*1)` = 2 (a union takes part in array arithmetic); `SUBTOTAL(9,IF(TRUE,(A1:A2,A3:A3),B1:B3))` = `#VALUE!`, `SUBTOTAL(9,CHOOSE(1,(A1:A2,A3:A3),B1))` = `#VALUE!`. Every ref-slot form is `#VALUE!` — `(A1:A2,A3:A3)`, `(A1:A2,A3)`, `(A1,A2,A3)`, `(A1:A1,A2:A2,A3:A3)`, `(A1:A2,B1:B2)`, `(A1:A3,B1:B3)`, `(A1:A3,A1:A3)`, codes 1/3/4/9/109, as `ref2` after `B1`, `AGGREGATE(9,4,…)`/`(9,0,…)`/`(3,4,…)`, and via the name `Un` — while the two-argument `SUBTOTAL(9,A1:A2,A3:A3)` = 14, `AGGREGATE(9,4,A1:A2,A3:A3)` = 14, the intersection `SUBTOTAL(9,(A1:A2 A1:A3))` = 5, the parenthesised range `SUBTOTAL(9,(A1:A3))` = 14, and every value-consuming function accepts the union (`SUM` 14, `COUNT` 3, `COUNTA` 3, `MAX` 9, `MIN` 0, `AVERAGE`, `SMALL` 0, `LARGE` 9, `MEDIAN` 5, `PRODUCT` 0, `RANK` 1, `AREAS` 2); the array slot accepts it (`AGGREGATE(15,6,(A1:A2,A3:A3),1)` = 0, `(14,4,…,1)` = 9). MySheet: a union is a reference everywhere — `SUBTOTAL` 14, `ISREF` TRUE, `SUMPRODUCT` 14, `COUNTIF` 2, `AREAS` 2, `ROWS` 1 (first area), mini-CSE `#VALUE!`.
      *Ruling:* record as a deliberate, user-approved divergence (pending the user's answer): keep the shipped union support, pin `SUBTOTAL(9,(A1:A2,A3:A3))` = 14, `SUBTOTAL(9,A1:A2,A3:A3)` = 14, `AGGREGATE(9,4,(A1:A2,A3:A3))` = 14 and `ISREF((A1:A2,A3:A3))` = TRUE in `ExcelCompatibilitySweepTests` under a comment that names the Aspose answers and this exception, and keep the master plan's open decision for the real-Excel `.xlsx`. If the user declines the exception, the item is: in `AggregateCodes.Gather` the `UnionReference` arm returns `Error.Value` for the REFERENCE form only (the array form's `Gather` call at `Aggregate.cs:159` must keep accepting a union — measured 0 / 9 — so the arm needs a caller flag or the array form must expand the union itself); flip `Subtotal_IgnoresNestedSubtotals_ThroughAUnionAndAnOpenRange`; and the docs rows for `SUBTOTAL`/`AGGREGATE`/`ISREF` in both languages say a union is rejected in `ref` slots while `SUBTOTAL(9,A1:A2,A3:A3)` is the supported spelling.

- [ ] **M6. `INDEX(…,0,…)` is a work item under the addendum, not a risk note — add item 13.** Nothing structural blocks it: `Index.Evaluate` (`Lookup/Index.cs:29-…`) and `Index.TryResolveReference` (`:160-226`) both reject `row < 1 || column < 1` and produce `#REF!`/`false`.
      *Measured evidence (Aspose / MySheet), `A1:A3` = 5,0,9, `B1:B3` = 1,2,3:* `SUBTOTAL(9,INDEX(A1:B3,0,1))` **14 / `#REF!`**; `SUM(INDEX(A1:B3,0,1))` 14 / `#REF!`; `ROWS(INDEX(A1:B3,0,1))` 3 / `#REF!`; `ISREF(INDEX(A1:B3,0,1))` TRUE / FALSE; `AREAS(…)` 1 / `#REF!`; `SUM(INDEX(A1:B3,2,0))` 2 / `#REF!`; `COLUMNS(INDEX(A1:B3,2,0))` 2 / `#REF!`; `SUM(INDEX(A1:B3,0,0))` 20 / `#REF!`; `SUBTOTAL(9,INDEX(A1:B3,0,0))` 20 / `#REF!`; `SUM(INDEX(A1:B3,,2))` (omitted = 0) 6 / `#REF!`; `SUM(INDEX(A1:A3,0))` 14 / `#REF!`; `SUBTOTAL(9,INDEX(A1:A3,0))` 14 / `#REF!`; bare `=INDEX(A1:B3,0,1)` at a non-intersecting row `#VALUE!` / `#REF!` and `INDEX(A1:B3,0,1)+0` `#VALUE!` / `#REF!` (the reference reaches the cell boundary and implicit intersection applies).
      *Item 13 text (after item 10, since it produces a `ComputedValueKind.Reference` through the same `CaptureValue` boundary item 10 pins):* in `Index.TryResolveReference`, after the range resolves and the numbers coerce, a `row` of 0 with a `column` ≥ 1 yields the `RangeReference` of that whole column of the range; a `column` of 0 with a `row` ≥ 1 the whole row; both 0, or the 2-argument form with 0, the whole range — built from `range.LeftColumn`/`TopRow` and the bounds, `SheetName` preserved. In `Index.Evaluate`, when the concrete-range path sees a 0 in either slot, return `ComputedValue.Reference(...)` of that same rectangle (call the override, do not duplicate the arithmetic); the array-form paths (`IndexIntoArray`, the open-column `ROW` identity) keep their current behaviour — not measured, out of scope. Pins: the fourteen rows above, the bare `=INDEX(A1:B3,0,1)` at `E2` → 0 (row intersection) in `CellBoundaryIntersectionTests`. Docs: the `INDEX` row (EN `:242`, pt-BR `:248`) gains the 0-selects-the-whole-row/column sentence with `SUM(INDEX(A1:B3,0,1))` = 14.

- [ ] **M7. Item 10(b)/(f)'s "an array condition must NOT become a reference" is Excel's rule and MySheet's, but Aspose contradicts it in a `ref` slot — record the divergence and pin the `#VALUE!` explicitly.**
      *Measured evidence (Aspose):* `SUBTOTAL(9,IF({TRUE;FALSE;TRUE},A1:A3,B1:B3))` = **14**, `SUBTOTAL(9,IF({FALSE;TRUE;TRUE},…))` = **6**, `ROWS(IF({TRUE;FALSE;TRUE},…))` = 3, `ISREF(IF({FALSE;TRUE},…))` = TRUE, CSE `SUBTOTAL(9,IF(A1:A3>4,A1:A3,B1:B3))` = 14 — the condition's FIRST element picks a branch and the branch's reference is returned; yet `SUM(IF({TRUE;FALSE;TRUE},A1:A3,B1:B3))` = 16 and CSE `SUM(IF(A1:A3>4,…))` = 16 are elementwise. MySheet: `SUM(IF(A1:A3>4,A1:A3,B1:B3))` = 16, `SUBTOTAL(9,IF(A1:A3>4,A1:A3,B1:B3))` = `#VALUE!` (Feed's `TryStream` gate, before `If.Evaluate` is ever reached — unchanged by item 9); array constants do not parse here (`{…}` is a `ParseException`), so the `{…}` rows cannot be pinned at all.
      *Ruling:* keep `#VALUE!` — Aspose's own `SUM` answer shows the elementwise semantics are the real ones and the 14 is the first-element artifact. Item 10(b) pins `SUBTOTAL(9,IF(A1:A3>4,A1:A3,B1:B3))` = `ErrorValue.NotValue` next to the 16.0, with the Aspose 14 named in the comment; add the row to "Open questions" for real Excel.

## Minor corrections (fold in while implementing)

- Item 4's "an empty argument" claim is confirmed and needs pins: `SUBTOTAL(9,A1:A3,)` = 14 today / `#VALUE!` Aspose, `AGGREGATE(9,4,A1:A3,)` 14 / `#VALUE!`, `SUBTOTAL(9,)` 0 / `#VALUE!` (a `BlankValue` node reaches the same `default:` arm); also `SUBTOTAL(9,A1+0)` = 5 today / `#VALUE!` (a computed scalar from a cell). And the no-regression pin that proves the new line sits BELOW the error propagation at `:214`: `SUBTOTAL(9,1/0)` = `#DIV/0!` on both sides (confirmed in code and by measurement).
- Item 3: `Gather_SkipsANestedSubtotalStoredAsASharedFormulaSlave` (`AggregateCodesTests.cs:357`) is NOT "unchanged": its third assertion (`:373-375`) names `NestedSkip.SubtotalAndAggregate` and will not compile after item 2 — delete that assertion, keep the other two.
- Item 1(a) anti-vacuity pins for the shapes item 2 fixes beyond the C/D/F triples, all measured: nested array-form `AGGREGATE(14,4,A1:A3,1)` in `H1` (=2) → `AGGREGATE(9,0,H1:H3)` = 10 [today 8]; all-three-nested `S1:S3` (`SUBTOTAL`, `AGGREGATE(9,4,…)`, `AGGREGATE(14,4,…,1)`) → `AGGREGATE(9,0,S1:S3)` = 5 [today 0], `SUBTOTAL(9,S1:S3)` = 5 [passes]; `AGGREGATE(15,0,C1:C3,1)` = 3 [today 5], `AGGREGATE(5,0,C1:C3)` = 3 [today 5], `AGGREGATE(1,0,C1:C3)` = 4 [today 5]; the nested `AGGREGATE(1,4,…)` (avg = 1) in `R1` → `AGGREGATE(9,0,R1:R3)` = 9 [today 8]; and item 3's rewritten slave test covers `AGGREGATE(9,0,U1:U3)` = 11 [today 5] (an `AGGREGATE` master with a `SharedFormulaSlave(master, 1, 0)` in `U2`, `U3` = 5).
- (e)'s table has "—" for MySheet on `COUNT(X*1)`; measured: name **0**, literal 3 (Aspose CSE 3 / 3) — the silent-zero class; add it to item 1(e).
- Item 1(e)/Verification Plan: the open-range name is not "landing on the intended outcome" — `SUM((MyCol<>0)*1)` = 1 today and stays 1 after item 8 (refused → the consumer's scalar path → `NameReference.Evaluate` is a reference value → TRUE), while the literal `SUM((A:A<>0)*1)` = `#VALUE!` and Aspose gives 0 non-CSE / 2 CSE. Pin the 1 as the cost-guard divergence, named as such, not as correct.
- The Verification Plan's pt-BR grep cannot find two of the four sentences: they are written `**não** é um valor medido` and `**não** é medido` (bold markers inside), so `não é um valor medido\|não é medido` never matches — measured: today's grep returns only the `palpite` and `prevalece` lines. Use `\*\*não\*\* é \(um valor \)\?medido\|não um palpite\|página documentada prevalece`.
- Docs: the `MODE.SNGL` rows (EN `:155`, pt-BR `:160`) ALREADY say "a tie resolves to the first value encountered" — the docs were right and the code wrong; items 11/12 add the 2,1,1,2 → 2 example and the row-major clause (B1), they do not "state" a missing rule. `Fold_Code13_IsTheModeInScanOrder` (`AggregateCodesTests.cs:54`, [3,3,2,2] → 3) does not flip but its comment ("PRIMEIRO valor que atinge") must be reworded; `CompatibilityAliasTests.Mode_MatchesTheModeSnglGoldenExample` (`:69`) has no tie.
- Array constants `{…}` do not parse in MySheet (`SUBTOTAL(9,{1,2,3})`, `MODE.SNGL({2;1;1;2})` → `ParseException`), so no `{…}` shape can be pinned; the (b)/(f) prose and the docs' "even the constant `SUBTOTAL(9,{1,2,3})`" describe an Aspose-only measurement. Item 1 must not cite one.
- `Sheet1!Rng` (a sheet-qualified name) is a `ParseException` here ("Expected a cell reference after '!'") and 3 on Aspose (CSE); out of this phase — note it in "Open questions".
- The `LET` open question is wrong as reasoned: `SUM(LET(r,A1:A3,(r<>0)*1))` = 1 today and stays 1 after item 8 — the consumer's argument is the `Let` node, which `Probe`'s `default:` treats as an opaque scalar, so no `NameReference` arm ever sees `r` (`SUM(LET(r,A1:A3,r*1))` = `#VALUE!` today). Aspose: 2, CSE and non-CSE. Under the addendum this is a P0 item of its own (a `Let` arm in `Probe`/`TryBuildOperand` that binds the scope and probes the body) — record it in the master plan's open decisions as DECIDED-shaped, not "pin whatever it does".
- (g) can also cite, all measured identical on both sides: the all-error population under option 6 for codes 4/5/6 → 0, 7/8/10/11 → `#DIV/0!`, 12 → `#NUM!`, 13 → `#N/A`, and `SUBTOTAL(9,…)` → `#DIV/0!`, `SUBTOTAL(3,…)` → 3.
- Aspose's `SUBTOTAL(9,E2:E2*1)` = `#DIV/0!` (a 1×1 computed array in a `ref` slot collapses to its element, and the error propagates) where MySheet's `Feed` gate gives `#VALUE!`; `SUBTOTAL(9,E1:E3*1)` is `#VALUE!` on both. Niche; record in "Open questions", do not act.
- The item-format script's contract holds: twelve `- [ ] **N.**` items, 24 six-space `*Files:*`/`*Why:*` lines, no stray indentation, no GUID anywhere in the file (none invented). Items 2b and 13 must follow the same format.
- Verification Plan baseline re-measured at `1b1e2d3` by the verifier: core **total 1393, failed 16, succeeded 1377**; Excel **90 / 90** — the file's figures hold. The "before items 2-9 land this command must FAIL" list gains B1's 2-D `MODE` (2), M1's `SUBTOTAL(9,X1:X3)` (11), M2's `SUBTOTAL(9,IFS(TRUE,A1:A3))` (`#VALUE!`), M3's `AGGREGATE(15,6,"7",1)` (`#NUM!`) and B2's `SUM((GhostName<>0)*1)` (1).

### What the verifier actively confirmed correct (do not re-verify)

The (a) matrix, all 48 cells, on both sides; every (b) row including `Sete`/`Rng`/`""`/`B1`/`A1` and the four array-slot no-regression cells; (c)'s three fixtures on `MODE.SNGL`/`MODE`/`AGGREGATE(13,4,…)`; every (d) row including the option-0 `#DIV/0!`, `k` = 2 / 0, the 1×1 text `#NUM!`, `E1:E2` → 5, `E2:E3` → 9, and the reference-form `0`s; (e)'s 1/3, 1/2, 0/5 pairs and `COUNT(ROW(Rng))` = 3; all fifteen (f) rows plus `ROW(IF(TRUE,A1:A3,B1:B3))` = 1, `SUBTOTAL(9,IF(1/0,…))` = `#DIV/0!` and `ISREF(IF(1/0,…))` = FALSE on both sides; the `CHOOSE` cell-boundary pair (E2 → 0, E9 → `#VALUE!`; Aspose's bare `=IF(TRUE,A1:A3,B1:B3)` at a non-intersecting row → `#VALUE!`, at `D1` → 5); every (g) row; `AGGREGATE(15,6,1/0,1)` = `#VALUE!` on Aspose (Phase 2's summary sentence is wrong); Gather's error propagation at `:214` sits above the `default:` return (so `SUBTOTAL(9,1/0)` stays `#DIV/0!` after item 4); `ProbePosition`/`ResolvePositionRange` already resolve a `NameReference` with a context (the master plan's "needs an `EvaluationContext`" precondition is indeed already met); `Choose.TryResolveReference` is the exact template for item 9; every cited test line range; the designer's EN grep finds exactly the three sentences; `SubtotalAndAggregate` appears only in the two source files, the two test files and Phase 2's record.

## USER RULING (2026-09-09) — two exceptions and one scope decision

- **Unions in a `ref` slot: exception GRANTED, do not match Aspose.** MySheet keeps `SUBTOTAL(9,(A1:A2,A3:A3))` = 14 and `ISREF((A1:A2,A3:A3))` = TRUE. Reason on the record: Aspose models a union as a flattened array and contradicts itself (`ISREF` FALSE and `ROWS(...)` `#VALUE!`, yet `ROWS(UnionName)` = 2; `SUBTOTAL(9,A1:A2,A3:A3)` = 14 as two arguments while the same cells as one union are `#VALUE!`), and matching would delete shipped, tested behaviour that the Microsoft page allows. Pin MySheet's answers as a DELIBERATE divergence, with the contradictory measurements quoted in the test comment and a line in the docs' known-divergence list. The verifier's fallback item text is not used.
- **`IF` with an ARRAY condition over reference branches (M7): keep `#VALUE!`**, same treatment — Aspose returns a reference for `SUBTOTAL(9,IF({T;F;T},A1:A3,B1:B3))` = 14 while `SUM(...)` = 16 on the same node, which is not one rule. Pin and record. (M2's scalar-condition `IFS`/`SWITCH`/`CHOOSE` work is unaffected and still lands.)
- **`MODE` scan order (B1): scope is the two MODE feeders ONLY.** Switch to a row-major walk inside `MODE`/`MODE.SNGL`'s feeders; do NOT change `Gather` or `RangeReference.Expand` engine-wide. Any other order-sensitive function (first-error reporting, `MATCH`, `LARGE` ties) stays as it is and is out of this phase's scope.

## Controller additions after verification (2026-09-09) — measured by Phase 8 Task 4

- [ ] **14.** `FIXED` and `DOLLAR` default decimals. Measured on Aspose.Cells 26.6.0 (2026-09-09) by Phase 8's Task 4: `FIXED(A1,,TRUE)` = `1` and `DOLLAR(A1,)` = `$1` (0 decimals) where MySheet answers `1.00` and `$1.00` (the Microsoft page's documented default of 2). Under the P0 addendum a measurement beats a page, so measure the full matrix first (omitted vs explicit `0`, negative decimals, `FIXED(A1)` with both trailing arguments omitted, `DOLLAR(A1)`, and the same through a lifted array) and match Aspose; if the matrix shows Aspose reading an omitted slot as 0 rather than "default 2", say so in the item and fix the omitted-slot decoding rather than the default.
      *Files:* `Danfma.MySheet/Expressions/Text/Fixed.cs`, `Dollar.cs` (verify the real paths), `tests/Danfma.MySheet.Tests/Parsing/TextFunctionTests.cs` (or the nearest sibling), `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* Phase 8's B1 guard proves the omitted slot reaches these functions as `BlankValue`; the divergence is in what they do with it. It is user-visible in every formatted-number workbook, not a 1900-only edge.

- [ ] **15.** `INDEX(<computed array>, 0)` on Aspose intersects the WHOLE array and returns its FIRST element, so the answer depends on the array: `INDEX(ROW(B2:B5),0)` = **2** (B2's row number) and `INDEX(ROW($A:$A),0)` = **1** — measured plain and CSE-entered, 2026-09-09. (An earlier controller note said "= 1" for both; that was wrong.) MySheet answers `#REF!` for every form — pinned as a deliberate divergence in `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs` (`Index_IntoRowVector_OutOfRange_IsRefError`, `Index_IntoWholeColumnRowNumbers_IsIdentity`), whose comments were corrected in bb1a0fc to state the measured values instead of claiming Excel parity. Fold this into item 13's whole-row/whole-column work: the same `row < 1 || column < 1` rejection owns both.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Index.cs`, `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`
      *Why:* Task 4 measured it while pinning INDEX over a lifted array; leaving it out would make item 13 fix the reference form and leave the computed-array form diverging.

## Controller additions after verification, part 2 (2026-09-09) — measured by Phase 9's Task 3

- [ ] **16.** `BondMath.Diff360Us` keeps a start-of-February rule Aspose lacks — the same shape as the one Phase 9
      removed from `DayCount`'s basis 0, but in the bond path, which Phase 9 deliberately left alone because no
      1900 row depends on it. Measured on Aspose.Cells 26.6.0 (PLAIN, 2026-09-09):
      `ACCRINT(DATE(2023,2,28),DATE(2023,8,31),DATE(2023,3,31),0.1,1000,2,0)` = **9.1667** (33/360) against
      **8.611** (31/360) here. Measure the whole 30/360 matrix for the bond and coupon family first (February
      starts and ends, day-31 starts and ends, both orders, every basis), then match. NOTE: after Phase 9
      `DAYS360` and `YEARFRAC` basis 0 deliberately DISAGREE on February-end / day-31 pairs because Aspose does —
      do not "unify" them.
      *Files:* `Danfma.MySheet/Expressions/Financial/BondMath.cs`, the bond and coupon test files, `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* A generic divergence on modern dates in the accrual functions, not a 1900 edge, so every real bond
      workbook is affected.

- [ ] **17.** `YEARFRAC` basis 1 (actual/actual) is wrong on ORDINARY MODERN spans, not just at the day zero.
      **Re-scoped 2026-09-09 after Phase 9's final reviewer measured it — my earlier text called it a day-zero
      artefact, which hid the real defect.** Measured on Aspose.Cells 26.6.0 (PLAIN, 2026-09-09), Aspose first:
      `YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)` = **1** against 1.0013680 here;
      `(DATE(2023,12,31),DATE(2024,1,1),1)` = 1/365 against 1/365.5;
      `(DATE(2023,3,1),DATE(2024,2,28),1)` = 0.99726 against 0.99590;
      `(DATE(2023,6,1),DATE(2024,6,1),1)` = 1 against 1.00137 — **7 of 12 measured modern rows differ**, i.e. any
      span of a year or less that crosses a year boundary. Cause: `DayCount` averages every calendar year the span
      touches regardless of the span's length. PRE-EXISTING (the pre-phase build gives the same numbers), so it is
      not a Phase 9 regression, but basis 1 is the actual/actual convention most financial models use, which makes
      this the highest-impact item in the sweep. Derive Excel's real rule from the oracle across the cases that
      distinguish it (span within one calendar year; span exactly one year crossing a boundary; span over a leap
      day; span longer than a year) before writing code, and keep the day-zero rows (`YEARFRAC(0,366,1)` = 1 on
      Aspose against 366/365 here) as the narrow tail of the same fix rather than its subject.
      *Files:* `Danfma.MySheet/Expressions/DayCount.cs`, `tests/Danfma.MySheet.Tests/Parsing/DateEpochTests.cs` or the nearest sibling, both docs twins
      *Why:* Every interest and accrual figure computed on basis 1 over a year boundary is currently off, and the
      error is small enough to pass unnoticed in a spreadsheet and large enough to matter in money.

- [ ] **18.** A malformed or wrong-length `WORKDAY.INTL` / `NETWORKDAYS.INTL` weekend mask is **`#NUM!`** on the
      oracle, not `#VALUE!`. Measured on Aspose.Cells 26.6.0 (PLAIN, 2026-09-09) by Phase 9 Task 4:
      `WORKDAY.INTL(45366,5,"00X0011")`, `(...,"000011")`, `NETWORKDAYS.INTL(45362,45376,"00X0011")`,
      `(...,"000011")` and `(...,"0000011X")` all answer `#NUM!`; MySheet answers `#VALUE!`. **This one contradicts
      a committed test that took its value from the Microsoft page** —
      `tests/Danfma.MySheet.Tests/Parsing/DateWorkdayTests.NetworkDaysIntl_InvalidWeekendArgs` — so under the P0
      addendum (a measurement beats a page) the test's expectation changes with the code, and its comment must
      record that the page said `#VALUE!` and the oracle says `#NUM!`.
      **Also measured by Phase 9's fix wave, same family:** with `days` = 0 Aspose answers the start serial even
      when the weekend argument is INVALID (`WORKDAY.INTL(45366,0,0)` = `WORKDAY.INTL(45366,0,"00X0011")` = 45366),
      while MySheet errors, because weekend parsing runs before `Advance`'s zero-days shortcut. Phase 9 put the
      zero-days shortcut ahead of the all-weekend guard only; this needs it ahead of weekend PARSING too. Pin both
      the valid-mask and invalid-weekend forms of `days` = 0.
      *Files:* `Danfma.MySheet/Expressions/Dates/WorkdayFunctions.cs`, `tests/Danfma.MySheet.Tests/Parsing/DateWorkdayTests.cs`, `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* Cheap, and it is the clearest live example of the addendum's rule: a page-sourced golden that the
      oracle contradicts.

- [ ] **19.** `WORKDAY` and `WORKDAY.INTL` CRASH instead of answering, for a whole interval of inputs rather than one
      literal value. Measured by Phase 9's fix wave: `WORKDAY(45362,-3000000000)` throws
      `OverflowException: Negating the minimum value of a twos complement number is invalid.` at
      `Danfma.MySheet/Expressions/Dates/WorkdayFunctions.cs:393` (`var remaining = Math.Abs(days);`), reached from
      `Workday.Evaluate` (:353). The cause is line 392's `(int)Math.Truncate(daysArg)`, which SATURATES every
      `days` at or below -2147483648 to `int.MinValue` — so the crash covers the entire open interval below that,
      not just the literal, and `WORKDAY.INTL` reaches the same line. An exception escaping evaluation into the
      host is worse than any wrong answer. The oracle is no guide: Aspose answers 45362 for this input and itself
      throws on other extremes. Set MySheet's own contract — no evaluation throws, an out-of-range result is
      `#NUM!` — and pin the interval ends plus a saturating value, for both functions.
      *Files:* `Danfma.MySheet/Expressions/Dates/WorkdayFunctions.cs`, `tests/Danfma.MySheet.Tests/Parsing/DateWorkdayTests.cs`
      *Why:* This is a robustness bug, not a compatibility one, and it is the only item in the sweep where
      matching the oracle is explicitly NOT the goal — Aspose crashes too.

## Controller additions after verification, part 4 (2026-09-09) — measured by Phase 9's Task 5

- [ ] **20.** `TEXT`'s `A/P` format prints two letters where the oracle prints one: measured on Aspose.Cells
      26.6.0 (PLAIN, 2026-09-09) `TEXT(0.5,"h:mm A/P")` = `12:00 P`, against `12:00 PM` here. Pre-existing and
      unrelated to the epoch (the old format translation did the same), so Phase 9 recorded it in a code comment
      and deliberately left it unpinned. Measure the whole `AM/PM` family (`AM/PM`, `A/P`, `a/p`, lowercase and
      mixed) and match.
      *Files:* `Danfma.MySheet/Expressions/Text/Text.cs`, `tests/Danfma.MySheet.Tests/Parsing/TextFormatTests.cs`, both docs twins
      *Why:* A one-character difference in every formatted time string that uses the short form.

- [ ] **21.** Lone `d` / `m` / `y` format tokens inside the 1900 window. Phase 9's token-substitution rendering
      fixed every serial at or above 61 (`TEXT(45366,"d")` = 15, `"m"` = 3, `"y"` = 24, `"s"` = 0, `"h"` = 12 —
      all now matching the oracle where several previously printed a whole date or `#VALUE!`), but below 61 the two disagree. **The rule is derivable** — Phase 9's
      reviewer found it, correcting the earlier "no rule" reading: Aspose reads a run of ONE letter off the raw
      DateTime map and a run of TWO or more off the day-zero/phantom-aware map. Measured (PLAIN, 2026-09-09):
      `TEXT(60,"d")` = 28 but `TEXT(60,"dd")` = 29, and `TEXT(60,"xdd")` = 29; `TEXT(0,"d")` = 31,
      `TEXT(0,"m")` = 12, `TEXT(0,"y")` = "99". MySheet uses the phantom-aware map at every run length and
      answers 29 / 29 / 0 / 1 / "00". All of it is now pinned in `Text_LoneFieldTokens`, both the parity rows and
      the divergence. Since the rule IS consistent, matching it is a small, well-defined change — unlike the
      working-day family, where the user's exception rests on Aspose being self-contradictory. Decide on that
      basis.
      *Files:* `Danfma.MySheet/Expressions/Text/Text.cs`, `tests/Danfma.MySheet.Tests/Parsing/TextFormatTests.cs`
      *Why:* Narrow (four serials) but it is the last unexplained 1900 divergence after Phase 9, and it is already
      pinned, so the decision is cheap to act on.

## Controller additions after verification, part 5 (2026-09-10) — measured by Phase 10

- [x] **22.** **MOVED TO PHASE 11a — DELIVERED 2026-09-10** ([`phase-11a-unblocking-slice.md`](phase-11a-unblocking-slice.md)): see that file for the UNION outcome, which B2 predicted wrong (a union name resolves to the Scalar outcome, so the whole expression is scalar-only and B2's `resolved.Evaluate` arm never runs — `SUM((UnN<>0)*1)` is still 1, not the literal's `#VALUE!`), and for the criteria gate's EXACT predicate (`!ArrayEvaluation.IsBareReferenceNode(argument) && ArrayEvaluation.IsArrayEligible(argument, context)` — `TryStream`'s first two conditions; a looser gate leaves two rows silently failing). The criteria family's answer for a computed argument is **`#REF!` CSE-entered and `#VALUE!` on plain
      entry**, and the docs state only the second as "Excel's answer", which is half true. Measured by Phase 10's
      Task 4 on Aspose.Cells 26.6.0 (2026-09-10) for all six shapes (`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`,
      `AVERAGEIF`, `AVERAGEIFS`); MySheet answers 0, `#VALUE!` and 0 depending on the form, two of which are
      SILENT. Decide the rule from the oracle and match it. (The DOCS half is already done: Phase 10's a9eced2 names both
      entry modes and both answers, so what remains here is the ENGINE — matching the oracle, and removing the
      two silent answers.)
      *Files:* `Danfma.MySheet/Expressions/CriteriaScan.cs`, the criteria test files, both docs twins
      *Why:* Two of the three MySheet answers are silent wrong numbers, which is the worst class, and the docs
      currently give a reader a value that is right only for one way of typing the formula.

- [ ] **23.** Divergences Phase 10 measured and handed on. **Partly documented since:** Phase 10's a9eced2 wrote
      up the `ROWS`/`COLUMNS`-over-a-computed-array row in the user docs as a measured deviation, so for that one
      the work left is the ENGINE and a pin; the rest below are still unwritten anywhere but here.
      Measured on Aspose.Cells 26.6.0 (2026-09-10, CSE unless noted): an OPEN range in a broadcast —
      `SUM(A:A*E1:E3)` = `#N/A` and `ROWS(A:A*1)` = 1048576, where MySheet treats an open range as an opaque
      scalar; a UNION in a broadcast — 792 and 36; `COUNTA(A1:C3*H1:H2)` = 9 against 1 here, before AND after
      Phase 10, so `COUNTA` is not a broadcast consumer at all; `SUM(IF(TRUE,A1:A3,B1:B3))` = 12, which is the
      reference-branch item 9 already owns; and `ROWS`/`COLUMNS` of a computed array = 3/3 on the oracle against
      `#VALUE!` here, which is what forced Phase 10 to pin a computed array's extent through `INDEX` bounds
      instead of asking for it directly.
      **Two test GAPS Phase 10's Task 5 found while documenting this, both cheap and both belonging here:**
      nothing fails if `ROWS`/`COLUMNS` of a computed array stops answering `#VALUE!` — the divergence is
      documented with no test pinning it, the only such bullet in that list; and `ROWS(ROW(A1:C3))` is 1 here
      against the oracle's 3, a smaller sibling of the same defect. Also unpinned, pre-existing from Phase 8's
      text: `SUM(ROUND(A1:A3,B1:B3))` = 6, verified on this tree but asserted nowhere.
      *Files:* to be decided per row when the item is briefed
      *Why:* Each is a real divergence that Phase 10 met and could not fix inside its scope; leaving them only in
      a task brief means they vanish when the brief does.

## Controller addition after the Phase 10 final review (2026-09-10) — measured by the GLM-5.3 reviewer

- [ ] **24.** `AREAS` over a computed array answers **0** on the oracle (CSE, Aspose.Cells 26.6.0, 2026-09-10)
      against `#VALUE!` here, measured for both `AREAS(ROW(A1:C3))` and `AREAS(A1:C3*E1:E3)`. Worse than a bare
      divergence: `docs/function-reference.md` (the `AREAS` row) documents `#VALUE!` as THE rule for those shapes,
      so the docs assert as Excel's behaviour something the oracle contradicts. Phase 10 named `AREAS` in B1's
      blast-radius list and nobody measured it; it fell through because `AREAS` is not a mini-CSE consumer and so
      no pin covered it. Measure the family (`AREAS` over a reference, a union, a name, a computed array, an open
      range), match the oracle, and correct both docs twins.
      *Files:* the `AREAS` implementation under `Danfma.MySheet/Expressions/`, its test file, `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* It is a documented claim about Excel that measurement disproves, which is the exact defect class this
      project has shipped thirteen times, and here it is in user-facing docs rather than a comment.

## Controller addition after the Phase 10 final review, part 2 (2026-09-10) — measured by the Fable reviewer

- [ ] **25.** A RECTANGLE as the shorter operand against a ROW vector: the oracle fills the uncovered column with
      **0**, not `#N/A`, whenever the rectangle has at least as many rows as the vector has columns. Measured on
      Aspose.Cells 26.6.0 (2026-09-10), CSE and plain agreeing with themselves: `SUM(A1:B3*E5:G5)` = **420** with
      `COUNT` 9 and `INDEX(...,1,3)` = **0**, against `#N/A` / 6 / `#N/A` here; likewise `SUM(A1:B4*E5:H5)` 740,
      `SUM(A1:C4*E5:H5)` 1640, `SUM(A1:B3*A1:C1)` 42, the composite twin `SUM((E5:F5*E1:E3)*E5:G5)` 3000, and
      `SUMPRODUCT(A1:B3*E5:G5)` 420. **Phase 10 never fixtured this class** — every shape it measured has the
      VECTOR as the shorter operand — which is why the rule shipped coherent and the DOCS shipped a false claim of
      agreement (fixed in the phase's own wave).
      **The oracle is self-inconsistent here, so do not match it blindly.** With fewer rows than the vector has
      columns it reverts to the documented rule (`SUM(A1:B3*E5:H5)` `#N/A`, `COUNT` 6), yet
      `INDEX(A1:B2*E5:G5,1,3)` is still 0 in that same mode while its own `COUNT` says the position is uncovered.
      The column-vector mirror never behaves this way (`SUM(A1:C2*J1:J4)`, `SUM(A1:B3*J1:J4)`, `SUM(A1:C2*E1:E3)`
      are all `#N/A` with `COUNT` 6, agreeing with us).
      **Sharper characterisation, measured by Phase 10's fix wave:** the 0-fill turns on the RECTANGLE being the
      uncovered operand, not on which axis is short — the reverse direction `A1:C3*E5:F5`, where the VECTOR is
      short on the same axis and the extent is the same 3x3, agrees with us at `#N/A` with `COUNT` 6. Two further
      measurements make the defect read plainly: in the fewer-rows case the oracle's own extent is 3x3 rather
      than the 2x3 both operands imply (`COUNT` is **4**, not 6, and `INDEX(...,3,1)` is `#N/A` where ours is
      `#REF!`), and with a 4x2 rectangle against a 1x3 vector it CLIPS the extent to 3x3 and silently drops the
      rectangle's fourth row (`SUM` 420, `COUNT` 9, `INDEX(4,1)` `#REF!`). An engine that drops a row of user
      data is not a rule to reproduce. Decide deliberately: match the oracle, or keep our coherent `#N/A` under P0's "genuinely cannot match"
      clause as the two-axis mismatch already does. Whichever you choose, a real-Excel fixture would settle it and
      is worth the trouble here, because this is not an exotic shape.
      *Files:* `Danfma.MySheet/Expressions/Broadcasting.cs` if matched, `tests/Danfma.MySheet.Tests/Expressions/VectorBroadcastingTests.cs`, both docs twins
      *Why:* An ordinary shape a user would write, where we and the oracle disagree by a whole column of values
      rather than by an error code.

- [ ] **26.** The oracle's `INDEX` over a UNARY or LIFTED composite disagrees with the oracle's own element-wise
      forms at the collision between a real error and an uncovered position. Measured (CSE, `A3` = `=1/0`):
      `INDEX(-(A1:C3)*H1:H2,3,1)`, and the `ABS`, `LEN` and `ROUND` twins, all answer `#N/A`, while
      `SUM(ISERR(-(A1:C3)*H1:H2)*1)` = 1 and `ISNA` = 2 say that same position is `#DIV/0!`. Leaf and
      binary-composite forms answer `#DIV/0!` and agree with their own sums. MySheet answers `#DIV/0!` everywhere,
      matching the element-wise forms. Same family as the `INDEX`-over-a-reference quirk already pinned in
      `VectorBroadcastingTests`. Recorded so that nobody later "fixes" our behaviour by pinning it through
      `INDEX`, which is the one form the oracle answers inconsistently.
      *Files:* none unless a decision changes behaviour; the note is the deliverable
      *Why:* It is a trap for a future implementer, not a defect of ours.

## Controller additions after the Phase 3 final review (2026-09-10)

- [ ] **27.** A huge column reference CRASHES instead of answering `#REF!`, and it reaches every cell. Phase 3's
      correction M2 bounded `TryParseColumn`, but the sibling accumulators are still unbounded: `CellAddress.Parse`
      (`CellAddress.cs:23-27`) and `CellAddress.TryGetColumnRow` (`:38-70`), the latter reached by EVERY plain cell
      or range reference through `Parser.cs`, whose own comment calls it unguarded. Measured on both this branch and
      `main` by two independent reviewers: `=COUNTA(FXSHRXW1:FXSHRXW1)`, `=A<24 letters>1` and
      `=COUNTA(A<24 letters>1:A<24 letters>5)` throw `IndexOutOfRangeException` out of `GetCellValue`. Pre-existing,
      so it blocked no merge, but an exception escaping evaluation into the host is worse than any wrong answer, and
      this is the second item of that class in the sweep (the other is `WORKDAY`'s overflow). Bound both siblings the
      way M2 bounded the first, answer `#REF!`, and pin the boundary on both sides through a real formula rather than
      through the parser unit only.
      *Files:* `Danfma.MySheet/CellAddress.cs`, `Danfma.MySheet/Parsing/Parser.cs` if the arm needs it, the nearest test file
      *Why:* Two reviewers found it independently while checking M2's bound, and M2 shipped closing one of three
      doors.

- [ ] **28.** `COUNTIF` counts a cell holding the TEXT `"TRUE"` as the boolean TRUE. Measured on Aspose.Cells 26.6.0
      (2026-09-10, both entry modes agreeing) while investigating a user report about `IF`:
      `COUNTIF(A1:A2,TRUE)` and `COUNTIF(A1:A2,"TRUE")` are **0** on the oracle against **1** here, with `A1` holding
      the text `"TRUE"`. `SUMPRODUCT(--(A1:A2=TRUE))` is 0 on both, so plain comparison already distinguishes the
      types and only the criteria matcher conflates them. SILENT wrong number. Phase 11b fixed the `IF` and `NOT`
      half of that report and deliberately left this one here, because it is a different rule — criteria-family type
      equality, next to the existing criteria item.
      *Files:* `Danfma.MySheet/Expressions/CriteriaScan.cs` or wherever criteria equality lives, the criteria test files, both docs twins
      *Why:* A count that silently includes a row Excel excludes, in the family this sweep already has an item for.

## Controller addition after Phase 11b (2026-09-10)

- [ ] **29.** Eight boolean FLAG slots reject the text `TRUE` and `FALSE` where the oracle coerces them. Measured by
      Phase 11b on Aspose.Cells 26.6.0 (2026-09-10, both entry modes agreeing) across every remaining caller of
      `ValueCoercion.CoerceToBool`: `VLOOKUP(5,D1:D3,1,"TRUE")`, `HLOOKUP`'s range_lookup,
      `ADDRESS(1,1,1,"TRUE")` = `$A$1`, `TEXTJOIN(",","TRUE",…)` = `a,b`, `FIXED(1234.567,1,"TRUE")` = 1234.6,
      `DAYS360(…,"TRUE")` = 60, `VDB`'s no_switch and a bond argument — the oracle accepts the two words
      case-insensitively and rejects other text with `#VALUE!`, exactly the rule Phase 11b implemented for the
      condition slots. Phase 11b deliberately did NOT widen its own scope to these; the fix is likely to route the
      eight through the same opt-in extension it added, which makes this cheap.
      *Files:* the eight call sites the Phase 11b summary names, their test files, both docs twins
      *Why:* Phase 11b's design assumed a shared fix would BREAK the folding functions and the measurement showed
      the opposite — those functions never reach the helper, and the eight slots that do all want the new rule.

## Controller additions after Phase 7's Task 5 (2026-09-10)

- [ ] **30.** `ROWS` and `COLUMNS` over a single cell holding an ERROR answer the error on the oracle and `1` here:
      measured `ROWS(E2)` and `ROWS(E2:E2)` with `E2` = `=1/0` give `#DIV/0!` there, `1` here (both entry modes).
      Pre-existing, on the reference path rather than the array path Phase 7 added, so that phase left it untouched.
      Note the oracle also distinguishes a 1x1 error by PROVENANCE, which is why this is not a one-line fix:
      `ROWS(SEQUENCE(1)*E2)` is `#DIV/0!` but `ROWS(SEQUENCE(1)/0)` is 1, and `ROWS(FILTER(A1:A3,A1:A3>100,1/0))` is
      1 array-entered against `#DIV/0!` typed. MySheet answers `#DIV/0!` for all of them. Measure the whole matrix
      before choosing a rule, and decide deliberately whether provenance is reproducible at all.
      *Files:* `Danfma.MySheet/Expressions/ReferencePosition.cs` and the `ROWS`/`COLUMNS` tests
      *Why:* Recorded by Phase 7 rather than fixed, because a rule that depends on where an error came from may be
      another case of the oracle being inconsistent rather than a behaviour to copy.

- [ ] **31.** The oracle reads `IF(range, …)` as REFERENCE-returning in a criteria slot, and we reject it. Measured
      by Phase 7's Task 5 (array-entered): `COUNTIF(IF(A1:A3>0,A1:A3),">0")` = 2 and `SUMIF(IF(A1:A3>0,A1:A3),">0")`
      = 14, while `A1:A3*1` in the same slot is `#REF!` on both engines. Phase 11a's criteria gate rejects the `IF`
      form too, so MySheet answers `#REF!`. This is the same family as the existing criteria item and as the
      dropped `IF`-returns-a-reference item, and nobody owns it — decide the three together rather than one at a
      time, because they are one question about whether `IF` preserves a reference.
      *Files:* `Danfma.MySheet/Expressions/Logical/If.cs`, `CriteriaScan.cs`, the criteria test files, both docs twins
      *Why:* Phase 11a deliberately dropped the `IF`-reference item as not blocking, and this measurement shows the
      criteria family is where it actually bites.

## Implementation items

- [ ] **1.** Create `tests/Danfma.MySheet.Tests/Expressions/ExcelCompatibilitySweepTests.cs` holding the acceptance pins for (a)-(f) ONLY, each carrying **Aspose's** value and each therefore failing on `1b1e2d3` with the MySheet value named in the comment. Reuse `MathAggregateTests`'s `Calc(formula, params (string Id, object Value)[] cells)` shape (it is the nearest sibling; copy the helper rather than making it public). Pins, with today's failing value in brackets: **(a)** on the (a) fixture — `AGGREGATE(9,o,C1:C3)` = 8 for o in 0..3 [today 5] and 11 for o in 4..7 [passes], `AGGREGATE(9,o,F1:F3)` = 11 for all o in 0..7 [today 8 at 0-3], `AGGREGATE(3,o,C1:C3)` = 2 for 0..3 [today 1] and 3 for 4..7, `AGGREGATE(3,o,F1:F3)` = 3 for all o [today 2 at 0-3], `AGGREGATE(9,o,D1:D3)` = 8 / 11 and `SUBTOTAL(9,C1:C3)` = 8, `SUBTOTAL(3,C1:C3)` = 2, `SUBTOTAL(9,F1:F3)` = 11 as no-regression pins [all pass today]. **(b)** `SUBTOTAL(9,7)`, `SUBTOTAL(9,A1:A3,7)`, `SUBTOTAL(2,7)`, `SUBTOTAL(3,7)`, `SUBTOTAL(9,"7")`, `SUBTOTAL(9,TRUE)`, `SUBTOTAL(9,A1:A3,"")`, `AGGREGATE(9,4,7)`, `AGGREGATE(9,6,7)`, `AGGREGATE(9,4,A1:A3,7)`, `AGGREGATE(9,0,A1:A3,7)` → `ErrorValue.NotValue` [today 7/21/1/1/0/0/14/7/7/21/21], plus the no-regression pins `AGGREGATE(9,4,A1:A3,B1)` = 15, `SUBTOTAL(9,A1)` = 5, `AGGREGATE(15,6,7,1)` = `AGGREGATE(15,4,7,1)` = `AGGREGATE(14,6,7,1)` = `AGGREGATE(16,6,7,0.5)` = 7, and `AGGREGATE(15,6,1/0,1)` → `ErrorValue.NotValue` [today `#DIV/0!`]. **(c)** `MODE.SNGL(A1:A4)` = 2 on 2,1,1,2 [today 1]; = 1 on 1,2,2,1 [today 2]; `MODE.SNGL(A1:A6)` = 3 on 3,1,2,1,2,3 [today 1]; and `MODE(...)` / `AGGREGATE(13,4,...)` equal to it on each. **(d)** the seven `#DIV/0!` rows of (d)'s table [today `#NUM!`], plus `AGGREGATE(15,0,E2:E2,1)` = `#DIV/0!`, `AGGREGATE(15,6,G1:G1,1)` = `#NUM!`, `AGGREGATE(15,6,E1:E1,1)` = 5, `AGGREGATE(15,6,E1:E2,1)` = 5, `AGGREGATE(15,6,E2:E3,1)` = 9, `AGGREGATE(9,6,E2)` = 0 as no-regression pins. **(e)** MOVED TO PHASE 11a — DELIVERED 2026-09-10 in `tests/Danfma.MySheet.Tests/Expressions/DefinedNameArrayEligibilityTests.cs` (the pins are `COUNT((Rng<>"")*1)` = 3 [was 1], `SUM((Rng<>0)*1)` = 2 [was 1], `SMALL(IF(Rng>0,Rng),1)` = 5 [was 0], each asserted EQUAL to its literal-range twin exactly as this half asked; do NOT duplicate them here). **(f)** all fifteen `IF`/`CHOOSE` rows of (f)'s table.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ExcelCompatibilitySweepTests.cs`
      *Why:* TDD is mandatory here (master plan, "Rules specific to this repo") and this phase is unusually
      exposed to it: six of the seven changes make a currently-GREEN assertion wrong, so without a red suite
      first there is no way to tell "the fix worked" from "the old pin was deleted". Every bracketed value was
      measured on `1b1e2d3`, not reasoned. Two pins are the anti-vacuity ones the master plan demands: the
      (a) fixture deliberately separates `D1:D3` (nested SUBTOTAL only) from `F1:F3` (nested AGGREGATE only)
      so a predicate wired to the wrong node kind cannot pass, and `COUNT(IF(TRUE,A1:A3,B1:B3))` = 3 is pinned
      because today it answers **0** with no error — the silent-wrong-number class.
      Head the file with the Microsoft citations in this suite's mandatory format (article title +
      support.microsoft.com GUID + fetch date), **fetched at implementation time** — AGGREGATE, SUBTOTAL,
      MODE.SNGL and IF. No GUID is supplied here; do not invent one. Add, right under them, the sentence that
      where the page and the oracle disagree the ORACLE wins (P0 addendum), naming Aspose.Cells 26.6.0 and
      the measurement date, because item 2 deliberately contradicts the AGGREGATE page.

- [ ] **2.** Collapse the nested-skip predicate. In `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`: delete `SubtotalAndAggregate` from `internal enum NestedSkip : byte` (`:19-24`, leaving `None, Subtotal`), delete the `Aggregate => skip == NestedSkip.SubtotalAndAggregate,` arm from `IsNested` (`:258-266`, leaving `SharedFormulaSlave` → recurse, `Subtotal` → true, `_` → false), and rewrite the enum's doc comment: both functions now skip exactly a nested `SUBTOTAL`, the predicate stays a parameter only because options 4-7 need `None`. In `Danfma.MySheet/Expressions/Mathematics/Aggregate.cs:63-66` change the skip selection to `(options & 4) == 0 ? AggregateCodes.NestedSkip.Subtotal : AggregateCodes.NestedSkip.None` and rewrite the `bit2` comment block (`:59-61`) to state the measurement: Microsoft's options table says "Ignore nested SUBTOTAL and AGGREGATE functions", Aspose.Cells 26.6.0 counts the nested AGGREGATE at every option and skips only the nested SUBTOTAL, and under the P0 addendum the measurement is the specification — cite the (a) fixture and both answers. `Subtotal.cs:53-57` already passes `NestedSkip.Subtotal` and does not change.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`, `Danfma.MySheet/Expressions/Mathematics/Aggregate.cs`
      *Why:* Deleting the enum member rather than leaving it unused is what makes the reversal
      non-reversible-by-accident: with `SubtotalAndAggregate` gone the compiler finds every caller, and no
      future contributor can re-derive it from the Microsoft page without also re-adding the member and
      confronting the comment that says why it was removed. The two-member enum is still not a `bool`,
      because `None` is a distinct third behaviour that options 4-7 need.

- [ ] **3.** Flip the pins that item 2 contradicts. `tests/Danfma.MySheet.Tests/Expressions/AggregateCodesTests.cs`: `Gather_SubtotalAndAggregate_DropsBothNestedNodeKinds` (`:297-316`) and `Gather_SkipsANestedAggregateStoredAsASharedFormulaSlave_OnlyOnTheWideArm` (`:377-…`) both name a member that no longer exists — rewrite them as ONE test proving a nested `AGGREGATE` survives under `NestedSkip.Subtotal` AND under `None`, and that the shared-formula slave wrapper is seen through for a nested `SUBTOTAL` only (keep `Gather_SkipsANestedSubtotalStoredAsASharedFormulaSlave` at `:357` unchanged, it is the surviving half). Also check `GatherOver`'s signature at `:261` and the parameterised rows at `:385`/`:393`. `tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs`: rewrite `Aggregate_NestedBit_SkipsBothSubtotalAndAggregateCells` (`:770-805`) — rename it (it no longer skips both), change the options 0-3 expectation on `B1:B3` from 5.0 to 8.0 and the `AGGREGATE(3,0,B1:B3)` expectation from 1.0 to 2.0, keep 11.0/3.0 for options 4-7 and `SUBTOTAL(9,B1:B3)` = 8.0, and REPLACE the long "DIVERGÊNCIA REGISTRADA … a página documentada vence o oráculo" comment block with the new ruling and its measurement. Verify `Aggregate_ArrayForm_OverAPlainRange_HonoursTheNestedSkip` (`:869`) still holds — its fixture nests a `SUBTOTAL`, so it should be untouched; if it nests an `AGGREGATE`, it flips too.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/AggregateCodesTests.cs`, `tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs`
      *Why:* These are the two tests Phase 2 wrote SPECIFICALLY to lock the page-over-oracle decision, comment
      and all; leaving either behind makes the suite red or, worse, leaves a comment in the repo asserting the
      opposite of the shipped behaviour. `Gather_SubtotalAndAggregate_…` will not even compile after item 2,
      which is the desired forcing function; `MathAggregateTests` will compile and fail, which is why its
      exact expectations are enumerated here.

- [ ] **4.** Reject a non-reference argument in a `ref` slot. In `AggregateCodes.Gather`'s `default:` arm (`:196-219`), replace the final `return accumulator.Add(computed);` with `return Error.Value;` and a comment: `ref1, ref2, …` are declared as references and Excel enforces it literally — a literal, a string, a boolean, a name bound to a constant and an empty argument are all `#VALUE!` (measured, with the values from (b)'s table), while a single cell, a range, a union, a name bound to one and a reference-returning function all resolve through the `TryGetReference` arm just above and are unaffected. Extend `Feed`'s doc comment (`:24-43`) so its "a non-`Reference` argument that the mini-CSE can evaluate element-wise is `#VALUE!`" sentence becomes "any argument that does not denote a reference is `#VALUE!` — a computed array through the `TryStream` gate, a bare scalar through `Gather`'s `default:` arm". **This item MUST land together with item 5**, which is the array form's compensating arm.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`
      *Why:* One line, in the one arm every non-reference shape already funnels through — the `TryGetReference`
      recursion at `:201` and the unconditional error propagation at `:214` are both above it and both keep
      their meaning. The alternative (a guard in `Feed` before `Gather`) would have to re-implement "does this
      denote a reference", which is exactly the duplicated-answer problem `Feed`'s own comment warns about.
      The blast radius is real and bounded: `AGGREGATE`'s ARRAY form calls `Gather` too (`Aggregate.cs:159`),
      and `AGGREGATE(15,6,7,1)` = 7 reaches this same arm today — without item 5 it would regress from 7 to
      `#VALUE!`, which is a divergence in the other direction.

- [ ] **5.** Rewrite the tail of `Aggregate.ArrayForm` (`Danfma.MySheet/Expressions/Mathematics/Aggregate.cs:105-165`) so the array slot keeps its own, different rule. After the `ReferenceGuard.MissingSheet` guard (`:122-125`) and BEFORE the `ArrayEvaluation.TryStream` gate (`:132`), add the **1×1 collapse** (item (d)): resolve the argument with `NamedReferences.TryResolveReference(array, context, out var reference, boundOpenRanges: false)`; if it succeeds and the reference denotes exactly one cell (`CellReference`, or a `RangeReference` whose `GetBounds()` is 1×1 — `RangeBounds` is at `Danfma.MySheet/Expressions/RangeReference.cs:173`), read that cell's value and, if it is an error, return it UNCONDITIONALLY, ignoring `ignoreErrors` and never reaching the `k` bound. Then, after the `TryStream` gate fails, replace the bare `AggregateCodes.Gather(array, …)` at `:159` with the two-way split: if the same `TryResolveReference` succeeded, `Gather(reference, …, skip)` exactly as today; otherwise the argument is a non-reference scalar — evaluate it once, return `ComputedValue.Error(Error.Value)` if it is an error (item (b)'s last row), else feed that single value into the accumulator so `AGGREGATE(15,6,7,1)` stays 7. Hoist the resolution so it happens once, and comment each of the three arms with its measured pair.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/Aggregate.cs`
      *Why:* Three of this phase's measurements land in this one method and they only make sense together: the
      array slot ACCEPTS a scalar (7 → 7 on both sides) where the ref slot rejects it, REJECTS an error scalar
      with `#VALUE!` rather than the error, and collapses a 1×1 reference to a scalar before the options. Doing
      them as separate edits would produce two intermediate states that are each wrong in a measurable way.
      The 1×1 arm goes BEFORE the stream gate on purpose: `E2` is a `CellReference`, which never streams, so
      placing it after would work by accident today and break the moment Phase 7's producers widen `TryStream`.

- [ ] **6.** Change the `MODE.SNGL` tie-break in `Danfma.MySheet/Expressions/StatisticsMath.cs` `Mode` (`:113-145`). The current single pass with the strict `>` cannot express "first ENCOUNTERED among the winners" — measured: on 2,1,1,2 it must answer 2, and no update rule on a running best can, because 1 legitimately holds the crown at index 2. Use two passes over the same `Dictionary<double,int>`: pass one fills the counts and tracks `maxCount`; pass two walks `values` again in scan order and returns the first whose count equals `maxCount`. Keep `bestCount < 2 → Error.NA` unchanged (measured unchanged) and keep the "never sorted" contract. Rewrite the XML doc (`:113-118`) from "ties going to the first value that REACHES the winning count" to "ties going to the value that appears FIRST in scan order", citing the three measured fixtures.
      *Files:* `Danfma.MySheet/Expressions/StatisticsMath.cs`
      *Why:* Two O(n) passes and no extra allocation, against a one-pass alternative that would have to store a
      first-seen index per value (a wider dictionary value, touching every read). `Mode` is called by
      `ModeSngl`, by `MODE` (the alias) and by `AGGREGATE(13,…)` — measured identical on all three on both
      sides, so one fold fixes all three, which is exactly why Phase 2 extracted it.

- [ ] **7.** Flip the `Mode` pins in `tests/Danfma.MySheet.Tests/Expressions/StatisticsMathFoldTests.cs`. `Mode_Tie_TakesTheFirstValueToReachTheWinningCount` (`:67-76`) is the tie-break assertion, comment and all — rename it `Mode_Tie_TakesTheFirstValueInScanOrder`, change `[2,1,1,2]` → 1.0 to → 2.0 and replace its explanatory comment. `Mode_DoesNotSortThePopulation` (`:79-87`) becomes VACUOUS under the new rule: `[3,5,5,3]` answers 3 first-encountered, which is also what sorting would answer — change its fixture to `[5,3,3,5]` (first-encountered 5; sorted `[3,3,5,5]` would answer 3) so it still discriminates. `Mode_MostFrequentValue_Wins` (`:57-64`) and `Mode_NoRepeatedValue_IsNAError` (`:48-55`) are unaffected. Then check `AggregateCodesTests.Fold_Code13_IsTheModeInScanOrder` (`:54`) and `CompatibilityAliasTests.Mode_MatchesTheModeSnglGoldenExample` (`:69`) — flip only if their fixtures tie.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/StatisticsMathFoldTests.cs`, `tests/Danfma.MySheet.Tests/Expressions/AggregateCodesTests.cs`, `tests/Danfma.MySheet.Tests/Parsing/CompatibilityAliasTests.cs`
      *Why:* The master plan's anti-vacuity rule applies verbatim here: `Mode_DoesNotSortThePopulation` exists
      only to prove the population is consumed unsorted, and after item 6 its current fixture proves nothing.
      A fixture whose value coincides with the wrong answer is an assertion that cannot fail.

- [x] **8.** **MOVED TO PHASE 11a — DELIVERED 2026-09-10** ([`phase-11a-unblocking-slice.md`](phase-11a-unblocking-slice.md)): see that file for the UNION outcome, which B2 predicted wrong (a union name resolves to the Scalar outcome, so the whole expression is scalar-only and B2's `resolved.Evaluate` arm never runs — `SUM((UnN<>0)*1)` is still 1, not the literal's `#VALUE!`), and for the criteria gate's EXACT predicate (`!ArrayEvaluation.IsBareReferenceNode(argument) && ArrayEvaluation.IsArrayEligible(argument, context)` — `TryStream`'s first two conditions; a looser gate leaves two rows silently failing). Close the defined-name mini-CSE gap. In `Danfma.MySheet/Expressions/ArrayEvaluation.cs` add ONE arm to each switch, placed immediately after the `OpenRangeReference` arm and before the `Row`/`Column` arms so the reference shapes stay grouped: in `Probe` (`:199-280`) `case NameReference:` → resolve with `NamedReferences.TryResolveReference(expression, context, out var reference, boundOpenRanges: false)`, then `false` → `(true, false)` (opaque scalar: a name bound to a constant or a formula keeps today's behaviour, and an unknown name keeps reporting `#NAME?` through the scalar path), `ReferenceGuard.MissingSheet(reference, context) is not null` → `(true, false)` (so the `#REF!` is what gets broadcast), `RangeReference` → `(true, true)`, `OpenRangeReference` → `(false, false)` (the cost guard, unchanged for `MyColumn`), anything else (a single cell, a union) → `(true, false)`; in `TryBuildOperand` (`:304-363`) the mirrored arm, ending `operand = BuildRange(range, context); return true;` for the `RangeReference` outcome and `operand = new ScalarOperand(expression.Evaluate(context)); return true;` for every scalar outcome. Extract the shared five-way classification into one private helper both arms call, in the shape `ResolvePositionRange` (`:494-…`) already uses, and extend `ArrayEvaluation`'s class doc (`:29-50`) and `IsArrayEligible`'s doc (`:126-134`) to name the new shape.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* The master plan sized this at "roughly three items" on the belief that `Probe` had no
      `EvaluationContext`. It does — Phase 1 shipped `IsArrayEligible(Expression, EvaluationContext)` and both
      switches already take a context, and `ProbePosition`/`ResolvePositionRange` already resolve a
      `NameReference` through exactly this call for `ROW(Rng)` (measured working: `COUNT(ROW(Rng))` = 3 while
      `COUNT((Rng<>"")*1)` = 1). So the real change is one arm per switch plus a shared helper. The five
      outcomes are not negotiable: dropping the `MissingSheet` degradation loses a `#REF!`, and mapping a
      union or a single cell to `(true, true)` would hand `BuildRange` a node it cannot build. Cost: one
      resolution in the probe and one in the build per name node per evaluation — the same double resolution
      `ResolvePositionRange` documents and accepts at `:126-135`, on nodes that are rare in hot loops.

- [ ] **9.** Make `IF` return its branch's reference. In `Danfma.MySheet/Expressions/Logical/If.cs` change both branch returns from `Arguments[n].Evaluate(context)` to `NamedReferences.CaptureValue(Arguments[n], context)` (the 2-argument false case still returns `ComputedValue.Boolean(false)`), and add a `TryResolveReference` override mirroring `Choose`'s (`Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs:41-…`): coerce the condition, pick the branch, delegate to that node's own resolution; return `false` if the condition errors or the taken branch is not a reference. Keep the short-circuit — only the taken branch is touched.
      *Files:* `Danfma.MySheet/Expressions/Logical/If.cs`
      *Why:* This is `Choose`'s shape, verbatim, on the node next to it — `CaptureValue`'s own doc names "a
      CHOOSE alternative" as one of its three callers and exists because evaluating a range node yields
      `#VALUE!` (issue #8). The `TryResolveReference` half is not optional: `Evaluate` alone fixes
      `SUM`/`SUBTOTAL`/`AGGREGATE`/`SUMPRODUCT`/`COUNT`, but `ROWS`, `COLUMNS`, `AREAS`, `ISREF`, `INDEX` and
      `MATCH` go through `NamedReferences.TryResolveReference` (`:134`) and would still answer `#VALUE!`/
      `#REF!`/`FALSE` — six of the fifteen measured rows. Doing only the `Evaluate` half would also leave
      `COUNT(IF(TRUE,A1:A3,B1:B3))` at its silent **0**.

- [ ] **10.** Pin the two integration points item 9 must NOT move, with one assertion each. (a) **Cell boundary**: add to `tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs` a cell at `E2` holding `=IF(TRUE,A1:A3,B1:B3)` → 0 (the row intersection, matching the `=CHOOSE(1,…)` twin at `D2` which already gives 0) and the same formula at `E9` → `ErrorValue.NotValue` (no intersection; Aspose gives `#VALUE!` for the bare formula at a non-intersecting cell — measured). (b) **The mini-CSE array-condition path**: add to `ExcelCompatibilitySweepTests.cs` `SUM(IF(A1:A3>4,A1:A3,B1:B3))` = 16.0 — MySheet's `IfOperand` answer today, unchanged by item 9, and deliberately NOT taken from the oracle, whose non-CSE `#VALUE!` is implicit intersection and carries no signal. Comment it as such.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs`, `tests/Danfma.MySheet.Tests/Expressions/ExcelCompatibilitySweepTests.cs`
      *Why:* Item 9 makes `If` produce a `ComputedValueKind.Reference`, and S4 forbids one escaping as a cell
      value. `Workbook` (`:349-362`) already captures and intersects, and `CHOOSE` proves the path works — but
      "proves" by inspection is what P0 forbids, and the failure mode (a reference leaking into a cell) is
      silent. The second pin exists because item 9 touches the SAME node the mini-CSE's `IfOperand` handles,
      on a path where the oracle cannot referee.

- [ ] **11.** Rewrite the affected `docs/function-reference.md` prose. **AGGREGATE row (`:40`)**: replace the sentence pair "So 0-3 skip a referenced cell whose own formula is a `SUBTOTAL` or an `AGGREGATE` … That 5 is the options table read literally and is **not** a measured value: Aspose.Cells 26.6.0 … answers 8 … A known divergence in which the documented page wins." with the measured rule — 0-3 skip a referenced cell whose own formula is a `SUBTOTAL`, and only a `SUBTOTAL`; a nested `AGGREGATE` is counted at every option; over the `C1:C3` fixture `AGGREGATE(9,0,C1:C3)` = 8 and `AGGREGATE(9,4,C1:C3)` = 11, `AGGREGATE(3,0,C1:C3)` = 2 and `AGGREGATE(3,4,C1:C3)` = 3, all measured on Aspose.Cells 26.6.0 (2026-09-09), noting in one clause that Microsoft's options table says "nested SUBTOTAL and AGGREGATE" and that the measurement is what MySheet implements. Rewrite the `ref`-slot sentence to cover a bare scalar (`AGGREGATE(9,4,7)`, `AGGREGATE(9,4,A1:A3,7)` and a name bound to a constant are `#VALUE!`, while a single cell, a union and a reference-returning function are not), and the array-slot sentence to state the three measured asymmetries: a plain scalar is accepted (`AGGREGATE(15,6,7,1)` = 7), an error scalar is `#VALUE!` (`AGGREGATE(15,6,1/0,1)`), and a 1×1 reference collapses to its scalar before the options so `AGGREGATE(15,6,E2,1)` is `#DIV/0!` while `AGGREGATE(15,6,E2:E3,1)` is 9. Add the reference-form all-error sentence from (g) (`AGGREGATE(9,6,…)` = `AGGREGATE(3,6,…)` = `AGGREGATE(2,6,…)` = 0, `AGGREGATE(1,6,…)` = `#DIV/0!`, only the array form gives `#NUM!`) and relabel the remaining measured-but-hedged claims. **SUBTOTAL row (`:98`)**: drop "so the narrow rule is Excel's, not a guess" and the whole "`AGGREGATE(9,0,…)` over that same range is 5 here, which is **not** measured … the documented page wins" clause; state instead that SUBTOTAL and AGGREGATE now skip the same thing and give the same 8 on that fixture, both measured; add the `ref`-slot scalar rule (`SUBTOTAL(9,7)`, `SUBTOTAL(9,A1:A3,7)`, `SUBTOTAL(9,"7")`, `SUBTOTAL(9,TRUE)` → `#VALUE!`). **`MODE.SNGL` / `MODE` rows**: state the tie-break as first-in-scan-order with the 2,1,1,2 → 2 example. **`IF` row**: add that an `IF` whose taken branch is a reference yields that reference, so `SUM(IF(TRUE,A1:A3,B1:B3))` = 14 and `ROWS(…)` = 3 — the same rule `CHOOSE` already documents. Every number cited must be pinned by a committed test.
      *Files:* `docs/function-reference.md`
      *Why:* After this phase there is no such thing as a documented-page-wins divergence in these rows, and
      the two sentences that say there is are the most-read statement of the decision this phase reverses.
      The `grep` that finds them: `grep -n 'a guess\|is \*\*not\*\* measured\|\*\*not\*\* a measured value'
      docs/function-reference.md` — it must return nothing when this item is done.

- [ ] **12.** Mirror item 11 in `docs/pt-BR/function-reference.md`: the AGGREGATE row at `:44` ("Esse 5 é a tabela de options lida ao pé da letra e **não** é um valor medido … na qual a página documentada prevalece") and the SUBTOTAL row at `:102` ("de modo que a regra estreita é a do Excel, e não um palpite" … "esse 5 **não** é medido … na qual a página documentada prevalece"), plus the MODE.SNGL/MODE and IF rows. Translate the new prose rather than transliterating it, matching the register of the surrounding rows.
      *Files:* `docs/pt-BR/function-reference.md`
      *Why:* `docs/pt-BR/` is a full mirror and every doc edit needs its twin (master plan, "Rules specific
      to this repo"). It is a separate item because it is a separate skill and because leaving a stale
      pt-BR row is the failure this repo has hit before — the pt-BR sentences carry the same page-wins ruling
      in a form no English `grep` will find.

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded, **0 Warning(s)**, 0 Error(s). Deleting `NestedSkip.SubtotalAndAggregate`
      makes every stale reference a compile ERROR, which is the point; a `CS8509` non-exhaustive-switch
      warning on `IsNested` means the `_ => false` arm was dropped with it.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier format . && dotnet csharpier check .`
      → expected: exit 0, no file listed.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/ExcelCompatibilitySweepTests/*"`
      → expected: "Passed!", failed: 0. **Before items 2-9 land this same command must FAIL**, and on the
      exact values item 1 brackets: 5 for `AGGREGATE(9,0,C1:C3)`, 7 for `SUBTOTAL(9,7)`, 1 for
      `MODE.SNGL` over 2,1,1,2, `#NUM!` for `AGGREGATE(15,6,E2,1)`, 1 for `COUNT((Rng<>"")*1)`, `#VALUE!` for
      `SUBTOTAL(9,IF(TRUE,A1:A3,B1:B3))` and **0** for `COUNT(IF(TRUE,A1:A3,B1:B3))`. Confirm those seven
      failure reasons individually before implementing anything.
- [ ] `… --treenode-filter "/*/*/AggregateCodesTests/*"` and `… "/*/*/MathAggregateTests/*"` and `… "/*/*/StatisticsMathFoldTests/*"` and `… "/*/*/CompatibilityAliasTests/*"`
      → expected: all "Passed!", failed: 0. These four hold every pin items 3 and 7 flip; a red one here after
      the flip means a fixture ties where the item assumed it did not.
- [ ] `… --treenode-filter "/*/*/MiniCseConsumerTests/*"` and `… "/*/*/ArrayEvaluationTests/*"` and `… "/*/*/CellBoundaryIntersectionTests/*"` and `… "/*/*/ElementwiseLiftingTests/*"`
      → expected: all "Passed!", failed: 0, totals ≥ today's. Item 8 widens `Probe`, which every mini-CSE
      consumer reads, and item 9 changes a node `IfOperand` also handles; `MiniCseConsumerTests` already
      defines `MyName`, `MyColumn`, `MyCell` and `GhostName` on its `OnPositionGrid` fixture, so it is the
      suite that proves the four name shapes (range / open range / single cell / missing sheet) still land on
      the four intended outcomes. `ElementwiseLiftingTests` is Phase 8's and reaches the same two switches.
- [ ] `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release --no-build`
      → expected: both "Passed!", failed: 0 — the pair the pre-push hook and `.github/workflows/ci.yml` run.
      Measured baseline at `1b1e2d3`: core **total 1393** (1377 passing + Phase 8's **16** red TDD pins, which
      Phase 8 turns green before this phase starts), Excel **90 / 90 passing**. Expected here: 1393 plus
      Phases 8-10's additions plus this phase's, all green; Excel unchanged at 90 — this phase touches no
      loader, no writer and no wire format.
- [ ] `grep -n 'a guess\|is \*\*not\*\* measured\|\*\*not\*\* a measured value' docs/function-reference.md` and
      `grep -n 'não é um valor medido\|não é medido\|não um palpite\|página documentada prevalece' docs/pt-BR/function-reference.md`
      → expected: **no output** from either. These are the exact strings items 11 and 12 exist to remove.
- [ ] `grep -rn 'SubtotalAndAggregate' Danfma.MySheet/ tests/ docs/ plans/structured-table-references-and-aggregate/phase-11-excel-compatibility-sweep.md`
      → expected: matches only inside this phase file and inside Phase 2's file (a historical record, left
      alone). Any match under `Danfma.MySheet/` or `tests/` is an item-2 or item-3 leftover.
- [ ] Oracle re-run (the numbers are only as good as the fixture): rebuild the Aspose.Cells 26.6.0 console
      probe with the (a) fixture and re-measure `AGGREGATE(9,0..7,C1:C3)`, `AGGREGATE(3,0..7,F1:F3)`,
      `SUBTOTAL(9,C1:C3)`, `MODE.SNGL` on 2,1,1,2, `AGGREGATE(15,6,E2:E2,1)` and
      `SUBTOTAL(9,IF(TRUE,A1:A3,B1:B3))` → expected 8/8/8/8/11/11/11/11, 3 for every option, 8, 2, `#DIV/0!`,
      14. If any answer differs, a newer Aspose changed the oracle and the affected item stops until it is
      re-decided — do not "fix" a test to match a rule this file cannot reproduce.

## Risks carried by this phase

- **Item 4 is a behaviour change to a SHIPPED function against real workbooks.** `SUBTOTAL(9,7)` = 7 has been MySheet's answer since SUBTOTAL landed; after item 4 it is `#VALUE!`. Any corpus workbook that passes a literal where a ref is declared goes from a number to an error. That is what P0 demands and the master plan decided, but it is the one item whose blast radius is outside this repo. It ships as `feat(eval):`, never `fix`, and the docs items must say plainly which shapes are now rejected.
- **Item 8 widens array eligibility for every consumer at once.** `NumericAggregation`, `OrderSelection`, `Index.TryResolveReference`, `SumProducts`, `CriteriaScan`'s neighbours and `AggregateCodes.Feed` all read `IsArrayEligible`/`TryStream`, and `Feed` uses it to REJECT. So a formula like `SUBTOTAL(9,(Rng<>0)*1)` moves from today's accidental scalar answer to `#VALUE!` — which is the oracle's answer for the literal-range twin, hence correct, but it is a second behaviour change riding on item 8. Pin it.
- **Item 9 makes `If` produce a `ComputedValueKind.Reference`.** Every consumer that pattern-matches on a `ComputedValue` from an `IF` now has a new case to consider. `CHOOSE` and `OFFSET` already produce one, so the paths exist and are tested — but `IF` is far more common, and the cell-boundary pin in item 10 is the only automatic guard that a reference does not escape.
- **`SUBTOTAL(9,(A1:A2,A3:A3))` is `#VALUE!` on Aspose and 3 in MySheet — DO NOT act on it in this phase.** Measured 2026-09-09: Aspose rejects a UNION in a `ref` slot while accepting one in the array slot (`AGGREGATE(15,6,(A1:A2,A3:A3),1)` = 0). That contradicts Microsoft's own SUBTOTAL page, contradicts MySheet's shipped and tested union handling (`AggregateCodes.Gather`'s `UnionReference` arm, pinned by `Subtotal_IgnoresNestedSubtotals_ThroughAUnionAndAnOpenRange`), and would be a large behaviour change on a single suspicious measurement. Item 4 must therefore reject only NON-reference arguments — a union stays a reference and is untouched. Recorded to the master plan's open decisions for a real-Excel `.xlsx` to settle.
- **`SUBTOTAL(9,INDEX(A1:B3,0,1))` is `#REF!` in MySheet and 14 on Aspose** — measured on the same run, and out of this phase's scope: it is `INDEX`'s whole-column form (`row_num` = 0), a different function's missing feature, not a `ref`-slot rule. Do not fold it into item 4 or 5; record it as its own P0 item.
- **The code-review-graph MCP server was NOT reachable while this file was written** (the session's configured MCP servers failed to connect), so every code path cited here was located by `grep` and read directly, not through the graph. Line numbers are `1b1e2d3` and the master plan's "do not trust a line number" rule applies — anchor on member names and comment text.
- **Items 2-5 all edit `AggregateCodes.cs` / `Aggregate.cs`, which Phase 2 owns and no later phase touches**, so conflicts are unlikely; but item 8 edits `ArrayEvaluation.cs`, which Phase 7 (dynamic arrays) and Phase 8 (elementwise lifting) both rewrite. This phase executes after both — rebase rather than merge, and re-read the two switches before inserting the arm, because their case ORDER will have changed.

## Open questions owned by this phase

- **What is the 2-D scan order for the first error?** Still unsettled and still engine-wide: Aspose returned the same answer for `AGGREGATE(9,4,A1:B2)`, `SUBTOTAL(9,A1:B2)` and `SUM(A1:B2)` on Phase 2's fixture, so it did not discriminate column-major from row-major, and no fixture in this sweep does either. Carried from Phase 2 unchanged.
- **Does Excel's `SUBTOTAL` really reject a union in a `ref` slot?** Aspose says yes (see the risks); the Microsoft page's own examples say no. Aspose ANSWERED, so this is not an "oracle could not tell" question — it is an "the oracle's answer is not credible enough to act on alone" question, and it needs the self-verifying `.xlsx` fixture (author `=SUBTOTAL(9,(A1:A2,A3:A3))` in real Excel, read the cached `<v>`).
- **Does `MODE.MULT` share the new tie-break?** Aspose returned 3 for `MODE.MULT(A1:A6)` over 3,1,2,1,2,3 — but that is the top-left of a spilled array read through a single cell, which says nothing about the rest of it. Not probed further; `MODE.MULT` is out of scope and unchanged by item 6.
- ~~**Is a `LET`-bound name array-eligible after item 8?**~~ **ANSWERED and MEASURED 2026-09-10 in [Phase 11a](phase-11a-unblocking-slice.md), both sides, and pinned:** yes when the name is bound to a RANGE (`LET(r,A1:A3,SUM((r<>0)*1))` 1 → 2, `LET(r,A1:A3,COUNT(r*1))` 0 → 3, `LET(r,A1:A3,INDEX(r*2,3))` `#REF!` → 18, oracle 2, 3 and 18 in both modes — `DefinedNameArrayEligibilityTests.LetBoundName_InANestedArrayPosition_ResolvesThroughTheScope`); no when it is bound to a computed ARRAY, because `CaptureValue` evaluates that binding as a scalar before the arm sees it (`LET(r,A1:A3*1,COUNT(r*1))` = 0, `SUM(r*1)` = `#VALUE!`, `INDEX(r*2,3)` = `#REF!`, oracle 3, 14 and 18 in both modes). The `Let` NODE in a consumer's own argument slot is untouched (`SUM(LET(r,A1:A3,(r<>0)*1))` = 1, `SUM(LET(r,A1:A3,r*1))` = `#VALUE!`, as :203 predicted), and the criteria-slot pair is Phase 11a's handed-on divergence 7. Phase 7's `LET` routing (its item M1) owns what is left.

## Phase Summary

_(write when phase completes)_
