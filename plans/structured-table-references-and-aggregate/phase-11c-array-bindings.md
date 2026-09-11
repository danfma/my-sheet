# Phase 11c — Array bindings: LET, CHOOSE, unary `+` and a defined name stop collapsing a computed array

A binding site that goes through `NamedReferences.CaptureValue` — a `LET` binding, the chosen branch of `CHOOSE`, the operand of unary `+`, a defined name's definition — collapses any computed array (a producer, an operator over a range, a lifted function, an array `IF`) to its top-left element, silently. This phase makes those four sites carry the array through to the consumer, and makes the criteria family's gate see through them, so the one deliberately red test Phase 7 shipped goes green for the right reason.

## For Future Agents

As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases are done, fill in **Final Recap** and **Deployment Plan**.

**Governing rule (P0 with its 2026-09-09 addendum):** MySheet matches Excel, and "Excel" means **Aspose.Cells 26.6.0 as MEASURED**. A measurement beats a Microsoft page. A divergence is a work item, never a documented limitation. The oracle probe is `/tmp/aspose-probe-fable`; copy it to a task-unique scratch folder before editing. PLAIN and array-entered (`SetArrayFormula`) columns differ, and **this engine implements the ARRAY-ENTERED rule**; every number in this file was measured in BOTH modes on 2026-09-10 and the two agree unless a row says otherwise.

**Process (user-chosen, 2026-09-10):** subagent TDD tasks with per-task review, then the two-part final review (Fable, GLM-5.3 via the z.ai endpoint — Copilot dropped by the user on 2026-09-11 after four phases of fabricated or unmeasured findings), a fix wave, and `git rebase main` + `git merge --ff-only`. Work happens in a worktree at `/Volumes/Work/Develop/MySheet-11c` on branch `feat/array-bindings`, off `main` @ `f87d4b1` or later. **Commit messages carry no AI attribution of any kind** — the harness supplies trailers in a system reminder and the user's global rule forbids them; `main` is clean and this branch stays clean.

**Release:** the same release as Phase 7, which is NOT yet cut — tag `v3.19.0` stops at Phase 3 and `main` is 29 commits past it. That release will be **3.20.0** (Phase 7's `feat:` commits force a minor bump) and it must not ship with a deliberately red test in the suite, which is why this phase runs before Phases 4, 5 and 6.

## The defect, measured on `main` @ `f87d4b1` (2026-09-10)

`NamedReferences.CaptureValue` (`NamedReferences.cs:59-80`) has a closed type list — `RangeReference`, `OpenRangeReference`, `UnionReference`, `AnchoredRangeReference`, `SharedFormulaSlave` — and everything else falls to `_ => expression.Evaluate(context)`. A producer's scalar rule is `FirstElement` (Phase 7 item 3), so a `FILTER` in a binding becomes its top-left with no error; an operator over a range or an array `IF` becomes `#VALUE!` loudly. Five callers reach that fall-through: `Let.cs:30` (each binding), `LookupFunctions.cs:34` (`CHOOSE`'s chosen branch), `UnaryOperation.cs:27` (unary `+`), `NamedReferences.cs:42` (a defined name's definition, via `EvaluateDefinition`) and `Workbook.cs:392` (the cell boundary — where a bare producer in a cell correctly shows its top-left and must keep doing so).

The mini-CSE cannot see through any of them either: `ArrayEvaluation.Probe` has no `Let`, `Choose` or `UnaryOperation{Plus}` arm (Phase 8 deliberately kept `+` opaque as "Excel's reference-preserving no-op", `ElementwiseLiftingTests.cs:396-405`), and `ResolveNameShape` (`ArrayEvaluation.cs`, Phase 11a's four-outcome oracle) answers `Opaque` for any definition that is not itself a reference, because it goes through `NamedReferences.TryResolveReference`. A LET-bound name is worse still: the scope stores a `ComputedValue` (`EvaluationContext.WithName` `:78`, `TryGetName` `:92`), so by the time `NameReference.Evaluate` (`NameReference.cs:14-25`) runs, the array is already gone.

### Every row, both engines

| formula | MySheet today | oracle | kind |
| --- | --- | --- | --- |
| `LET(f,FILTER(A1:A3,A1:A3>0),SUM(f))` | **5** | 14 | silent |
| `LET(f,FILTER(…),ROWS(f))` | **1** | 2 | silent |
| `LET(f,FILTER(…),INDEX(f,2))` | `#REF!` | 9 | loud |
| `LET(f,FILTER(…),SUM(f)+ROWS(f))` | **6** | 16 | silent |
| `LET(f,FILTER(…),SUM(f*2))` | **10** | 28 | silent |
| `LET(f,FILTER(…),SUM(LEN(f)))` | **1** | 2 | silent |
| `LET(a,FILTER(…),b,SORT(a),SUM(b))` | **5** | 14 | silent |
| `LET(f,FILTER(…),g,f*2,SUM(g))` | **10** | 28 | silent |
| `LET(x,A1:A3*2,SUM(x))` | `#VALUE!` | 28 | loud |
| `LET(x,IF(A1:A3>0,A1:A3),SUM(x))` | `#VALUE!` | 14 | loud |
| `SUM(CHOOSE(1,FILTER(…)))` | **5** | 14 | silent |
| `SUM(CHOOSE(2,0,FILTER(…)))` | **5** | 14 | silent |
| `ROWS(CHOOSE(1,FILTER(…)))` | **1** | 2 | silent |
| `SUM(+FILTER(…))` | **5** | 14 | silent |
| `ROWS(+FILTER(…))` | **1** | 2 | silent |
| `SUM(+(A1:A3*2))` | `#VALUE!` | 28 CSE / 10 plain | loud |
| `SUM(+IF(A1:A3>0,A1:A3))` | `#VALUE!` | 14 CSE / 5 plain | loud |
| `SUM(ProdName)` — name defined as `FILTER(…)` | **5** | 14 | silent |
| `ROWS(ProdName)` | **1** | 2 | silent |
| `SUM(OpName)` — name defined as `A1:A3*2` | `#VALUE!` | 28 | loud |
| `SUM(LET(f,FILTER(…),f))` | **5** | 14 | silent |
| `LET(f,FILTER(…),COUNTIF(f,">0"))` | **1** | `#REF!` | silent |
| `LET(f,FILTER(…),SUMIF(f,">0"))` | **5** | `#REF!` | silent |
| `COUNTIF(LET(r,A1:A3,r*1),">0")` and `SUMIF(…)` | **0** | `#REF!` | silent (Phase 11a's standing limit) |
| `COUNTIF(ProdName,">0")` | **1** | `#REF!` | silent |
| `COUNTIF(CHOOSE(1,FILTER(…)),">0")` | **1** | `#REF!` | silent |
| `COUNTIF(+FILTER(…),">0")` | **1** | `#REF!` | silent |

Fixture throughout: `A1:A3` = 5, 0, 9, so `FILTER(A1:A3,A1:A3>0)` is the two-row 5, 9 (sum 14). `ProdName` is a workbook defined name whose definition is `FILTER(Sheet1!$A$1:$A$3,Sheet1!$A$1:$A$3>0)`; `OpName` is `Sheet1!$A$1:$A$3*2`.

### What must NOT move — the guards

| formula | both engines | why it is a guard |
| --- | --- | --- |
| `=LET(f,FILTER(…),f)` bare in a cell | 5 | the `@` top-left rule passes through a LET binding; the cell boundary keeps its top-left |
| `=LET(f,FILTER(…),f+1)` bare in a cell | 6 | and through an operator on the bound name — the scalar path is top-left + 1, not an array |
| `LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))` | 0 | a LET binding is evaluated ONCE; an array binding must be built once and read twice, never rebuilt with a second draw |
| `LET(n,ROWS(FILTER(…)),n)` | 2 | a scalar binding whose expression merely CONTAINS a producer stays a scalar |
| `LET(f,A1:A3,SUM(FILTER(f,f>0)))` | 14 | a binding that IS a range keeps `CaptureValue`'s reference path — Phase 11a's Rule A, unchanged |
| `SUM(+A1:A3)` | 6 (`UnaryOperationTests.cs:137`) and 356 on the `OnLengths` grid (`ElementwiseLiftingTests.cs:405`) | `+` over a bare reference still carries the REFERENCE (`SUM(+A1:A3)` reads cells); making `+` transparent to the mini-CSE must not turn it into a lift |
| `SUM(-(+A1:A3))` | measure before pinning — see item 2 | Phase 8 documented this as a gap because `+` was opaque; a transparent `+` may move it, and the oracle decides where |

## Design decision

**No new `ComputedValueKind`, no materialized array value in a cell** — Phase 7's design of record (`plans/dynamic-array-functions.md`) holds. What changes is the binding SITE and the three places that ask "what shape is this name".

1. **The evaluation scope can hold an array binding.** `EvaluationContext` gains a second binding form beside `WithName(string, ComputedValue)`: a name bound to an `ArrayOperand` that was built ONCE at binding time. `Let.Evaluate` decides per binding: if the binding expression is array-eligible in the binding's scope (`ArrayEvaluation.IsArrayEligible(expr, scope)` — which already excludes a bare reference, so a range binding keeps today's `CaptureValue` path), it calls `TryBuildOperand` once and binds the operand; otherwise it calls `CaptureValue` exactly as today. Evaluate-once falls out by construction, which is what makes `LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))` stay 0. The operand is valid for the scope's lifetime (one evaluation): a `RangeOperand` holds a dense handle and bounds, an `AxisSelectionOperand` its source plus an `int[]`.
2. **`NameReference` answers the binding's shape in each of its three roles.** `Evaluate` on an array binding returns the operand's top-left (`FirstElement`, the same rule a bare producer follows), which is what keeps `=LET(f,FILTER,f)` at 5 and `f+1` at 6. `Probe`/`TryBuildOperand`'s `NameReference` arm consults the scope BEFORE `ResolveNameShape` and answers `(true, true)` / the stored operand for an array binding. `TryResolveReference` answers false for it — an array binding is not a reference.
3. **The "bare reference" predicate becomes context-aware.** `ArrayEvaluation.IsBareReferenceNode(Expression)` is `expression is Reference or NameReference`, and it is the ONE predicate the top-level gates share: `TryStream` (`ArrayEvaluation.cs:233`), `CriteriaScan.RejectComputedArray` (`CriteriaScan.cs:235`), `Index`'s gate (`Index.cs:175`), plus the two If-branch uses (`:420`, `:431`, `:972`). A `NameReference` bound to an array must answer FALSE to it — it is not a reference, it is an array — so `COUNTIF(f,">0")` rejects with `#REF!` and `ROWS(f)` answers 2 through the mini-CSE gate. Add an overload `IsBareReferenceNode(Expression, EvaluationContext)` that returns false for a name whose scope binding is an array, migrate every gate to it, and keep the context-free one only where no context exists (the If-branch probe already has one; use it).
4. **A defined name whose definition is a computed-array expression is an array.** `ResolveNameShape` gains a fifth outcome: when `NamedReferences.TryResolveReference` fails, probe the DEFINITION expression in the definition's context (through the same recursion guard `EvaluateDefinition` uses at `NamedReferences.cs:36-44`); if array-eligible, the name is an array and the build builds that definition. `SUM(ProdName)` → 14, `SUM(OpName)` → 28, `COUNTIF(ProdName,">0")` → `#REF!`. The scalar path (`EvaluateDefinition` → `CaptureValue`) stays top-left, which is what a bare `=ProdName` in a cell must show.
5. **`CHOOSE` and unary `+` get `Probe`/`TryBuildOperand` arms in the shape Task 8 established for a scalar-condition `IF`.** `Choose`: evaluate `index_num` once (the caller-built scalar, like `IF`'s condition), probe/build the CHOSEN branch only, hand a taken scalar back as a 1x1 (never a `ScalarOperand`, so the consumer cannot fall back and re-evaluate — Task 8's volatility lesson), and a bare-reference branch resolves through `WrapScalar` (`BuildRange` for a reference-kind value). `UnaryOperation{Plus}`: probe/build the operand; a bare reference operand builds as a range (so `SUM(+A1:A3)` stays 6 and is NOT lifted), a computed array streams. The scalar paths (`LookupFunctions.cs:34`, `UnaryOperation.cs:27`) stay `CaptureValue` → top-left.
6. **The criteria gate follows from 1-5 with no code of its own.** A `Let`/`Choose`/`+` node in a range slot is now array-eligible → `RejectComputedArray` → `#REF!`. A LET-bound array name in that slot fails the context-aware bare-reference test → `#REF!`. `PositionalRange.RejectComputedArray` (`COUNTBLANK`) gets the same for free.

**Why not the alternative.** Capturing the binding EXPRESSION lazily (a closure re-probed by each consumer) would keep the scope value-only but breaks evaluate-once for a volatile binding unless the built operand is memoized on the closure — at which point it IS an operand in the scope with extra steps. Binding the operand is the honest form.

## Phase 1: The pins, red on arrival

Status: Complete

- [x] **1.** Create `tests/Danfma.MySheet.Tests/Expressions/ArrayBindingTests.cs` holding every row of the two tables above as an assertion with the ORACLE's value, so the whole file is red on `main` except the guard rows. Reuse `SelectionProducerFixture.Grid()` for `A1:A3` = 5, 0, 9 and add `ProdName`/`OpName` through `Workbook.DefineName` inside a local fixture. Head the file with the citation block this suite requires (Microsoft LET, CHOOSE and UNIQUE pages, GUID + fetch date) and the sentence that the oracle wins where a page disagrees. Every silent row's comment names today's wrong value beside the expected one.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ArrayBindingTests.cs`
      *Why:* TDD is mandatory here and the class is silent: without a red suite first there is no way to tell "the fix worked" from "the pin was flipped".
- [x] **2.** Measure on the oracle, in BOTH modes, the rows this file does not yet carry a number for, and add them to item 1's file with the column named: `SUM(-(+A1:A3))` and `SUM(-(+FILTER(…)))` (a transparent `+` under a lifted `-`), `=CHOOSE(1,FILTER(…))` bare in a cell, `=+FILTER(…)` bare in a cell, `=ProdName` bare in a cell, `LET(f,FILTER(…),f)*2` bare in a cell, `LET(f,FILTER(…),COUNTA(f))`, `LET(f,FILTER(…),SUMPRODUCT(f))`, `LET(f,FILTER(…),COUNTBLANK(f))`, `CHOOSE(1/0,FILTER(…))`, `CHOOSE(3,FILTER(…))` (out of range), a nested `LET(f,FILTER(…),LET(g,SORT(f),SUM(g)))`, and `LET(f,FILTER(A1:A3,A1:A3>100),SUM(f))` (an empty producer bound: `#CALC!` must survive the binding). Report any row where the oracle contradicts itself between modes before pinning it, per the entry-mode lesson in `tasks/lessons.md`.
      *Files:* the same test file; a task-unique copy of `/tmp/aspose-probe-fable`
      *Why:* A row an implementer needs and nobody measured becomes a guess in a comment; the bare-cell rows in particular decide what `Workbook.cs:392` must keep doing.
- [x] **3.** The three rows of `DynamicArrayTests.ALetBoundProducer_InACriteriaSlot_IsAKnownLimitHandedOverByPhase11a` are already at the oracle's values (`#REF!`, 14, 2) — leave them, but RENAME the test to `ALetBoundProducer_StreamsWhole_AndIsRefusedInACriteriaSlot` and rewrite its comment: it is no longer a known limit, and the sentence "the phase's ONE standing red pin" becomes the history of why it was red. Flip the two Phase 11a pins in `CriteriaComputedArgumentTests.cs:315-317` from `0` to `ErrorValue.Reference` with both numbers in the comment, and rename that test too.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/DynamicArrayTests.cs`, `tests/Danfma.MySheet.Tests/Expressions/CriteriaComputedArgumentTests.cs`
      *Why:* A test's NAME is an assertion (lessons.md, 2026-09-10): a method still called `IsAKnownLimit` after the limit is gone is the defect class this project has recorded sixteen times.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/ArrayBindingTests/*"` — every non-guard row FAILS with today's value named in the message; every guard row passes.
- Full core suite: **1927 + N total**, where N is the new file's test count, with **exactly** the new red rows plus the LET pin's three and the two 11a pins failing, and nothing else. Record the number.
- `dotnet csharpier check .` clean; Release build 0 warnings.

### Phase Summary

**Complete — `feat/array-bindings` @ `6ad7cea`** (`test(arrays): the Phase 11c pins, red on arrival, and the
two Phase 11a limits flipped to #REF!`), tree clean, no AI attribution trailer. Files:
`tests/Danfma.MySheet.Tests/Expressions/ArrayBindingTests.cs` (new, 51 tests: 36 red, 15 guards green),
`tests/Danfma.MySheet.Tests/Parsing/DynamicArrayTests.cs` (LET pin renamed
`ALetBoundProducer_StreamsWhole_AndIsRefusedInACriteriaSlot`, values unchanged, comment rewritten as
history), `tests/Danfma.MySheet.Tests/Expressions/CriteriaComputedArgumentTests.cs` (renamed
`LetBoundComputedArray_InARangeSlot_IsRefused`, FOUR pins `0` → `#REF!`, class header no longer lists a LET
half as must-not-move).

**Counts.** Before: core 1927 / 1, Excel 93 / 0. After: core **1978 / 38** (36 new red rows + the renamed LET
pin + the flipped 11a test — nothing else), Excel 93 / 0. Release build 0 warnings, csharpier clean.

**Oracle rows measured for the first time** (Aspose 26.6.0, 2026-09-11, plain / CSE 1x1, formula alone in
H20; today's MySheet value): `SUM(-(+A1:A3))` `#VALUE!` / -14 (today `#VALUE!`); `SUM(-(+FILTER(…)))` -14 /
-14 (today -5); `=CHOOSE(1,FILTER(…))` bare 5 / 5 (today 5); `=+FILTER(…)` bare 5 / 5 (today 5); `=ProdName`
bare 5 / 5 (today 5); `=LET(f,FILTER(…),f)*2` bare 10 / 10 (today 10); `COUNTA(f)` 2 / 2 (today 1);
`SUMPRODUCT(f)` 14 / 14 (today 5); `COUNTBLANK(f)` `#REF!` / `#REF!` (today 0); `CHOOSE(1/0,FILTER(…))`
`#DIV/0!` / `#DIV/0!` bare and under SUM (today the same); `CHOOSE(3,FILTER(…))` `#VALUE!` / `#VALUE!` bare
and under SUM (today the same); nested `LET(f,FILTER(…),LET(g,SORT(f),SUM(g)))` 14 / 14 (today 5);
`LET(f,FILTER(A1:A3,A1:A3>100),SUM(f))` `#CALC!` / `#CALC!` (today `#CALC!`).

**Finding against the brief.** The brief's item 3 and the ledger said "the two Phase 11a pins at
`CriteriaComputedArgumentTests.cs:315-317`"; the method holds FOUR assertions (COUNTIF/SUMIF over
`LET(r,A1:A3,r*1)` and `LET(r,A1:A3*1,COUNTIF/SUMIF(r,…))`), all measured `#VALUE!` plain / `#REF!` CSE, so
all four flipped — a two-row flip would have left a test asserting `0` against a measured `#REF!`. It is ONE
test method, so it counts as ONE failure, not two.

## Phase 2: The scope carries an array, and the gates know it

Status: Complete

- [x] **4.** Add the array-binding form to `EvaluationContext`: a `WithName(string name, ArrayOperand operand)` overload and a `TryGetArrayBinding(string name, out ArrayOperand operand)` lookup, stored in the same O(1)-per-binding linked structure `WithName` already uses (`EvaluationContext.cs:12-14` explains why not a dictionary). `TryGetName` keeps answering scalars only; an array binding is not a `ComputedValue`.
      *Files:* `Danfma.MySheet/Expressions/EvaluationContext.cs`
      *Why:* This is the one structural change; everything else reads it.
- [x] **5.** In `Let.Evaluate` (`Let.cs:20-34`), for each binding: if `!ArrayEvaluation.IsBareReferenceNode(expr) && ArrayEvaluation.IsArrayEligible(expr, scope)` and `TryBuildOperand` succeeds, bind the OPERAND; otherwise bind `CaptureValue(expr, scope)` exactly as today. Build once, before the next binding is evaluated, so a later binding that refers to an earlier array name (`LET(a,FILTER(…),b,SORT(a),SUM(b))`) finds it. Pin evaluate-once with the `SEQUENCE(3,1,RAND(),0)` row and with a counting custom function that proves the binding expression is evaluated exactly once when the name is read twice.
      *Files:* `Danfma.MySheet/Expressions/Logical/Let.cs`, `ArrayBindingTests.cs`
      *Why:* Evaluate-once is LET's semantics and the reason the operand, not the expression, is what gets bound.
- [x] **6.** `NameReference`: `Evaluate` returns the array binding's top-left through the same `FirstElement` helper a producer uses (`ArrayEvaluation.FirstElement`), checked BEFORE `TryGetName`; `TryResolveReference` returns false for an array binding. In `ArrayEvaluation.Probe` and `TryBuildOperand`, the `NameReference` arm checks `context.TryGetArrayBinding` first and answers `(true, true)` / the operand, then falls through to `ResolveNameShape` as today.
      *Files:* `Danfma.MySheet/Expressions/NameReference.cs`, `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* The name's three roles — scalar, reference, array — must agree on what the binding is, or `=LET(f,FILTER,f)` and `SUM(f)` disagree about the same `f`.
- [x] **7.** Add `IsBareReferenceNode(Expression, EvaluationContext)`: false when the expression is a `NameReference` whose scope binding is an array, otherwise the context-free answer. Migrate the gates at `ArrayEvaluation.cs:233` (`TryStream`), `CriteriaScan.cs:235` (`RejectComputedArray`), `Index.cs:175`, and the If-branch uses at `ArrayEvaluation.cs:420/431/972` to the overload. Grep for every remaining caller of the context-free form and justify each one in a comment or migrate it. Update `CriteriaScan.cs:92-99`'s doc, which states the predicate verbatim.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`, `CriteriaScan.cs`, `Lookup/Index.cs`, `PositionalRange.cs` if it has its own use
      *Why:* This is what turns `COUNTIF(f,">0")` into `#REF!` and `ROWS(f)` into 2 at the same time: one predicate, two consumers, opposite answers, both right.
- [x] **8.** Confirm `Workbook.cs:392` (the cell boundary) needs NO change and PIN the reason: a bare `=LET(f,FILTER,f)` reaches the boundary as a `Let` node whose `Evaluate` returns the top-left scalar, so the boundary never sees an array. Assert the two bare-cell rows (5 and 6) through `Workbook.GetCellValue`, not through `Expression.Evaluate`.
      *Files:* `ArrayBindingTests.cs`
      *Why:* The cell path and the expression path have disagreed before in this project (`INDIRECT` measured 14 through the cell and `#REF!` bare); a boundary row pinned only through `Evaluate` proves nothing about cells.

### Verification Plan
- `ArrayBindingTests` filtered run: every LET row green, every CHOOSE/`+`/defined-name row still red (they belong to Phase 3), every guard green.
- `DynamicArrayTests.ALetBoundProducer_*` and the two `CriteriaComputedArgumentTests` pins green.
- Full core suite: only Phase 3's rows red. Excel suite 93 / 0. csharpier clean, 0 warnings.
- Mutation, backed up and restored from the backup: revert item 5's array branch to `CaptureValue` and confirm the LET rows go back to 5 / 1 / 1 and the evaluate-once row still passes (it passes for the wrong reason without item 5 — say so in the report).

### Phase Summary

**Complete — `feat/array-bindings` @ `1fd687d`** (`feat(arrays): a LET binding carries a computed array,
built once and read by every consumer`), tree clean, no AI attribution trailer. `ArrayBindings.Capture` is
the one helper; the context-free `IsBareReferenceNode` was REMOVED (zero callers after migrating the gates),
not kept dead.

**Counts.** Before: core 1978 / 38. After: core **1988 / 15**, with exactly Phase 3's fifteen rows red, Excel
93 / 0.

**Mutation.** Disabling the array branch turns 37 rows red; the `SEQUENCE(3,1,RAND(),0)` evaluate-once row
passes for the wrong reason under that mutation — its comment says so and the counting-function row is the
real guard.

**Ruling, measured by the controller** (Aspose 26.6.0, 2026-09-11, H1 and H5, both modes — CSE column is the
target), from a finding in Phase 1: the SCALAR reading of an array-eligible binding is the built array's
TOP-LEFT, at every binding site, not `Evaluate`'s `#VALUE!`. Bare in a cell: `=OpName` 10,
`=LET(x,A1:A3*2,x)` 10, `=+(A1:A3*2)` 10, `=CHOOSE(1,A1:A3*2)` 10, `=LET(x,IF(A1:A3>0,A1:A3),x)` 5 (PLAIN
intersects per row for the last four, `#VALUE!` in H5 — mode split, CSE pinned). MySheet before this phase:
`#VALUE!` for all five. Generalised into one helper (`ArrayBindings.Capture`) used by `Let`, `Choose`,
`UnaryOperation{Plus}` and `EvaluateDefinition`, with a `TopLeft` scalar reading. `Workbook.cs`'s cell
boundary keeps calling plain `CaptureValue` and does NOT change — bare `=A1:A3*2` in a cell stays `#VALUE!`
(pinned in `CellBoundaryIntersectionTests`, a separate sweep decision).

**Design item 6 was wrong**: the criteria gate for a `Let` NODE did not "follow with no code" — `Probe` had
no `Let` arm, so `ProbeLet`/`TryBuildLet` were added. Two more Phase 11a pins flipped
(`DefinedNameArrayEligibilityTests`: `SUM(LET(r,Rng,(r<>0)*1))` 1 → 2; `LET(r,A1:A3*1,…)` rows to 3 / 14 /
18). `PositionalRange` is a struct inside `CriteriaScan.cs`.

**Reviewer finding**, pinned as a known divergence under sweep item 32: probe and build can classify a
rebinding chain over `IF(TRUE,Rng,0)` differently; the probe is only ever more conservative.

## Phase 3: Defined names, CHOOSE and unary `+`

Status: Complete

- [x] **9.** `ResolveNameShape` gains the outcome `Array`: when `TryResolveReference` fails and the definition is a workbook defined name whose expression is array-eligible in the definition's context (evaluate through the same recursion guard `EvaluateDefinition` uses, `NamedReferences.cs:36-44`, so a self-referential name still answers `#REF!`), the probe answers `(true, true)` and the build builds the definition. The scalar path is untouched: `=ProdName` bare shows the top-left.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`, `Danfma.MySheet/Expressions/NamedReferences.cs`
      *Why:* `SUM(ProdName)` = 5 today is the same silent collapse one site over; Phase 11a's Rule A covered a name bound to a RANGE and stopped there.
- [x] **10.** `Choose` arm in `Probe` and `TryBuildOperand`, modelled on `TryBuildScalarConditionIf`: `index_num` is the caller-built scalar (evaluated once); an error or out-of-range index is the loud 1x1 the scalar path already answers; the CHOSEN branch alone is probed/built; a taken scalar comes back as a 1x1, a bare-reference branch through `WrapScalar`. The probe answers for the union of branches exactly as `ProbeIfBranches` does, with the same condition-aware treatment of a bare reference (`ProbeIfBranches`'s `conditionIsArray` is `false` here: CHOOSE's index is always scalar).
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`, `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs` (comment only)
      *Why:* `CHOOSE` is `Entry<Choose>` in the registry (`FunctionRegistry.cs:1518`, verified 2026-09-10), i.e. `Consumes`, so the lift arm does not shadow this one and the arm is reachable.
- [x] **11.** `UnaryOperation{Plus}` arm: probe/build the operand. A bare reference operand → `BuildRange` (the reference is preserved, `SUM(+A1:A3)` stays 6 and 356, and `+` itself is NOT a lift). A computed array → its operand. **Two Phase 8 pins FLIP with this item and the commit body must carry both numbers:** `ElementwiseLiftingTests.cs:398-399` pins `SUM(+LEN(A1:A3))` and `SUM(LEN(+A1:A3))` at `#VALUE!` while stating the oracle answers 6 for both — a transparent `+` makes them 6 (the lift happens BELOW or ABOVE the `+`, not at it). Rewrite that test's rationale (`:394-397`): the reason `+` is not a lift survives, the mechanism (opaque to the probe) is what changes. Pin whatever item 2 measured for `SUM(-(+A1:A3))` (documented today as a `-6` gap) and `SUM(-(+FILTER(…)))` against the oracle's array-entered column.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`, `Danfma.MySheet/Expressions/UnaryOperation.cs` (comment), `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs`
      *Why:* `SUM(+FILTER(…))` = 5 is the same collapse; Phase 8's "keep `+` opaque" was a decision about lifting, not about producers, which did not exist yet.
- [x] **12.** The criteria-gate rows for CHOOSE, `+` and a defined name (`COUNTIF(CHOOSE(1,FILTER(…)),">0")`, `COUNTIF(+FILTER(…),">0")`, `COUNTIF(ProdName,">0")`) go green with NO gate change — confirm that, and if any needs a change, report it before making one.
      *Files:* none expected
      *Why:* The gate keys on `IsArrayEligible`, which items 9-11 extend; a gate change here would mean the design is wrong somewhere.

### Verification Plan
- `ArrayBindingTests` filtered run: **all green**, guards included.
- Full core suite **1927 + N / 0** — zero failures for the first time since Phase 7 registered its producers. Record the exact total. Excel 93 / 0. csharpier clean, 0 warnings.
- Mutation: mis-flag nothing, but DELETE the `Choose` arm and confirm exactly the CHOOSE rows go red; delete the `Plus` arm and confirm exactly the `+` rows go red; revert item 9 and confirm exactly the defined-name rows go red. Each restored from a backup, never with `git checkout`.
- Volatility: `LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))` over 200 seeds is 0 on every seed; `SUM(CHOOSE(IF(RAND()<0.5,1,2),FILTER(…),SEQUENCE(3)))` over 200 seeds answers only 14 or 6, never 5 or 1 (the Task 8 collapse shape, one site over).

### Phase Summary

**Complete — `feat/array-bindings` @ `ceb7539`** (`feat(arrays): a defined name, CHOOSE and unary + carry a
computed array too`), tree clean, no AI attribution trailer. Files: `Danfma.MySheet/Expressions/ArrayEvaluation.cs`,
`NamedReferences.cs`, `UnaryOperation.cs`, `Lookup/LookupFunctions.cs`, `CriteriaScan.cs` (doc only),
`NumericAggregation.cs` (doc only), `ArrayBindings.cs` (doc only), plus `ArrayBindingTests.cs`,
`ArrayEvaluationTests.cs`, `ElementwiseLiftingTests.cs`, `ElementwiseLiftingMechanismTests.cs`,
`MiniCseConsumerTests.cs`/`MiniCseVolatileTaintTests.cs` (a renamed method in comments).

**Counts.** Before: core **1988 / 15**. After: core **2001 / 0** — zero failures for the first time since
Phase 7 registered its producers — Excel **93 / 0**, Release build 0 warnings, csharpier clean. Thirteen
rows added: the three bare-cell scalar readings the controller's ruling names (`=OpName`, `=+(A1:A3*2)`,
`=CHOOSE(1,A1:A3*2)` = 10), CHOOSE's chosen bare range beside an array branch (3),
`SUM((+A1:A3)*B1:B3)` = 32, the `+`-over-a-reference criteria guards (4), and two seeded volatility rows.

**Mutations**, each backed up to the scratchpad and restored FROM THE BACKUP, tree verified clean after
each: delete the `Choose` arms → exactly **4** red (the three `AChosenArray_…` rows +
`COUNTIF(CHOOSE(1,FILTER(…)),">0")`), plus the new volatility rows. Delete the `Plus` arms (probe, build and
the `IsBareReferenceNode` arm) → exactly **11** red: the 7 `+` rows the brief names, `SUM((+A1:A3)*B1:B3)`,
and the three Phase 8 pins this phase flipped. Revert item 9 (both halves) → exactly **5** red: the three
`ADefinedNameOverAComputedArray_…` rows, `COUNTIF(ProdName,">0")` and the `=OpName` bare-cell row. Extra
mutation for finding (1) below: keep the `Array` outcome and revert ONLY `IsBareReferenceNode`'s name clause
→ all **4** defined-name consumer rows still red — the proof that the predicate extension is load-bearing.

**Seeded runs, 200 fresh workbooks each.** `LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))` → `0:200`.
`SUM(CHOOSE(IF(RAND()<0.5,1,2),FILTER(…),SEQUENCE(3)))` → `14:94, 6:106` and nothing else (never 5, never 1).
Two extra shapes: `SUM(+CHOOSE(…))` → `14:104, 6:96`; `SUM(-CHOOSE(…))` → `-14:103, -6:97`. The CHOOSE run is
now a permanent test (`AVolatileChooseIndex_NeverCollapsesTheBranchItTakes`).

**Findings against the brief/plan.** (1) Design item 4 and brief item 9 were INCOMPLETE: `ResolveNameShape`'s
new `Array` outcome alone cannot make `SUM(ProdName)` 14, because `TryStream`'s FIRST condition
(`IsBareReferenceNode`) still admitted the name. Fixed by extending `IsBareReferenceNode`'s `NameReference`
arm to also answer `false` for a name whose definition is a computed array (the oracle agrees:
`ISREF(ProdName)` is FALSE). (2) The plan's item 11 as written would have created **FIVE** new divergences —
`COUNTIF(+A1:A3,">0")` is 2 on the oracle, `SUMIF` 14, `AVERAGEIF` 7, `ISREF` TRUE, `COUNTBLANK` 0 — fixed
inside the existing predicate: `IsBareReferenceNode` answers for the OPERAND of a `+`, so `+` is transparent
BELOW a lift or an operator and still a reference at the top. All five are pinned in
`AUnaryPlusOverABareReference_IsStillAReferenceAtTheTopLevel` (`COUNTIF`/`SUMIF`/`AVERAGEIF`/`COUNTBLANK`) and
its sibling `AUnaryPlusOverABareReference_IsStillAReferenceForIsref` (`ISREF`, added by the final-review fix
wave — both were measured from the start but not separately pinned until then). (3) Item 11 undercounted the
Phase 8 pins:
THREE flip, not two — besides `ElementwiseLiftingTests.cs`'s `SUM(+LEN(A1:A3))`/`SUM(LEN(+A1:A3))`
(`#VALUE!` → 6, renamed `LiftedCall_UnderATransparentUnaryPlus_IsLifted`), two mechanism pins asserted the
opacity directly and were rewritten (`ArrayEvaluationTests`, `ElementwiseLiftingMechanismTests`).
(4) `ElementwiseLiftingTests`' "three known divergences this phase LEAVES OPEN" preamble was already stale
before this task — rewritten to say which phase closed each. (5) `SUM(-(+A1:A3))` on the 1, 2, 3 grid is now
-6, the number `docs/workbook-and-expressions.md` named as Excel's — that gap is closed and the bullet was
false as written (Task 4's job). (6) Three docs bullets and their pt-BR twins were now false and only
partly covered by Task 4's numbered items — the extra sites are named in the ledger and closed by Task 4.
(7) Item 10's "a bare-reference branch resolves through `WrapScalar`" needed CHOOSE's OWN capture, not IF's:
`SUM(CHOOSE(1,A1:A3,FILTER(…)))` is 14 (oracle, both modes) via `Choose`'s own reference capture since Onda
3, not IF's `branch.Evaluate` (which would have given `#VALUE!`). Pinned as
`AChosenBareRange_BesideAnArrayBranch_StillCarriesItsCells`. (8) Minor, not changed: `Choose.TryResolveReference`
still carries its own copy of the index rule, duplicated with `TryChoose`'s equivalent one; left alone and
named here.

**Refactors** (behaviour-neutral, suite-verified): `NamedReferences`' cycle guard becomes a reusable
`DefinitionGuard` `using` scope; `ProbeIfBranches` becomes `ProbeBranches(ReadOnlySpan<Expression>, …)` shared
by `IF` and `CHOOSE`; `Wrap`/`WrapScalar` hoisted out of `TryBuildScalarConditionIf` and shared with the
`CHOOSE` arm; `Choose.TryChoose` holds the index rule for both of `CHOOSE`'s paths; `ResolveNameShape`'s
`out` becomes `Expression?` (the resolved reference, or the DEFINITION for the `Array` outcome).

## Phase 4: Docs, the design of record, and the bookkeeping

Status: Complete

- [x] **13.** `docs/workbook-and-expressions.md:638-640` and its pt-BR twin: the bullet "A producer bound by `LET`, passed through `CHOOSE`, or through a unary `+`, collapses to its top-left" is DELETED from the known-divergence list and a positive sentence is added to the producers section: a binding carries the array, evaluated once, with the `LET(f,FILTER(…),SUM(f))` = 14 and `SUM(f)-SUM(f)` = 0 examples. Grep both twins for `LET` and `CHOOSE` afterwards and diff the twin against the original paragraph by paragraph.
      *Files:* `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`
      *Why:* The sentence is the project's most-repeated defect class the moment the code no longer carries it.
- [x] **14.** `plans/dynamic-array-functions.md`: the "standing gaps" section that names M1 option (a) as taken is rewritten — option (b) is what shipped, in this phase, and the section says so with the row table. `docs/function-reference.md` rows for `LET`, `CHOOSE` and the unary operators (and pt-BR) gain the one clause each: an array binding / a chosen array / a `+` over an array streams whole in a consumer and shows its top-left bare.
      *Files:* `plans/dynamic-array-functions.md`, `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* The design of record currently records the limitation as the decision; a reader would conclude the collapse is intended.
- [x] **15.** Master plan and sweep bookkeeping: the Phase 7 row's "one deliberate red pin" clause gains "closed by Phase 11c"; the Phase 11a row's "handed on ONE standing limit" gains the same; `phase-11-excel-compatibility-sweep.md` item 31/32's mention of "the dropped `IF`-returns-a-reference item" is untouched (that question is NOT decided here — a bare `A1:A3` in an `IF` branch is still `#VALUE!`), but any sweep line that lists the LET/CHOOSE/`+` collapse as a divergence is closed with this phase named. `tasks/lessons.md` gets whatever this phase teaches.
      *Files:* `plans/structured-table-references-and-aggregate.md`, `plans/structured-table-references-and-aggregate/phase-11-excel-compatibility-sweep.md`, `tasks/lessons.md`
      *Why:* Three plan files currently describe the collapse as a limit somebody else owns.

### Verification Plan
- Every `.md#anchor` added or touched resolves (script it, as Phase 7's docs task did).
- `grep -rn "collapses to its top-left" docs/ plans/` returns only history sentences that name this phase as the fix.
- pt-BR twins: full Português do Brasil orthography, and a structural count (bullets, code fences, table rows) equal to the EN file for every touched section.

### Phase Summary

**Complete** — docs/plans/lessons only, no engine change, tree unchanged in behaviour (core **2001 / 0**,
Excel **93 / 0**, unchanged from Phase 3's baseline). Sites changed, beyond the items' own file lists:

- `docs/workbook-and-expressions.md` / `docs/pt-BR/workbook-and-expressions.md`: item 13's bullet deleted, a
  positive paragraph added to "Dynamic array producers"; the two opaque-`+` bullets at `:723-727`/`:761` and
  `:859-865`/`:907-909` (both false since Task 3 made `+` transparent) removed and folded into the same
  positive paragraph; the defined-name "Known divergences" bullet's LET-related sub-clauses (three shapes →
  two) rewritten, since the LET-node-in-a-consumer's-slot and LET-bound-computed-array halves are now closed
  too — found by grepping the LET-bound-name test file, not named in the ledger's own "extra sites" list.
- `docs/function-reference.md` / pt-BR: `LET` and `CHOOSE` rows gain the one clause each. **Finding:** the
  brief's third target — "the unary operators" — does not exist as a row in this file; unary operators are
  not registered `FunctionRegistry` entries, so there is nothing there to add a clause to. Left alone rather
  than inventing a row; the mechanism is documented in `workbook-and-expressions.md` instead. The pt-BR `LET`
  row was also missing a whole sentence the EN row has (the range-stays-a-range one) — brought to parity
  while touching the row for the new clause.
- `plans/dynamic-array-functions.md`: "Standing gaps handed on" rewritten — option (b) shipped, with a row
  table; the "name bound to a computed array" bullet also closed (it shares the `CaptureValue` cause with
  the first bullet, and Phase 11c fixes both).
- `plans/structured-table-references-and-aggregate.md`: Phase 7 row's clause gains "closed by Phase 11c"
  (and its stale test-name citation fixed); Phase 11a row's "handed to Phase 11" clause gains the same; the
  Phase 11c row itself marked **Complete** with counts.
- `plans/structured-table-references-and-aggregate/phase-11-excel-compatibility-sweep.md`: **checked, no
  edit** — items 31/32 discuss the unrelated `IF`-returns-a-reference question, correctly left alone, and a
  repo-wide grep for the LET/CHOOSE/`+` collapse phrase found no other sweep line to close. **Correction
  (final-review fix wave):** that claim was false — the grep was for a phrase, and item 15's actual
  instruction was any sweep line that LISTS the collapse as a divergence. Two do, `:203` and `:658`, both
  still stating the pre-11c numbers as current fact; both are now marked **CLOSED BY PHASE 11C** with the
  measured numbers (`3`, `14`, `18` and `2`).
- `plans/structured-table-references-and-aggregate/phase-11a-unblocking-slice.md`: three stale citations
  fixed, not two — the brief named lines 138 and 143, but line 148 (`CriteriaComputedArgumentTests`'s old
  name) cites the OTHER renamed test and was equally stale. All three paragraphs also had their substance
  updated (closed, not just renamed), since the paragraphs assert current behaviour, not only a name.
- `plans/structured-table-references-and-aggregate/phase-8-elementwise-lifting.md:307`: the open question
  answered — transparent, and the blanks/text worry never arises because the top-level gate keeps a bare
  `+range` on the range path.
- `tasks/lessons.md`: one dated section, five lessons (see the file).

**Verification.** `grep -rn "collapses to its top-left" docs/ plans/` returns **six** hits, not four, all
legitimate: three in Phase 7's own historical plan file (the option (a) description in the M1 correction
block, the same phrase reused where that correction's own text says the option was never the description
that shipped, and the unrelated pre-existing cell-boundary paragraph at `:457`) and three in this phase's own
file (item 13 quoting the OLD bullet as an instruction, the Verification Plan line quoting the grep, and
this sentence quoting the grep again). Every anchor this
task added or touched (`#dynamic-array-producers`, `#produtores-de-array-dinâmico`, `#named-ranges`,
`#intervalos-nomeados`) resolves against an existing heading — no new heading was added, so no new anchor
needed generating. pt-BR structural parity checked globally: headings 23/23, bullets **63/63**, code fences
38/38, table rows 49/49, all equal EN/pt-BR after the edits (the LET occurrence-count gap, 34 EN / 31 pt-BR,
is pre-existing — 36/33 before this task's edits — and unchanged in size). Every pt-BR paragraph touched was
diffed against its English original paragraph by paragraph; one mechanical defect was caught and fixed this
way — a scripted-style deletion had dropped the blank line between two paragraphs in BOTH files at the same
spot (`**Which factory a new built-in uses…**` glued onto the previous sentence) — see `tasks/lessons.md`.

## Risks carried by this phase

- **A gate migrated to the context-aware predicate that had no context.** If a caller of `IsBareReferenceNode(Expression)` runs before any `EvaluationContext` exists (the parser, the dependency extractor), it cannot know about LET bindings and must keep the context-free answer — which is correct for it, because a LET binding does not exist at parse time. Item 7 must list each such caller and say why it stays.
- **A transparent `+` changes three Phase 8 answers.** `SUM(+LEN(A1:A3))` and `SUM(LEN(+A1:A3))` are PINNED at `#VALUE!` (`ElementwiseLiftingTests.cs:398-399`) with the oracle's 6 named in the comment, and `SUM(-(+A1:A3))` is documented as a `-6` gap — all three under the assumption that `+` is opaque. Item 2 measures the third on the oracle before item 11 touches anything; the oracle's array-entered column is the target for all three, and each flip is stated with both numbers in the commit body.
- **Evaluate-once and volatility taint.** An array binding built at LET time must taint the cell exactly as the scalar binding does today (`MiniCseVolatileTaintTests` covers array-condition IFs and, since Phase 7's fix wave, the scalar-condition path — not LET). Phase 3's verification adds the seeded runs; if a volatile inside a bound producer does not refresh the cell, that is a defect of item 5, not of the taint machinery.

## Open questions owned by this phase

- None that block dispatch. The one decision the user might revisit — whether `+` should stay fully opaque and only LET/CHOOSE/names change — is answered by the measurement: the oracle answers 14 for `SUM(+FILTER(…))` in both modes, and under P0 that is a work item.

## Final Recap

A `LET` binding, `CHOOSE`'s chosen branch, unary `+`'s operand and a defined name's definition all used to
collapse a computed array (a producer, an operator over a range, a lifted call, an array `IF`) to its
top-left the moment it crossed one of those four sites, silently for a producer and loudly (`#VALUE!`) for
everything else — the reason Phase 7 shipped with one deliberately red pin and Phase 11a handed over a
standing `LET` limit. Four tasks closed it without a new `ComputedValueKind` or a materialized array value in
a cell: the evaluation scope gained a second binding form (a name bound to an `ArrayOperand`, built once);
`NameReference` answers the binding's shape consistently in its three roles (scalar top-left, reference-no,
array-yes); the "bare reference" predicate that every top-level array gate shares became context-aware, so
the SAME name answers `#REF!` to `COUNTIF` and `2` to `ROWS`; and `CHOOSE`/unary `+`/a defined name's
definition each got the same `Probe`/`TryBuildOperand` treatment through one shared helper
(`ArrayBindings.Capture`). Two design assumptions were wrong and caught by running the suite rather than by
reasoning about it: the criteria gate for a bare `Let`/`Choose`/`+` node needed its OWN `Probe` arms (item 6
did not "follow for free"), and making `+` transparent everywhere would have created four new criteria-family
divergences (fixed by keeping the SAME predicate's answer for a `+` over a bare reference at a consumer's top
level). The suite went from core **1927 / 1**, Excel 93 / 0 at the phase's start to core **2001 / 0**, Excel
**93 / 0** — zero failures for the first time since Phase 7 registered its producers — across four tasks and
three engine commits (`6ad7cea`, `1fd687d`, `ceb7539`) plus this task's docs/plans/lessons commit(s), with
`dotnet csharpier check .` clean and a Release build carrying 0 warnings throughout. **Six** test methods
were renamed as their pins moved from "wrong on purpose" to "right" (`git diff 75f5e84..HEAD -- tests/`
confirms the count), and every renamed test's OLD name was meant to be grepped out of `docs/`, `plans/` and
`tasks/lessons.md` before this phase closed — the final review found one surviving bare citation of a
renamed name (`plans/…/phase-7-dynamic-arrays.md`, closed by this fix wave) alongside the three citations
the controller's own briefs had already enumerated and fixed. One divergence outside this phase's scope is explicitly
unmoved and named as such wherever it appears: a bare-reference branch under a scalar-condition `IF`
(`SUM(IF(TRUE,A1:A3,0))`) is still `#VALUE!` against the oracle's 14, owned by sweep items 31/32's
"does `IF` return a reference?" question, not by this phase.

## Deployment Plan

The release is manual through `.github/workflows/release.yml` (`workflow_dispatch`), which runs versionize,
tags, packs both packages and publishes to NuGet.org via Trusted Publishing. This phase ships inside the
same **3.20.0** release as Phase 7 (Phase 7's `feat:` commits already force the minor bump past `v3.19.0`).

1. Rebase `feat/array-bindings` (head `ceb7539` plus this task's docs commit) onto the current `main` and
   confirm `dotnet build Danfma.MySheet.slnx -c Release` is 0 warnings, `dotnet csharpier check .` is clean,
   and both suites are green (`dotnet run --project tests/Danfma.MySheet.Tests/... -c Release` and the Excel
   project's equivalent, NOT `dotnet test`) — expect core **2001 / 0**, Excel **93 / 0** on `main`'s HEAD
   too, since nothing on `main` past `75f5e84` should have moved either number.
2. Fast-forward merge (`git merge --ff-only`) into `main`; do not squash — the commit history is the audit
   trail for which task did what and the per-task review already happened.
3. Confirm `main`'s HEAD carries no deliberately red test — this phase's entire reason for existing before
   Phases 4/5/6 run — with the same suite command as step 1.
4. Dispatch the release workflow once Phase 7's own commits are confirmed on the same `main` HEAD (do not
   dispatch from the worktree branch). The CHANGELOG's Phase 11c entries come from this branch's `feat:`
   subjects: `test(arrays): the Phase 11c pins, red on arrival, and the two Phase 11a limits flipped to
   #REF!`, `feat(arrays): a LET binding carries a computed array, built once and read by every consumer`,
   `feat(arrays): a defined name, CHOOSE and unary + carry a computed array too`, and this task's
   `docs(arrays):` commit(s) — write each for a reader of the changelog, not for the controller.
5. Remove the `/Volumes/Work/Develop/MySheet-11c` worktree and delete `feat/array-bindings` only after the
   ff-merge is confirmed reachable from `main` (`git merge-base --is-ancestor feat/array-bindings main`).
