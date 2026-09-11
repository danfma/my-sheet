# Phase 8: Elementwise lifting of unary operators and scalar functions in the mini-CSE

Status: Complete   <!-- Not started | In progress | Complete -->

Adversarial verifier verdict: **needs-revision** (1 blocker, 3 majors, 7 minors — folded in below; every one was MEASURED with probes against the Release build, not argued). Executes AFTER Phase 2 and BEFORE Phase 7.

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `elementwise-lifting`. Design dependencies: `reference-semantics` (Phase 1 — **landed**, commits `32a67f6..c04e78b`; its Phase Summary supersedes the Phase 1 plan text). This phase executes **AFTER Phase 2 (`aggregate`)** — not because it needs AGGREGATE, but because both phases edit `FunctionRegistry.Entries` and this one rewrites ~180 of its lines; landing it first guarantees a conflict. Phase 2 landed while this file was being written (`7d968a3 feat(eval): AGGREGATE(function_num, options, ref1, [k])`, `224a739 test: …`), so `AGGREGATE` is present in the registry and its union tag is allocated (323 attributes, tags 0-322). It executes **BEFORE Phase 7 (`dynamic-arrays`)**, which consumes it: FILTER/SORT/UNIQUE/SEQUENCE become array PRODUCERS, and every operand this phase adds must already accept a producer as an argument (`LEN(FILTER(...))`) without a second design pass. This phase adds **no** node type, **no** MemoryPack union tag and **no** public API.

Line numbers and counts in this file were measured at commit `b043ade`; the registry counts were re-checked at `224a739` (Phase 2 complete: **306** entries, 323 union tags). Several cited files were being edited concurrently while this was written. Anchor edits on member names, `case` labels and heading text, and re-read before editing.

## Design decision

**The bug.** The mini-CSE (`Danfma.MySheet/Expressions/ArrayEvaluation.cs`) recognizes exactly five array-producing shapes — `RangeReference`, `AnchoredRangeReference`, `BinaryOperation`, `If` with an array condition, and `ROW`/`COLUMN` over a reference. Everything else falls to `Probe`'s `default: return (true, false)` (:226-227) and `TryBuildOperand`'s `default: operand = new ScalarOperand(expression.Evaluate(context))` (:307-309): an OPAQUE SCALAR, evaluated once and broadcast. `UnaryOperation` and all ~180 pure-scalar built-ins land there, and evaluating one over a range yields `#VALUE!` (a range has no scalar value) which then poisons every element. Measured on today's tree (probe at `/tmp/mysheet-lift-probe`, `dotnet run -c Release`):

| formula | today | Excel |
| --- | --- | --- |
| `SUMPRODUCT(--(E6:E8>1))` | `#VALUE!` | 2 |
| `SUM(--(E6:E8>1))` | `#VALUE!` | 2 |
| `SUM(-(E6:E8>1))` | `#VALUE!` | -2 |
| `SUM(LEN(D7:F9))` | `#VALUE!` | 7 |
| `SUM(LEN(TRIM(D7:F9)))` | `#VALUE!` | 6 |
| `SUM(ABS(E6:E8*-1))` | `#VALUE!` | 6 |
| `SUM(IF(LEN(D7:F9)>0,1,0))` | `#VALUE!` | 3 |
| `IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))>0,"Show","Hide")` | `#VALUE!` | `Show` |

**And three SILENT wrong answers the brief did not list, measured on the same fixture** — worse than the errors above because no user sees a failure: `SUM(ISNUMBER(E6:E8)*1)` → **0** (Excel 3), `SUM(IFERROR(E6:E8,0))` → **0** (Excel 6), `COUNT(LEN(D7:F9))` → **0** (Excel 9). In each case the scalar `#VALUE!` is *consumed* by the enclosing node (a comparison, IFERROR's error arm, COUNT's non-numeric skip) instead of propagating.

**The classification: an explicit per-entry flag on `RegistryEntry`, defaulting to `Consumes`.** Not a marker interface / `ScalarFunction : Function` base record, for one measured reason that settles it: **`Function` declares no `Arguments` member** (`Danfma.MySheet/Expressions/Function.cs` is one line: `public abstract record Function : Expression;`). A generic lift cannot reach a function's arguments except through `RegistryEntry.GetArguments`, and cannot rebuild the node except through `RegistryEntry.Create`. The registry lookup is therefore **mandatory regardless of where the flag lives**, so the flag rides along for free, while a marker interface would need 180 node-file edits *on top of* the same lookup. (Cost of that lookup, measured: `FunctionRegistry.ByType.TryGetValue(node.GetType(), …)` = **16.5 ns**, versus 0.32 ns for an `is Function` type test. It is paid once per function NODE per EVALUATION, never per element; the same lookup already runs per node in `FormulaWriter.Call`, `DependencyExtractor.VisitArguments` and `AnchoredFormulaSupport.TryGetArguments`. A 50k-cell `SUM(A1:A50000)` costs 2.4 ms, so one extra lookup is under 0.001%.)

The flag **defaults to `Consumes`** and lifting is opt-in, because the two failure modes are not symmetric: a contributor who forgets the flag on a new scalar function loses the feature (today's behaviour, no regression), while a contributor who forgets it on a new range-aware function gets **silently wrong numbers**. To keep the registry header's "adding a function is one line here" contract intact, the opt-in is a sibling factory — `Elementwise<Len>("LEN", 1, 1, …)` instead of `Entry<Len>("LEN", 1, 1, …)` — not a sixth positional argument, so a lifted entry stays one line and `grep -c 'Elementwise<'` is the count.

**The list, derived and verified, not guessed.** Two independent signals were computed over all 305 registered built-ins (306 with Phase 2's AGGREGATE): (1) a marker grep of each node's own record body for `Reference`, `ArrayEvaluation`, `PositionalRange`, `RangeValueSequence`, `ArgumentFlattening`, `NumericAggregation`, `ReferenceGuard`, `TryResolveReference`, `CriteriaScan`, `RangeSnapshot`, `IsArrayEligible` → **63** consumers; (2) an executable range-awareness oracle — for every argument position, compare `F(…,A1:A3,…)` against `F(…,C4:C6,…)` where the two rectangles differ in position AND contents; a scalar-blind node evaluates both to `#VALUE!` and answers identically → **76** consumers. Their union is **97**. Both signals have documented false NEGATIVES (the oracle cannot see `ROWS`/`COLUMNS`/`AREAS`/`OFFSET`/`ISREF`, which answer the same for two rectangles; the body grep cannot see helpers in the same file), so **18 more were added by hand and are named individually in item 5**. Final: **180 `Elementwise` / 126 `Consumes`** (125 + AGGREGATE). All 180 were then smoke-driven through the prototype: **176 produced a correct 3-element array; 4 (`ODDFPRICE`/`ODDFYIELD`/`ODDLPRICE`/`ODDLYIELD`) failed only because the probe capped its filler arity at 6 and they need 7-8** — a probe artifact, not a mechanism failure.

**The mechanism: a REUSED node over mutable `ScratchLiteral` slots (candidate c), chosen on numbers.** All three candidates were prototyped in `/tmp/mysheet-lift-probe` against the Release build and driven over `SUM(LEN(A1:A50000))` (50 000 elements, cells pre-warmed, 10 reps):

| mechanism | ms/eval | alloc/eval | ns/elem | lift overhead vs. raw range read |
| --- | --- | --- | --- | --- |
| baseline `RangeOperand` only | 2.37 | 0 B | 47.5 | — |
| (a) typed literal nodes + `entry.Create` per element | 8.53 | **3.81 MB** | 170.5 | 123.0 ns |
| (c-lite) one `ComputedValueLiteral` per arg + `entry.Create` per element | 7.63 | **4.58 MB** | 152.6 | 105.1 ns |
| **(c) `ScratchLiteral` slots, node built once** | **5.01** | **0 B** | **100.3** | **52.8 ns** |

Two arguments (`SUM(ROUND(B1:B50000,2))`) widen the gap: (c) 73.8 ns/elem and 0 B, (a) 207.0 ns/elem and 112 B/elem (5.6 MB per evaluation). Allocation per element for (a) decomposes exactly: `Expression[1]` 32 B + `NumberValue` 24 B + `Len` 24 B = 80 B.

Beyond the numbers, (a) is **not total**: there is no `ValueExpression` literal for `ComputedValueKind.Reference`, so a reference-kind element (from `ScalarOperand(+A1:A3)`, or a host custom function) cannot be represented as a typed literal at all. A literal carrying a `ComputedValue` is total over the kind enum. The cost of (c) is that `ScratchLiteral` is a MUTABLE record inside an engine whose contract is "expression trees are immutable". That is acceptable here and only here: the instance is created inside `LiftedFunctionOperand`'s constructor, is reachable only from that operand's private array, is discarded with the operand, is never serialized (no `[MemoryPackable]`, no union tag), and is never hashed or compared — verified by grep: the engine holds no `Dictionary<Expression,…>`, `HashSet<Expression>` or `ConditionalWeakTable` anywhere. The classification is what guarantees a lifted node only ever calls `Evaluate` on its argument nodes and never stores them.

**Unary.** `UnaryOperand` mirrors `BinaryOperand` for `Negate` and `Percent` only. `Plus` is deliberately NOT lifted: it is Excel's reference-preserving Lotus no-op (`UnaryOperation.Evaluate` routes it through `NamedReferences.CaptureValue`, `UnaryOperation.cs:16-27`), and it already works — **measured `SUM(+A1:A3)` = 6 today**, through the reference value path, and unchanged after this phase (the prototype reports `+A1:A3` as `SCALAR:RangeReference{…}`, i.e. still an opaque scalar carrying the reference). `--x` is two nested `Negate` nodes and lifts twice; **measured through the prototype: `--(E6:E8>1)` → 2, `-(E6:E8>1)` → -2, `E6:E8%` → 0.06, `-(-(E6:E8))` → 6.**

**Volatiles are unaffected, by construction.** A lifted function's arguments are built as OPERANDS first, so a scalar argument is a `ScalarOperand` evaluated exactly once at build time and broadcast — the rule `ArrayEvaluation`'s class comment already states for IF's branches. **Measured with a counting custom function:** `ROUND(E6:E8,TICK())`, `IFERROR(E6:E8,TICK())` and `ROUND(E6:E8*TICK(),0)` each call TICK **exactly once** through the lifted path, matching today's `SUM(IF(A1:A3>0,TICK(),0))` = 1 call. `RAND`/`NOW`/`TODAY` take zero arguments and can never lift. `RANDBETWEEN` *could* (two scalar arguments) and is therefore classified `Consumes` — see open questions.

**`Probe` still never evaluates.** The new arm inspects only the registry classification and recurses into the ARGUMENTS' eligibility. **Measured: `IsArrayEligible("ROUND(E6:E8,TICK())")`, `IsArrayEligible("IFERROR(TICK(),E6:E8))")` and `IsArrayEligible("LEN(TRIM(D7:F9))")` each call TICK zero times.** The invariant `IsArrayEligible ⇒ TryEvaluateStream succeeds` holds because the new probe arm and the new build arm walk the identical argument list from the identical registry entry, with no oracle in between (unlike `ROW`/`COLUMN`, which must resolve).

**The cell boundary does NOT change, and that is a deliberate documented gap.** `Workbook.EvaluateCell` (`Workbook.cs:327-380`) calls `NamedReferences.CaptureValue` and then handles only `value.TryGetReference(...)` → `ImplicitIntersection.Apply`. The mini-CSE is entered **only** by consumers that explicitly call `IsArrayEligible`/`TryEvaluateStream` (`NumericAggregation.cs:114`, `OrderSelection.cs:73`, `CriteriaScan.OpenArrayOrRange:180`, `Index.cs:24`/`:173`, `AggregateCodes.cs:45`, `Aggregate.cs:134`) — the boundary is not one of them. A lifted function therefore cannot reach the boundary as an array: `Len.Evaluate` is untouched, so bare `=LEN(A1:A3)` in a cell still evaluates the scalar body and still answers `#VALUE!` after this phase, exactly as `=A1:A3*2` and `=IF(B2:B5="Show",1,0)` do today (`ImplicitIntersection`'s own class comment: "the ARRAY half … has no producer in this engine yet and is deliberately absent"). **Correction to Phase 7's design text**, which claims the bare-array case belongs to "Phase 5's FIX B": FIX B is **Phase 1's**, it landed (commit `b89b856`, `feat(eval): implicit intersection at the cell boundary`), and it implemented the RANGE half only. Adding the array half means giving `EvaluateCell` a second arm that runs `ArrayEvaluation.IsArrayEligible(expression, context)` on the WHOLE cell expression and takes `ElementAt(0)` — a new call site on the cell hot path, for every cell, and it must not fire for a plain `=A1:A3` (already handled by the reference arm). That belongs to Phase 7, which is the phase that actually creates producers users will type bare; **this phase does not add it, and item 12 pins the current answer so Phase 7 changes it deliberately rather than by accident.**

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** The mechanism's central safety claim ("a lifted node only ever calls `Evaluate` on its argument nodes and never stores them") is the WRONG invariant. Replacing every argument node with a `ScratchLiteral` also requires that no lifted function PATTERN-MATCHES its argument nodes — and 10 of the 180 do, to detect an OMITTED optional argument, which the parser represents as a literal `BlankValue` node: FIXED, DOLLAR, NUMBERVALUE (`Text/TextFormatting.cs` `Arguments[1] is not BlankValue` etc.), TEXTBEFORE/TEXTAFTER (`DelimiterSplit`), VALUETOTEXT, REGEXEXTRACT/REGEXREPLACE/REGEXTEST (`Text/RegexFunctions.cs`), ADDRESS (`Lookup/LookupFunctions.cs`).
      *Measured evidence:* with the designer's own `LiftedFunctionOperand`: scalar `FIXED(A1,,TRUE)` = "1234.50" but lifted `FIXED(A1:A3,,TRUE)` = ["1235"|"1235"|"1235"] — decimals silently become 0 (a `ScratchLiteral` is not a `BlankValue`, so the omitted branch is skipped and the slot's Blank is coerced to 0); `DOLLAR(A1:A3,)` = ["$1,235"|…] silent; `NUMBERVALUE(B1:B3,,)`, `TEXTBEFORE(C1:C3,"-",,1)`, `TEXTAFTER(C1:C3,"-",,,1)`, `ADDRESS(D1:D3,2,,,)`, `ADDRESS(D1:D3,2,,FALSE)` → `#VALUE!` per element. FIXED/DOLLAR are the silent-wrong-number class this phase exists to remove. `FormulaWriter.ToFormula` round-trips these formulas byte-for-byte, so the engine already reads/writes them. The derivation probe missed it because it filled every optional slot with `1` and never omitted one.
      *Correction (minimal, keeps 0 B/elem):* in the `LiftedFunctionOperand` constructor take the original `Expression[] arguments` and substitute a scratch slot ONLY where the slot is not a `BlankValue` — `slots[j] = arguments[j] is BlankValue ? arguments[j] : _scratch[j]` (an omitted slot is always a literal, hence never per-element). Then add the omitted-argument guard test from M2. Alternative (rejected: loses lifting for ADDRESS/TEXT family, which Excel lifts): move the 10 to `Consumes`.

## Major corrections

- [ ] **M1.** Item 7(b)/item 8 say "return `(false,false)` on the first argument whose Probe does not succeed" — that turns a TOLERATED open range into a hard refusal that unwinds the whole enclosing tree and REGRESSES a formula that works today.
      *Measured evidence:* today `=SUM(IF(A1:A3>0,1,LEN(B:B)))` = 3 (`LEN(B:B)` lands on Probe's `default: (true,false)` opaque-scalar arm). Through the specified arms `IsArrayEligible(IF(...))` becomes false, NumericAggregation takes the scalar path, and `IF(...).Evaluate` sees a range in a boolean slot → `#VALUE!`. Same shape reaches SMALL/SUMPRODUCT/AGGREGATE.
      *Correction:* in BOTH new arms (Function and UnaryOperation), treat "an argument's build refuses" as the OPAQUE-SCALAR answer, not a refusal — Probe returns `(true, false)`, `TryBuildLift` returns `new ScalarOperand(function.Evaluate(context))`. Preserves `IsArrayEligible ⇒ TryEvaluateStream succeeds`, keeps `SUM(LEN(A:A))` = `#VALUE!` as the risk section states, removes the regression.
- [ ] **M2.** Item 9(b)'s classification guard test is 169/180 VACUOUS and tests a property the mechanism does not need. Measured over all 180 on its fixture: 169 rows have `#VALUE!` on both sides (they assert only that `#VALUE!` propagates), 11 assert a constant of that same `#VALUE!`. FIXED/DOLLAR/NUMBERVALUE/TEXTBEFORE/TEXTAFTER/ADDRESS all PASS it and all break under lifting (B1).
      *Correction:* per `Elementwise` entry with `MaxArgs > MinArgs`, build `F(<range>, …, , <trailing>)` with an OMITTED middle slot, lift it, and assert element i equals the scalar `F(<cell i>, …, , <trailing>)`. Fails today for the 7 measured functions, passes after B1's constructor fix. Optionally a source-level assertion that no `Elementwise` node's file contains `Arguments[` followed by anything other than `.Evaluate(`/`.Length`.
- [ ] **M3.** Item 2's date pin `=SUM(MONTH(E6:E8))` → 3.0 can never go green: on THIS engine serial 1 is 1899-12-31, not Excel's 1900-01-01, so `MONTH(1)`=12, `MONTH(2)`=1, `MONTH(3)`=1 → SUM = **14**. Pin 14.0, or (better) move the date pin to a fixture with real dates (`DATE(2024,3,15)` cells) so it stops encoding an unrelated epoch divergence — which is itself a pre-existing P0 item, recorded in the master plan's open decisions.

## Minor corrections (fold in while implementing)

- The two-argument perf figure (73.8 ns/elem) does not reproduce and is implausible (lower than the one-argument 100.3); re-measured `SUM(ROUND(B1:B50000,2))` ≈ **165 ns/elem, 408 B/eval** — the conclusion (≈0 B/elem, ~2× baseline) holds; only that figure is wrong.
- Derivation arithmetic: "63 + 76, union 97, +18 by hand" vs "automatic signals cover 108" and a 19-name hand list — the LISTS are correct (mechanically diffed: 180 names exist, complement is exactly 126 incl. AGGREGATE); fix the narrative counts.
- "176 produced a CORRECT 3-element array" overstates: the probe asserted `n == 3` and no throw, no values. Say so.
- Item 7's placement rationale names `If` among the arms the new arm must follow; in the real source `case If` sits AFTER `case BinaryOperation` in both Probe and TryBuildOperand. Harmless (IF is `Consumes`), but drop `If` from that sentence.
- Item 14's "count reads 304" — it reads **305** (EN and pt-BR); stale by one, not two.
- The eager-branch risk is overstated for the common case: measured `IFERROR(E6:E8, COST())` calls COST ONCE for 3 elements (a scalar fallback is a `ScalarOperand` evaluated at build); the cost exists only when the fallback is itself an ARRAY. Reword.
- `grep -c 'Elementwise<'` = 181 is fragile if item 3's doc comment writes `Elementwise<T>` literally; anchor as `grep -c '^\s*Elementwise<'`. The sibling `Entry<…>(` count is correct (307 today = 306 entries + factory → 127 after converting 180).
- Line numbers cited for the six TryStream consumers and the docs anchors are ~2 lines stale at HEAD 224a739 — anchor on names/headings.

### What the verifier actively confirmed correct (do not re-verify)

All 15 adversarially-sampled range-taking-non-aggregate functions are `Consumes` (MATCH, LOOKUP, XLOOKUP, COUNTBLANK, CHOOSE, TEXTJOIN, CONCAT, SUMSQ, NPV, IRR, AREAS, XMATCH, SUBTOTAL; FREQUENCY/TRANSPOSE/HYPERLINK/SEQUENCE not registered); optional-range functions conservatively `Consumes`; the fixture split and its goldens (7/6/3/2; merged 13/7; two-space E9 8/6); `=LEN(A1:A3)` bare stays `#VALUE!` (EvaluateCell never enters the mini-CSE — Phase 7's); per-element error/blank semantics (Blank survives the `ScratchLiteral` as Blank: `ISBLANK(blank)` = TRUE); volatiles drawn once per evaluation and Probe evaluation-free for the sampled shapes; reentrancy cut by the cycle guard; no `Expression` is ever a dictionary key anywhere (Excel loader's cache keys by string); `ScratchLiteral : ValueExpression` builds clean with no MemoryPack diagnostic; `RegistryEntry` constructed in exactly one place; DependencyExtractor/AnchoredFormulaSupport/FormulaWriter need no change; Phase 2's `TryStream` sites change nothing structural; the RED step is real (`SUM(ISNUMBER(E6:E8)*1)` = 0, `SUM(IFERROR(E6:E8,0))` = 0, `COUNT(LEN(D7:F9))` = 0 at HEAD).

## Implementation items

- [ ] **1.** Create `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs` with the acceptance pins ONLY (they must all fail first, with `#VALUE!` or the silent wrong number named below). Two DISTINCT local fixtures, because the brief's single fixture is self-contradictory (see *Why*): `Numeric()` — `E6:E8` = 1,2,3 and `F6:F8` = 10,20,30, nothing else; `Textual()` — `D7`="abc", `D8`="def", `E9`=`" "` (ONE space), every other cell of `D7:F9` empty. Pins, each `ExpressionParser.Parse("=…", sheet).Evaluate(workbook).AsObject()`: on `Numeric()` — `=SUMPRODUCT(--(E6:E8>1))` → 2.0, `=SUM(--(E6:E8>1))` → 2.0, `=SUM(-(E6:E8>1))` → -2.0, `=SUM(ABS(E6:E8*-1))` → 6.0, `=SUM(-E6:E8)` → -6.0, `=SUM(E6:E8%)` → 0.06, `=SUM(ISNUMBER(E6:E8)*1)` → 3.0 (today **0.0**, silent), `=SUM(IFERROR(E6:E8,0))` → 6.0 (today **0.0**, silent), `=SUM(+E6:E8)` → 6.0 (today 6.0 — the no-regression pin for unary `+`); on `Textual()` — `=SUM(LEN(D7:F9))` → 7.0, `=SUM(LEN(TRIM(D7:F9)))` → 6.0, `=SUM(IF(LEN(D7:F9)>0,1,0))` → 3.0, `=COUNT(LEN(D7:F9))` → 9.0 (today **0.0**, silent), and the production formula `=IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))>0,"Show","Hide")` → `"Show"`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs`
      *Why:* TDD is mandatory here (master plan, "Rules specific to this repo"). The two-fixture split is a
      **blocking correction to the brief**: its stated fixture sets `E6:E8` and `F6:F8` *and* says "rest of
      `D7:F9` blank", but `E7`, `E8`, `F7`, `F8` lie inside `D7:F9`. Measured through the prototype on the
      brief's merged fixture, `SUM(LEN(D7:F9))` = **13**, not 7 (3+3+0 for D, 1+1+1 for E, 2+2+0 for F), and
      `SUM(IF(LEN(D7:F9)>0,1,0))` = **7**, not 3. On the split `Textual()` fixture the prototype returns
      exactly **7, 6, 3 and 2** — the brief's Excel numbers. This is the same class of defect as Phase 2's
      correction M1: a golden measured on a fixture the plan does not state. Note also that the brief writes
      `E9="  "` (two spaces) while its own golden of 7 requires ONE space — measured: two spaces gives 8 for
      `SUM(LEN(...))` and still 6 for the TRIM form.
      Follow `MiniCseConsumerTests`'s style exactly (it is the nearest sibling): TUnit, `[Test] public async
      Task …`, `await Assert.That(Num(F("=…"))).IsEqualTo(2.0)` for numbers and
      `await Assert.That(F("=…")).IsEqualTo(ErrorValue.NotValue)` for errors, with a private static
      fixture-plus-evaluate helper per fixture.
      Head the file with the Microsoft citation in this suite's mandatory format (article title +
      support.microsoft.com GUID + fetch date), **fetched at implementation time** — the relevant pages are
      "Guidelines and examples of array formulas" and "Implicit intersection operator: @". No GUID is supplied
      here; do not invent one.

- [ ] **2.** Add the shape/edge pins to the same new file, still failing: `LEFT(D7:F9,E6:E8)` in a consumed position (3x3 against 3x1) → every element `#VALUE!`, so `=SUM(LEFT(D7:F9,E6:E8))` is `ErrorValue.NotValue`; `=SUM(ROUND(E6:E8,F6:F8))` → 6.0 (two arrays of the SAME shape pair elementwise); `=SUM(ROUND(E6:E8,0))` → 6.0 (a scalar argument broadcasts); blank elements — `=SUM(-(A1:A3))` over three empty cells → 0.0 (blank coerces to 0 under `-`) and `=SUM(LEN(A1:A3))` → 0.0; error elements — with `C1` = `=1/0`, `=SUM(ABS(1/(E6:E8-2)))` → `ErrorValue.DivByZero` (first error in scan order wins) and `=SUM(LEN(C1:C3))` → `ErrorValue.DivByZero`; text elements under a numeric lift — with `B2`="x", `=SUM(-B1:B3)` and `=SUM(ABS(B1:B3))` → `ErrorValue.NotValue`; missing sheet — `=SUM(LEN(Ghost!A1:A3))` → `ErrorValue.Reference`. Per-category coverage, one pin each: text `=SUM(LEN(D7:F9))` (item 1), math `=SUM(ABS(E6:E8*-1))` (item 1), information `=SUM(ISNUMBER(E6:E8)*1)` (item 1), logical `=SUM(IFERROR(C1:C3,7))` with `C1`=`=1/0`, `C2`=5, `C3`=2 → 14.0, date `=SUM(MONTH(E6:E8))` → 3.0.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs`
      *Why:* Every value here was measured through the prototype, not reasoned: the mismatch fill (`#VALUE!` at
      element 1), the blank→0 rule, the first-error-wins order, the text→`#VALUE!` rule, and — the one
      surprise — `LEN(Ghost!A1:A3)` yields **`#REF!` per element**, not `#VALUE!`, because `RangeOperand.At`
      reads through `Workbook.GetCellValueDense`, which owns the missing-sheet rule. That is Excel's answer
      and is *better* than today's `#VALUE!`; it also contrasts with the documented `SUM(ROW(Ghost!A1:A3))` =
      6 divergence, which survives because `PositionNumbersOperand` never touches a cell.
      The IFERROR fixture is deliberately `C1`=`=1/0`, `C2`=5, `C3`=2 and NOT the naive `C2`/`C3` blank: with
      blanks the lifted answer (7) and today's broken answer (7, from the broadcast scalar) **coincide**, and
      a fixture whose value equals the wrong answer is an assertion that cannot fail — the master plan's
      anti-vacuity rule. Measured both ways.

- [ ] **3.** Add `internal enum ArrayLifting : byte { Consumes, Elementwise }` to `Danfma.MySheet/Parsing/FunctionRegistry.cs` (immediately above `RegistryEntry`), add a seventh member `ArrayLifting Lifting` to `internal readonly record struct RegistryEntry` (`FunctionRegistry.cs:36-43`), and add a sibling factory next to `Entry<T>` at the bottom of the file (`FunctionRegistry.cs:2148-2165`): `private static RegistryEntry Elementwise<T>(string name, int minArgs, int maxArgs, Func<Expression[], Expression> create, Func<Function, Expression[]> getArguments) where T : Function => new(name, minArgs, maxArgs, create, typeof(T), getArguments, ArrayLifting.Elementwise);` and give `Entry<T>` the trailing `ArrayLifting.Consumes`. Extend the class doc comment (`FunctionRegistry.cs:15-21`) so "adding a function is one line here plus the node itself and its union tag" reads "… one line here — `Entry<T>` for a function that consumes ranges/arrays itself, `Elementwise<T>` for a pure-scalar one the mini-CSE may lift per element — plus the node itself and its union tag", and state the default-deny rationale in one sentence.
      *Files:* `Danfma.MySheet/Parsing/FunctionRegistry.cs`
      *Why:* `Consumes` must be the enum's zero value so `default(RegistryEntry)` and any future
      construction path are safe-by-default. Two factories rather than a sixth positional argument keeps each
      entry one line (the header's stated contract) and makes the classification greppable
      (`grep -c 'Elementwise<'`). `RegistryEntry` is constructed in exactly ONE place — `Entry<T>` — so
      widening the record struct touches nothing else; verified by grep (`new RegistryEntry(` appears
      nowhere).

- [ ] **4.** Convert `Entry<…>` → `Elementwise<…>` for these **180** entries and nothing else. **Text (31):** CHAR, CLEAN, CODE, DOLLAR, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, NUMBERVALUE, PROPER, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, T, TEXT, TEXTAFTER, TEXTBEFORE, TRIM, UNICHAR, UNICODE, UPPER, VALUE, VALUETOTEXT. **Mathematics (57):** ABS, ACOS, ACOSH, ACOT, ACOTH, ARABIC, ASIN, ASINH, ATAN, ATAN2, ATANH, BASE, CEILING, CEILING.MATH, CEILING.PRECISE, COMBIN, COMBINA, COS, COSH, COT, COTH, CSC, CSCH, DECIMAL, DEGREES, EVEN, EXP, FACT, FACTDOUBLE, FLOOR, FLOOR.MATH, FLOOR.PRECISE, INT, ISO.CEILING, LN, LOG, LOG10, MOD, MROUND, ODD, POWER, QUOTIENT, RADIANS, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, SEC, SECH, SIGN, SIN, SINH, SQRT, SQRTPI, TAN, TANH, TRUNC. **Financial (49):** ACCRINT, ACCRINTM, AMORDEGRC, AMORLINC, COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, CUMIPMT, CUMPRINC, DB, DDB, DISC, DOLLARDE, DOLLARFR, DURATION, EFFECT, FV, INTRATE, IPMT, ISPMT, MDURATION, NOMINAL, NPER, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD, PDURATION, PMT, PPMT, PRICE, PRICEDISC, PRICEMAT, PV, RATE, RECEIVED, RRI, SLN, SYD, TBILLEQ, TBILLPRICE, TBILLYIELD, VDB, YIELD, YIELDDISC, YIELDMAT. **Dates (19):** DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, SECOND, TIME, TIMEVALUE, WEEKDAY, WEEKNUM, YEAR, YEARFRAC. **Information (12):** ERROR.TYPE, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISTEXT, N. **Statistical (6):** FISHER, FISHERINV, PERMUT, PERMUTATIONA, PHI, STANDARDIZE. **Logical (5):** IFERROR, IFNA, IFS, NOT, SWITCH. **Lookup (1):** ADDRESS.
      *Files:* `Danfma.MySheet/Parsing/FunctionRegistry.cs`
      *Why:* This is the union of "not flagged by the body-marker grep" and "not flagged by the executable
      range-awareness oracle", minus the 18 hand-added exclusions in item 5. All 180 were driven through the
      prototype with a range in argument slot 0: **176 produced a correct 3-element array**, and the 4
      apparent failures (ODDFPRICE/ODDFYIELD/ODDLPRICE/ODDLYIELD) were the probe capping filler arity at 6
      against their 7-8 argument minimum — not a mechanism failure. Four entries deserve their reasoning on
      the record because they are not obviously scalar: **`N`** and **`T`** are value converters, scalar in
      Excel and lifted in array context (`N(A1:A3)` = `{1;2;3}`); **`IFERROR`/`IFNA`** are lifted by Excel —
      `SUMPRODUCT(IFERROR(A1:A3/B1:B3,0))` is the canonical idiom — and their short-circuit is preserved
      per-element (the node's own `Evaluate` still runs per element), only the *operands* are built eagerly,
      which is precisely what `TryBuildIf` already does for IF's branches; **`ERROR.TYPE`** returns a number
      per error and is scalar. `IFS`/`SWITCH` are lifted on the same grounds as IFERROR — see the open
      question about their eager branch operands.

- [ ] **5.** Leave these **126** entries as `Entry<…>` (`Consumes`). Automatic signals already cover 108 of them; the following **18 are hand-added because both signals miss them** and each MUST stay `Consumes`: the two-population statistics **CORREL, COVARIANCE.P, COVARIANCE.S, PEARSON, RSQ, SLOPE, STEYX, FORECAST.LINEAR** and their compatibility aliases **COVAR, FORECAST** (the oracle's filler argument errored on both sides, hiding the difference); **PERCENTILE.EXC** and **TRIMMEAN** (an out-of-range filler `k` made both sides `#NUM!`); **MAXIFS, MINIFS** (criteria family); **SUMX2MY2, SUMX2PY2, SUMXMY2** (paired arrays); **`IF`** (it already owns a dedicated `IfOperand` arm and must never reach the generic one); **`RANDBETWEEN`** (volatile with two scalar arguments — lifting it would draw per element; see the open question). Also confirm the eight zero-argument entries (PI, TRUE, FALSE, NA, SHEETS, NOW, TODAY, RAND) stay `Entry` — with no arguments they can never lift, so the flag is moot, and `Consumes` states that. Finally, **AGGREGATE** (Phase 2) is `Consumes`.
      *Files:* `Danfma.MySheet/Parsing/FunctionRegistry.cs`
      *Why:* The oracle's false negatives are systematic and must be named or they will be silently
      mis-flagged by a future contributor: a function whose answer depends only on a rectangle's SHAPE or
      POSITION (`ROWS`, `COLUMNS`, `AREAS`, `ROW`), or that only inspects a reference (`ISREF`, `TYPE`,
      `ISFORMULA`, `SHEET`, `FORMULATEXT`, `OFFSET`, `INDIRECT`), answers identically for two different
      rectangles and so looks scalar-blind. Those eleven are caught by the body-marker grep; the eighteen
      above are caught by neither, which is exactly why this phase ships an explicit list rather than a
      generated one. `IF` is the load-bearing one: the generic arm sits AFTER `case If ifNode` in both
      switches, so a mis-flag would be masked in `Probe`/`TryBuildOperand` but would still change
      `Index.TryResolveReference`'s eligibility probe.

- [ ] **6.** Add three types to `Danfma.MySheet/Expressions/ArrayOperands.cs`, after `IfOperand`. (a) `internal sealed record ScratchLiteral : ValueExpression { public ComputedValue Value; public override ComputedValue Evaluate(EvaluationContext context) => Value; }` — with a comment stating that it is the ONE mutable node in the engine, that it never escapes `LiftedFunctionOperand`, that it carries NO `[MemoryPackable]` attribute and NO union tag, and why a `ComputedValue`-carrying literal is required rather than the typed `NumberValue`/`StringValue`/… nodes (there is no literal node for `ComputedValueKind.Reference`). (b) `internal sealed class UnaryOperand : ArrayOperand` mirroring `BinaryOperand` field-for-field: `(UnaryOperator @operator, ArrayOperand inner, int rows, int columns)`, the same `_rows != rows || _columns != columns → #VALUE!` mismatch guard, then `inner.At(index, _rows, _columns)` → `CoerceToNumber` (propagating its `Error?`) → `Negate` = `-number`, `Percent` = `number / 100`, and an unreachable `_ => ComputedValue.Error(Error.Value)` — the exact ladder of `UnaryOperation.Evaluate`'s non-Plus path (`UnaryOperation.cs:29-42`). (c) `internal sealed class LiftedFunctionOperand : ArrayOperand` holding `Expression _node`, `ScratchLiteral[] _scratch`, `ArrayOperand[] _arguments`, `EvaluationContext _context`, `_rows`, `_columns`; its constructor takes `Func<Expression[], Expression> create` plus the operand array, fills a fresh `Expression[]` with fresh `ScratchLiteral`s and calls `create(slots)` ONCE; `At` runs the same mismatch guard, then `for (var j = 0; j < _arguments.Length; j++) _scratch[j].Value = _arguments[j].At(index, _rows, _columns);` and returns `_node.Evaluate(_context)`.
      *Files:* `Danfma.MySheet/Expressions/ArrayOperands.cs`
      *Why:* `ArrayOperands.cs` is the file Phase 1's summary created precisely so Phase 2 and 3 could add
      operands without growing `ArrayEvaluation.cs` past 650 lines. The reuse-the-node design is what buys
      the measured **0 B/element vs 80 B/element (1 arg) / 112 B/element (2 args)** and the 1.7x-2.8x speedup;
      it is safe because nothing in the engine hashes, compares, serializes or stores an `Expression`
      (grepped: no `Dictionary<Expression`, no `HashSet<Expression>`, no `ConditionalWeakTable`), and because
      the `Elementwise` classification is exactly the guarantee that a lifted node only calls `Evaluate` on
      its arguments. `UnaryOperand` must NOT handle `Plus`: `UnaryOperation.Evaluate` routes `Plus` through
      `CaptureValue` so the operand comes back type-preserved and a range comes back as a REFERENCE value —
      `SUM(+A1:A3)` = 6 today, and it must stay 6.

- [ ] **7.** Add the two `Probe` arms in `ArrayEvaluation.cs`, placed **after** `case Row {…}` / `case Column {…}` and **before** `case BinaryOperation binary:` (`:180`). (a) `case UnaryOperation { Operator: not UnaryOperator.Plus } unary: return Probe(unary.Operand, context);` — a `Negate`/`Percent` is an array exactly when its operand is, and it refuses exactly when its operand refuses. (b) `case Function function when TryGetLift(function, out var arguments):` — loop the arguments, `Probe` each, return `(false, false)` on the first that does not succeed, and return `(true, anyArgumentIsArray)`. Add the private helper `TryGetLift(Function function, out Expression[] arguments)`: `FunctionRegistry.ByType.TryGetValue(function.GetType(), out var entry) && entry.Lifting is ArrayLifting.Elementwise` → `arguments = entry.GetArguments(function); return arguments.Length > 0;` else `arguments = []; return false;`. Extend `ArrayEvaluation`'s class doc comment (:29-50) and `IsArrayEligible`'s doc (:112-134) to name the two new shapes.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* Order matters twice. The `UnaryOperation` arm must exclude `Plus` by pattern (not by an `if`
      inside), so `+range` still reaches `default:` and stays the opaque scalar that carries the reference.
      The `Function` arm must sit after the `Row`/`Column`/`If` arms so those keep their dedicated
      handling — `Row`/`Column`/`If` are classified `Consumes` anyway, which makes the ordering
      belt-and-braces rather than load-bearing, but a `when` guard that silently shadowed `If` would be a
      hard bug to find. `arguments.Length > 0` is what keeps a zero-argument entry (PI, RAND, …) out of the
      arm without a special case. **Measured: this arm evaluates nothing** — `IsArrayEligible` over
      `ROUND(E6:E8,TICK())`, `IFERROR(TICK(),E6:E8)` and `LEN(TRIM(D7:F9))` calls the counting function zero
      times.

- [ ] **8.** Add the mirrored `TryBuildOperand` arms in the same order (after `case Column {…}`, before `case BinaryOperation binary:` at `:299`), plus two builders and an N-ary shape fold. `case UnaryOperation { Operator: not UnaryOperator.Plus } unary: return TryBuildUnary(unary, context, out operand);` — build the inner operand, return `false` if it refuses, `new ScalarOperand(unary.Evaluate(context))` if it is not an array (the whole unary is then an opaque scalar, evaluated once), else `new UnaryOperand(unary.Operator, inner, inner.Rows, inner.Columns)`. `case Function function when TryGetLift(function, out var liftArguments): return TryBuildLift(function, liftArguments, context, out operand);` — build every argument operand in order, return `false` on the first refusal, track whether ANY is an array and fold the result shape as `ResultShape` already does (first array's shape; then per-axis `Math.Max` against each further array — identical to the two-operand rule at `:571-586` for two operands, so extract a `Fold(ref rows, ref columns, ArrayOperand)` helper or add an overload rather than writing a second rule); if no argument is an array, `new ScalarOperand(function.Evaluate(context))` (today's behaviour, unchanged, evaluated once); otherwise `new LiftedFunctionOperand(entry.Create, operands, context, rows, columns)`, re-reading the entry via `FunctionRegistry.ByType`.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* The `!IsArray → ScalarOperand(node.Evaluate(context))` fallback in both builders is what
      preserves the "eligible iff the build succeeds as an array, and the build is the SINGLE evaluation"
      contract (`ArrayEvaluation.cs:112-121`) — `Probe` returned `(true, false)` for that shape, so the build
      must succeed as a non-array. The per-axis-max mismatch shape is what makes
      `LEFT(D7:F9,E6:E8)` fill with `#VALUE!` through `LiftedFunctionOperand`'s own guard, matching
      `BinaryOperand` exactly — **measured through the prototype: `#VALUE!` at element 1.** Both new arms
      were exercised end-to-end in the prototype (a shadow copy of this file at
      `/tmp/mysheet-lift-probe/Shadow.cs`) and returned every acceptance value in items 1-2.

- [ ] **9.** Add a classification guard test to `ElementwiseLiftingTests.cs` that enumerates `FunctionRegistry.Entries` (internals are visible to the test assembly): for every entry whose `Lifting` is `ArrayLifting.Elementwise`, assert (a) `MaxArgs > 0`, and (b) that the function is scalar-blind on the ORDINARY scalar path — build `F(A1:A3, 1, 1, …)` and `F(C4:C6, 1, 1, …)` with `MinArgs` arguments on a fixture where `A1:A3` = 1,2,3 and `C4:C6` = 100,"x",blank, evaluate both through `ExpressionParser.ParseFormulaBody(...).Evaluate(workbook)` and assert the two results are equal. Add a second test that asserts `SUM`, `SUMPRODUCT`, `INDEX`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `AREAS`, `OFFSET`, `INDIRECT`, `ISREF`, `TYPE`, `MATCH`, `VLOOKUP`, `SUBTOTAL`, `AGGREGATE`, `IF`, `LET`, `RANDBETWEEN` and the whole `*IF`/`*IFS` family each report `ArrayLifting.Consumes`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs`
      *Why:* This is the regression guard for the next contributor. A range-aware function mistakenly flagged
      `Elementwise` produces *silent wrong numbers*, the worst failure mode this codebase has (see the three
      measured silent cases in the Design decision). Test (b) is the executable form of the oracle that
      produced the list, so it re-runs the derivation on every build instead of trusting a frozen list; it
      catches any newly added function that reads its range argument. It does NOT catch the shape/position
      family (`ROWS`-like), which is why the explicit assertion list in the second test exists.

- [ ] **10.** Add the consumer-level cases to `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`, next to the existing `Sum_*`/`Small_*` blocks: `=SMALL(LEN(D7:F9),1)`, `=INDEX(LEN(D7:F9),1)` and `=AGGREGATE(15,6,LEN(TRIM(D7:F9)),1)` over the `Textual()` fixture, plus `=SUMIFS(LEN(A1:A3),A1:A3,">0")` asserting the documented **`#VALUE!`** (the criteria family still refuses computed arrays — `CriteriaScan.Open` is deliberately not `OpenArrayOrRange`). Add the eager-vector twin to `tests/Danfma.MySheet.Tests/Expressions/ArrayEvaluationTests.cs`: `ArrayEvaluation.TryEvaluate` over a hand-built `Len([RangeReference])` and over `new UnaryOperation(UnaryOperator.Negate, range)`, asserting `Rows`/`Columns` and each element, so the eager and lazy shapes are pinned to agree.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`, `tests/Danfma.MySheet.Tests/Expressions/ArrayEvaluationTests.cs`
      *Why:* The mini-CSE has SIX consumers (`NumericAggregation.cs:114`, `OrderSelection.cs:73`,
      `CriteriaScan.OpenArrayOrRange:180`, `Index.cs:24` and `:173`, `AggregateCodes.cs:45`,
      `Aggregate.cs:134`) and this phase changes what every one of them accepts. `ArrayEvaluationTests` is the
      only place that drives the EAGER `TryEvaluate` path; the class comment promises the two shapes are
      bit-for-bit identical, and nothing would notice if `LiftedFunctionOperand` broke that (the eager path
      calls the same `At`, but it calls it in a fresh loop — a scratch-slot bug that depended on call order
      would show up here first).

- [ ] **11.** Pin `Index.TryResolveReference`'s widened rejection. `Index.cs:167-178` returns `false` — falling back to normal evaluation — when `Arguments[0] is not Reference && ArrayEvaluation.IsArrayEligible(Arguments[0], context)`. After this phase more shapes are eligible, so `INDEX(LEN(A1:A3), 2)` now takes the ARRAY path where it previously took the reference path and failed. Add to `ElementwiseLiftingTests.cs`: `=INDEX(LEN(D7:F9),1)` → 3.0 and `=ROWS(INDEX(LEN(D7:F9),1))` → 1.0 (a scalar counts as 1x1, per Phase 1 item 2). Do NOT try to make `INDEX` return a reference for a lifted argument — it has no cell address.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/ElementwiseLiftingTests.cs`
      *Why:* This is the one integration point where widening eligibility changes an existing decision rather
      than adding a new one, and `ArrayEvaluation`'s own remark (:129-133) already flags
      `Index.TryResolveReference` as the call site that "probes only to REJECT the array forms and never
      builds". Widening the eligible set widens the rejection; the pin makes that intentional and visible.

- [ ] **12.** Pin the four integration points that must NOT change, with one assertion each, and record why. (a) **Cell boundary**: `=LEN(A1:A3)` typed bare in a cell still yields `ErrorValue.NotValue`, and `=A1:A3` in the row-intersecting cell still yields the intersected value — add to `tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs` with a comment naming Phase 7 as the owner of the array half. (b) **`DependencyExtractor`**: a cell holding `=SUM(LEN(A1:A3))` reports `A1:A3` as a dependency — its `case Function function:` arm (`DirtyGraph/DependencyExtractor.cs:207-213`) already visits arguments through `FormulaWriter.Call`, so no code change; add the assertion to the existing dirty-graph suite. (c) **`AnchoredFormulaSupport.IsFullyAnchored`**: a shared-formula master `=SUM(LEN($A$1:$A$3))` still reports fully-anchored — its `Function function =>` arm (`Parsing/AnchoredFormulaSupport.cs:63-65`) accepts a function when every argument is, again through `FormulaWriter.Call`. (d) **`FormulaWriter`**: `=SUM(LEN(A1:A3))` and `=SUM(-(A1:A3>1))` round-trip parse→write unchanged — add both strings to the flat identity list in `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs`, `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`, the dirty-graph test class
      *Why:* All four were verified by reading, and three need no production change at all: the writer, the
      dependency extractor and the anchored-formula check reach a function's arguments through the SAME
      registry accessor (`FormulaWriter.Call`, `FunctionRegistry.cs`), and this phase adds no node type, so
      there is nothing new for any of them to learn. (a) is the one that is a deliberate GAP rather than a
      non-event: `Workbook.EvaluateCell` (`Workbook.cs:355-365`) only ever inspects
      `value.TryGetReference(...)`, and a lifted function is not a reference — it is `Len.Evaluate`'s
      untouched `#VALUE!`. Pinning it now means Phase 7 must change the assertion deliberately when it adds
      the S4 array half, instead of discovering the behaviour by accident.

- [ ] **13.** Update `docs/workbook-and-expressions.md` §"Implicit array arguments" (`:386-446`) and its pt-BR twin §"Argumentos implícitos de array" (`docs/pt-BR/workbook-and-expressions.md:401-…`). In **Supported** (`:400-409`): after "or `ROW`/`COLUMN` over a rectangle", add the two new producers — a **unary `-`/`%`** over an array (`SUM(-(A1:A3>1))` = -2), and **any pure-scalar built-in** whose arguments include an array, lifted element by element with scalars broadcast (`SUM(LEN(A1:A3))`, `SUM(ROUND(A1:A3,2))`, `SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))`), stating that mismatched shapes fill with `#VALUE!` exactly as an operator does, and naming the count (180 of 306 built-ins are liftable; the range-aware rest — `SUM`, `COUNT`, `INDEX`, `ROW`, `VLOOKUP`, `OFFSET`, `INDIRECT`, `SUMPRODUCT`, `SUBTOTAL`, `AGGREGATE`, the criteria family — are never lifted because they already consume arrays themselves). In **Not supported** (`:411-439`): the "dry cell" bullet at `:413-418` keeps its `#VALUE!` verdict but must now name the new case explicitly — bare `=LEN(A1:A3)` in a cell is still `#VALUE!`, because lifting happens inside the CONSUMERS and the cell boundary is not one of them; add a bullet stating that unary `+` is deliberately NOT lifted (it stays the reference-preserving no-op, so `SUM(+A1:A3)` = 6 through the range path) and that `-(+A1:A3)` is therefore still `#VALUE!`; keep the criteria-family bullet as is but add that a LIFTED argument is refused there for the same reason (`SUMIFS(LEN(A1:A3),…)` → `#VALUE!`). Every number cited must be pinned by a committed test (master plan rule).
      *Files:* `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`
      *Why:* Both the "Supported" and "Not supported" lists become factually wrong the moment item 8 lands —
      the section currently states the eligible set as exactly five shapes and states the dry-cell rule in a
      way that a reader will now read as covering `=LEN(A1:A3)` by accident rather than by design. `docs/pt-BR/`
      is a full 11-file mirror and every doc edit needs its twin (master plan rule).

- [ ] **14.** Add ONE note to `docs/function-reference.md` and its pt-BR twin, not per-function rows: a short paragraph in the introduction, right after the `MySheet implements **NNN built-in functions**` sentence (`docs/function-reference.md:3`, `docs/pt-BR/function-reference.md:5`), saying that the per-function rows are unchanged but that a pure-scalar function given a range in an array-consuming position is applied element by element, with a link to `workbook-and-expressions.md#implicit-array-arguments`. Do NOT add a per-function marker to 180 rows, and do NOT touch the function COUNT — this phase adds and removes no function. (Note in passing: that count reads **304** while the registry holds **306** entries as of `224a739`; it is stale by two and fixing it belongs to whichever phase added them, not here.)
      *Files:* `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* A per-function note would be 180 rows of churn that restates one rule, and it would go stale the
      moment a function is added — exactly the duplication the project's principles forbid. The single
      cross-reference is the honest form. The header counts and the ✅/⬜ lists stay untouched because no
      function is added or removed by this phase.

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded, **0 Warning(s)**, 0 Error(s). Watch specifically for a MemoryPack source-
      generator diagnostic on `ScratchLiteral` (a new `Expression` subtype with no `[MemoryPackable]` and no
      union tag). None is expected — `ValueExpression` itself carries no `[MemoryPackable]` attribute and the
      union list lives on `Expression` — but a `MEMPACK…` warning here is the one build-time signal that the
      mutable-literal design needs the fallback in the risks section.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier format . && dotnet csharpier check .`
      → expected: exit 0, no file listed. CSharpier 1.3.0 is pinned in `dotnet-tools.json` and both git hooks
      run `check`; converting 180 registry entries touches enough lines that formatting will drift.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/ElementwiseLiftingTests/*"`
      → expected: "Test run summary: Passed!", failed: 0. Before item 6-8 land, this same command must FAIL
      with `#VALUE!` on the eight error pins and with `0` on the three silent pins (`SUM(ISNUMBER(E6:E8)*1)`,
      `SUM(IFERROR(E6:E8,0))`, `COUNT(LEN(D7:F9))`) — confirm those exact failure reasons before implementing.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/MiniCseConsumerTests/*"`
      → expected: "Passed!", failed: 0. Every pre-existing case must still pass — in particular
      `Small_OfIfArray_ErrorAfterKthElement_StillPropagates` and
      `Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap`, which pin behaviours the new arms sit
      next to.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/ArrayEvaluationTests/*"`
      → expected: "Passed!", failed: 0, total greater than today's. Proves the eager `TryEvaluate` vector and
      the lazy `ArrayStream` still agree element-for-element through `LiftedFunctionOperand`.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/UnaryOperationTests/*"`
      → expected: "Passed!", failed: 0. `UnaryOperand` must not have changed the scalar path: unary `+` on
      text/boolean/error (issue #8) and `-`/`%` coercion are all pinned here.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/CellBoundaryIntersectionTests/*"` and `… "/*/*/FormulaWriterTests/*"`
      → expected: both "Passed!", failed: 0. The boundary suite proves bare `=LEN(A1:A3)` is still `#VALUE!`
      (the deliberate gap, Phase 7's to close); the writer suite proves no round-trip moved.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release --no-build`
      → expected: both "Passed!", failed: 0 — the exact pair the pre-push hook and `.github/workflows/ci.yml`
      run. Core baseline at the start of this phase is whatever Phase 2 left (Phase 1 ended at 1280); this
      phase only ADDS cases. Confirms in particular that no serialization golden moved — `ScratchLiteral` must
      never reach the wire.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -c 'Elementwise<' Danfma.MySheet/Parsing/FunctionRegistry.cs && grep -o 'Entry<[A-Za-z0-9_.]*>(' Danfma.MySheet/Parsing/FunctionRegistry.cs | wc -l`
      → expected: **181** (180 entries + the `Elementwise<T>` factory declaration) and **126 + 1** remaining
      `Entry<…>(` occurrences (125 consumers + AGGREGATE + the `Entry<T>` factory declaration). If the second
      number is not `total − 180`, an entry was converted that item 4 does not name, or one it names was
      missed. Re-derive rather than adjust: the classification is the risk, not the count.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -c 'MemoryPackUnion(' Danfma.MySheet/Expressions/Expression.cs && grep -n 'ScratchLiteral' Danfma.MySheet/Expressions/Expression.cs`
      → expected: the count UNCHANGED from before this phase (323 with Phase 2 landed), and **no output** for
      the second grep. This phase adds no node type and must not add a union tag; a `ScratchLiteral` line in
      `Expression.cs` means someone registered a node that must never be serialized.
- [ ] Allocation re-measurement (the number that chose the mechanism, re-confirmed on the real build): a
      50 000-row column of text, `SUM(LEN(A1:A50000))`, evaluated 10 times.
      → expected: **~5 ms per evaluation and ~0 bytes per element** (the prototype measured 5.01 ms and 308
      bytes TOTAL per evaluation — the operand tree, built once). If the per-element allocation is nonzero,
      the node is being rebuilt per element and mechanism (a) was implemented by mistake; that costs 3.81 MB
      per evaluation and 1.7x the time.

## Risks carried by this phase

- **A mis-flagged range-aware function is a SILENT wrong number, not an error.** This is the phase's defining risk and it is already demonstrable in the current code: `SUM(ISNUMBER(E6:E8)*1)` = 0, `SUM(IFERROR(E6:E8,0))` = 0 and `COUNT(LEN(D7:F9))` = 0 today (Excel: 3, 6, 9) — all measured, all error-free, all wrong. Flag a consumer as `Elementwise` and you get the same class of failure in the other direction: `SUM` lifted over a range would return the top-left cell repeated. The default-deny flag, item 9's guard test and item 5's named exclusion list are the three mitigations; none of them is complete on its own.
- **`ScratchLiteral` is a mutable node in an engine whose stated contract is immutable expression trees.** It is safe only as long as no lifted function's `Evaluate` stores or defers its argument node. That holds today by classification (a pure-scalar function only calls `Evaluate` on its arguments) and by grep (nothing hashes, compares, serializes or caches an `Expression`), but it is a property a future function could violate silently — the symptom would be a stale value from a previous element, which no test asserts against directly. Fallback if it ever bites: mechanism (c-lite), an immutable `ComputedValueLiteral` rebuilt per element, at 4.58 MB and 1.5x the time for 50k elements.
- **Eager branch operands in lifted `IFERROR`/`IFNA`/`IFS`/`SWITCH`.** `LiftedFunctionOperand.At` must fill EVERY scratch slot before calling `Evaluate`, so the fallback argument's operand is evaluated at every element even where the primary argument did not error. This changes no VALUE (volatile scalars are still `ScalarOperand`s drawn once at build time — measured, one draw) and no error propagation (an unused element is written to a slot and ignored), but it does change COST: `IFERROR(A1:A50000, <expensive array>)` computes the expensive side 50 000 times. `IfOperand` does better (it calls `At` only on the taken branch); the lifted path cannot, because it hands the node a filled argument list. Documented, not fixed.
- **A defined name in a lifted argument is still NOT array-eligible** — `SUM(LEN(MyName))` stays `#VALUE!` where `SUM(LEN(A1:A3))` becomes 7 (measured both). This is the pre-existing gap the master plan already records ("`COUNT((Rng<>"")*1)` = 1 where the literal range gives 3") and Phase 3 closes for table references; this phase makes the inconsistency MORE visible by fixing the literal-range half of many more formulas. Do not close it here.
- **Open/whole-column arguments are refused, so a lifted function over one stays broken.** `SUM(LEN(A:A))` remains `#VALUE!` (measured: `LEN(E:E)` probes as not-array and evaluates to the scalar `#VALUE!`). That is the existing cost guard, correctly inherited — but a user who fixes `SUM(LEN(A1:A3))` and then drags it to a whole column gets the old failure back with no explanation. The docs bullet in item 13 must say so.
- **`FunctionRegistry.cs` is edited by Phase 3 as well** (table-related entries), and Phase 7 adds FILTER/SORT/UNIQUE/SEQUENCE there. Converting 180 lines in the same file will conflict with either if they land concurrently. Phase 2's AGGREGATE entry is already in (`7d968a3`); rebase rather than merge against the others.
- **The `Probe` hot path gains a `FrozenDictionary<Type>` lookup for every function node.** Measured at 16.5 ns per lookup versus 0.32 ns for a type test — a 50x relative regression on that one operation. It is per NODE per EVALUATION (never per element) and the same lookup already runs in three other places per node, so the absolute cost is negligible; but a future change that calls `IsArrayEligible` per element would turn it into a real cost. `Index.TryResolveReference` (`Index.cs:173`) is the one call site that probes and then throws the answer away.

## Open questions owned by this phase

- **Does Excel draw a lifted volatile once or per element?** MySheet will draw ONCE for a volatile SCALAR argument of a lifted function (`ROUND(A1:A3, RAND())` — one draw, broadcast; measured with a counting function, and identical to today's `SUM(IF(A1:A3>0,TICK(),0))` = 1 call), and `RANDBETWEEN` is classified `Consumes` precisely so the question does not arise for a volatile that could itself be lifted. Excel's array semantics suggest per-element for `RANDBETWEEN(A1:A3, 10)`. Settle with a real .xlsx: `A1:A3` = 1,2,3, `B1` = `=SUMPRODUCT(RANDBETWEEN(A1:A3,1000))` recalculated a few times — three independent draws sum very differently from 3x one draw. If Excel draws per element, `RANDBETWEEN` moves to `Elementwise` and the "scalar operands are evaluated once" rule in `ArrayEvaluation`'s class comment needs an exception clause.
- **Should unary `+` be transparent in an array position? ANSWERED BY PHASE 11C: yes, and the blanks/text worry below never arises.** This plan kept it opaque (the brief's instruction and the safe choice at the time): `SUM(+A1:A3)` = 6 then, through the reference value path. The cost was that `SUM(-(+A1:A3))` stayed `#VALUE!` where Excel gives -6 — measured both then and through the prototype. Making `Plus` recurse into its operand for array purposes looked like a two-line change that would move `SUM(+A1:A3)` from the RANGE path to the ARRAY path, and those two paths do not agree on blanks and text (measured: `SUM(B1:B3)` = 1 with `B2`="x" and `B3` blank; whether the array path folds the same needed its own pin). Phase 11c made `+` transparent to the probe — `ArrayEvaluation.Probe`/`TryBuildOperand` gained a `UnaryOperation{Plus}` arm that sees through to the operand — but kept the top-level gate (`IsBareReferenceNode`'s own `Plus` arm) answering TRUE for a `+` over a bare reference, so `SUM(+A1:A3)` never leaves the RANGE path: it still reads the cells, blanks and text included, exactly as `SUM(A1:A3)` does. The worry above is about a `+` at a consumer's TOP level moving to the array path; what actually happens is narrower — the probe now looks THROUGH a `+` that is not itself the top-level argument (so a `-` or a lifted call around it, or an operator like `(+A1:A3)*B1:B3`, reaches what is inside), while a bare `+range` at the top stays on the range path where the blanks/text rule already agreed with it. `SUM(-(+A1:A3))` is now -6, matching Excel, and `SUM(+LEN(A1:A3))`/`SUM(LEN(+A1:A3))` are now 6 (`ArrayBindingTests`, `ElementwiseLiftingTests.LiftedCall_UnderATransparentUnaryPlus_IsLifted`). The one shape the plan worried about — `COUNTIF(+A1:A3,">0")` and its siblings taking the ARRAY path and losing the criteria family's reference semantics — was measured and fixed the same way: `IsBareReferenceNode` answers for the OPERAND of a `+` too, so `COUNTIF(+A1:A3,">0")` = 2, `SUMIF(+A1:A3,">0")` = 14 and `COUNTBLANK(+A1:A3)` = 0, all matching the oracle and pinned in `ArrayBindingTests.AUnaryPlusOverABareReference_IsStillAReferenceAtTheTopLevel`; `AVERAGEIF(+A1:A3,">0")` = 7 and `ISREF(+A1:A3)` = TRUE were measured the same way (Aspose.Cells 26.6.0, 2026-09-11) but are not separately pinned.
- **Is `TYPE(A1:A3)` 64 (array) in an array context?** MySheet classifies `TYPE` as `Consumes` (its body reads a `Reference`), so `TYPE` over a range keeps today's answer. Excel documents `TYPE` returning 64 for an array. Not verified; low value, but the docs must not claim Excel parity for it.
- **Does Excel lift `CHOOSE` elementwise over an array `index_num`?** MySheet classifies `CHOOSE` as `Consumes` because it returns a `Reference` through `TryResolveReference` (verified by the oracle: it is range-aware today). Excel's `CHOOSE({1;2;1}, "a", "b")` does lift. Lifting it here would break the reference-returning half. Left as-is; needs an Excel probe before anyone changes it.
- **What does Excel do with `IFS`/`SWITCH` over an array condition — lift, or refuse?** This plan lifts them on the same grounds as `IFERROR`, but unlike `IFERROR` they are variadic and short-circuiting, and the eager-operand rule (see risks) makes their unused branches run at every element. If Excel refuses them, move both to `Consumes`; the change is two registry lines. Not verified.
- **`SUM(LEN(Ghost!A1:A3))` = `#REF!` while `SUM(ROW(Ghost!A1:A3))` = 6.** Both measured. The first is Excel-correct (the lifted path reads cells through `Workbook.GetCellValueDense`, which owns the missing-sheet rule); the second is the divergence Phase 1 documented and left open, because `PositionNumbersOperand` never touches a cell. This phase makes the inconsistency between two adjacent array shapes visible in the docs for the first time. Whether to close it (a guard on the literal-rectangle fast path, on the hot path) was already deferred to Phase 3, which lands on the same arms.
- **What should bare `=LEN(A1:A3)` in a cell return?** This phase leaves it `#VALUE!` (nothing changes; the mini-CSE is not entered at the boundary). S4's array half says TOP-LEFT, i.e. `LEN(A1)`, and Excel's legacy `@` behaviour agrees. Phase 7 owns the change and its own open question already asks top-left vs `#VALUE!` for `FILTER`/`SEQUENCE`; whatever it decides must cover lifted functions too, since after this phase they are the most COMMON computed array a user can type bare.

## Phase Summary

**Status: Complete** — branch `feat/elementwise-lifting`, commits `3d030ae..dc29d30` (26 commits: 13 from
Tasks 1-5 and their review fix rounds, 5 from the consolidated wave after the three-part final review, 8 the
controller's `docs(plans)`/`docs(lessons)` commits about Phases 9-11 and the standing Aspose rule), pending
fast-forward merge to `main`. Gates at the head: csharpier clean (355 files), Release build 0 warnings, core
**1490/1490** (1349 → +141 cases), Excel **90/90**, no MemoryPack union tag added (`ScratchLiteral` is
deliberately not serializable), no public API removed, function counts unchanged at **306**.

**What it fixes.** The mini-CSE recognised only five array-producing shapes, so a unary operator or any of the
~180 pure-scalar built-ins over a range fell to the opaque-scalar arm: evaluated once over a range it yielded
`#VALUE!`, which then poisoned every element — or, worse, was silently CONSUMED by the enclosing node. The
user's production formula `=IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$22))>0))>0,"Show","Hide")` returned `#VALUE!`
where the oracle returns `Show`, and 43 formulas in one workbook used that shape. Three answers were silently
wrong rather than erroring: `SUM(ISNUMBER(E6:E8)*1)` = 0, `SUM(IFERROR(E6:E8,0))` = 0, `COUNT(LEN(D7:F9))` = 0.
All ten rows of the report now match Aspose.Cells 26.6.0's CSE column and are pinned.

**How.** An explicit `ArrayLifting { Consumes, Elementwise }` flag on `RegistryEntry`, defaulting to `Consumes`
and opted into by a sibling factory `Elementwise<T>` so a lifted entry stays one line: 180 lifted, 126
consuming. `LiftedFunctionOperand` builds the node ONCE through `entry.Create` over mutable `ScratchLiteral`
slots and rebinds them per element — 0.009 B/element measured against 3.81 MB/evaluation for per-element node
construction — with `UnaryOperand` mirroring it for `Negate` and `Percent` only. Unary `+` is deliberately not
lifted: it is Excel's reference-preserving no-op and already worked.

**Decisions that shaped it, each measured.** An omitted argument slot keeps the ORIGINAL `BlankValue` node, or
a function silently loses its own default (`FIXED(A1:A3,,TRUE)`). An argument the mini-CSE cannot build makes
the whole call an opaque scalar rather than failing the enclosing formula, so `SUM(IF(A1:A3>0,1,LEN(B:B)))`
stays 3. `ProbeLift` probes before building, because building first double-evaluated a scalar nested inside a
lifted node. Element KIND is preserved, not just the number: `ISBLANK`, `ISTEXT`, `ISNUMBER`, `N`, `T` and
`ERROR.TYPE` over a range with a blank match the oracle's CSE column exactly.

**The guard is the deliverable as much as the mechanism.** A range-aware function wrongly flagged `Elementwise`
returns silently wrong numbers, so the classification is pinned by NAME against a committed roster (a count is
not enough) and by a probe that sweeps EVERY argument position with number, text, boolean and range fillers
across arities. The final review earned this: the first guard walked only the `Consumes` half, and a
contributor-style mutation — wrong flag, deleted name row, bumped counts, every edit one makes on purpose when
adding a function — shipped a wrongly-lifted `HLOOKUP` with the suite green at 1485/0. After the fix that
mutation fails naming HLOOKUP twice, and even adding the name to the roster (the "fix the failing test" move)
still fails naming it, because the widened sweep tries the function in its later slots. 21 entries remain blind
to the sweep by nature and every one of them is named in a test; the closing assertion is that the
blind-and-unnamed set is EMPTY.

**Left open deliberately, all pinned as divergences with the oracle's answer in the comment.** The equal-shape
rule stays: Aspose BROADCASTS a vector across a rectangle (`SUM(A1:C3*E1:E3)` = 108) — Phase 10. Bare
`=LEN(A1:A3)` in a cell stays `#VALUE!`: the mini-CSE is entered only by consumers that ask for it, and the
cell boundary's array half is Phase 7's, which must implement Aspose's per-operand intersection rather than
S4's top-left. The defined-name gap now covers every lifted shape (`SUM(LEN(MyName))` = `#VALUE!` here, 6 on
the oracle, and the comparison shapes answer a SILENT 1 against 2) — Phase 11. Also Phase 11: unary `+` inside
a lift, per-slot lifting of the consuming family, `INDEX(<computed array>,0)`, and `FIXED`/`DOLLAR` default
decimals. Sixteen `Elementwise` flags are inferred from the node bodies rather than measured, because Aspose
answers `#NAME?` for them; they are labelled as such.

**Three-part final review.** Fable "Yes with fixes" (2 Important, 3 Minor — it found the guard hole), GLM-5.3
"Yes" (2 Minor, same items by its own derivation), Copilot "Yes" (no findings, but plan mode blocked its
mutations so its section C is not independent). Consolidated at
`.superpowers/sdd/phase-8-elementwise-lifting/final-review/CONSOLIDATED.md`; the wave fixed all five findings
and the scoped re-review came back clean, having re-run the exploit and its "fix the test" variant.
