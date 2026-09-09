# Phase 2: AGGREGATE(function_num, options, ref1, [k]) on a shared SUBTOTAL/order-statistics core

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `aggregate`. Design dependencies: `reference-semantics`, `table-model-registry`, `lexer-parser`. Those dependencies are for its END-TO-END acceptance test only (the corpus formula needs Phase 1's ROW fix and Phase 5's array-eligible table node). Everything else in this phase is self-contained and can land first. Adversarial verifier verdict: **needs-revision** (1 blocker, 1 major, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

AGGREGATE is registered as one new `Function` node (`Danfma.MySheet/Expressions/Mathematics/Aggregate.cs`,
`MinArgs=3`, `MaxArgs=int.MaxValue`) that disambiguates Excel's two forms purely by `function_num` at
EVALUATION time — n≤13 is the reference form (every argument from index 2 on is a ref), n≥14 is the array form
and requires exactly 4 arguments (`Arguments.Length != 4` → `#VALUE!`, matching Microsoft's documented "if a
second ref argument is necessary but not provided, AGGREGATE returns #VALUE!"). Rather than duplicating
SUBTOTAL's per-cell scan or the order statistics' top-k heap, the phase EXTRACTS three shared helpers and
makes SUBTOTAL a thin caller of them: a new `AggregateCodes` (the parameterized nested-node-skipping scan +
the 1-13 code map + the 14-19 positional map), the promotion of `file static class OrderSelection` to an
`internal` class in its own file so its `BoundedHeap` streaming k-th selection is reachable, and four new
`StatisticsMath` folds (`Median`, `Mode`, `QuartileInclusive`, `QuartileExclusive`) extracted from the four
nodes that currently inline them. Error-skipping for options 2/3/6/7 needs NO new machinery on the reference-
form path: the accumulator already drops error cells from the numeric population and only reports the first
one through a return channel, so "ignore errors" is one flag that suppresses that return (plus one extra rule
for COUNTA, which today *does* count error cells — measured: `SUBTOTAL(3,A1:A3)` with `A2=#DIV/0!` → 3, so
`AGGREGATE(3,6,…)` must return 2). CORRECTION TO THE BRIEF: Microsoft's options table ignores nested
SUBTOTAL/AGGREGATE for options 0/1/2/3 (NOT 2/3/6/7 — those ignore error VALUES); the bits decompose as
bit0=hidden rows (a no-op here, the S6 caveat), bit1=ignore errors, bit2=do NOT ignore nested, so
`ignoreErrors = (opts & 2) != 0` and `skipNested = (opts & 4) == 0`. The phase is self-contained and can land
first; only its end-to-end acceptance test is gated on the table node being mini-CSE-eligible and on ROW
accepting `INDEX(...)`.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Item 12 step (4): the reference form (function_num 1-13) accumulates each `Arguments[2..]` through `AggregateCodes.Gather` only — no mini-CSE gate. Item 10's rationale simultaneously declares that arm "mandatory, not an optimization: an array-eligible argument like `(A1:A3<>"")*1` reached through Gather's `default:` (Subtotal.cs:188-195) would be `Evaluate`d to a SCALAR".
      *Measured evidence:* Two contradictions, both measured (probe at /tmp/verify-aggregate, `dotnet run`
      against Danfma.MySheet). (a) The scalarization is real and silent for the `ROW(range)` shape:
      `=SUBTOTAL(9,ROW(A1:A3))` -> **1** while `=SUM(ROW(A1:A3))` -> **6**. SUBTOTAL and the specified
      AGGREGATE share the identical `Gather` path, so `AGGREGATE(9,0,ROW(A1:A3))` returns 1 — and the phase's
      own end-to-end target is built out of `ROW(A1:A3)-ROW(A1)+1`. (b) For the shape item 10 actually names
      it is worse than "a scalar": `=SUBTOTAL(9,(A1:A3<>0)*1)` -> **#VALUE!** (Subtotal.cs:189
      `argument.Evaluate(context)`; BinaryOperation over a range has no scalar value), and item 7 turns the
      default arm into `return ignoreErrors ? null : error`, so under options 2/3/6/7 that #VALUE! is
      SWALLOWED and the population is empty: `Fold(9, [])` -> 0 (measured `=SUBTOTAL(9,B1:B3)` over an empty
      range -> 0). So `AGGREGATE(9,6,(A1:A3<>0)*1)` -> **0**. Same for `AGGREGATE(3,6,…)` -> 0 and
      `AGGREGATE(2,6,…)` -> 0 (measured `=SUBTOTAL(3,(A1:A3<>0)*1)` -> 1, `=SUBTOTAL(2,(A1:A3<>0)*1)` -> 0
      today, i.e. the #VALUE! scalar being counted). Under either candidate Excel answer (Microsoft's page
      defines `array` as "An array, an array formula, or a reference to a range of cells" only for the ARRAY
      form, and `ref1` as "The first numeric argument" — so Excel either folds the array or returns #VALUE!)
      the values 1 and 0 are wrong. Nothing in `openQuestions` covers this, which P0 requires.
      *Correction:* Apply the same gate in step (4) that step (5) uses — for each ref, `arg is not Reference
      && ArrayEvaluation.IsArrayEligible(arg) && ArrayEvaluation.TryEvaluateStream(arg, context, out var s)`
      -> `AggregateCodes.CollectStream(s, ref accumulator)`, else `Gather(...)` — so codes 1-13 match
      SUM/COUNT/AVERAGE, which already do exactly this via `NumericAggregation.Fold`'s default arm; OR decide
      the reference form rejects a non-Reference argument with #VALUE! and say so. Either way state the Excel
      behaviour being matched, and add the unresolved half to openQuestions. Also correct item 10's factual
      claim: `(A1:A3<>"")*1` through Gather's default becomes #VALUE!, not a scalar number; `ROW(range)` is
      the shape that silently scalarizes.

## Major corrections

- [ ] **M1.** Item 14, second half: "Add the row-index idiom over a plain range as the standalone half of the end-to-end target: `=AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0)),ROWS($B$2:B2))` -> 1.0 and `…,ROWS($B$2:B3))` -> 3.0", placed "next to `Small_OfIfArray_ErrorAfterKthElement_StillPropagates` (:85-114) whose local-`Calc` fixture … is the right shape".
      *Evidence:* That fixture (tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs:94-103) sets
      A2=1, A3=2, A4=3, A5=`=1/0` and leaves **A1 blank**. The item's golden values were measured on a
      different fixture (its own rationale says "on A1=5/A2=0/A3=9"). Measured against the file's actual
      fixture: `=(A1<>"")` -> False, so the array `(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0))` is
      [#DIV/0!, 2, 3] — `INDEX(...,1)` -> #DIV/0!, `INDEX(...,2)` -> 2, `INDEX(...,3)` -> 3. With option 6
      skipping the error, k=`ROWS($B$2:B2)`=1 yields **2.0**, not the asserted 1.0 (k=2 coincidentally yields
      3.0).
      *Correction:* State the fixture explicitly in the item: a new local `Calc` with A1=5, A2=0, A3=9 (the
      values the goldens were measured on), not the neighbouring test's fixture. Keep the option-4-vs-6 IF-
      array pair on the existing fixture — those four values (#DIV/0!, 1.0, 3.0, #NUM!) are correct there.

## Implementation items

- [ ] **1.** Move `file static class OrderSelection` (Danfma.MySheet/Expressions/Statistical/OrderStatistics.cs:446-690, including its nested `private struct BoundedHeap` :603-689) VERBATIM into a new file Danfma.MySheet/Expressions/OrderSelection.cs with `namespace Danfma.MySheet.Expressions;`, changing only `file static class` → `internal static class`. Move `using System.Buffers;` (OrderStatistics.cs:1) into the new file and delete it from OrderStatistics.cs — after the move nothing left in that file uses ArrayPool. Do NOT touch any call site: `Median`/`PercentileInc`/`PercentileExc`/`PercentRankInc`/`QuartileInc`/`QuartileExc`/`TrimMean`/`Large`/`Small` all say `OrderSelection.X(...)` and C# namespace lookup walks up from `Danfma.MySheet.Expressions.Statistical` to `Danfma.MySheet.Expressions`.
      *Files:* `Danfma.MySheet/Expressions/OrderSelection.cs`, `Danfma.MySheet/Expressions/Statistical/OrderStatistics.cs`
      *Why:* `BoundedHeap` (O(n log k), ArrayPool-rented, OrderStatistics.cs:603-689) is exactly what
      AGGREGATE 14/15 needs and is unreachable from another file while its owner is `file static`. A move +
      visibility widening is the smallest change: the walk-up-namespace trick is already proven in this very
      file, which calls `StatisticsMath` (OrderStatistics.cs:36) and `ArrayEvaluation` (:508) — both in the
      parent `Danfma.MySheet.Expressions` — with no using directive. Duplicating the heap instead would
      violate the project's shared-helper-over-duplication rule; leaving it `file static` and writing
      AGGREGATE's own heap would duplicate ~90 lines.
- [ ] **2.** In the moved Danfma.MySheet/Expressions/OrderSelection.cs, extract `public static ComputedValue KthOfSorted(IReadOnlyList<double> sorted, double k, bool largest)` from the tail of `KthValue` (was OrderStatistics.cs:520-527: truncate k, `sorted.Count == 0 || position < 1 || position > sorted.Count` → `Error.Num`, else `largest ? sorted[^position] : sorted[position - 1]`) and make `KthValue` call it. AGGREGATE 14/15 reuses it for its non-streaming (reference/collected) path.
      *Files:* `Danfma.MySheet/Expressions/OrderSelection.cs`
      *Why:* Without this, `AggregateCodes.Positional` would re-implement the 4-line bounds contract that
      SMALL/LARGE already own, and the two could drift on the `k > n` / `k < 1` / empty edges (all three
      verified live: `=SMALL(B1:B3,1)` over an empty range → #NUM!).
- [ ] **3.** In Danfma.MySheet/Expressions/OrderSelection.cs change `private static ComputedValue KthValueStreaming(ArrayEvaluation.ArrayStream stream, Expression kArgument, EvaluationContext context, bool largest)` (was OrderStatistics.cs:537-542) to `public static ... KthValueStreaming(..., bool largest, bool ignoreErrors = false)`, and change the error arm of its scan loop (was :560-562) from `if (element.TryGetError(out var cellError)) { arrayError ??= cellError; }` to `if (element.TryGetError(out var cellError)) { if (!ignoreErrors) arrayError ??= cellError; }`. Nothing else in the method changes: the scan still visits EVERY element, and the error-precedence ladder (array error > k coercion error > `count == 0 || position < 1 || position > count` → #NUM!, :573-586) already produces AGGREGATE's required precedence once errors are skipped.
      *Files:* `Danfma.MySheet/Expressions/OrderSelection.cs`
      *Why:* One defaulted flag on the existing loop beats a parallel loop: the ladder at :573-586 is
      precisely AGGREGATE's contract (all-elements-error → count 0 → #NUM!; k past the post-skip count →
      #NUM!; empty array → #NUM!), and `count` is already incremented only for numeric elements, so skipping
      errors automatically makes the k bound the POST-SKIP population. A separate loop would duplicate the
      ArrayPool rent/return `try/finally` (:548-596) and the subtle 'scan every element even after k is
      satisfiable' rule that MiniCseConsumerTests.cs:85-114 pins for SMALL.
- [ ] **4.** Add four folds to Danfma.MySheet/Expressions/StatisticsMath.cs, each `public static Error? …(…, out double result)` in the file's existing style: `Median(IReadOnlyList<double> sorted, out double result)` (empty → `Error.Num`; `middle = Count/2`; odd → `sorted[middle]`, even → `(sorted[middle-1]+sorted[middle])/2` — lifted from Median.Evaluate at OrderStatistics.cs:45-54); `Mode(IReadOnlyList<double> values, out double result)` (the Dictionary<double,int> count loop with the strict `>` scan-order tie-break lifted from ModeSngl.Compute at OrderStatistics.cs:73-90, `bestCount < 2` → `Error.NA`); `QuartileInclusive(IReadOnlyList<double> sorted, double quart, out double result)` (truncate, `< 0 or > 4` → `Error.Num`, else delegate to `PercentileInclusive(sorted, quart/4, out result)` — lifted from QuartileInc.Compute at OrderStatistics.cs:368-377); `QuartileExclusive(IReadOnlyList<double> sorted, double quart, out double result)` (truncate then delegate to `PercentileExclusive(sorted, quart/4, out result)` with NO extra range check — lifted verbatim from QuartileExc.Evaluate at OrderStatistics.cs:401-405, whose 0/4 → #NUM! comes from PercentileExclusive's own `k is <= 0 or >= 1` guard at StatisticsMath.cs:136).
      *Files:* `Danfma.MySheet/Expressions/StatisticsMath.cs`
      *Why:* AGGREGATE codes 12/13/17/19 have no reusable entry point today: MEDIAN's logic is inline in
      `Median.Evaluate` (OrderStatistics.cs:16-55), QUARTILE.EXC's inline in `QuartileExc.Evaluate`
      (:386-406); only 13/16/17 expose an `internal static Compute` and those take `Expression[]`, not a
      population. StatisticsMath is already documented (:3-8) as 'shared numeric machinery … reused by
      STDEV*/VAR*/SUBTOTAL', so this is the established home. Verified live so the extraction cannot change
      behaviour: `=MEDIAN(B1:B3)` (empty) → #NUM!, `=MODE.SNGL(A1:A3)` (no repeats) → #N/A,
      `=QUARTILE.EXC(A1:A3,0)` → #NUM!, `=QUARTILE.EXC(A1:A3,2)` → 5.
- [ ] **5.** Rewrite the four nodes to call the new folds, keeping every observable outcome: Median.Evaluate (OrderStatistics.cs:45-54) → `StatisticsMath.Median(values, out var m) is { } e ? ComputedValue.Error(e) : ComputedValue.Number(m)` (the snapshot branch at :22-33 already yields a SORTED list, the collect branch already sorts at :41 — pass both to the same fold); ModeSngl.Compute (:73-90) → `StatisticsMath.Mode(values, out var m)`; QuartileInc.Compute (:368-377) → `StatisticsMath.QuartileInclusive(sorted, quart, out var r)`; QuartileExc.Evaluate (:401-405) → `StatisticsMath.QuartileExclusive(sorted, quart, out var r)`.
      *Files:* `Danfma.MySheet/Expressions/Statistical/OrderStatistics.cs`
      *Why:* Pure extract-method, guarded by the existing OrderStatisticTests suite — it is the anti-
      duplication half of the previous item. Skipping it would leave AGGREGATE 12/13/17/19 as a second
      implementation of four Excel definitions, exactly the branching-and-duplication the project's principles
      forbid.
- [ ] **6.** Create Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs, `internal static class AggregateCodes` in `namespace Danfma.MySheet.Expressions.Mathematics`, holding four members MOVED (not copied) out of Subtotal.cs: (a) `internal enum NestedSkip : byte { None, Subtotal, SubtotalAndAggregate }`; (b) `internal struct Accumulator(int code, bool ignoreErrors)` from Subtotal.SubtotalAccumulator:204-260, adding `public readonly List<double> Numbers => _numbers;` and the two error rules below; (c) `public static Error? Gather(Expression argument, EvaluationContext context, ref Accumulator accumulator, NestedSkip skip)` from Subtotal.GatherSkippingSubtotals:59-197 with every `expression is Subtotal` probe (:80, :114, :143) replaced by `IsNested(expression, skip)` and every recursive call (:166, :185, :194) threading `skip`; (d) `private static bool IsNested(Expression? e, NestedSkip skip) => skip switch { NestedSkip.None => false, NestedSkip.Subtotal => e is Subtotal, _ => e is Subtotal or Aggregate }`. Add `using Danfma.MySheet.Expressions.Statistical;` only if needed — `StatisticsMath`/`OrderSelection` resolve via the parent namespace.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`, `Danfma.MySheet/Expressions/Mathematics/Subtotal.cs`
      *Why:* Option (a) of the three offered: widen and share. (b) duplicate is out — the scan is 140 lines
      with five reference shapes (RangeReference dense walk :67-101, OpenRangeReference index walk :103-129,
      CellReference :131-161, UnionReference :163-172, Anchored* re-dispatch :181-186, default evaluate-then-
      re-dispatch :188-195) and each arm carries a hard-won invariant (measured: the nested skip fires through
      a union AND an open range — `=SUBTOTAL(9,(A1:A2,A3:A3))` and `=SUBTOTAL(9,A:A)` both → 3 with A3 =
      SUBTOTAL). (c) co-locating AGGREGATE inside Subtotal.cs is out — Subtotal.cs is a 346-line public node
      file, not a helper file, and the codebase's file-scoping precedents (`file static class OrderSelection`,
      `file static class SumOfPairs`) are for helpers with ONE consumer; this one has two. Mathematics is the
      right namespace because it needs both node types by name.
- [ ] **7.** In AggregateCodes.Accumulator.Add (from Subtotal.cs:210-244) make exactly two error-rule changes and widen the code range to 1-13: in the `case 3:` (COUNTA) arm, return early WITHOUT incrementing when `ignoreErrors && value.Kind == ComputedValueKind.Error` (today an error cell IS counted — measured: `=SUBTOTAL(3,A1:A3)` with A2 = `=1/0` → 3); in the `default:` arm change `if (value.TryGetError(out var error)) { return error; }` to `if (value.TryGetError(out var error)) { return ignoreErrors ? null : error; }`. Leave `case 2:` (COUNT) alone — it only tallies `ComputedValueKind.Number`, so it already ignores errors either way (measured: `=SUBTOTAL(2,A1:A3)` with an error cell → 2).
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`
      *Why:* This IS the error-skipping mechanism for the reference form — no new loop, no new fold. The
      accumulator already excludes error cells from `_numbers` (Subtotal.cs:231-240 returns before
      `_numbers.Add`), so 'ignore errors' is purely the suppression of the propagation channel. The COUNTA
      rule is a genuine behavioural difference the measurement exposes: without it `AGGREGATE(3,6,range)`
      would count the very cells option 6 says to ignore.
- [ ] **8.** In AggregateCodes, rename Subtotal.Aggregate:262-345 to `public static ComputedValue Fold(int code, List<double> numbers)`, and CHANGE ITS `default:` ARM: today `default: // 11: VAR.P` (:338-343) is the catch-all, so passing 12 would silently compute VAR.P. Convert it to `case 11:` and add `case 12:` (`numbers.Sort(); StatisticsMath.Median(numbers, out var median)`), `case 13:` (`StatisticsMath.Mode(numbers, out var mode)` — do NOT sort: the documented tie-break is FIRST value in scan order, OrderStatistics.cs:82), and a real `default: return ComputedValue.Error(Error.Value);` unreachable guard.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`
      *Why:* The silent-fallthrough is the single most likely defect in this phase: Subtotal.cs:338's
      `default` is safe only because Subtotal.Evaluate:27-30 rejects codes outside 1-11 first, and AGGREGATE
      deliberately admits 12-19. Sorting for 12 but not 13 is load-bearing — MODE.SNGL's tie-break is
      documented as scan-order and pinned by OrderStatisticTests.
- [ ] **9.** Add `public static ComputedValue Positional(int code, IReadOnlyList<double> sorted, Expression kArgument, EvaluationContext context)` to AggregateCodes: coerce k first (`kArgument.Evaluate(context).CoerceToNumber(out var k) is { } e → Error(e)`), then a switch STATEMENT (not expression — each arm needs its own `out` local): 14 → `OrderSelection.KthOfSorted(sorted, k, largest: true)`, 15 → same with `largest: false`, 16 → `StatisticsMath.PercentileInclusive`, 17 → `StatisticsMath.QuartileInclusive`, 18 → `StatisticsMath.PercentileExclusive`, default(19) → `StatisticsMath.QuartileExclusive`, each wrapping `Error? → ComputedValue.Error` / `ComputedValue.Number(result)`.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`
      *Why:* The 14-19 half of the code map, mirroring Fold's 1-13 half, so the node file holds only form
      selection. Coercing k AFTER the population is gathered preserves the ordering
      OrderSelection.SortedArrayAndScalar:463-477 already documents (the array's first error precedes the
      scalar's), which matters for options 0/1/4/5 where errors still propagate.
- [ ] **10.** Add `public static Error? CollectStream(ArrayEvaluation.ArrayStream stream, ref Accumulator accumulator)` to AggregateCodes: `foreach (var element in stream) { if (accumulator.Add(element) is { } error) return error; } return null;`. The array form uses it for codes 16-19 (which need the whole sorted population, so the bounded heap buys nothing) and for any code whose array argument is mini-CSE-eligible.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/AggregateCodes.cs`
      *Why:* One accumulator, three feeds (Gather for references, CollectStream for mini-CSE arrays,
      KthValueStreaming for 14/15) keeps every option rule — ignoreErrors, COUNTA's error exclusion — in ONE
      place. This arm is mandatory, not an optimization: an array-eligible argument like `(A1:A3<>"")*1`
      reached through Gather's `default:` (Subtotal.cs:188-195) would be `Evaluate`d to a SCALAR, and the
      mini-CSE gate is the only correct reading of it.
- [ ] **11.** Rewrite Danfma.MySheet/Expressions/Mathematics/Subtotal.cs to keep ONLY the record and Evaluate:13-51, delegating: hoist `var refs = Arguments[1..];` once (today allocated twice, :35 and :42), `new AggregateCodes.Accumulator(code, ignoreErrors: false)`, `AggregateCodes.Gather(argument, context, ref accumulator, AggregateCodes.NestedSkip.Subtotal)`, and `Finish()` now calling `AggregateCodes.Fold`. Keep the code-range guard `code is < 1 or > 11` → #VALUE! (:27-30) and the 101-111 mapping (:22-25) unchanged, and keep the file's explanatory comments by moving them with the code they describe. Do NOT change SUBTOTAL's nested skip to also skip `Aggregate` — see the open question.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/Subtotal.cs`
      *Why:* Subtotal drops from 346 to ~50 lines with zero behaviour change, which is what makes the
      extraction a net simplification rather than an addition. Leaving SUBTOTAL's skip predicate at
      `NestedSkip.Subtotal` is deliberate: Microsoft's SUBTOTAL page documents only 'nested subtotals are
      ignored', while the 'nested SUBTOTAL and AGGREGATE' wording appears solely in AGGREGATE's own options
      table — I could not verify the converse and will not guess it into the engine (see openQuestions).
- [ ] **12.** Create Danfma.MySheet/Expressions/Mathematics/Aggregate.cs: `[MemoryPackable] public sealed partial record Aggregate(Expression[] Arguments) : Function`. `Evaluate` in order: (1) `Arguments[0].Evaluate(context).CoerceToNumber(out var rawCode)` → propagate; `code = (int)Math.Truncate(rawCode)`; `code is < 1 or > 19` → `Error.Value`. (2) `Arguments[1].Evaluate(context).CoerceToNumber(out var rawOptions)` → propagate; `options = (int)Math.Truncate(rawOptions)`; `options is < 0 or > 7` → `Error.Value` (an OMITTED options arrives as `BlankValue.Instance` per Parser.cs:646-654 and coerces to 0 at ValueCoercion.cs:19-21 — Excel's '0 or omitted'). (3) `var ignoreErrors = (options & 2) != 0; var skip = (options & 4) == 0 ? AggregateCodes.NestedSkip.SubtotalAndAggregate : AggregateCodes.NestedSkip.None;` — bit0 (hidden rows) is intentionally unread, the S6 caveat. (4) if `code <= 13`: `var refs = Arguments[2..]`; `ReferenceGuard.MissingSheet(refs, context)` → #REF!; accumulate each ref through `AggregateCodes.Gather(…, skip)`; `Finish()`. (5) else: `Arguments.Length != 4` → `Error.Value`; `ReferenceGuard.MissingSheet(Arguments[2], context)` → #REF!; if `Arguments[2] is not Reference && ArrayEvaluation.IsArrayEligible(Arguments[2]) && ArrayEvaluation.TryEvaluateStream(Arguments[2], context, out var stream)` then `code is 14 or 15` → `OrderSelection.KthValueStreaming(stream, Arguments[3], context, largest: code == 14, ignoreErrors)`, else `CollectStream` + `Numbers.Sort()` + `AggregateCodes.Positional`; otherwise `Gather(Arguments[2], …, skip)` + `Numbers.Sort()` + `Positional`.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/Aggregate.cs`
      *Why:* Disambiguation by function_num alone is the only rule that can work —
      `AGGREGATE(9,6,A1:A3,B1:B3)` and `AGGREGATE(15,6,A1:A3,2)` are both 4-argument calls, so nothing
      syntactic separates them. The `is not Reference` half of the mini-CSE gate copies
      OrderSelection.KthValue:506-513 verbatim so a plain range keeps the reference path (and its nested-skip
      walk) while a computed array streams. The up-front `ReferenceGuard.MissingSheet` is load-bearing under
      option 6: without it a missing-sheet range would aggregate to an empty population and return 0/#NUM!
      instead of #REF! — the exact failure ReferenceGuard.cs:6-9 was written to prevent for the error-ignoring
      COUNT family.
- [ ] **13.** Register the node: add `[MemoryPackUnion(<next free tag>, typeof(Aggregate))]` immediately before `public abstract partial record Expression` (Danfma.MySheet/Expressions/Expression.cs:353), after the 321 line at :352. Count the attributes first (`grep -c 'MemoryPackUnion(' Danfma.MySheet/Expressions/Expression.cs` — currently 322, tags 0-321, so 322 is next) exactly as the policy comment at :14-16 instructs, and refresh that comment's stale '319+'. Then add an `Entry<Aggregate>("AGGREGATE", 3, int.MaxValue, static arguments => new Aggregate(arguments), static f => ((Aggregate)f).Arguments)` to FunctionRegistry.Entries, next to the SUBTOTAL entry (Danfma.MySheet/Parsing/FunctionRegistry.cs:1550-1556).
      *Files:* `Danfma.MySheet/Expressions/Expression.cs`, `Danfma.MySheet/Parsing/FunctionRegistry.cs`
      *Why:* `MinArgs=3` makes `AGGREGATE(15,6)` a parse-time ParseException (Parser.cs:634-642 — verified
      live for `=SUBTOTAL(9)`: "Function 'SUBTOTAL' does not accept 1 argument(s)"), matching Excel rejecting
      it at entry; `MaxArgs=int.MaxValue` matches SUBTOTAL's entry and defers the array form's 'exactly 4'
      rule to Evaluate, which is where Excel's own #VALUE! for a missing k lives. Naming the ctor parameter
      `Arguments` (not `Expressions`) keeps it out of the `Sum` special case documented at
      FunctionRegistry.cs:33-37. Un-parse and dependency extraction then cost NOTHING: FormulaWriter.Call
      resolves any Function through `FunctionRegistry.ByType` (:2142) and DependencyExtractor.cs:207-214 has a
      generic `case Function` arm that visits arguments via that same map — verified live that
      `=SUBTOTAL(109,A1:A3)` round-trips through `ToFormula` with no per-node writer arm.
- [ ] **14.** Add AGGREGATE tests to tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs next to the SUBTOTAL block (:175-243), reusing the existing `Calc` harness (:11-31) and its `"=1/0"` idiom for an error cell: (a) every function_num 1-19 against a fixture, each asserted EQUAL to the same-population call of its namesake function (the anti-vacuity pattern already used at :217-241 for SUBTOTAL 7/8/10/11); (b) the options matrix — `AGGREGATE(9,4,A1:A3)` with A2=`=1/0` → `ErrorValue.DivByZero` vs `AGGREGATE(9,6,A1:A3)` → the sum of the survivors, and `AGGREGATE(3,4,…)` → 3 vs `AGGREGATE(3,6,…)` → 2; (c) the S6 caveat — 1≡0, 3≡2, 5≡4, 7≡6 pairwise equal; (d) nested skip — a cell holding `=SUBTOTAL(9,…)` and a cell holding `=AGGREGATE(9,0,…)` are both skipped at options 0-3 and both COUNTED at options 4-7; (e) form/arity errors — `AGGREGATE(20,0,A1:A3)` and `AGGREGATE(9,8,A1:A3)` → `ErrorValue.NotValue`, `AGGREGATE(15,6,A1:A3)` → `ErrorValue.NotValue` (k required), `AGGREGATE(15,6,A1:A3,4)` → `ErrorValue.Number` (k past the population), all-error population with option 6 → `ErrorValue.Number`, `AGGREGATE(9,6,Ghost!A1:A3)` → `ErrorValue.Reference`; (f) reference vs array form on the same shape — `AGGREGATE(9,6,A1:A3,B1:B3)` sums both refs while `AGGREGATE(15,6,A1:A3,2)` treats 2 as k. Head the block with the Microsoft AGGREGATE page citation in the file's mandatory format (see :175-177): the article title, its support.microsoft.com GUID and the fetch date must be FETCHED at implementation time — do not copy a GUID from this plan, none is supplied.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs`
      *Why:* MathAggregateTests already owns SUBTOTAL, shares the fixture style, and carries the golden-value
      convention; a new file would fork the `Calc` helper a fourth time. The options matrix is the only way to
      pin the bit decomposition, and (d) is the only test that would catch `IsNested` being wired to the wrong
      option range — the mistake the brief's own (incorrect) 2/3/6/7 claim would have produced.
- [ ] **15.** Add the array-consumer tests to tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs, next to `Small_OfIfArray_ErrorAfterKthElement_StillPropagates` (:85-114) whose local-`Calc` fixture (A2:A4 = 1,2,3 and A5 = `=1/0`) is the right shape: assert `=AGGREGATE(15,4,IF(B2:B5="Show",A2:A5),1)` → `ErrorValue.DivByZero` (option 4 keeps SMALL's propagation) and `=AGGREGATE(15,6,IF(B2:B5="Show",A2:A5),1)` → 1.0 with `…,3)` → 3.0 and `…,4)` → `ErrorValue.Number` (option 6 skips the trailing error, so k is bounded by the POST-SKIP count of 3). Add the row-index idiom over a plain range as the standalone half of the end-to-end target: `=AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0)),ROWS($B$2:B2))` → 1.0 and `…,ROWS($B$2:B3))` → 3.0.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`
      *Why:* AGGREGATE becomes the fifth mini-CSE consumer (after NumericAggregation.Fold,
      OrderSelection.KthValue and Index), and this file's docstring (:7-12) names the consumers explicitly.
      The option-4-vs-6 pair on the SAME fixture is the direct regression for the one-line change to
      KthValueStreaming's error arm — it proves the flag flips the behaviour without disturbing SMALL, whose
      own assertion sits ten lines above. All four values are MEASURED, not derived:
      `=INDEX((ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0)),i)` yields 1, #DIV/0!, 3 for i=1,2,3 on
      A1=5/A2=0/A3=9, and `ROWS($B$2:B2)`/`ROWS($B$2:B3)` yield 1/2.
- [ ] **16.** Add `"AGGREGATE(9,6,A1:A3)"` and `"AGGREGATE(15,6,A1:A3,2)"` to the flat parse→write identity list in tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs, next to `"SUBTOTAL(9,A1:A3)"` (:268) and `"SMALL(A1:A3,2)"` (:272).
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`
      *Why:* Both forms must round-trip, and the two-entry pair is what proves the registry accessor
      (`((Aggregate)f).Arguments`) is wired for the variable-arity node. FormulaWriter needs no code change —
      `Call` (:433-445) is table-driven — so this test is the whole verification of the write half.
- [ ] **17.** Update docs/function-reference.md: `**304 built-in functions**` → 305 at :3; `## Math and trigonometry (74)` → (75) at :31; insert the AGGREGATE row ALPHABETICALLY between the `ACOTH` row (:39) and the `ARABIC` row (:40) — NOT after SQRTPI:96 — documenting both forms, the 1-19 map, the 0-7 options table, the 1/3/5/7 ≡ 0/2/4/6 hidden-row caveat, `#VALUE!` for an invalid function_num/options or a missing k, and `#NUM!` for an out-of-range k; `<strong>Math and Trigonometry</strong> — 74/82` → 75/82 at :447; move `AGGREGATE` from the ⬜ list (:451) into the ✅ list (:449, first position, before `ABS`). Also extend the SUBTOTAL row (:97) so it says a nested `AGGREGATE` is NOT skipped by SUBTOTAL (the documented-limit note). Mirror every edit in docs/pt-BR/function-reference.md at :5, :35, between :43 and :44, :101, :456, :460 and its ✅ list.
      *Files:* `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* Verified that the Math table at :35-45 IS alphabetical (`ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`,
      `ARABIC`, …), so the brief's 'after SQRTPI/before SUBTOTAL' insertion point is wrong and would break the
      table's ordering. The pt-BR mirror is mandatory per docs/pt-BR/README.md:3, and its line numbers differ
      by +4 because of the translation banner.
- [ ] **18.** Add AGGREGATE's new union tag to the compatibility section of docs/serialization.md — either as its own '### Forward-compatibility' subsection after the shared-formula one (:191-214, the template to copy) or folded into the shared subsection the table-registry phase writes, whichever lands second. State the same two-way asymmetry that subsection states: a workbook whose cells hold an AGGREGATE formula cannot be opened by a build predating the tag; an older file never contains it and loads unchanged. Mirror in docs/pt-BR/serialization.md. NOTE: adding a union tag does NOT touch tests/Danfma.MySheet.Tests/CellStoreTests.cs:20's frozen base64 (that golden breaks only on a new Workbook MEMBER) and is safe against Fixtures/workbook-pre-namespaces.msgpack.bin.
      *Files:* `docs/serialization.md`, `docs/pt-BR/serialization.md`
      *Why:* S7 requires the boundary to be documented, and the tag number is only knowable at integration
      time because whichever of this phase and the table phase lands first claims 322 — Expression.cs:14-16
      explicitly warns to count the attributes rather than trust a written number. Spelling out what does NOT
      break saves the implementer from regenerating a golden unnecessarily.


> **Landed differently (2026-09-09, Phase 1 Task 2):** the ROW/COLUMN axis mirror this item describes was collapsed at implementation time. The live symbols are `PositionNumbersOperand(origin, PositionAxis, rows, columns)`, `ProbePosition`/`TryBuildPositionOperand`, `ResolvePositionRange` and `PositionArgumentShape` in `Danfma.MySheet/Expressions/ArrayEvaluation.cs`. `RowNumbersOperand`, `ColumnNumbersOperand`, `ResolveRowRange` and `RowArgumentShape` no longer exist — read the file, not this text.

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded, 0 Error(s). A CS0103/CS0122 on `OrderSelection`, `BoundedHeap` or
      `ArrayPool` means the file-scoped→internal move in item 1 was incomplete (or `using System.Buffers;` did
      not travel with it).
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier check .`
      → expected: Exit code 0 and no file listed as needing formatting. CSharpier 1.3.0 is pinned in dotnet-
      tools.json and both the pre-commit and pre-push hooks run this, so a non-zero exit blocks the commit.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/OrderStatisticTests/*"`
      → expected: "Test run summary: Passed!" with failed: 0. This is the regression gate for items 1-5:
      SMALL/LARGE/MEDIAN/MODE.SNGL/PERCENTILE.*/QUARTILE.*/PERCENTRANK.*/TRIMMEAN must be byte-for-byte
      unaffected by the OrderSelection move, the KthOfSorted extraction, the ignoreErrors parameter and the
      four StatisticsMath extractions.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/MathAggregateTests/*"`
      → expected: "Test run summary: Passed!" with failed: 0 and total ≥ 19 (it is 13 today, measured, all
      passing in 243ms). A failure in the pre-existing Subtotal_* tests means the item-11 delegation changed
      SUBTOTAL's behaviour; a failure only in the new Aggregate_* tests is AGGREGATE's own logic.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/MiniCseConsumerTests/*"`
      → expected: "Test run summary: Passed!" with failed: 0. Specifically
      Small_OfIfArray_ErrorAfterKthElement_StillPropagates must still pass (proves `ignoreErrors: false` is
      the unchanged default) alongside the new option-4-vs-6 AGGREGATE pair.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/FormulaWriterTests/*"`
      → expected: "Test run summary: Passed!" with failed: 0 — both new AGGREGATE strings parse and un-parse
      to themselves through the registry-driven writer.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release --no-build && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release --no-build`
      → expected: Both runs report "Passed!" with failed: 0 — the exact pair the pre-push hook and
      .github/workflows/ci.yml:26-30 run. Confirms in particular that no serialization golden
      (CellStoreTests.PreChangeCellsWireGolden, MemoryPackCompatibilityTests,
      ContainerVersionCompatibilityTests) moved from the new union tag.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -c 'MemoryPackUnion(' Danfma.MySheet/Expressions/Expression.cs && grep -n 'typeof(Aggregate)' Danfma.MySheet/Expressions/Expression.cs`
      → expected: 323 attributes (was 322, measured) and exactly one line matching `typeof(Aggregate)` whose
      tag equals 322 (or the next free number if another phase landed first). No duplicate and no gap.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -rn 'file static class OrderSelection\|private static ComputedValue Aggregate(int code' Danfma.MySheet/`
      → expected: No output. Both the file-scoped OrderSelection and Subtotal's private `Aggregate(int,
      List<double>)` must be GONE — if either still exists the extraction was a copy, not a move, and the two
      implementations will drift.

## Risks carried by this phase

- SILENT WRONG ANSWER if the table node is not mini-CSE-eligible. MEASURED with a defined name as the stand-in for `T[Col]`: `=SUM((A1:A3<>0)*1)` → 2 but `=SUM((T<>0)*1)` → 1, where T is a defined name for Sheet1!$A$1:$A$3. `ArrayEvaluation.Probe` (ArrayEvaluation.cs:123-201) whitelists only RangeReference/AnchoredRangeReference/`Row{[RangeReference]}`; every other range-producing node falls to `default: return (true, false)` at :199-200, becomes an OPAQUE SCALAR evaluated once, implicitly intersects to a single cell and broadcasts. The end-to-end target would then return 1 at BOTH k=1 and k=2 instead of 1 and 3 — an error-free wrong number. The table phase MUST add a range-leaf arm for the table node to Probe (near :206-212) and TryBuildOperand (:420-481), plus a `Row{[table node]}` arm mirroring :220-226 / RowNumbersOperand:302, and must preserve the documented `IsArrayEligible ⇒ TryBuildOperand succeeds as an array` invariant (ArrayEvaluation.cs:113-117) even when the table name cannot be resolved.
- HARD DEPENDENCY on the ROW fix, failing loudly rather than silently. MEASURED: `=ROW(INDEX(A1:A3,1,1))` → #VALUE! today, and `=INDEX((ROW(A1:A3)-ROW(INDEX(A1:A3,1,1))+1)/((A1:A3<>"")*(A1:A3<>0)),1)` → #VALUE! for EVERY element, because the failed opaque scalar broadcasts into the whole array. Under option 6 all elements are then skipped, the population is empty and AGGREGATE returns #NUM! instead of 1. Note the good news: after Row.Evaluate (Danfma.MySheet/Expressions/Lookup/Row.cs:13-35) learns to resolve reference-producing arguments, `Row{[Index]}` stays an OPAQUE SCALAR in Probe — which is exactly right for the anchor term — so no ArrayEvaluation change is needed for that half.
- `Subtotal.Aggregate`'s `default: // 11: VAR.P` (Subtotal.cs:338-343) is a catch-all. Moving it to a shared `Fold` that AGGREGATE calls with 12 and 13 will compute VAR.P for MEDIAN and MODE.SNGL unless it is converted to `case 11:` with a real default (item 8). The bug is silent — VAR.P returns a plausible number.
- MODE.SNGL's tie-break scan order deviates from Excel for a 2-D reference, and AGGREGATE(13,…) inherits it. `AggregateCodes.Gather`'s RangeReference arm walks COLUMN-outer/row-inner (Subtotal.cs:74-77), as does `RangeReference.ExpandComputedValues`, while Excel scans row-major; the documented 'first value in scan order' tie-break (OrderStatistics.cs:82) therefore picks a different winner when two values tie in a multi-column range. Pre-existing in MODE.SNGL, not introduced here — do not attempt to fix it inside this phase.
- The reference-form population uses REFERENCED semantics for directly-passed scalars, unlike SUM. MEASURED: `=SUBTOTAL(9,"abc")` → 0 and `=SUBTOTAL(9,TRUE)` → 0, whereas `=SUM("abc")` → #VALUE! and `=SUM(TRUE)` → 1, because the accumulator's default arm (Subtotal.cs:230-243) only reads TryGetError/TryGetNumber and never routes through `NumericAggregation.AddDirect` (NumericAggregation.cs:282-326). AGGREGATE inherits this. It is defensible (AGGREGATE's arguments are refs) but it IS a deviation from Excel for literal arguments; do not 'fix' it here — it would change SUBTOTAL's shipped behaviour.
- Options 0-3 make the array form take the per-cell `Gather` walk rather than the shared per-epoch sorted view (`Workbook.TryGetRangeSnapshot` + `RangeSnapshot.SortedNumericValues`, RangeValueCache.cs:624), which `OrderSelection.SortedArrayAndScalar:463-477` uses for SMALL/LARGE/PERCENTILE over a whole column. `AGGREGATE(16,0,A:A,0.5)` dragged down a column therefore re-collects and re-sorts per formula. A snapshot fast path is available and correct ONLY when `skip == NestedSkip.None` (options 4-7), and `SortedNumericValues` already excludes error cells and reports the first through an out-parameter, so discarding it IS the ignore-errors reading. Deliberately deferred: adding it forks the array form on options, and `Workbook.TryGetRangeSnapshot` has stateful second-use admission (CriteriaScan.cs:52-60) so the result must be threaded, never re-fetched.
- The union tag number cannot be fixed in this plan. Both this phase and the table-model phase append tags, and whichever lands second must re-count (Expression.cs:14-16 says so explicitly). If both are written as '322' independently, MemoryPack will fail at type-initialization with a duplicate-tag error rather than at compile time.

## Open questions owned by this phase

- Does Excel's SUBTOTAL ignore a referenced cell whose formula is an AGGREGATE? The brief asserts yes; I could NOT verify it. Microsoft's SUBTOTAL page documents only 'If there are other subtotals within ref1, ref2,… (or nested subtotals), these nested subtotals are ignored'; the phrase 'nested SUBTOTAL and AGGREGATE functions' appears only in AGGREGATE's own options table, which describes AGGREGATE's behaviour, not SUBTOTAL's. This plan therefore leaves SUBTOTAL's skip predicate at `NestedSkip.Subtotal`. Settle it with a real .xlsx: A1=1, A2=2, A3=`=AGGREGATE(9,0,A1:A2)`, A4=`=SUBTOTAL(9,A1:A3)`; open in Excel (or Aspose.Cells) — A4=3 means SUBTOTAL skips it and item 11 should flip to `NestedSkip.SubtotalAndAggregate`, A4=6 means it does not.
- Does the ARRAY form (function_num 14-19) honour the nested SUBTOTAL/AGGREGATE skip at options 0-3? This plan assumes yes on the grounds that Microsoft states the options table once for the whole function, and it costs nothing because the array-over-a-Reference path already goes through `Gather`. Same experiment shape: A1:A3 = 1, 2, `=SUBTOTAL(9,A1:A2)`, then `=AGGREGATE(14,0,A1:A3,1)` — 3 means the skip applies (LARGE over {1,2}), 3 also being the nested value makes this ambiguous, so use A1=1, A2=2, A3=`=SUBTOTAL(9,A1:A2)`=3 and compare `AGGREGATE(14,0,…,1)` (expect 2 if skipped) with `AGGREGATE(14,4,…,1)` (expect 3).
- Under options 2/3/6/7, does Excel ignore an error passed as a whole non-reference argument, e.g. `AGGREGATE(9,6,1/0)`? 'Ignore error values' plausibly covers it, and this plan's accumulator does ignore it (returns 0), but the docs speak of error values in the aggregated data. MySheet's current SUBTOTAL propagates it (measured: `=SUBTOTAL(9,1/0)` → #DIV/0!), so the two functions will differ here by design. Worth a probe against Excel before the docs row is written.
- Under options 4/5 ('ignore nothing') over a 2-D reference holding two DIFFERENT error values, which error does Excel return? MySheet returns the first in column-major order (Subtotal.cs:74-77); Excel scans row-major. Only observable with a multi-column range and two distinct errors. Pre-existing for SUBTOTAL, inherited here — low value to chase, but it should not be claimed as Excel-exact in the docs.
- Does Excel truncate a non-integer function_num (e.g. `AGGREGATE(9.7,0,A1:A3)` → SUM) or reject it as #VALUE!? This plan truncates, following `Subtotal.Evaluate`'s `(int)Math.Truncate(rawCode)` at Subtotal.cs:20. Same question for options.

## Phase Summary

_(write when phase completes)_
