# Phase 7: FILTER / SORT / UNIQUE / SEQUENCE as mini-CSE producers

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `dynamic-arrays`. Design dependencies: `aggregate`, `reference-semantics`, `resolution-and-graph`. Adversarial verifier verdict: **needs-revision** (2 blockers, 4 majors, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Controller note (binding, 2026-09-09) — the cell-boundary array half is NOT top-left

Measured by Phase 10's verifier on Aspose.Cells 26.6.0: a bare formula in a cell applies implicit intersection to EACH range operand BEFORE the operator (`=E1:E3*10` in L2 = 20, L3 = 30, L5 = `#VALUE!`; `=E1:E3*E5:G5` in F2 = 40 = E2×F5; `=A1:C3*E1:E1` = `#VALUE!` because a 2-D operand has no single intersection; the top-left element is reachable only through `INDEX(...,1,1)`). This phase's design text ("bare in a cell follows S4's array half", `ElementAt(0)` on the whole expression) must be replaced by per-operand intersection at the boundary under the P0 addendum ("Excel" = Aspose measured). Bare `=LEN(A1:A3)` and bare `=SEQUENCE(5)` need re-measuring on the oracle under the same rule before the design is briefed. See the master plan's S4 correction.

## Design decision

The decision: each of the four becomes an `ArrayOperand` producer inside the existing mini-CSE operand tree,
reached through ONE new arm in `ArrayEvaluation.Probe`/`TryBuildOperand` that dispatches on a new `internal
interface IArrayProducer` — not a new `ComputedValueKind`, not a materialized array value. I verified by probe
that this is the ONLY producer channel: `ArrayEvaluation.Probe` (ArrayEvaluation.cs:123-201) accepts NO
function except `Row{[RangeReference]}` and `If`; `ABS/SQRT/N/IFERROR/SMALL/INDEX/COLUMN/TRANSPOSE` over
`A1:A3` all measure `IsArrayEligible=false`. Because `ArrayOperand` is `internal abstract` with `public
abstract` members (:215-230) I proved cross-assembly that each operand can live in its own function's file, so
`ArrayEvaluation.cs` gains exactly two arms, ever. The second insight is that FILTER, SORT and UNIQUE are the
SAME operation — a selection or permutation of one axis of a source array — so they share ONE
`AxisSelectionOperand` over an `int[]`, and they reuse `ValueCoercion.Compare` (ValueCoercion.cs:171-186,
already exactly Excel's classic number<text<FALSE<TRUE order, case-insensitive) and `ValueCoercion.AreEqual`
(:123-165) instead of a new comparer. Variable-length FILTER IS representable: the include vector is read once
at BUILD time (the laziness contract at ArrayEvaluation.cs:83-91 already sanctions build-time scalar
evaluation), the shape is then constant, and the VALUES stay on demand. `SUM(FILTER(..))`,
`INDEX(SORT(..),1)`, `COUNT(UNIQUE(..))` and `SUM(SEQUENCE(5))` then work with zero changes to any consumer,
and `DependencyExtractor`, `AnchoredFormulaSupport` and `FormulaWriter` need no change at all (measured).
Bare-in-a-cell is Excel's `@` rule, VERIFIED verbatim against Microsoft's "Implicit intersection operator: @"
article: for a RANGE "@ returns the value from the cell on the same row or column as the formula" (S4/Phase
5's half), and for an ARRAY Excel "picks the top-left value" — so `=FILTER(...)` bare yields the top-left,
exactly what Excel writes as `=@FILTER(...)`.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Items 9/10 (SORT, UNIQUE) build an AxisSelectionOperand from `source` without ever checking `source.IsArray`, so a scalar/single-cell argument produces the 0-extent array that item 5 exists to forbid.
      *Measured evidence:* ArrayEvaluation.cs:238-240 `ScalarOperand`: `public override bool IsArray => false;
      public override int Rows => 0; public override int Columns => 0;`. Probe's `default: return (true,
      false);` (ArrayEvaluation.cs:198-199) makes ANY single cell, literal or non-eligible function a
      ScalarOperand — measured `IsArrayEligible(A1)` -> False, `IsArrayEligible(INDEX(A1:A3,1))` -> False.
      Item 9 says "build source; ... materialize the key vector by reading `source.At` along the key column
      ... build `int[] permutation = [0..n)`" — with source.Rows == 0, n == 0, so `new
      AxisSelectionOperand(source, Rows, int[0])` has Rows == 0. Measured in /tmp probe: a 0x1 operand gives
      `ArrayStream.Length=0 enumerated=0` (ArrayStream.Length is `Rows * Columns`, ArrayEvaluation.cs:604). So
      `SUM(SORT(A1))` returns 0 where Excel returns A1's value. Item 8 (FILTER) escapes this only
      incidentally, via its axis-match check. Item 10 (UNIQUE) hits its "nothing kept yields 1x1 #CALC!"
      branch, so `SUM(UNIQUE(A1))` is #CALC! where Excel gives A1's value. The spec's own risk #4 names this
      "the exact class of silent-wrong-answer bug this phase must not ship".
      *Correction:* In items 8, 9 and 10, first line of `TryBuildArrayOperand`: after building `source`, `if
      (!source.IsArray) { operand = new SingletonArrayOperand(source.Scalar); return true; }`. Add it to the
      item 5 invariant text as a named precondition, and add `SUM(SORT(A1))`, `SUM(UNIQUE(A1))`,
      `SUM(FILTER(A1,TRUE))` to item 22's test list.
- [ ] **B2.** Item 15's ROWS/COLUMNS gate returns `ComputedValue.Number(array.Rows)` unconditionally, so it reports 1 for the 1x1 error singleton item 5 mandates for the empty/invalid case — silently wrong in exactly the headline use case item 15 justifies itself with.
      *Measured evidence:* Item 5 requires "an empty result is expressed as a 1x1 `SingletonArrayOperand`
      carrying `#CALC!`" and item 7 requires 1x1 `#CALC!`/`#NUM!` for bad SEQUENCE arguments. Item 15's gate
      is verbatim `if (Arguments[0] is not Reference && ArrayEvaluation.IsArrayEligible(...) &&
      ArrayEvaluation.TryEvaluateStream(..., out var array)) return ComputedValue.Number(array.Rows);` — no
      error inspection. So `ROWS(FILTER(A1:A3,A1:A3>100))` -> 1, `ROWS(SEQUENCE(-1))` -> 1,
      `ROWS(SEQUENCE(0))` -> 1, `COLUMNS(SEQUENCE(1e7))` -> 1. Excel propagates: FILTER itself is #CALC!, and
      ROWS(#CALC!) is #CALC!. Item 15's own rationale: "`ROWS(FILTER(range, cond))` is THE Excel idiom for
      'how many rows matched'" — the zero-match answer is the single most likely value a user reads from it.
      *Correction:* In both gates: `if (array.Length == 1 && array.ElementAt(0).TryGetError(out var
      arrayError)) return ComputedValue.Error(arrayError); return ComputedValue.Number(array.Rows);`. This is
      correct for every existing shape too — a multi-element all-error array (`ROWS(IF(A1:A3>0,1/0))`) still
      counts 3, matching Excel. Pin `ROWS(FILTER(A1:A3,A1:A3>100))` = #CALC! in item 22.

## Major corrections

- [ ] **M1.** The spec's completeness list omits `NamedReferences.CaptureValue`'s closed type list — the cross-cutting site the review brief names explicitly — so LET, CHOOSE and unary-plus will silently collapse a producer to its top-left element.
      *Evidence:* NamedReferences.cs:59-69 `CaptureValue` is `RangeReference or OpenRangeReference or
      UnionReference => ComputedValue.Reference(...)`, `AnchoredRangeReference anchored => ...`, `_ =>
      expression.Evaluate(context)`. `Filter`/`Sort`/`Unique`/`Sequence` are `Function`, not `Reference`, so
      they hit `_` and get item 3's `FirstElement` — the top-left. Let.cs:29-31 calls
      `NamedReferences.CaptureValue(Arguments[i + 1], scope)`; LookupFunctions.cs (CHOOSE) and
      UnaryOperation.cs do the same. Measured today with the analogous array shapes: `SUM(IF(A1:A3>0,A1:A3))`
      -> 14 but `LET(x,IF(A1:A3>0,A1:A3),SUM(x))` -> #VALUE!, `LET(x,A1:A3*2,SUM(x))` -> #VALUE!,
      `SUM(CHOOSE(1,A1:A3*2))` -> #VALUE!, `SUM(+IF(A1:A3>0,A1:A3))` -> #VALUE! (fixture A1=5,A2=0,A3=9). So
      today the composition fails LOUDLY; after this phase `LET(x, FILTER(A1:A3,A1:A3>0), SUM(x))` returns 5
      instead of 14 with no error — a regression in failure mode, not just a missing feature. The spec asserts
      "`DependencyExtractor`, `AnchoredFormulaSupport` and `FormulaWriter` need no change at all (measured)"
      and never mentions CaptureValue.
      *Correction:* Add an explicit item: either (a) document the limitation in plans/dynamic-array-
      functions.md and docs/workbook-and-expressions.md ("an array producer bound by LET/CHOOSE or passed
      through unary + collapses to its top-left; use it directly as the consumer's argument") plus a pinned
      test, or (b) route `Probe`/`TryBuildOperand` through `Let`/`Choose`/`UnaryOperation{Plus}` so the mini-
      CSE sees through them. (a) is the honest scope; either way it must not be silent.
- [ ] **M2.** Item 6 adds `#CALC!` but leaves `ERROR.TYPE` — which already exists and is documented — returning #N/A for it, where Excel returns 14.
      *Evidence:* Danfma.MySheet/Expressions/Information/InformationFunctions.cs:218-258 `ErrorType.Evaluate`
      is a closed if-chain ending `return error == Error.NA ? ComputedValue.Number(7) :
      ComputedValue.Error(Error.NA);`. Measured: `ERROR.TYPE` over a `#CALC!` cell -> #N/A today, and item 6
      changes nothing here, so after the phase `ERROR.TYPE(FILTER(A1:A3,A1:A3>100))` is #N/A. Excel's
      ERROR.TYPE mapping continues 8 #GETTING_DATA, 9 #SPILL!, 10 #CONNECT!, 11 #BLOCKED!, 12 #UNKNOWN!, 13
      #FIELD!, 14 #CALC! [Likely — Microsoft's ERROR.TYPE mapping table; not fetched in-session, worth
      confirming]. P0 makes this an observable wrong answer for the very error the phase introduces. Item 21's
      doc list also misses the two places the error set is enumerated for users: docs/function-
      reference.md:258 (`| ERROR.TYPE | ... #NULL!=1 ... #N/A=7 ...`) and docs/computed-value.md:112-120 (the
      `Error.*` -> Display table, currently `Error.Null` through `Error.NA`).
      *Correction:* Extend item 6 with: `if (error == Error.Calc) return ComputedValue.Number(14);` in
      ErrorType.Evaluate; update docs/function-reference.md:258 and docs/computed-value.md:112-120 plus both
      pt-BR twins; add an `ERROR.TYPE` golden to item 22. If the 14 cannot be confirmed against an oracle,
      move it to openQuestions rather than shipping #N/A silently.
- [ ] **M3.** Item 21's replacement doc bullet would state an `@` rule the engine does not implement for the array shapes that section is actually about.
      *Evidence:* Item 21: "REPLACE the 'dry cell keeps #VALUE!' bullet at :358-361 with the `@` rule: a bare
      dynamic-array formula yields the array's TOP-LEFT element". Item 3 only changes `Evaluate` on the four
      new records; `BinaryOperation.Evaluate` and `If.Evaluate` are untouched. Measured in cell D2:
      `=ROW(A1:A3)` -> Number:1 (top-left, as the spec claims), but `=A1:A3*2` -> #VALUE! and
      `=IF(A1:A3>0,1,0)` -> #VALUE!. The doc section being rewritten (docs/workbook-and-
      expressions.md:337-371) uses `=IF(B2:B5="Show",1,0)` as its own worked example of the bullet in
      question, so the rewritten text would be false for its own example.
      *Correction:* Scope the rewritten bullet: the `@`/top-left rule applies to FILTER/SORT/UNIQUE/SEQUENCE
      and `ROW(range)`; a bare `IF(range…)` or range comparison still yields #VALUE! (stated as a known
      inconsistency, with a pointer to the follow-up that would route both through
      `ArrayEvaluation.FirstElement`). Or extend item 3 to `BinaryOperation`/`If` and price that in.
- [ ] **M4.** Producer OUTPUT values are never blank-normalized, so FILTER/SORT/UNIQUE disagree with Excel on blank source cells — and risk #8 describes the opposite of what items 4/8/10 specify.
      *Evidence:* Item 4 defines `AxisSelectionOperand.At` as "maps the row-major index through `selection`
      ... and pulls from `source.At(...)`" — raw. `ArrayShaping.Normalize` is applied only to SORT's key
      vector (item 9) and UNIQUE's key tuples (item 10); item 8 (FILTER) never mentions it. Measured that
      RangeOperand preserves Blank: `A5:A8` -> `[Number:7, Blank, Text:t, Number:7]` (4x1), exactly as item
      4's rationale states. Excel's dynamic arrays return 0 for a blank source cell, so `COUNTA(FILTER(A5:A8,
      …))` and `COUNTA(UNIQUE(A5:A8))` will skip the blank where Excel counts a 0. Risk #8 asserts
      "`COUNTA(A1:A4)` and `COUNTA(UNIQUE(A1:A4))` can disagree on a blank cell" — with raw output they agree
      with each other and both differ from Excel, so the risk as written describes a seam that does not exist
      while missing the one that does.
      *Correction:* Apply `ArrayShaping.Normalize` in `AxisSelectionOperand.At` (and to FILTER's kept values),
      which keeps the change confined to the four producers exactly as item 4 intends, and rewrite risk #8 to
      say the seam is between a RANGE argument (Blank preserved) and a PRODUCER result (Blank -> 0), which is
      Excel's own seam.

## Implementation items

- [ ] **1.** Add `internal interface IArrayProducer` to Danfma.MySheet/Expressions/ArrayEvaluation.cs (immediately above `internal abstract class ArrayOperand` at :215) with two members: `(bool Succeeds, bool IsArray) ProbeArray();` (pure syntax, NEVER evaluates) and `bool TryBuildArrayOperand(EvaluationContext context, out ArrayEvaluation.ArrayOperand operand);`. Widen `ArrayEvaluation.Probe` (:123) and `ArrayEvaluation.TryBuildOperand` (:420) from `private static` to `internal static` so an implementation in another file can recurse into its own child arguments. Because the interface and `ArrayOperand` are internal while the four function records are public, every implementation MUST use EXPLICIT interface implementation (`bool IArrayProducer.TryBuildArrayOperand(...)`) — an implicit public method mentioning `ArrayOperand` is a CS0050 inconsistent-accessibility error.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* One interface + two switch arms is the whole extension point; every future array producer
      (TRANSPOSE, SORTBY, TAKE, HSTACK) then needs zero edits to ArrayEvaluation.cs. Mirrors the codebase's
      existing extension-by-shared-abstraction style (`INumericFold` at NumericAggregation.cs:9-12). Proved
      subclassable: a probe assembly derived from `ArrayEvaluation.ArrayOperand` and constructed
      `ArrayEvaluation.ArrayStream` successfully (2x3 SEQUENCE-like operand enumerated [1,2,3,4,5,6],
      ElementAt(4)=5).
- [ ] **2.** Add the two dispatch arms: in `ArrayEvaluation.Probe` insert `case IArrayProducer producer: return producer.ProbeArray();` immediately BEFORE `default:` at ArrayEvaluation.cs:198, and in `ArrayEvaluation.TryBuildOperand` insert `case IArrayProducer producer: return producer.TryBuildArrayOperand(context, out operand);` immediately BEFORE `default:` at :477. Position matters: after the `OpenRangeReference` cost-guard arms (:139/:441) and after the `Row`/`If` arms, so the new nodes never shadow them.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* `Probe`'s documented invariant at :113-117 is `IsArrayEligible == (build succeeds as array)`;
      NumericAggregation.Fold:103-124 relies on it to avoid double-evaluating a volatile argument. Two
      mirrored arms keep that lockstep mechanically, and the arm order preserves the existing open-range
      refusal.
- [ ] **3.** Add `internal static ComputedValue FirstElement(Expression expression, EvaluationContext context)` to ArrayEvaluation.cs next to `IsArrayEligible` (:118): build the operand, return `operand.At(0, operand.Rows, operand.Columns)` when `IsArray`, else `ComputedValue.Error(Error.Value)`. This is Excel's `@`-on-an-array rule and becomes the single `Evaluate` body of all four records.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* VERIFIED verbatim from support.microsoft.com "Implicit intersection operator: @": for an array
      Excel "picks the top-left value"; for a range it "returns the value from the cell on the same row or
      column as the formula". MySheet already does exactly this in the one case it can today — measured
      `=ROW(A1:A3)` in D2 returns 1, the array's top-left. One shared helper means the rule cannot drift
      across the four records, and Phase 5 (S4 FIX B) can reuse it for the array half of the cell boundary.
- [ ] **4.** Create Danfma.MySheet/Expressions/ArrayShaping.cs holding three shared pieces, all `internal`: (a) `enum ArrayAxis { Rows, Columns }`; (b) `sealed class SingletonArrayOperand : ArrayEvaluation.ArrayOperand` — a 1x1 array wrapping one `ComputedValue` (`IsArray => true`, `Rows => 1`, `Columns => 1`); (c) `sealed class AxisSelectionOperand : ArrayEvaluation.ArrayOperand` — (source operand, axis, `int[] selection`) whose `At(index, rows, columns)` maps the row-major index through `selection` on the chosen axis and pulls from `source.At(...)`, returning `ComputedValue.Error(Error.Value)` on the shape-mismatch guard exactly as `RangeOperand.At` does at ArrayEvaluation.cs:284-287. Also add `static ComputedValue Normalize(in ComputedValue value)` mapping `Blank` to `Number(0)` and passing everything else through.
      *Files:* `Danfma.MySheet/Expressions/ArrayShaping.cs`
      *Why:* FILTER, SORT and UNIQUE are the same operation — a selection or permutation of one axis — so one
      operand class serves all three; only their `int[]` builders differ. `SingletonArrayOperand` is required
      because `ScalarOperand` (:232-244) has `IsArray => false` and therefore can never be a RESULT.
      `Normalize` encodes Excel's dynamic-array blank-to-zero rule; it is confined to the four new functions
      because `RangeOperand.At` deliberately preserves `Blank` (measured: `A5:A8` yields [Number:7, Blank,
      Text:t, Number:7]) and changing that would alter existing AVERAGE/COUNT results.
- [ ] **5.** Establish and comment the hard invariant in ArrayShaping.cs: an `ArrayOperand` produced by an `IArrayProducer` MUST have `Rows >= 1 && Columns >= 1` — an empty result is expressed as a 1x1 `SingletonArrayOperand` carrying `#CALC!`, never as a 0-extent array. Add `Debug.Assert` in both operand constructors.
      *Files:* `Danfma.MySheet/Expressions/ArrayShaping.cs`
      *Why:* Measured directly: a 0x1 operand gives `ArrayStream.Length == 0` and `foreach` yields nothing, so
      a SUM consumer would silently return 0 instead of an error — the exact class of silent-wrong-answer bug
      this phase must not ship. `ArrayStream.Length` is `Rows * Columns` (ArrayEvaluation.cs:604) with no
      emptiness channel, so the invariant is the only defence.
- [ ] **6.** Add `#CALC!` to the error set: in Danfma.MySheet/Error.cs append `"#CALC!"` at index 7 of the `Displays` array (:14-24), add `public static readonly Error Calc = new(7);` after `NA` (:33), add `"#CALC!" => Calc` to `FromDisplay` (:62-74), and add `7 => ErrorValue.Calculation` to `ToErrorValue` (:77-86). In Danfma.MySheet/Expressions/ErrorValue.cs add `public static readonly ErrorValue Calculation = new("#CALC!");` alongside the existing singletons (:7-16).
      *Files:* `Danfma.MySheet/Error.cs`, `Danfma.MySheet/Expressions/ErrorValue.cs`
      *Why:* Microsoft's FILTER article states verbatim: "If your dataset has the potential of returning an
      empty value, then use the 3rd argument ([if_empty]). Otherwise, a #CALC! error will result, as Excel
      does not currently support empty arrays." Returning #VALUE! instead would be observably wrong (P0). The
      change is contained: `ExcelExport.ErrorText` (ExcelExport.cs:296-298) writes `error.ToString()`, so xlsx
      gets a valid `t="e"` `#CALC!`; on load `WorksheetStreamLoader.GetOrAddError` (:717-740) falls to its
      per-load cache and `Error.FromDisplay` maps it back, so the round-trip closes. Measured today:
      `Error.FromCode(7).Display` is `"#ERR?"`, so a pre-3.17 build reading a warm snapshot with
      `CachedCellValue.ErrorCode == 7` degrades to #ERR? rather than crashing — fold this into S7's one-way
      boundary subsection.
- [ ] **7.** Create Danfma.MySheet/Expressions/Mathematics/Sequence.cs: `public sealed partial record Sequence(Expression[] Arguments) : Function, IArrayProducer`, `[MemoryPackable]`, `Evaluate => ArrayEvaluation.FirstElement(this, context)`. `IArrayProducer.ProbeArray() => (true, true)` (arity is already registry-checked at parse time). `IArrayProducer.TryBuildArrayOperand` evaluates rows/[columns]/[start]/[step] ONCE each via `.Evaluate(context).CoerceToNumber(out …)`, defaulting every omitted optional to 1; any coercion error becomes a 1x1 `SingletonArrayOperand(ComputedValue.Error(thatError))`; `Math.Truncate` rows and columns; `rows < 1 || columns < 1` yields 1x1 `#CALC!`; `rows > 1_048_576 || columns > 16_384 || rows * columns > 1_048_576` yields 1x1 `#NUM!` (MySheet's cost guard). Add a `file sealed class SequenceOperand : ArrayEvaluation.ArrayOperand` in the same file whose `At(index, rows, columns)` returns `ComputedValue.Number(start + index * step)` after the shape-mismatch guard. Implement this FIRST as the walking skeleton.
      *Files:* `Danfma.MySheet/Expressions/Mathematics/Sequence.cs`
      *Why:* SEQUENCE is the only one of the four with no input array, so it proves the whole producer channel
      end to end with nothing else in play: once `SUM(SEQUENCE(5))` returns 15 and `INDEX(SEQUENCE(2,3),2,2)`
      returns 5, the interface, both switch arms, the registry entry, the union tag and the un-parse are all
      validated. Row-major `start + index*step` is confirmed against Excel's documented `=SEQUENCE(4,5)`
      layout and matches my probe of a 2x3 operand ([1,2,3,4,5,6]). Missing optional arguments defaulting to 1
      is verbatim from the SEQUENCE article ("Any missing optional arguments will default to 1").
      Zero/negative to #CALC! per Microsoft's #CALC! guidance ("SEQUENCE can't spill an array with 0 or
      negative values"). The size cap is an explicit, documented MySheet deviation justified by the absence of
      a spill/grid bound — without it `SEQUENCE(1e6,1e4)` hangs the consumer, since `ArrayStream` is lazy but
      the consumers iterate every element.
- [ ] **8.** Create Danfma.MySheet/Expressions/Lookup/Filter.cs: `public sealed partial record Filter(Expression[] Arguments) : Function, IArrayProducer`. `ProbeArray()` returns `(false,false)` if `ArrayEvaluation.Probe` of args[0], args[1] or (when present) args[2] does not `Succeed`, else `(true,true)`. `TryBuildArrayOperand`: build source and include via `ArrayEvaluation.TryBuildOperand` (return false only when a child build fails); pick the axis — `ArrayAxis.Rows` when `include.Rows == source.Rows && include.Columns == 1`, else `ArrayAxis.Columns` when `include.Columns == source.Columns && include.Rows == 1`, else 1x1 `#VALUE!`; walk the include vector once, propagating the FIRST element error as a 1x1 array and turning a `CoerceToBool` failure into that same 1x1 error; collect kept positions into `int[]`; when nothing is kept return `new SingletonArrayOperand(Arguments.Length == 3 ? Arguments[2].Evaluate(context) : ComputedValue.Error(Error.Calc))`; otherwise `new AxisSelectionOperand(source, axis, kept)`.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Filter.cs`
      *Why:* Signature FILTER(array, include, [if_empty]) and the empty/error rules are verbatim from
      support.microsoft.com's FILTER article (include is "a Boolean array whose height or width matches the
      array"; "If any value of the include argument is an error (#N/A, #VALUE, etc.) or cannot be converted to
      a Boolean, the FILTER function will return an error"). The height-or-width wording is why both axes are
      supported. Dimension mismatch to #VALUE! is from secondary sources, not the MS article — see
      openQuestions. Reading the include vector at BUILD time is what makes the data-dependent length
      representable at all: `ArrayOperand.Rows`/`Columns` are properties read after construction
      (ArrayEvaluation.cs:217-219), so the shape must be fixed by then. The build must NEVER return false for
      a semantic error, because `IsArrayEligible` returning true followed by a failed build makes
      NumericAggregation.Fold:103-124 fall through to `argument.Evaluate(context)` — a second evaluation of
      the same subtree.
- [ ] **9.** Create Danfma.MySheet/Expressions/Lookup/Sort.cs: `public sealed partial record Sort(Expression[] Arguments) : Function, IArrayProducer`. `TryBuildArrayOperand`: build source; read `[by_col]` (default FALSE) to choose the axis; read `[sort_index]` (default 1, truncated) and return 1x1 `#VALUE!` when it is `< 1` or exceeds the cross-axis extent; read `[sort_order]` (default 1) and return 1x1 `#VALUE!` when it is neither 1 nor -1; materialize the key vector by reading `source.At` along the key column (or row), returning the FIRST key error as a 1x1 array; build `int[] permutation = [0..n)` and `Array.Sort` it with a comparer of `ArrayShaping.Normalize`d keys via `ValueCoercion.Compare`, negated for descending, tie-broken by the original index for stability; return `new AxisSelectionOperand(source, axis, permutation)`.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Sort.cs`
      *Why:* Signature and defaults are verbatim from the SORT article: sort_index default 1, sort_order
      default 1 ("1 for ascending, -1 for descending"), by_col default FALSE ("By default Excel will sort by
      row, and will only sort by column where by_col is TRUE"). #VALUE! for a sort_order other than 1/-1 is
      stated on the sibling SORTBY page. Reusing `ValueCoercion.Compare` (ValueCoercion.cs:171-186) is the key
      reuse win: its doc comment already describes Excel's classic order — "number < text < boolean (FALSE
      before TRUE); same-type text compares case-insensitively; blank counts as 0" — so no new comparer is
      invented and SORT's ordering is automatically consistent with `<`/`>`, MATCH and VLOOKUP's approximate
      match. Propagating a key error up front honours that method's stated contract ("Callers propagate errors
      first", :169). Explicit index tie-break because `Array.Sort` is not stable.
- [ ] **10.** Create Danfma.MySheet/Expressions/Lookup/Unique.cs: `public sealed partial record Unique(Expression[] Arguments) : Function, IArrayProducer`. `TryBuildArrayOperand`: build source; read `[by_col]` (default FALSE) and `[exactly_once]` (default FALSE); materialize the key tuples along the chosen axis with `ArrayShaping.Normalize`, propagating the first key error as a 1x1 array; group positions by first appearance using a `Dictionary<int, List<int>>` keyed on a hash of the normalized tuple with a bucket-local equality check via `ValueCoercion.AreEqual`; select the first index of each group (or, when `exactly_once`, the index of each group of size 1) in first-appearance order; nothing kept yields 1x1 `#CALC!`; otherwise `new AxisSelectionOperand(source, axis, kept)`.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Unique.cs`
      *Why:* Signature and both flags are verbatim from the UNIQUE article ("TRUE will compare columns against
      each other and return the unique columns"; "TRUE will return all distinct rows or columns that occur
      exactly once"). `ValueCoercion.AreEqual` gives Excel's `=` semantics (case-insensitive text, type-
      distinct so `1` and `"1"` stay separate) at zero cost. The `ArrayShaping.Normalize` step is load-
      bearing, not cosmetic: `AreEqual` makes blank equal to BOTH `0` and `""` (:130-138) while `0` and `""`
      are unequal, so raw equality is NON-TRANSITIVE and no hash bucketing is sound. Normalizing blank to 0
      first — Excel's own dynamic-array rule — restores transitivity and makes the O(n) hashed dedupe correct;
      without it the only correct algorithm is an O(n^2) scan.
- [ ] **11.** Register the four in Danfma.MySheet/Parsing/FunctionRegistry.cs `Entries` (:48-2135) with `Entry<T>(name, minArgs, maxArgs, create, getArguments)`: `("FILTER", 2, 3)`, `("SORT", 1, 4)`, `("UNIQUE", 1, 3)`, `("SEQUENCE", 1, 4)` — SORT/UNIQUE/FILTER next to the other Lookup entries, SEQUENCE next to the Math ones, each `static arguments => new X(arguments)` and `static f => ((X)f).Arguments`.
      *Files:* `Danfma.MySheet/Parsing/FunctionRegistry.cs`
      *Why:* This single registration is what makes the Parser accept the names (Parser.cs:629), what makes
      `FormulaWriter.Call` un-parse them via `FunctionRegistry.ByType` (FormulaWriter.cs:434-445), what makes
      `DependencyExtractor.VisitArguments` (:223-247) enumerate their arguments, and what makes
      `AnchoredFormulaSupport.IsFullyAnchored` (:57-58) accept them — four capabilities from one line each.
      Arity is enforced at PARSE time, so `FILTER(A1:A3)` throws `ParseErrorKind.InvalidArgumentCount` and an
      xlsx load keeps degrading to the cached value plus the UnparsableFormula warning.
- [ ] **12.** Reserve four consecutive `[MemoryPackUnion]` tags for `Filter`, `Sort`, `Unique`, `Sequence` immediately after `[MemoryPackUnion(321, typeof(SharedFormulaSlave))]` at Danfma.MySheet/Expressions/Expression.cs:352, with a comment naming this phase. Do NOT hard-code 322-325 while writing this phase's code: assign the concrete numbers in ONE coordinated edit once the table-model, aggregate and reference phases have taken theirs, and record the allocation in plans/dynamic-array-functions.md as the single source of truth.
      *Files:* `Danfma.MySheet/Expressions/Expression.cs`, `plans/dynamic-array-functions.md`
      *Why:* Tags 0..321 are contiguous with no gaps and the policy comment at :14-16 ("Add new tags at 319+")
      is already stale. If two phases each independently claim 322 the MemoryPack union registration collides
      — loudly, not silently, but it costs a rebase across four phases. One coordinated edit at the end is
      cheaper than four independent guesses.
- [ ] **13.** Add the three `ReferenceGuard` arms: in Danfma.MySheet/Expressions/ReferenceGuard.cs insert `case Filter f: return MissingSheet(f.Arguments[0], context);`, `case Sort s: return MissingSheet(s.Arguments[0], context);` and `case Unique u: return MissingSheet(u.Arguments[0], context);` immediately BEFORE `default: return null;` at :96, with a comment that these three stand for their source array the way the `UnaryOperation{Plus}` arm (:82-84) stands for its operand. SEQUENCE needs no arm (no reference argument).
      *Files:* `Danfma.MySheet/Expressions/ReferenceGuard.cs`
      *Why:* Without them the error-IGNORING COUNT family silently reports an empty result for a deleted sheet
      — exactly the failure the file's own doc comment (:5-13) says the guard exists to prevent. Measured pre-
      existing hole: `SUM(IF(Ghost!A1:A3>0,1))` correctly gives #REF! (the per-cell errors reach SUM's error
      channel) but `COUNT(IF(Ghost!A1:A3>0,1))` gives 0. These arms make
      `COUNT(FILTER(Ghost!A1:A3,Ghost!B1:B3>0))` a structural #REF! as Excel gives. The equivalent
      `If`/`BinaryOperation` hole is pre-existing and out of scope for this phase.
- [ ] **14.** Add the mini-CSE gate to `Rows.Evaluate` (Danfma.MySheet/Expressions/Lookup/Rows.cs:11, after the `ReferenceGuard.MissingSheet` guard and before the `NamedReferences.TryResolveReference` ternary) and to `Columns.Evaluate` (Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs:307, same position): `if (Arguments[0] is not Reference && ArrayEvaluation.IsArrayEligible(Arguments[0]) && ArrayEvaluation.TryEvaluateStream(Arguments[0], context, out var array)) return ComputedValue.Number(array.Rows);` (`array.Columns` for COLUMNS). Refactor each `Evaluate` from its expression body to a statement body to accommodate the guard.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Rows.cs`, `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`
      *Why:* `ROWS(FILTER(range, cond))` is THE Excel idiom for "how many rows matched", and without this gate
      FILTER is largely unusable. Measured today: `ROWS((A1:A3>0)*1)` returns 1 and `COLUMNS(A1:B3*2)` returns
      1 — both wrong for any array. The gate is copied verbatim in shape from `Index.Evaluate`
      (Index.cs:22-29), the established precedent for a reference-or-array consumer. Note this also FIXES
      existing array forms (`ROWS(IF(A1:A3>0,A1:A3))` goes 1 to 3), so it is a behaviour change that needs a
      changelog line. Distinct from S8's ROW/COLUMN work, which is about reference-PRODUCING expressions, not
      computed arrays.
- [ ] **15.** Add the mini-CSE arm to `ArgumentFlattening.FlattenComputedValues` (Danfma.MySheet/Expressions/ArgumentFlattening.cs:63, at the top of the `default:` branch, before `argument.Evaluate(context)`): `if (ArrayEvaluation.IsArrayEligible(argument) && ArrayEvaluation.TryEvaluateStream(argument, context, out var array)) { foreach (var element in array) yield return element; break; }`. Extract it as a `private static bool TryStreamArray(Expression, EvaluationContext, out ArrayEvaluation.ArrayStream)` helper in the same file so the aggregate phase can call the identical helper from `ExpandComputedValues`.
      *Files:* `Danfma.MySheet/Expressions/ArgumentFlattening.cs`
      *Why:* `FlattenComputedValues` has exactly five call sites — CountA.cs:21, CountBlank.cs:18,
      Concat.cs:20, Concatenate.cs:20, TextJoin.cs:29 — so one arm lights all five up at once, which is the
      project's stated preference for a shared helper over per-function branching. It is what makes
      `COUNTA(UNIQUE(range))`, Excel's canonical distinct-count idiom, work at all: `COUNT(UNIQUE(...))`
      counts only numbers and returns 0 for text. Measured today: `COUNTA(IF(A1:A3>0,A1:A3))` gives 1 (Excel:
      3), `CONCAT(A1:A3*2)` and `TEXTJOIN(",",TRUE,A1:A3*2)` give #VALUE! (Excel: "10018" / "10,0,18").
      Behaviour changes, all in the Excel-correct direction; needs tests plus a changelog line. The extracted
      helper is the coordination point with the aggregate phase's SUMPRODUCT gate (S8), which needs the same
      arm in `ExpandComputedValues`.
- [ ] **16.** Write plans/dynamic-array-functions.md as the design of record: the producer-channel decision and why no `ComputedValueKind` was added; the AxisSelection unification of FILTER/SORT/UNIQUE; the per-function Excel citations; the build-time-read extension to the laziness contract; the `@` top-left rule with its Microsoft citation; the union-tag allocation table; and the four documented deviations (no spill, open-range cost guard, blank-to-zero confined to these four, SEQUENCE size cap).
      *Files:* `plans/dynamic-array-functions.md`
      *Why:* plans/ is the established convention (38 files) and plans/mini-cse-array-arguments.md is the
      direct predecessor this phase extends; the executing agent needs the Excel citations and the deviation
      list in one place rather than reconstructing them.
- [ ] **17.** Create tests/Danfma.MySheet.Tests/Parsing/DynamicArrayTests.cs with a local `private static object? Calc(string formula, params (string Id, object Value)[] cells)` harness copied from MathAggregateTests.cs:11-31, and a file header citing the four Microsoft pages plus the fetch date, per the mandatory golden-value convention (MathAggregateTests.cs:6-8). Cover per function: SEQUENCE shape and row-major order (`SUM(SEQUENCE(5))`=15, `INDEX(SEQUENCE(2,3),2,2)`=5, `SUM(SEQUENCE(3,1,10,5))`=45), zero/negative to #CALC!, the size cap to #NUM!; FILTER over a column (`SUM(FILTER(A1:A3,A1:A3>0))`, `ROWS(FILTER(...))`), empty with and without if_empty (#CALC! vs the supplied value), dimension mismatch to #VALUE!, an error inside include propagating; SORT ascending/descending, `INDEX(SORT(A1:A3),1)`, sort_index out of range and sort_order 0 both #VALUE!; UNIQUE first-appearance order, case-insensitive text collapse, `exactly_once`, `COUNTA(UNIQUE(...))`.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/DynamicArrayTests.cs`
      *Why:* Function-behaviour tests live in tests/Danfma.MySheet.Tests/Parsing/<Family>Tests.cs with a per-
      file duplicated `Calc` harness and no shared base class; the Microsoft-page-plus-fetch-date citation is
      mandatory in this repo. The #CALC!/#VALUE! cases are the ones most likely to be quietly implemented as
      the wrong error, so they need explicit golden assertions.
- [ ] **18.** Extend tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs with the PRODUCER direction: each of the four inside SUM, COUNT, AVERAGE, MIN/MAX, SMALL/LARGE, MEDIAN and INDEX; the compositions `SUM(SORT(FILTER(A1:A5,A1:A5>0)))`, `SUM(FILTER(A1:A5,A1:A5>0)*2)` and `INDEX(UNIQUE(A1:A5),1)` (proving the recursive builder composes producers with each other and with BinaryOperand); the new ROWS/COLUMNS/COUNTA/TEXTJOIN gates; and the open-range refusal `SUM(FILTER(A:A,A:A>0))` pinned as #VALUE!. Extend MiniCseVolatileTaintTests.cs with `SUM(FILTER(A1:A3,A1:A3>RAND()))` asserting the cell is volatile-tainted and refreshed by `Recalculate()`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`, `tests/Danfma.MySheet.Tests/Expressions/MiniCseVolatileTaintTests.cs`
      *Why:* MiniCseConsumerTests.cs is where an array-consuming/producing function belongs and its existing
      tests only cover the CONSUMER direction. The composition tests are the cheapest proof that the design's
      main claim holds. The volatile test matters because FILTER evaluates its include vector at BUILD time —
      the taint must still land in the enclosing cell frame, which is the property ArrayEvaluation.cs:88-90
      promises for scalar operands and this phase extends to predicate vectors. The open-range case must be
      PINNED, not left implicit, because it is a deliberate deviation from Excel.
- [ ] **19.** Extend tests/Danfma.MySheet.Tests/DirtyGraph/DependencyExtractorTests.cs to PIN that no change is needed there: `SUM(FILTER(A1:A3,B1:B3>0))` contributes exactly the A1:A3 and B1:B3 `RangeDep`s with `AlwaysDirty == false`, and `SUM(SEQUENCE(5))` contributes no ranges and is not always-dirty — contrasting with `Indirect_IsAlwaysDirty` at :68. Extend the shared-formula tests to pin `AnchoredFormulaSupport.IsFullyAnchored` accepting a master containing `FILTER($A$1:$A$3, $B$1:$B$3>0)`. Add the four identity strings (`"FILTER(A1:A3,B1:B3)"`, `"SORT(A1:A3)"`, `"UNIQUE(A1:A3)"`, `"SEQUENCE(3)"`) to the flat parse-to-write list in tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs near :268.
      *Files:* `tests/Danfma.MySheet.Tests/DirtyGraph/DependencyExtractorTests.cs`, `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`
      *Why:* I measured that `DependencyExtractor.Visit`'s generic `case Function function:` (:207) plus
      `VisitArguments` (:223-247) already produce the correct RangeDeps for `SUM(IF(A1:A3>0,B1:B3))` with
      AlwaysDirty=false, so the four need ZERO extractor work — but that is load-bearing and non-obvious (the
      `default: return;` at :216-217 is a silent lost-dependency trap for a node the registry does NOT cover),
      so it must be pinned by a test rather than left to a future reader. FILTER's data-dependent LENGTH does
      not affect its dependency SET — the whole source range is read either way — which is precisely the
      super-approximation the file documents at :38-42, so AlwaysDirty would be a needless pessimization.
- [ ] **20.** Update docs/function-reference.md: recompute the header count at :3 from `FunctionRegistry.ByName.Count` rather than assuming 304+4 (it measures 305 today, so the doc is already off by one); change `## Lookup and reference (16)` at :230 to (19) and add FILTER/SORT/UNIQUE rows; change `## Math and trigonometry (74)` at :31 to (75) and add a SEQUENCE row; update the coverage paragraph's "implements 304 of the ~520" (~:409); in the coverage table move FILTER/SORT/UNIQUE from the Lookup-and-Reference not-yet list to implemented and change `16/40` to `19/40`, and move SEQUENCE in Math and Trigonometry changing `74/82` to `75/82` — coordinating with the aggregate phase, which moves AGGREGATE in the same block. Mirror every change in docs/pt-BR/function-reference.md.
      *Files:* `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* The exact strings and counts are confirmed in the file: `**304 built-in functions**` at :3, `##
      Logical (12)`/`## Math and trigonometry (74)` at :14/:31, `## Lookup and reference (16)` at :230,
      `Lookup and Reference — 16/40` and `Math and Trigonometry — 74/82` in the coverage block, with
      FILTER/SORT/UNIQUE and SEQUENCE all currently in the not-yet lists. Excel's own categorisation puts
      FILTER/SORT/UNIQUE in Lookup and Reference and SEQUENCE in Math and Trigonometry, matching where the
      not-yet entries already sit. docs/pt-BR is a confirmed full mirror.
- [ ] **21.** Rewrite docs/workbook-and-expressions.md's "Implicit array arguments" section (:336-371): add the four as PRODUCERS with runnable examples, list the full consumer set (SUM/COUNT/AVERAGE/MIN/MAX/PRODUCT/MEDIAN/STDEV*/percentiles, SMALL/LARGE, INDEX, ROWS/COLUMNS, COUNTA/COUNTBLANK/CONCAT/CONCATENATE/TEXTJOIN), and REPLACE the "dry cell keeps #VALUE!" bullet at :358-361 with the `@` rule: a bare dynamic-array formula yields the array's TOP-LEFT element, i.e. exactly what Excel writes as `=@FILTER(...)`, and MySheet does NOT spill. Add an explicit NO-SPILL limitation block naming the four deviations (no spill; open-range argument refused; blank-to-zero inside these four only; the SEQUENCE size cap). Mirror in docs/pt-BR/workbook-and-expressions.md. Add the new union tags and the `#CALC!`/`CachedCellValue.ErrorCode == 7` note to S7's one-way compatibility subsection in docs/serialization.md and its pt-BR twin.
      *Files:* `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`, `docs/serialization.md`, `docs/pt-BR/serialization.md`
      *Why:* The current bullet at :358-361 states "A dry cell whose whole formula is the array keeps #VALUE!"
      — directly contradicted by the `@` rule this phase adopts, and by S4/Phase 5 for ranges, so it must be
      rewritten in the same pass. The no-spill block is required by S5's own terms: `=SEQUENCE(5)` returning 1
      rather than spilling 1..5 is the single most surprising consequence of this phase and must be stated
      plainly rather than discovered. `docs/pt-BR/README.md:3` establishes English as authoritative but the
      mirror must stay complete (11 files each).

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded, 0 Error(s). A CS0050 "inconsistent accessibility" error on any of the four
      records means IArrayProducer was implemented implicitly instead of explicitly (see item 1). A MemoryPack
      duplicate-union-tag error means the tag allocation collided with another phase.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier check .`
      → expected: No files listed as needing formatting; exit code 0.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj`
      → expected: All tests pass, 0 failed. DynamicArrayTests, MiniCseConsumerTests,
      MiniCseVolatileTaintTests, DependencyExtractorTests and FormulaWriterTests must all be green. Pre-
      existing failures are expected in exactly two places and must be fixed as part of this phase, not
      ignored: any test asserting ROWS/COLUMNS returns 1 for an array-eligible argument, and any test
      asserting COUNTA/CONCAT/TEXTJOIN over an array-eligible argument.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj`
      → expected: All tests pass, 0 failed. Nothing in this phase touches the xlsx loader or exporter, so a
      failure here means Error.Calc broke the error round-trip in WorksheetStreamLoader.GetOrAddError or
      ExcelExport.ErrorText.
- [ ] `cd /tmp/probe-dynamic-arrays && dotnet build -c Release -v q --nologo && dotnet run -c Release --no-build 2>&1 | sed -n '/=== C\./,/=== D\./p'`
      → expected: After this phase: ROWS((A1:A3>0)*1) -> 3, COLUMNS(A1:B3*2) -> 2, COUNTA(IF(A1:A3>0,A1:A3))
      -> 3, CONCAT(A1:A3*2) -> "10018", TEXTJOIN(",",TRUE,A1:A3*2) -> "10,0,18". BEFORE this phase these
      measure 1, 1, 1, #VALUE!, #VALUE! respectively — the probe at /tmp/probe-dynamic-arrays is the
      before/after diff for the gates in items 15 and 16.
- [ ] `cd /Volumes/Work/Develop/MySheet && printf 'SUM(SEQUENCE(5))=15\nINDEX(SEQUENCE(2,3),2,2)=5\nSUM(FILTER(A1:A3,A1:A3>0))=14\nROWS(FILTER(A1:A3,A1:A3>0))=2\nINDEX(SORT(A1:A3),1)=0\nCOUNTA(UNIQUE(A1:A3))=3\n' && grep -c 'MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs`
      → expected: grep prints 326 (322 existing + 4 new). The printed list is the acceptance oracle for the
      fixture A1=5, A2=0, A3=9 and must be asserted in DynamicArrayTests.cs; SUM(FILTER(...)) = 5+9 = 14 and
      ROWS(FILTER(...)) = 2 because A2=0 fails the predicate.

## Risks carried by this phase

- The uncomfortable one first: bare `=SEQUENCE(5)` in a cell will return 1, not spill 1..5. That is Excel's verified `@`-on-an-array behaviour (support.microsoft.com "Implicit intersection operator: @": Excel "picks the top-left value"), and it is exactly what Excel writes when it upgrades a legacy formula — but it is NOT what modern Excel does with the un-prefixed formula. A user who types `=FILTER(A:A,B:B>0)` and sees one value will read it as a bug. This must be the most prominent line in the docs, and #VALUE! is a defensible alternative (see openQuestions).
- SORT's ordering of MIXED types is the least verifiable semantic in this phase. Microsoft's SORT article says nothing about how numbers, text, logicals and blanks interleave; I am reusing `ValueCoercion.Compare` (ValueCoercion.cs:171-186) whose documented order is number < text < FALSE < TRUE with blank as 0. If Excel differs, `INDEX(SORT(x),1)` silently returns the wrong element — a wrong value, not an error. Real Excel usage is overwhelmingly a homogeneous key column, where the order is unambiguous, so I recommend shipping; but SORT should be implemented LAST of the four and its mixed-type behaviour marked as unverified in the docs.
- `ArrayOperand.Rows`/`Columns` are properties read AFTER construction (ArrayEvaluation.cs:217-219), so FILTER/SORT/UNIQUE must read their predicate or key vector at BUILD time. That extends the laziness contract documented at ArrayEvaluation.cs:83-91 from "scalar sub-expressions evaluated once" to "predicate/key vectors read once". Cost: FILTER pays one pass over include plus an `int[]` (200KB at 50k rows); SORT pays one pass over the key column plus a `ComputedValue[]`; UNIQUE materializes ALL columns of its input (~1.2MB at 50k rows x 1 column). Only UNIQUE genuinely loses the no-allocation property the mini-CSE was built for. This must be stated in plans/dynamic-array-functions.md rather than discovered by a benchmark regression.
- A 0-extent array is a silent-wrong-answer trap I measured directly: a 0x1 operand yields `ArrayStream.Length == 0`, `foreach` yields nothing, and a SUM consumer returns 0 with no error. If any implementer returns a 0-row `AxisSelectionOperand` for an empty FILTER instead of the 1x1 `#CALC!` singleton, `SUM(FILTER(A1:A3,A1:A3>100))` becomes 0 instead of #CALC! and no test outside DynamicArrayTests will catch it.
- Returning `false` from `TryBuildArrayOperand` for a SEMANTIC error (bad sort_order, dimension mismatch) instead of a 1x1 error array breaks the `IsArrayEligible => build succeeds` invariant that NumericAggregation.Fold:103-124 depends on, and the consumer then falls through to `argument.Evaluate(context)` — a SECOND evaluation of the same subtree, which double-draws a volatile. The only legitimate `false` is a child build failing on the open-range cost guard.
- `COUNT(...)` discards NumericAggregation.Fold's error channel by design (Count.cs:16-25), so a 1x1 error array is invisible to it: `COUNT(SEQUENCE(-1))` will return 0 where Excel gives #CALC!. Measured precedent for the same class: `COUNT(IF(Ghost!A1:A3>0,1))` returns 0 today. There is no error channel out of `ArrayStream` to fix this without changing `TryEvaluateStream`'s signature; accept and document.
- `SUM(FILTER(A:A, A:A>0))` will be #VALUE!, because ArrayEvaluation refuses an open range in an array position (the pre-existing cost guard at :139/:441) and my `ProbeArray` propagates that refusal. Excel works. This is likely the second-most-reported gap after no-spill, and whole-column FILTER is a very common real-world shape.
- `ArrayShaping.Normalize` (blank to 0) applies inside the four new functions but NOT to `RangeOperand.At`, which measurably preserves `Blank` (`A5:A8` yields [Number:7, Blank, Text:t, Number:7]). So `COUNTA(A1:A4)` and `COUNTA(UNIQUE(A1:A4))` can disagree on a blank cell. The inconsistency is deliberate — normalizing RangeOperand would change existing AVERAGE/COUNT results across the whole engine — but it is a genuine seam a user can hit.
- Naming: the record `Sort` in namespace `Danfma.MySheet.Expressions.Lookup` shadows the bare identifier `Sort` inside that namespace. Nothing in Lookup/ currently calls a bare `Sort`, but `Array.Sort`/`List<T>.Sort` inside Sort.cs itself must stay qualified or instance-invoked or the compiler will bind to the record.

## Open questions owned by this phase

- Bare-in-a-cell: top-left (`@` semantics) or #VALUE!? I verified Excel's `@` returns the array's top-left value and recommend it, because it composes with S4 and is what Excel writes into upgraded legacy files. But the dimension brief says "bare-in-a-cell -> #VALUE! per S4" and S4's own fall-through ("no intersection, or a 2-D range") arguably covers a computed array, which has no worksheet position to intersect with. This is a one-line change in `ArrayEvaluation.FirstElement` either way, but it must be decided BEFORE the docs are written. A user decision, not an experiment.
- Does Excel's SORT propagate an error found in the KEY column, or sort the error into position? I propagate it as a whole-array error, which honours `ValueCoercion.Compare`'s stated contract ("Callers propagate errors first", ValueCoercion.cs:169) but may differ from Excel, which I suspect places errors last. Same question for a key error in UNIQUE. Settleable only by a real .xlsx opened in Excel or by Aspose.Cells; no Microsoft doc covers it.
- Is Excel's SORT stable for tied keys (first-appearance order preserved)? I implement an explicit index tie-break. Undocumented; low risk but unverified.
- Is UNIQUE's text comparison case-INSENSITIVE? The UNIQUE article is silent. I assume insensitive via `ValueCoercion.AreEqual`, consistent with every other comparison in Excel and in this engine. Needs an oracle check.
- FILTER dimension mismatch: is it #VALUE!? The Microsoft FILTER article does NOT state an error for mismatched `array`/`include` dimensions (it only says include's "height or width matches the array"); the #VALUE! answer comes from secondary sources. Also unresolved: does a SCALAR include (`FILTER(A1:A10, TRUE)`) count as a mismatch, or does Excel broadcast it? I treat it as a mismatch.
- SEQUENCE with 0 or negative rows/columns: #CALC! or #VALUE!? Microsoft's #CALC! guidance says SEQUENCE "can't spill an array with 0 or negative values", which reads as #CALC! for both, but I did not find the negative case stated explicitly. Also: are non-integer rows/columns truncated or rounded? I truncate.
- MySheet's SEQUENCE size cap (`rows*columns > 1_048_576` -> #NUM!) is an invented behaviour with no Excel counterpart — Excel bounds SEQUENCE by the grid via #SPILL!, which MySheet has no model for. Is #NUM! the right error, should it be #VALUE!, or should the cap simply be higher? Needed to avoid `SUM(SEQUENCE(1e6,1e4))` hanging the consumer.
- Should FILTER/SORT/UNIQUE BOUND an open-range argument (via `NamedReferences.TryResolveReference(..., boundOpenRanges: true)`, as ROWS and VLOOKUP already do) instead of refusing it? That would make the very common `FILTER(A:A, B:B>0)` work. I did not design it because array and include would be bounded INDEPENDENTLY and could disagree on row count, producing a spurious #VALUE!. Resolving it properly needs a shared-bounds rule and is a candidate follow-up phase.
- Coordination, not a semantic question: the aggregate phase (S8/SUMPRODUCT) needs the same array arm in `ArgumentFlattening.ExpandComputedValues` that this phase adds to `FlattenComputedValues`. Whichever phase lands first should extract the shared `TryStreamArray` helper; if both add it independently the second will conflict. Same for the four union tags and the docs/function-reference.md Math counts (this phase 74->75 for SEQUENCE, the aggregate phase ->76 for AGGREGATE).

## Phase Summary

_(write when phase completes)_
