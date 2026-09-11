# Dynamic array functions — FILTER, SORT, UNIQUE and SEQUENCE as mini-CSE producers

The design of record for Phase 7 of
[Structured table references, AGGREGATE, and the blocking reference-semantics gaps](structured-table-references-and-aggregate.md)
(phase file: [`phase-7-dynamic-arrays.md`](structured-table-references-and-aggregate/phase-7-dynamic-arrays.md)).
It is the direct successor of [`mini-cse-array-arguments.md`](mini-cse-array-arguments.md), which built the
element-wise evaluation this phase plugs into, and of Phases 8 (elementwise lifting) and 10 (vector
broadcasting), which widened it.

This file exists so that the union-tag allocation, the Excel citations and the deviation list live in ONE
place instead of being reconstructed from four function files. Where it states a behaviour, it names the
test that enforces it or the measurement that produced it.

**Governing rule (P0 with its 2026-09-09 addendum).** MySheet matches Excel, and "Excel" means
**Aspose.Cells 26.6.0 as measured** — a measurement beats a Microsoft page. A divergence is a work item,
never a documented limitation, unless the engine structurally cannot match. Two exceptions were granted in
this phase and both are recorded below with the oracle's own numbers, because the oracle contradicts itself
there. Every oracle figure in this file was measured in BOTH entry modes — `plain` (`Cell.Formula`) and
`CSE` (`Cell.SetArrayFormula`) — on 2026-09-10; where one label is given, the two modes agreed. The engine
implements the ARRAY-ENTERED rule everywhere, so a plain-entry figure is never compared against a
MySheet answer.

## What shipped

Four public records, each `[MemoryPackable]`, each implementing the internal `IArrayProducer`:

| Function | Node | Registry arity | Union tag |
| --- | --- | --- | --: |
| `FILTER(array, include, [if_empty])` | `Danfma.MySheet/Expressions/Lookup/Filter.cs` | 2-3 | 323 |
| `SORT(array, [sort_index], [sort_order], [by_col])` | `Danfma.MySheet/Expressions/Lookup/Sort.cs` | 1-4 | 324 |
| `UNIQUE(array, [by_col], [exactly_once])` | `Danfma.MySheet/Expressions/Lookup/Unique.cs` | 1-3 | 325 |
| `SEQUENCE(rows, [columns], [start], [step])` | `Danfma.MySheet/Expressions/Mathematics/Sequence.cs` | 1-4 | 326 |

Supporting types: `IArrayProducer` beside `ArrayOperand` in
`Danfma.MySheet/Expressions/ArrayOperands.cs`; `ArrayShaping`, `ArrayAxis`, `SingletonArrayOperand` and
`AxisSelectionOperand` in `Danfma.MySheet/Expressions/ArrayShaping.cs`; the three selectors' shared source,
flag and axis helpers in `Danfma.MySheet/Expressions/Lookup/SelectionProducers.cs`; `SequenceOperand` beside
its function. `Error.Calc` (`#CALC!`) in `Danfma.MySheet/Error.cs` and `ErrorValue.Calculation` in
`Danfma.MySheet/Expressions/ErrorValue.cs`.

## Decision 1 — ONE producer channel, and no new `ComputedValueKind`

A producer is an `ArrayOperand` **inside** the existing lazy operand tree, reached by exactly two switch
arms:

- `ArrayEvaluation.Probe` → `case IArrayProducer producer: return producer.ProbeArray(context);`
- `ArrayEvaluation.TryBuildOperand` → `case IArrayProducer producer: return producer.TryBuildArrayOperand(context, out operand);`

Both sit **last before `default:`** and after every other arm. The position is load-bearing and the code
says so: the open-range cost guard and the `NameReference` arm come first because they match different node
types, and the `Function when TryGetLift(...)` arm comes first because a producer is a `Function` too —
only its `ArrayLifting.Consumes` classification keeps it out of the lift. An `Elementwise` producer would be
lifted per element and answer from one cell. Phase 5's `TableReference` arm lands ABOVE the producer arm.

**Why no `ComputedValueKind.Array`.** `ComputedValue` is a struct with no array kind and the per-cell cache
assumes one scalar per cell; an array value escaping into a cell would break that contract. The mini-CSE was
built on exactly that constraint — arrays live only inside one formula's evaluation — and this phase does not
relax it. The consequence is the whole point of the channel: `SUM`, `COUNT`, `AVERAGE`, `INDEX`,
`ROWS`/`COLUMNS`, `AGGREGATE`'s array form and the flattening family read a producer with **no change of
their own**, because they were already reading the operand tree.

**Why an interface rather than four arms.** Every later producer — `TRANSPOSE`, `SORTBY`, `TAKE`, `HSTACK`,
`RANDARRAY` — needs zero edits to `ArrayEvaluation.cs`. It mirrors the codebase's existing
extension-by-shared-abstraction style (`INumericFold`).

**The one mechanical trap.** `IArrayProducer` and `ArrayOperand` are `internal` while the four records are
`public`, so every implementation MUST be EXPLICIT (`bool IArrayProducer.TryBuildArrayOperand(…)`). An
implicit public method mentioning `ArrayOperand` is a CS0050 inconsistent-accessibility error. `Probe` and
`TryBuildOperand` were widened from `private static` to `internal static` so an implementation in another
file can recurse into its own children through them rather than pattern-matching the child node — which is
why a defined name, a table column (Phase 5) or another producer in a child slot is whatever those two say
it is.

**The twins must agree.** The mini-CSE's documented contract is `IsArrayEligible == (the build succeeds as an
array)`, and `NumericAggregation.Fold` relies on it to evaluate a volatile argument exactly once. So
`ProbeArray` answers from SHAPE alone (it never evaluates the node, though like `Probe` it may resolve a
name) and `TryBuildArrayOperand` returns `false` ONLY where the probe answered `Succeeds = false`. A
SEMANTIC failure — a bad `sort_order`, an include of the wrong shape, a `SEQUENCE` size below 1 — is a 1x1
error operand, never a refused build; returning `false` there would make the consumer fall through to a
SECOND evaluation of the same subtree.

**Refusal follows the `BinaryOperation`/`If` side of Phase 8's split, not the lift's.** A refused child (an
open range below — the cost guard) makes the producer refuse: `(false, false)` from the probe, `false` from
the build. The consumer then keeps its scalar path and reaches `ArrayEvaluation.FirstElement`, which answers
`#VALUE!`. A lifted function, by contrast, becomes an opaque scalar. Both sides now exist, so the choice had
to be stated: producers propagate the refusal.

## Decision 2 — FILTER, SORT and UNIQUE are ONE operation

All three are a **selection or permutation of one axis of a source array**, so all three share
`AxisSelectionOperand(source, axis, int[] selection)` and differ only in how `selection` is computed at build
time (kept positions, a sort permutation, first appearances). `SelectionProducers` carries what is left over:
the source rule, the omitted-argument rule, the flag coercion and the axis arithmetic, once instead of three
times.

`AxisSelectionOperand.At` projects through `Broadcasting.TryProject` first, like every array operand, then
decomposes its own index into (row, column) of the SELECTED shape, maps the selected axis through
`selection[]`, and reads the source at the SOURCE's own extent — so a composite source (`A1:A3*2`, another
producer) is read exactly as it would be anywhere else in the tree. A position the selection does not cover
is `#N/A`, Phase 10's rule: `SUM(FILTER(A1:A3,A1:A3>0)*B1:B3)` is `#N/A` with `COUNT` 2 (a 2x1 selection
against a 3x1 range).

`SEQUENCE` has no input array and therefore its own operand: `SequenceOperand.At` = project, then
`start + own * step`. It was implemented FIRST as the walking skeleton — once `SUM(SEQUENCE(5))` was 15 and
`INDEX(SEQUENCE(2,3),2,2)` was 5, the interface, both arms, the registry entry, the union tag and the
un-parse were all validated with nothing else in play.

**Reuse instead of new comparers.** `SORT` orders through `ValueCoercion.Compare` — already exactly Excel's
classic number < text < FALSE < TRUE order, text case-insensitively — so `SORT` cannot drift from `<`, `>`,
`MATCH` and `VLOOKUP`'s approximate match. `UNIQUE` does **not** use `ValueCoercion.AreEqual`: see the
measurement note under `UNIQUE` below.

**Nothing else needed changing.** `DependencyExtractor` (its generic `case Function` plus `VisitArguments`),
`AnchoredFormulaSupport.IsFullyAnchored` and `FormulaWriter.Call` all work off the registry entry, so
registering the four is what makes their ranges enumerable, their anchored form shareable and their text
round-trip. That is load-bearing and non-obvious — the extractor's `default: return;` is a silent
lost-dependency trap for a node the registry does not cover — so it is pinned by a test that compares each
of the four against an UNREGISTERED twin (`FILTERX`, `SEQUENCEX`), which IS always-dirty over the same
ranges. The "not always dirty" verdict is therefore the classification and not a default.

## Decision 3 — the build-time read, and what it extends

`ArrayOperand.Rows`/`Columns` are properties read AFTER construction, so a producer's SHAPE must be fixed by
the time the constructor returns. That is what makes a data-dependent length representable at all:

- `FILTER` reads its `include` vector ONCE at build time, fixing the kept count; the kept VALUES stay on
  demand.
- `SORT` reads its key column (or row) once and fixes the permutation; the values stay on demand.
- `UNIQUE` materializes the key tuples along the chosen axis once; the values stay on demand.
- `SEQUENCE` evaluates `rows`, `columns`, `start` and `step` once each.

This **extends** the laziness contract from "scalar sub-expressions are evaluated once at build" to
"scalar sub-expressions AND predicate/key vectors are read once at build". The volatile-taint property is
unchanged and pinned: a `RAND()` inside a predicate draws once and its taint lands in the enclosing cell
frame.

**Cost, stated here rather than discovered by a benchmark regression.** `FILTER` pays one pass over
`include` plus an `int[]`; `SORT` pays one pass over the key vector plus a `ComputedValue[]`; `UNIQUE`
materializes ALL columns of its input along the compared axis. Only `UNIQUE` genuinely gives up the
no-allocation property the mini-CSE was built for.

## Decision 4 — the cell boundary: Excel's `@` on an array is the TOP-LEFT

`ArrayEvaluation.FirstElement` is the single `Evaluate` body of all four records: build the operand, return
`At(0, Rows, Columns)` when it is an array, `#VALUE!` when the build was refused or produced no array. The
rule cannot drift across the four because there is one implementation.

The citation is Microsoft's "Implicit intersection operator: @" article (no GUID is published for it; cited
by title, as elsewhere in this repo): for a RANGE, `@` "returns the value from the cell on the same row or
column as the formula"; for an ARRAY, Excel "picks the top-left value". So a bare `=FILTER(...)` yields the
top-left — exactly what Excel writes when it upgrades a legacy formula to `=@FILTER(...)`.

**The phase file carried a controller note saying the array half is NOT top-left. That note is about RANGE
operands, and the measurement settles the split.** Measured 2026-09-10 on Aspose.Cells 26.6.0 with
`A1:A3` = 5, 0, 9, one formula per workbook, in two cells whose rows differ (`D2` is inside the range's rows,
`D5` is outside), plain and array-entered:

| Bare formula in a cell | `D2` | `D5` | Mode |
| --- | --- | --- | --- |
| `=SEQUENCE(5)` | 1 | 1 | both |
| `=SEQUENCE(2,3,7,1)` | 7 | 7 | both |
| `=FILTER(A1:A3,A1:A3>0)` | 5 | 5 | both |
| `=SORT(A1:A3)` | 0 | 0 | both |
| `=UNIQUE(A1:A3)` | 5 | 5 | both |
| `=FILTER(A1:A3,A1:A3>100)` | `#CALC!` | `#CALC!` | both |
| `=A1:A3*2` | 0 / 10 | `#VALUE!` / 10 | plain / CSE |
| `=IF(A1:A3>0,1,0)` | 0 / 1 | `#VALUE!` / 1 | plain / CSE |

A PRODUCER answers its top-left in every cell and in both modes — position-independent, so it is not
intersection. A RANGE operand under an operator is position-DEPENDENT on plain entry (per-operand implicit
intersection) and top-left on array entry. The same contrast holds for a lifted call: `=-A1:A3` plain is
-5 in `C1`, 0 in `C2`, -9 in `C3` and `#VALUE!` in `C5`, while CSE-entered it is -5 in every one of them
(same fixture, same session). So the producers implement the array half; the RANGE half of the array
boundary is untouched by this phase and its `#VALUE!` stays pinned by `CellBoundaryIntersectionTests`.

MySheet's own answers, measured on this tree through `Workbook.GetCellValue` (the only path that crosses
`Workbook.EvaluateCell`), are identical to the producer rows above in both cells, and `#VALUE!` for every
range row.

## Decision 5 — the shape invariant, and `#CALC!` as the eighth error code

**Invariant.** An `ArrayOperand` handed back by `IArrayProducer.TryBuildArrayOperand` has `IsArray` and
`Rows >= 1 && Columns >= 1`. An EMPTY result — nothing kept by `FILTER`, nothing left by
`UNIQUE(…, exactly_once)` — is a 1x1 `SingletonArrayOperand` carrying the error, NEVER a 0-extent array. A
1x1 SOURCE is likewise wrapped in a singleton, because `ScalarOperand` reports 0x0 and can never be a
RESULT.

**Why it is enforced rather than asserted.** `ArrayStream.Length` is `Rows * Columns` with no emptiness
channel, so a 0x1 operand enumerates nothing and every consumer answers as if the argument were an empty
range: `SUM` 0, `COUNT` 0, `AVERAGE` `#DIV/0!` — a silent wrong number where the oracle answers `#CALC!`.
`ArrayShaping.RequireProducerShape` therefore THROWS at construction (once per build, never per element),
because a `Debug.Assert` is compiled out of the Release build the gates run on and would be decorative. The
invariant's breakage is itself pinned by `ArrayProducerContractTests`, which constructs a raw 0-extent test
operand and watches the guard fire.

**A 1x1 result broadcasts like a scalar** (Phase 10 rule 1), which is what makes an empty result COMPOSE
rather than collapse: `SUM(FILTER(A1:A3,A1:A3>100,7)*A1:A3)` is 98 — the singleton's 7 against every cell —
and a 1x1 `#CALC!` under an operator fills every element with its error. That is Excel's rule; it is not a
bug to fix.

**`#CALC!`.** `Error.Calc` is code **7** and the eighth entry of the `Displays` array, and `ERROR.TYPE`
answers **14** for it. Excel's `ERROR.TYPE` table runs 8 `#GETTING_DATA` … 13 `#FIELD!` in between and the
engine has no code for any of those, so 14 is the only mapping past `#N/A` = 7. Measured:
`ERROR.TYPE(FILTER(A1:A3,A1:A3>100))` = 14 on the oracle in both modes, and 14 on this tree.

The change is contained and its boundaries are pinned: `ExcelExport.ErrorText` writes `error.ToString()`, so
an xlsx round-trip closes through `Error.FromDisplay`; an error display the engine does not know still folds
onto `#VALUE!` (`Error.FromDisplay("#SPILL!")`), which is what proves `#CALC!` now EXISTS rather than being
conjured by the fold; and a CODE past the table displays as `#ERR?`, which is how a pre-3.17 build degrades
when it reads a warm snapshot whose `CachedCellValue.ErrorCode` is 7. See
[`docs/serialization.md`](../docs/serialization.md) for the one-way boundary that follows from it.

## Union-tag allocation — the single source of truth

`[MemoryPackUnion]` tags on `Expression` are APPEND-ONLY: never renumbered, reordered or reused. Tag 322
(`Aggregate`, Phase 2) was the last one on `main` when this phase started, so the four took 323-326 in ONE
coordinated edit (`Danfma.MySheet/Expressions/Expression.cs:356-359`):

| Tag | Type | Phase |
| --: | --- | --- |
| 322 | `Mathematics.Aggregate` | 2 (already on `main`) |
| 323 | `Lookup.Filter` | 7 |
| 324 | `Lookup.Sort` | 7 |
| 325 | `Lookup.Unique` | 7 |
| 326 | `Mathematics.Sequence` | 7 |

**The next free tag is 327.** The file now carries **327** `[MemoryPackUnion]` attributes, tags 0-326 with no
gaps, and the policy comment at the top of the file reads "Add new tags at 327+". Note that
`grep -c MemoryPackUnion Expression.cs` prints **328**, not 327: the policy comment itself contains the
word. Count `^\[MemoryPackUnion` if you want the attribute count. Two phases appending tags independently
collide at MemoryPack type initialization, not at compile time, so re-count before claiming a number.

## Registry classification

The four are registered with `Entry<T>` — `ArrayLifting.Consumes`, the enum's zero — never `Elementwise<T>`.
Counts on this tree, each asserted by a test rather than stated: **310** built-ins, **180** `Elementwise`,
**130** `Consumes` (`FunctionRegistryClassificationTests.TheClassificationSplitsThe310BuiltInsInto180Elementwise_And130Consumes`,
and `ElementwiseLiftingTests`, which walks all 130).

`FILTER`, `SORT` and `UNIQUE` take a range in their first slot, so Phase 8's three-rectangle sweep sees them
and would catch a wrong flag by diagnostic. `SEQUENCE` takes no range at all, so the sweep is blind to it for
good and it must be NAMED: it is the 22nd member of the blind set that
`ElementwiseLiftingTests.TheShapeAndPositionAndCriteriaFamilies_StayConsumes` names by hand, and that set is
pinned by name (not by count) so gaining a member fails the suite naming the newcomer.

`ReferenceGuard` gained one arm each for `Filter`, `Sort` and `Unique`, standing for their source array the
way the existing `UnaryOperation{Plus}` arm stands for its operand — without them the error-IGNORING `COUNT`
family reports an empty result for a deleted sheet. `SEQUENCE` needs no arm: it has no reference argument.

## Per-function Excel citations

The four official pages, all on support.microsoft.com, all fetched **2026-09-10** (the GUIDs are the ones in
`DynamicArrayTests`' header, checked by fetching rather than copied from circulation — two GUIDs found in
older material 404):

| Function | Page GUID |
| --- | --- |
| FILTER | `f4f7cb66-82eb-4767-8f7c-4877ad80c759` |
| SORT | `22f63bd0-ccc8-492f-953d-c20e8e44b86c` |
| UNIQUE | `c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e` |
| SEQUENCE | `57467a98-57e0-4817-9f14-2eb78519ca90` |

### FILTER

From the page, verbatim: `include` is "a Boolean array whose height or width matches the array" — which is
why BOTH axes are supported; "If any value of the include argument is an error (#N/A, #VALUE, etc.) or
cannot be converted to a Boolean, the FILTER function will return an error"; and "If your dataset has the
potential of returning an empty value, then use the 3rd argument ([if_empty]). Otherwise, a #CALC! error
will result, as Excel does not currently support empty arrays."

Where the measurement decided instead:

- **The page's two clauses SPLIT.** An ERROR element is the producer's whole answer
  (`SUM(FILTER(A1:A3,E1:E3>0))` = `#DIV/0!`), but an element that merely cannot be converted is simply NOT
  KEPT — over `" TRUE "`, `"yes"`, `"1"` nothing is kept and the answer is the empty result's `#CALC!`, not a
  coercion error. A ONE-element include is the exception: there the unconvertible text IS `#VALUE!`, even
  with `if_empty`, and a falsy one keeps nothing as `#VALUE!` rather than `#CALC!` unless `if_empty` answers.
- **The include COERCES the text words `TRUE`/`FALSE`** (case-insensitive, no trimming, those two words
  only) through `ValueCoercion.CoerceToBoolAllowingTextWords` — the extension Phase 11b measured for `IF`,
  `NOT` and `IFS` and could not pin here because the function did not exist yet.
- **A scalar include broadcasts over a VECTOR source only.** `SUM(FILTER(A1:A3,TRUE))` = 14 and
  `SUM(FILTER(A1:C1,TRUE))` = 6, but `SUM(FILTER(A1:B3,TRUE))` is `#VALUE!`. Correction B1's "a scalar
  include broadcasts" is narrower than it was written.
- **An EMPTY `if_empty` slot is a blank VALUE, not an omitted argument** — `FILTER(A1:A3,A1:A3>100,)` is a
  1x1 blank (`ISBLANK` TRUE, `ROWS` 1) — the opposite of `SEQUENCE`'s rule and the one slot in the four that
  does not follow `SelectionProducers.IsOmitted`.
- **`if_empty` is not probed**, because it is built lazily; a refused `if_empty` is a 1x1 `#VALUE!` from the
  build rather than a refusal.

### SORT

From the page, verbatim: `sort_index` defaults to 1; `sort_order` defaults to 1, "1 for ascending, -1 for
descending"; `by_col` defaults to FALSE — "By default Excel will sort by row, and will only sort by column
where by_col is TRUE". `#VALUE!` for a `sort_order` other than 1/-1 is stated on the sibling SORTBY page.

Where the measurement decided instead — and these two contradict the phase design:

- **BLANKS sort LAST in both directions**, not as 0. Over 7, blank, `"t"`, 7 the ascending order is
  7, 7, `"t"`, blank and the descending order is `"t"`, 7, 7, blank. `ValueCoercion.Compare` on its own would
  have put the blank first ascending.
- **ERRORS are SORTED, not propagated** — after every value ascending, before every value descending, in
  their source order among themselves. Over 2, `#DIV/0!`, 3, `#N/A`, `#DIV/0!` the ascending order is
  2, 3, `#DIV/0!`, `#N/A`, `#DIV/0!`. `SUM(SORT(E1:E3))` is `#DIV/0!` only because `SUM` propagates what it
  is handed.
- **The sort is STABLE in both directions**, a fact the page does not state: descending is not a reversal.
- **An empty, blank-cell or text `sort_order` is NOT measurable.** Aspose.Cells 26.6.0 throws a
  `NullReferenceException` out of `CalculateFormula` for `SORT(A1:A3,1,)`, `SORT(A1:A3,1,A9)` and
  `SORT(A1:A3,1,"x")`. Those three therefore follow the page alone — omitted → 1, blank cell → 0 →
  `#VALUE!`, text → coercion `#VALUE!` — and the reason is recorded so nobody re-measures it.

### UNIQUE

From the page, verbatim: `by_col` "TRUE will compare columns against each other and return the unique
columns"; `exactly_once` "TRUE will return all distinct rows or columns that occur exactly once".

Where the measurement decided instead:

- **UNIQUE is case-SENSITIVE.** `"a"`, `"A"`, `"b"` keeps all three with `"A"` second, while
  `COUNTIF(C1:C3,"a")` = 2 and `MATCH("A",UNIQUE(C1:C3),0)` = 1 stay case-INSENSITIVE beside it. The page as
  fetched 2026-09-10 says nothing about case in either direction — and the GUID older material cited for it
  404s — so the measurement stands alone rather than overriding a page.
- **A BLANK is its own key**, equal to neither 0 nor `""` nor FALSE: 0, blank, `""`, FALSE are four rows.
  This is why `ValueCoercion.AreEqual` is NOT used: it equates blank with BOTH 0 and `""` while 0 and `""`
  are unequal, which makes raw equality non-transitive and hash bucketing unsound. Comparing keys by KIND
  and value restores transitivity without normalizing anything — see the deviation list for why
  `ArrayShaping.Normalize` was never added.
- **An ERROR is a key like any other**, equal to the same error code, and an error row is kept.

### SEQUENCE

From the page, verbatim: "Any missing optional arguments will default to 1". Microsoft's `#CALC!` guidance
adds that SEQUENCE "can't spill an array with 0 or negative values".

Where the measurement decided instead:

- **Every bad SEQUENCE argument is `#VALUE!`, never `#CALC!`** — `SEQUENCE(0)`, `SEQUENCE(-1)`,
  `SEQUENCE(0,0)`, `SEQUENCE(0.5)`, with `ERROR.TYPE(SEQUENCE(-1))` = 3. That is correction B2, and the
  `#CALC!` guidance describes the spill error a no-spill engine cannot have.
- **`rows`/`columns` truncate** (`SUM(SEQUENCE(2.7))` = 3, `SUM(SEQUENCE(2,2.9))` = 10) while `start`/`step`
  are kept as given, a zero or negative step included.
- **An omitted optional is absent OR the parser's `BlankValue` for an empty slot** (`SUM(SEQUENCE(2,2,,))` =
  10, `=SEQUENCE(2,,,)` = 1), while a blank CELL is a value that coerces to 0
  (`SUM(SEQUENCE(2,2,A9))` = 6, `SUM(SEQUENCE(A9))` = `#VALUE!`).

## Where the oracle lost — the two rulings

Both are exceptions granted under the P0 addendum's "the oracle contradicts itself" clause, the same clause
as Phase 9's 1900 working-day window and Phase 11's union ruling. Neither is a limitation; both are
decisions, recorded with the oracle's own numbers so they can be revisited.

**1. `UNIQUE(…, exactly_once)` follows the page.** Over `Q1:Q4` = 9, 5, 9, 0 the distinct rows are 9, 5, 0
and the rows occurring exactly once are 5 and 0. MySheet answers those two rows: `ROWS` 2, `SUM` 5. The
oracle keeps the DISTINCT-count SHAPE and pads it by repeating the last kept value: `ROWS` **3** with `SUM`
**5** — three rows that are 5, 0, 0 — so a UNIQUE result holds a duplicate, which its own row count
contradicts (measured 2026-09-10, both modes). The same padding is why the oracle answers a 1-row 4 for
`UNIQUE(S1:S2,FALSE,TRUE)` over 4, 4, where nothing occurs once; here that is the empty result, a 1x1
`#CALC!`.

**2. `AVERAGE` over `UNIQUE` follows the page.** Over the same `Q1:Q4`, MySheet answers 14/3 = 4.666…, which
is its own `SUM` 14 over its own `COUNT` 3. The oracle answers `SUM` **14** and `COUNT` **3** and `AVERAGE`
**0** (measured 2026-09-10, both modes) — its own three answers cannot all be right. Isolated over six
fixtures the pattern is that it drops every element up to and including the last ZERO from the numerator
while keeping the full count, which is what produces the 0 here. Microsoft's AVERAGE page is explicit that
the mean is the sum over the count and that zeros are included, so the page wins; a no-zero control sits
beside the pin.

## Documented deviations

Four, and only the first is structural. Each is pinned by a test — with the oracle's number in the test
wherever the two engines differ — so closing one is a deliberate edit.

1. **No spill.** A cell holds one scalar; a bare producer answers its top-left and never fills its
   neighbours. `=SEQUENCE(5)` shows 1, not 1..5. Note precisely what this is a deviation FROM: it is
   **agreement** with the oracle, which answers the top-left in both entry modes and in every cell (the
   table under Decision 4), and it is what Excel itself writes as `=@FILTER(...)` when it upgrades a legacy
   formula. It is a deviation from what a user of MODERN Excel expects when typing the un-prefixed formula,
   and that expectation is the single most surprising consequence of this phase, so it is stated first in
   the user docs rather than discovered.
2. **An open-range argument is REFUSED.** The mini-CSE's pre-existing cost guard refuses a whole-column or
   open range in an array position, and a producer propagates that refusal, so `SUM(FILTER(A:A,A:A>0))` is
   `#VALUE!` here, pinned by `DynamicArrayTests.Filter_OverAnOpenRange_IsRefused_AKnownDivergence`.
   Measured on the oracle 2026-09-10 with `A1:A3` = 5, 0, 9 and nothing else in column A:
   14 in both modes. The refusal has a **SILENT half** worth naming: `COUNT` discards the error channel, so
   `COUNT(FILTER(A:A,A:A>0))` is 0 here against the oracle's 2 on that fixture — a wrong number rather than
   an error. Whole-column `FILTER` is a very common real-world shape; making it work needs a shared-bounds
   rule (`array` and `include` would otherwise be bounded independently and could disagree on row count) and
   is a candidate follow-up, not a one-line change.
3. **The SEQUENCE size cap is MySheet's own.** `rows > 1048576`, `columns > 16384` or
   `rows * columns > 1048576` answers a 1x1 `#NUM!`. The oracle has NO cap in a consumed position —
   `ROWS(SEQUENCE(1048577))` = 1048577, `COLUMNS(SEQUENCE(1,16385))` = 16385 — so the justification is not
   "Excel bounds SEQUENCE by the grid" (it does not, in this position); it is that `ArrayStream` is lazy
   while every consumer iterates every element, so without the cap `SEQUENCE(1e6,1e4)` hangs the consumer.
   Exactly at the cap is allowed. Pinned by
   `DynamicArrayTests.Sequence_BeyondTheGrid_IsNumError_TheOnePinnedDivergence`.
4. **`COUNT` and `COUNTA` cannot see a producer's error.** `NumericAggregation.Fold`'s error channel is
   discarded by the counting consumers by design, so a 1x1 error array is invisible to them:
   `COUNT(FILTER(A1:A3,A1:A3>100))` and `COUNT(SEQUENCE(-1))` are both 0. For the empty `FILTER` that
   MATCHES the oracle (0 in both modes); for the bad `SEQUENCE` it is the same class of blindness with no
   error channel out of `ArrayStream` to fix it without changing `TryEvaluateStream`'s signature.

**One deviation the phase design listed and the code deliberately does NOT have: blank-to-zero.** Item 4 of
the phase file specified an `ArrayShaping.Normalize` mapping `Blank` to `Number(0)`, confined to these four
functions, on the premise that "Excel's dynamic arrays return 0 for a blank source cell". The oracle
contradicts that premise on every axis it was checked: `COUNTA(FILTER(A5:A8,A5:A8<>"zzz"))` = 3 =
`COUNTA(A5:A8)`, `COUNTA(UNIQUE(A5:A8))` = 2, `COUNT(UNIQUE(A5:A8))` = 1, `ISBLANK` of a kept blank is TRUE,
`ROWS(UNIQUE(A5:A8))` = 3 over 7, blank, `"t"` (the blank is its own key, merged with neither 0 nor `""`),
and blanks sort LAST rather than as 0. So `Normalize` was never added: blanks are PRESERVED through all
three selectors with no normalization seam anywhere, RANGE and PRODUCER agree, and risk #8's "`COUNTA(range)`
and `COUNTA(UNIQUE(range))` can disagree on a blank cell" describes a seam that does not exist. All three
facts are pinned by `DynamicArrayTests.TheSelectors_PreserveBlank_TheyDoNotNormalizeItToZero`, `SortTests`
and `UniqueTests`.

## Consumers — what reads a producer, and the two that do not

No consumer needed a change to READ a producer; two needed one to report its SHAPE, and one family already
rejected it.

**Read it as elements** (all through `ArrayEvaluation.TryStream`, the one shared gate): the numeric
aggregators `SUM`/`COUNT`/`AVERAGE`/`MIN`/`MAX`/`PRODUCT` and, through the same fold, `MEDIAN`, `STDEV*`,
the percentiles/quartiles, `SMALL`/`LARGE`; `INDEX`; `SUMPRODUCT`; `AGGREGATE`'s array form (`function_num`
14-19). Measured on this tree over `A1:A3` = 5, 0, 9: `SUM(FILTER(A1:A3,A1:A3>0))` = 14, `COUNT` 2,
`AVERAGE` 7, `MIN` 5, `MAX` 9, `PRODUCT` 45, `SMALL(…,1)` 5, `LARGE(…,1)` 9, `INDEX(SORT(A1:A3),1)` = 0,
`MEDIAN(SEQUENCE(5))` = 3, `PERCENTILE(SEQUENCE(5),0.5)` = 3, `SUMPRODUCT(SEQUENCE(3))` = 6,
`AGGREGATE(15,6,FILTER(A1:A3,A1:A3>0),1)` = 5.

**Report its shape** — item 14, this phase, not the sweep: `ROWS` and `COLUMNS` gained the same gate.
`ROWS(FILTER(A1:A3,A1:A3>0))` = 2 and `COLUMNS(SEQUENCE(2,3))` = 3. The gate also fixes shapes that had
nothing to do with producers (`ROWS(LEN(A1:A3))` and `ROWS(IF(A1:A3>0,A1:A3))` went 1 → 3), and a producer's
own 1x1 ERROR is reported as itself rather than as a shape of 1 — `ROWS(FILTER(A1:A3,A1:A3>100))` =
`#CALC!`, `ROWS(SEQUENCE(-1))` = `#VALUE!` (correction B2) — while an error ELEMENT inside a rectangle does
not hide the shape (`ROWS(SORT(E1:E3))` = 3).

**Stream it element by element** — item 15, one arm in `ArgumentFlattening.FlattenComputedValues`:
`COUNTA(UNIQUE(Q1:Q4))` = 3, `CONCAT(SEQUENCE(3))` = "123",
`TEXTJOIN(",",TRUE,FILTER(A1:A3,A1:A3>0))` = "5,9".

**TWO of that family's five callers do NOT stream a producer, by two DIFFERENT mechanisms**, and the docs
must not flatten them into one rule:

- `CONCATENATE` is the single caller that passes `streamArrays: false`. It joins scalars, so a producer
  argument reaches its `Evaluate` and it answers the producer's TOP-LEFT: `CONCATENATE(SEQUENCE(3))` = "1"
  where `CONCAT(SEQUENCE(3))` = "123". That is the oracle's answer in both entry modes.
- `COUNTBLANK` REJECTS a computed array up front through `PositionalRange.RejectComputedArray` and never
  reaches the streaming arm at all: `COUNTBLANK(SEQUENCE(3))` is `#REF!`.

**Rejects it, and that IS the Excel answer** — Phase 11a's Rule B, which needed no work here: the criteria
family (`SUMIF`/`SUMIFS`, `COUNTIF`/`COUNTIFS`, `AVERAGEIF`/`AVERAGEIFS`, `MAXIFS`/`MINIFS`) answers `#REF!`
for a producer in any range slot, which is what the oracle answers in BOTH entry modes.
`COUNTIF(FILTER(A1:A3,A1:A3>0),">5")` = `#REF!`. `SUMPRODUCT`'s opt-in factory is deliberately not gated,
which is why it reads a producer.

**Composition needs no code.** A producer under a lifted function, a unary, a binary, an `IF`, or under
another producer, is reached by `TryBuildOperand`'s recursion. Measured on this tree, and equal to the
oracle in both modes: `SUM(SORT(FILTER(A1:A3,A1:A3>0)))` = 14, `SUM(FILTER(A1:A3,A1:A3>0)*2)` = 28,
`SUM(LEN(FILTER(A1:A3,A1:A3>0)))` = 2, `SUM(FILTER(SEQUENCE(5),SEQUENCE(5)>2))` = 12,
`ROWS(UNIQUE(FILTER(A1:A3,A1:A3>0)))` = 2.

## Standing gaps handed on

- **A producer bound by `LET`, passed through `CHOOSE`, or through unary `+`, collapses to its top-left.**
  `NamedReferences.CaptureValue`'s closed type list treats a `Function` node as a value, so
  `SUM(LET(x,FILTER(A1:A3,A1:A3>0),x))` is 5 here against the oracle's 14 in both modes (measured
  2026-09-10; `ROWS(LET(x,FILTER(…),x))` = 2 there, `SUM(CHOOSE(1,FILTER(…)))` and `SUM(+FILTER(…))` = 14).
  This is correction M1 taken at option (a) — documented, pinned, not silent — and the honest scope note is
  that option (b), a `Probe`/`TryBuildOperand` arm through `Let`/`Choose`/`UnaryOperation{Plus}`, is what
  actually closes it. The loudest instance is the criteria slot, where a `LET` escapes Rule B's gate from
  both sides: `LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,">0"))` is **1** here — `COUNTIF` over the single
  collapsed element — against the oracle's `#REF!` in BOTH modes, with `SUM(f)` 5 against 14 and `ROWS(f)` 1
  against 2 in the same shape. All three rows are RED under that ONE cause, and they are the phase's single
  deliberately failing pin
  (`DynamicArrayTests.ALetBoundProducer_InACriteriaSlot_IsAKnownLimitHandedOverByPhase11a`), red so that the
  fix turns them green rather than being discovered.
- **A BARE-REFERENCE branch under a scalar-condition `IF` is deliberately left outside the array path.**
  Since `aeda1d7` a producer under a scalar-condition `IF` streams whole (`SUM(IF(TRUE,SEQUENCE(3),0))` = 6,
  where it used to collapse to 1), and so does a computed branch (`SUM(IF(TRUE,A1:A3*2,0))` = 28,
  `SUM(IF(TRUE,LEN(A1:A3),0))` = 3). A branch that is a bare RANGE is not moved: `SUM(IF(TRUE,A1:A3,0))` is
  `#VALUE!` here against the oracle's 14 in both modes (measured 2026-09-10). Streaming it would answer the
  "does `IF` return a reference?" question for that one shape while its siblings stay unanswered, so it is
  recorded as sweep item 32 and pinned by
  `MiniCseConsumerTests.ABareReferenceBranch_UnderAScalarConditionIf_IsUnmovedAndStillDiverges`. A
  single-cell defined name in the same position (`SUM(IF(TRUE,MyCell,0))` = 0 on the oracle, `#VALUE!` here)
  is pre-existing and rides along with it.
- **A name bound to a computed ARRAY still does not resolve inside an array position.** Phase 11a closed the
  range-bound cases (`LET(r,A1:A3,COUNT(r*1))` = 3), but `CaptureValue` evaluates the binding as a scalar
  before the arm can see it, so `LET(r,A1:A3*1,COUNT(r*1))` is 0. It is the same `CaptureValue` gap as the
  first bullet.
- **An oracle crash, recorded so nobody re-measures it.** Array-entered `ROWS(IF(FALSE,SEQUENCE(3)))` throws
  `CellsException: IndexOutOfRangeException` inside `Workbook.CalculateFormula()` and takes down every other
  formula in the same workbook. Probes for these rows must evaluate ONE formula per workbook.

## Two claims from the phase file that did NOT reproduce

Both were repeated across the design and must not be repeated again.

1. **`ShapeFold`'s 0-row premise.** Phase 10's `ShapeFold` comment said "a 0-row array is a legitimate shape
   once an empty FILTER result exists (Phase 7)". It never exists: the invariant above forbids a 0-extent
   producer result and the oracle backs the invariant (`SUM(FILTER(A1:A3,A1:A3>100))` = `#CALC!`,
   `ROWS(…)` = `#CALC!`, `COUNT(…)` = 0). The comment was rewritten while keeping the `_seen` flag for the
   `Axis(0,1)` case it also names.
2. **`SUM(FILTER(A:A,A:A>0))` "is `#VALUE!` on plain entry".** Not on its own terms. Measured 2026-09-10:
   with column A holding only `A1:A3` = 5, 0, 9 the oracle answers **14 in BOTH modes**; the `#VALUE!`-plain
   / **28**-CSE split the phase file recorded reproduces only once the wider fixture is present
   (`A5` = 7, `A7` = `"t"`, `A8` = 7 → four positive cells). So the split is fixture-dependent, not a rule
   about entry mode, and quoting it without its fixture makes the open-range deviation look like a mode
   disagreement when it is not. `ROWS(FILTER(A:A,A:A>0))` is 2 and 5 respectively, both modes.
