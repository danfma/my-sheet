# Phase 11b: text "TRUE" and "FALSE" in a boolean slot

Status: Complete   <!-- Not started | In progress | Complete -->

A user bug report of 2026-09-10, measured on both sides before any code was written. Runs in PARALLEL with
Phase 3's tail and Phase 7; it shares no engine file with either.

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — read that master plan first for the governing principle P0 and its 2026-09-09 addendum ("Excel" means Aspose.Cells 26.6.0 as MEASURED; a measurement beats a documentation page).

## Design decision

**The report.** `=IF("TRUE",1,0)` answers `#VALUE!` here and `1` in Excel; likewise `"FALSE"` → 0, and a CELL
holding that text behaves the same way.

**What the measurement adds, and why it changes the shape of the fix.** All rows below were measured on
Aspose.Cells 26.6.0 on 2026-09-10, and **plain and array-entered entry agree on every one of them**, so entry
mode is not a factor in this phase.

| formula | MySheet | oracle | verdict |
| --- | --- | --- | --- |
| `IF("TRUE",1,0)` / `IF("FALSE",1,0)` | `#VALUE!` | 1 / 0 | **defect** |
| `IF(A1,1,0)` with `A1` = `"TRUE"` / `"FALSE"` | `#VALUE!` | 1 / 0 | **defect** |
| `IF(A3,1,0)` with `A3` = `"true"` | `#VALUE!` | 1 | **defect** — the match is case-INSENSITIVE |
| `NOT("TRUE")` / `NOT("FALSE")` | `#VALUE!` | FALSE / TRUE | **defect** — the same coercion site |
| `COUNTIF(A1:A2,TRUE)` / `COUNTIF(A1:A2,"TRUE")` | 1 | 0 | **defect, and SILENT** — belongs to the sweep, not here |
| `IF(" TRUE ",1,0)` | `#VALUE!` | `#VALUE!` | correct — whitespace is NOT trimmed |
| `IF("yes",1,0)` / `NOT("yes")` | `#VALUE!` | `#VALUE!` | correct — only the two words coerce |
| `IF("1",1,0)`, `IF(A6,1,0)` with `A6` = `"1"` | `#VALUE!` | `#VALUE!` | correct — numeric TEXT is not a boolean |
| `AND("FALSE",TRUE)` | TRUE | TRUE | correct — text is **IGNORED**, not coerced |
| `OR("TRUE",FALSE)` | FALSE | FALSE | correct — ignored |
| `AND("yes",TRUE)` | TRUE | TRUE | correct — ignored, not an error |
| `XOR("TRUE",FALSE)` / `XOR("TRUE","FALSE")` | FALSE / `#VALUE!` | FALSE / `#VALUE!` | correct — ignored; an all-text call errors |
| `"TRUE"+0` / `--"TRUE"` | `#VALUE!` | `#VALUE!` | correct — arithmetic does not coerce |
| `IF("TRUE"=TRUE,1,0)` | 0 | 0 | correct — comparison does not coerce |
| `SUMPRODUCT(--(A1:A2=TRUE))` | 0 | 0 | correct |

**The decision this forces.** This is NOT "fix boolean coercion" in a shared helper. `AND`, `OR` and `XOR`
**ignore** a text argument, which is the opposite of coercing it: `AND("FALSE",TRUE)` is TRUE, where coercion
would give FALSE. A shared fix would silently break those three. The coercion belongs to exactly two places —
`IF`'s condition and `NOT`'s argument — and the rule is narrow: the two words `TRUE` and `FALSE`, compared
case-insensitively, with NO trimming, and nothing else. Arithmetic, comparison and the criteria family are
untouched.

`COUNTIF`'s boolean-versus-text matching is a different rule (criteria-family type equality) and a SILENT wrong
number; it goes to the compatibility sweep beside the existing criteria item, not into this slice.

## Implementation items

- [ ] **1.** Create `tests/Danfma.MySheet.Tests/Expressions/BooleanTextCoercionTests.cs` with the acceptance pins,
      all RED, and the must-not-move guards, all GREEN, taken verbatim from the table above: the five `IF` rows and
      the two `NOT` rows are RED; the ten guard rows (`IF` with whitespace, `"yes"`, `"1"`, a cell holding `"1"`,
      `AND`/`OR`/`XOR`'s ignore behaviour including `AND("FALSE",TRUE)` = TRUE, arithmetic, comparison and
      `SUMPRODUCT`) are GREEN and must stay so. Cite the oracle version, the date, and that both entry modes agree.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/BooleanTextCoercionTests.cs`
      *Why:* `AND("FALSE",TRUE)` is the row that distinguishes "ignore" from "coerce"; without it a shared-helper
      fix looks correct.

- [ ] **2.** Find every boolean-coercion site and report the list before changing any of them: grep for the helper
      `IF` uses for its condition and for `NOT`'s argument, and confirm which other callers share it. If `IF` and
      `NOT` share a helper with `AND`/`OR`/`XOR`, the fix must NOT go in the shared helper — add the text rule at
      the two call sites, or give the helper an explicit opt-in flag and pass it only from those two.
      *Files:* to be named by the implementer from the grep
      *Why:* This is the whole risk of the phase. The measurement says the three folding functions must keep
      ignoring text.

- [ ] **3.** Implement the rule at the two sites: the two words `TRUE` and `FALSE`, case-insensitive
      (`OrdinalIgnoreCase`), with no trimming and no other accepted spelling. Everything else in a boolean slot
      keeps answering `#VALUE!`.
      *Files:* the sites item 2 names, likely `Danfma.MySheet/Expressions/Logical/If.cs` and `Not.cs`
      *Why:* The oracle accepts `"true"` and rejects `" TRUE "`, so the rule is a case-insensitive exact match.

- [ ] **4.** Docs: state the rule and its boundary in `docs/function-reference.md`'s `IF` and `NOT` rows and in the
      pt-BR twins — the two words only, case-insensitive, not trimmed, and that `AND`/`OR`/`XOR` IGNORE text
      instead. Every number labelled as measured with the version, date and the fact that both entry modes agree.
      *Files:* `docs/function-reference.md`, `docs/pt-BR/function-reference.md`
      *Why:* A reader who learns that `IF` coerces will reasonably assume `AND` does too, and it does not.

## Verification Plan

- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/BooleanTextCoercionTests/*"` — all green when item 3 lands.
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` — core at its baseline plus this file's tests, 0 failed. Baseline at branch creation: **1636**.
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` — **93 / 0**.
- `dotnet csharpier check .` clean; `dotnet build Danfma.MySheet.slnx -c Release` 0 warnings.

## Risks carried by this phase

- A shared-helper fix breaks `AND`, `OR` and `XOR` in a way no existing test catches, because their current answers
  coincide with the ignore rule. Item 1's `AND("FALSE",TRUE)` guard is the only thing standing there.
- `IF` is consumed by the mini-CSE and the cell boundary; a change to its condition path must not move those. Run
  the full suite, not the filtered class. **CORRECTED 2026-09-10 by the review: `AggregateCodes` has no coupling to
  `IF`'s condition at all, so naming it here was an unsubstantiated risk. The two real condition paths are
  `If.Evaluate` for a scalar condition and `IfOperand.At` for an array one, and there is no third — `Entry<If>` is
  not `Elementwise`, so no generic lift bypasses them.**

## Open questions owned by this phase

- Whether any OTHER function takes a boolean argument that should coerce the two words (`IFS`, `SWITCH`'s
  condition, `FILTER`'s include). Measure them; extend the pins if the oracle coerces there too, and record it if
  it does not.

## Phase Summary

**Status: Complete** — branch `feat/boolean-text`, three commits, core **1676 / 0** (1636 baseline + 40 tests),
Excel **93 / 0**, Release 0 warnings, csharpier clean.

`IF`, `NOT` and `IFS` now accept the text `TRUE` and `FALSE` in a condition slot, compared case-insensitively with
no trimming and no other accepted spelling. The fix is an opt-in extension `CoerceToBoolAllowingTextWords` beside
`ValueCoercion.CoerceToBool`, a **pure addition of 37 lines with 0 deletions**, called from exactly four sites, so
the nine other callers of the shared helper are provably untouched.

**What the design got wrong, and what saved it anyway.** This file argued that a shared-helper fix would break
`AND`, `OR` and `XOR` because they coerce text. Measured on the tree, they do not reach the helper at all:
`LogicalReduction` skips `Text` and `Blank` before the call, and its reference arm switches on kind. So the stated
mechanism was wrong — but the conclusion held for a larger reason the implementer found: the helper has **13 call
sites**, and eight are unrelated boolean FLAG slots (`VLOOKUP` and `HLOOKUP`'s range_lookup, `ADDRESS`'s a1,
`TEXTJOIN`'s ignore_empty, `FIXED`'s no_commas, `DAYS360`'s method, `VDB`'s no_switch, a bond argument), which a
shared change would have moved at once and unmeasured.

**Two things the implementer found that the design missed.** `IF` has TWO condition sites, not one — `If.Evaluate`
and `IfOperand.At`, the array-condition path — so fixing only the first would have left the scalar and array paths
disagreeing. And `IFS` shares the rule, measured, so it is fixed too.

**The open question, answered by measurement** (Aspose.Cells 26.6.0, 2026-09-10; plain and array-entered agree on
every row): `IFS` coerces and is fixed; `SWITCH` does NOT, because its expression is not a boolean slot, and it was
already correct here and is now pinned as a guard; `FILTER`'s include slot DOES coerce, which Phase 7 must apply
when it builds that function, since it does not exist yet.

**Handed to the compatibility sweep, both measured:** the eight boolean flag slots above, where the oracle coerces
the two words and rejects other text while MySheet rejects everything; and `COUNTIF` counting a cell holding the
text `TRUE` as the boolean, which is a criteria-family type-equality rule rather than a coercion one.
