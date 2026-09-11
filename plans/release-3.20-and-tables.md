# Release 3.20.0 and the table chain — execution plan

Goal: turn `main`'s one deliberately red test green (Phase 11c), cut release 3.20.0 on a fully green suite, then deliver the structured-table chain (Phases 4, 5, 6) for 3.21.0. This file sequences the four phase files; it does not repeat their items.

## For Future Agents

As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases are done, fill in **Final Recap** and **Deployment Plan**.

Rules that bind every phase here, decided by the user on 2026-09-10 and 2026-09-11:

- **The suite is never red at a merge.** `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` and the Excel twin must report `failed: 0` on `main` before any release is cut. Never `dotnet test`.
- **Cadence per phase:** subagent TDD tasks from the generated briefs, per-task review, then a TWO-part final review — Fable (subagent, model `fable`) and GLM-5.3 (the z.ai endpoint; export the seven env vars from `/Users/danfma/bin/zclaude` and call `$HOME/.local/bin/claude --print … < brief`, never the wrapper, which drops its flags). Copilot is out. Then a fix wave, `git rebase main`, `git merge --ff-only`.
- **No AI attribution in commit messages.** The harness supplies trailers in a system reminder; the user's global rule forbids them and `main` is clean.
- **"Excel" means Aspose.Cells 26.6.0 as measured**, array-entered column. A divergence is a work item.
- Ledgers and briefs live under `.superpowers/sdd/<phase>/` (gitignored); each ledger has a RESUME HERE section.
- **Model per task (user, 2026-09-11): Opus for the complex ones, Sonnet only where the work is mechanical.** Opus: anything that measures a new oracle rule, inverts a design decision, touches `ArrayEvaluation.cs`/`Parser.cs`, or writes an engine arm — Phase 4 T2 and T5, Phase 5 T1, T2, T3 and T5, Phase 11c T3. Sonnet: shifting anchors, pinning numbers another task already measured, and doc sweeps — Phase 4 T6, Phase 5 T4, Phase 6 T2 and T3, Phase 11c T4. A doc sweep on Sonnet still carries the twin-diff requirement: after any scripted edit, grep for a distinctive token from every paragraph you did not mean to touch.
- **Fable ran out of usage credits on 2026-09-11** and killed three subagents mid-flight (nothing lost — every worktree was clean at its last accepted commit). The two-part final review was planned as Fable plus GLM-5.3; until Fable's credits reset, the second opinion is an Opus subagent, since GLM runs on the separate z.ai endpoint and is unaffected.

## Phase A: 11c — array bindings, suite to zero failures

Status: Complete

Plan: `structured-table-references-and-aggregate/phase-11c-array-bindings.md`. Briefs: `.superpowers/sdd/phase-11c-array-bindings/task-{1..4}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-11c`, branch `feat/array-bindings`.

- [x] Task 1 — the pins, red on arrival (items 1-3), plus the twelve unmeasured oracle rows. `6ad7cea`, core 1927/1 → 1978/38.
- [x] Task 2 — the scope carries an array; gates consult it (items 4-8). `1fd687d`, core → 1988/15 (the 15 are Task 3's).
- [x] Task 3 — defined names, CHOOSE, unary `+` (items 9-12). `ceb7539`, core 1988/15 → **2001 / 0**.
- [x] Task 4 — docs, design of record, bookkeeping (items 13-15). `199fb98`.
- [x] Two-part final review (an Opus subagent and GLM-5.3; both Yes with fixes, same Critical found independently), fix wave, ff-merged at `47087c6`.
- [x] Add the missing Phase 11b row to the master plan's phase table. `a99ab43`.

### Verification Plan
- On `main` after the merge: core `failed: 0` (total recorded in the ledger), Excel `93 / 0`, `dotnet csharpier check .` clean, Release build 0 warnings.
- `git log main -40` carries no `Co-Authored-By: Claude` / `Claude-Session` line.
- STOP HERE and tell the user: `main` is green, 3.20.0 can be cut (`release.yml`, `workflow_dispatch`). Phase B starts only after the user says so or cuts the release.

### Phase Summary

**Complete, merged at `47087c6`.** `main` is green for the first time since Phase 7 registered its producers: core **2008 / 0**, Excel **93 / 0**, csharpier clean, Release 0 warnings, no AI trailer in 45 commits. **3.20.0 can be cut** (`release.yml`, `workflow_dispatch`); `main` is 38 commits ahead of `origin/main` and needs a push first.

Four binding sites stopped collapsing a computed array to its top-left. The scope holds an `ArrayOperand` built once (`ArrayBindings.Capture`); `NameReference` answers that binding in each of its three roles; the one bare-reference predicate every gate shares became context-aware, and the context-free overload was removed for want of callers; `ResolveNameShape` gained an `Array` outcome; `Choose`, `UnaryOperation{Plus}` and `Let` gained probe/build arms. No new `ComputedValueKind`. Three Phase 8 pins and six Phase 11a pins flipped, each with both numbers in its commit body.

Five of the controller's own claims were measured false by the implementers and are worth carrying forward: a design item saying a gate "follows with no code of its own" was wrong twice (the criteria gate needed `ProbeLet`/`TryBuildLet`); item 11 as written would have created four new divergences, because making `+A1:A3` array-eligible at a top level makes the criteria gate refuse what the oracle reads (`COUNTIF(+A1:A3,">0")` is 2 there); item 10 needed CHOOSE's own capture rather than IF's; the scalar reading of a binding is the operand's top-left at every site, not `#VALUE!`; and `TryGetRegion` would have been a lie as a `Try*` returning three outcomes.

The review's most useful finding was not a defect but a coverage gap: a contributor-style revert of the predicate's defined-name clause left the suite fully green with four measured divergences silently restored. The unary-`+` half survived the same mutation through mechanism pins in two other files; the defined-name half had none and now does.

## Phase B: 4 — foundation task first, the rest in parallel

Status: Complete

Status corrected 2026-09-11: the worktree was created off `main` EARLY (tags 323-326 were already there), so Phase B ran concurrently with Phase A. It still merges AFTER Phase A.

Plan: `structured-table-references-and-aggregate/phase-4-lexer-parser.md` (re-verified 2026-09-10, three rulings). Briefs: `.superpowers/sdd/phase-4-lexer-parser/task-{1..7}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-p4`, branch `feat/lexer-parser`.

- [x] Task 1 — foundation: token, error kinds, `TableReference` with the six-member enum and tag 327, scanner. `d6e4062` + `3dbd6b2`, core 1927/1 → 1973/1, tag count 327 → 328.
- [x] Task 3 — lexer. `ed336c5`, `TokenizerTests` 9 → 17 cases, core → 1981/1.
- [x] Task 7 — name-validator repoint. `eaaa36d`, core → 1993/1.
- [x] Task 4 — anchored support. `e2761c2`, core → 2008/1.
- [x] Task 2 — grammar + writer. `fd5fe81` + `40bcc93` + `f4217ed` + `f12e1c1` + `eefd87d` + `f53e367` (six commits, TDD red-first, both suites green at every boundary). core → 2369/0 on the final branch head.
- [x] Task 5 — parser arms (MERGED with T6, the reason is measured: the arms re-arm eleven `Load_*` tests, so both halves landed together). `621e11a` + `d224b80` (core 2369/0, Excel 93/11 → 93/0).
- [x] Task 6 — folded into T5 (loader vehicle + the eleven `Load_*` tests) and `a26cdbd` (the serialization.md tag sentences, fixed after the two reviews of the partial edit).
- [x] Final review, fix wave, ff-merge. Reviewer parts: GLM-5.3 via z.ai (before the user's ruling; Yes with fixes, all registration-only), the controller's own subagent (Yes with fixes — measured), Copilot (Yes, zero actionable findings, discounted). Kimi 403 (monthly). Every finding verified in the tree before acting.

### Verification Plan
- After Task 1's merge: core and Excel suites `failed: 0`; `grep -c '^\[MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` = 328.
- After the phase: `FormulaWriterTests` 51 → 73, `TokenizerTests` 9 → 13, `ParseExceptionTests` 14 → 25, both suites 0 failures.

### Phase Summary

**Complete, merged at `e3dc21f` (fix-wave heads `8f94b18` + `e3dc21f` after a CLEAN 13/13 rebase — the expected master-plan conflict never materialized: branch and main share no file).** main verified after the merge: core **2370 / 0**, Excel **93 / 0**, csharpier clean (389 files), Release 0 warnings, `^\[MemoryPackUnion` = **328** with 327 = `TableReference`, no AI trailer in any of the thirteen commits. Worktree and branch removed.

Three reviewers, all findings verified in the tree by the controller before acting: GLM-5.3 via z.ai (Yes with fixes, all registration), the controller's GLM-5.3-Flash subagent (Yes with fixes — ~87 oracle rows, an independent property pass 3017/474/0 holes, two contributor-style mutations both caught by their pins), and Copilot (Yes, zero findings — its gate re-runs are real, its "oracle measurements" are transcribed from tests and one of its distinctive claims is false; zero actionable findings, discounted per the user's standing verify-everything rule). Kimi was 403 (monthly limit).

The fix wave (`final-review/fix-wave.md` holds the full record): ONE behavioural fix both code reviewers converged on — `DecodeName` swallowed the next char after every `'`, so the raw spelling `Tabela1[a'b]` silently summed the WRONG existing column (red-first pin: 20 where the oracle's rule answers 10); the escape is honoured only as `''`→`'` and before `[ ] # ' @` (`40fc43b`, core 2369 → 2370/0, the 3598/1271 property counts scanner-level and unmoved) — plus nine registration fixes (`0f2bbbf`): the `[[#Data],[Valor]]` collapse re-registered as a divergence (the oracle PRESERVES `[#Data]` beside a column, on both storage surfaces), a stale rejection claim past-tensed, the "only `[Valor]` is rejected outside a table" sentence corrected (the whole bare implicit family is set-thrown), two surviving "next free tag is 327" sentences in `plans/` → 328, the serialization.md twins' self-contradicting bullet rewritten with parity 18/18/20/20/5/5, escape-set and `Tabela1[]` comments now naming their storage surface, a stale test header, and two honesty rewordings. Deferred, recorded: the TableInteropTests warning asserts and the "42" coincidence (no false pass reachable).

### Phase Summary
_(write when phase completes)_

## Phase C: 5 and 6 concurrently with Phase B's tail

Status: Complete

Plans: `phase-5-resolution-and-graph.md` (its items 1-4 are deleted by the 2026-09-10 ruling; it depends only on Phase B Task 1) and `phase-6-excel-loader.md`. Both need re-verification against `main` before briefs are generated — Phase 5's anchors are as old as Phase 4's were, and Phase 6 has never been audited.

- [x] Re-verify Phase 5 against `main`; rulings R1-R4 written into the phase file (`53010e4`); five briefs; worktree `/Volumes/Work/Develop/MySheet-p5`, branch `feat/resolution-and-graph` off `feat/lexer-parser`.
- [x] Re-verify Phase 6; nine rulings written into the phase file (`9bef4c7`); three briefs; worktree `/Volumes/Work/Develop/MySheet-p6`, branch `feat/excel-loader`.
- [x] Phase 6 Task 1 — the `<table>` reader, nine committed Aspose fixtures, export/merge pins. `ae6d82a`, Excel 93/0 → 125/0.
- [x] Phase 5 Task 1 — six-area geometry plus the `Empty` outcome. `5879a48`, core 1994/1 → **2092/1** (+98), Excel 93/0. `TryGetRegion` renamed to `GetRegion`, returns `TableRegionOutcome { Resolved, Empty, Absent }`.
- [x] Execute Phase 5 (its cadence), ff-merge. T2 `0654194` and T3 `c43799c`+`9f681f7` accepted (core 2436/0); T4 accepted 2026-09-11 session 2 (four commits, core **2451/0**, Excel 93/0, ledger has the details); T5 accepted (`82746fc` R3 + M4 + item 24b, `544c53e` docs, core **2552/0**). Review (controller subagent: **Yes with fixes** — both mutation batteries run in full; Copilot: "Yes" zero findings, discounted for fabricating the sweep-32 seam agreement and unexecuted probes), fix wave (one docs-only Important: the Named-ranges blockquote's stale "neither the structured-reference syntax" clause, both twins), clean 13/13 rebase, **ff-merged at `313a6c8`**. The INDIRECT canaries flipped on the rebased branch before T5 (`d8048a8`), measured 60 before flipping.
- [x] Execute Phase 6 (its cadence), ff-merge. T1 accepted earlier (`ae6d82a`); T2 accepted (`9be9703`+`2f50fcb`+`b6090ec` — the two merge-consequence flips first, then items 12-19 with three brief corrections measured; Excel **133/0**); T3 accepted (`84ea708` docs, 6 files, both twins; the handed-over SMALL row measured as ClosedXML's column — no edit). Review (controller subagent: **"Yes"**, zero Critical/Important — the full END-TO-END with an Aspose-authored file it authored itself matched every shape, containment mutation 13/29 red, streaming invariant re-measured; Copilot: "Yes", real probe artifacts this time), fix wave (two Minors: fixture-provenance sentence scoped; the export trap's `<f>` text pinned in Formulas mode), clean 6/6 rebase, **ff-merged at `b778ab7`**. **The release blocker is gone: an Aspose-authored `.xlsx` with a real table and `SUM(Tabela1[Valor])` loads and answers the oracle's number (42), verified end-to-end by the review.**

### Verification Plan
- Both suites 0 failures on `main` after each merge.
- An `.xlsx` produced by Aspose with a real table and `=SUM(Tabela1[Valor])` loads and answers the oracle's number (the end-to-end pin Phase 6 must add).

### Phase Summary

**Complete. Phase 5 merged at `313a6c8`, Phase 6 at `b778ab7`; main verified after each merge: core 2552 / 0, Excel 133 / 0 (93 → 133 over Phase 6), csharpier clean, Release 0 warnings, no AI trailers.** All worktrees and feature branches removed.

Phase 5 gave the structured reference its resolution and its consumers: one primitive (`TableReference.TryResolveRange` over `Table.GetRegion` returning `TableRegionOutcome {Resolved, Empty, Absent}`), ruling R2's error-VALUE semantics pinned against mutation, the mini-CSE CSE-column arms with the sweep-32 seam recorded, the range-value cache normalized on the resolved rectangle (DynamicRange deliberately excluded), ISFORMULA/FORMULATEXT measured off their switches' `#VALUE!`, and bare `=Tabela1` resolving through the name path at evaluation time (R3a) with the M4 cell-boundary pin and the docs of finding 8. Phase 6 closed the release blocker: the loader reads `<table>` parts in the per-sheet loop (`displayName`, `ref` geometry with `?? 1`/`?? 0` defaults, ordered raw column names + `_xHHHH_` decode), contains every malformed part behind ONE `InvalidTableDefinition` warning with the table skipped, keeps `SaveAsExcel` table-free, and pins evaluation over the nine committed Aspose-authored fixtures — ending with the review's own end-to-end: an Aspose-authored file loads, registers, and answers 42 with liveness.

The reviewers who were worth their cost: the controller's GLM-5.3-Flash subagent found the `DecodeName` apostrophe divergence (Phase 4, fixed), ran both Phase 5 mutation batteries to red, and executed the Phase 6 end-to-end; Copilot produced zero actionable findings across all three phases despite real artifacts on the last one — its Phase 5 review fabricated agreement on the sweep-32 seam, and every one of its findings was discounted after verification, exactly the standing rule. The controller's own brief was measured wrong twice by implementers/reviewers (item 11's INDIRECT claim in Phase 5 T4; the `=ZZZ999` classification example) — recorded where the next reader will see it.

## Phase D: release 3.21.0

Status: Ready — awaiting the user's release run

- [x] Master plan rows for 4, 5, 6 marked Complete with counts.
- [x] Tell the user `main` is green and 3.21.0 can be cut.

### Verification Plan
- Both suites 0 failures; no AI trailer in `git log`; `CHANGELOG.md` untouched (versionize owns it).

### Phase Summary

**`main` @ `b778ab7` is green and 37 commits ahead of `origin/main` (push first): core 2552 / 0, Excel 133 / 0, csharpier clean, Release 0 warnings, `^\[MemoryPackUnion` = 328 with 327 = `TableReference`, no AI trailer anywhere in the epic's commits, `CHANGELOG.md` untouched. The release blocker is closed: a real Aspose-authored `.xlsx` with a table and `SUM(Tabela1[Valor])` loads and answers the oracle's number.** Phases 4, 5 and 6 are Complete above; the user cuts 3.21.0 via `release.yml` `workflow_dispatch`.

## Final Recap

**The structured-table epic is done: a real `.xlsx` table now round-trips, loads, and answers.** Across Phases 4, 5 and 6 (eleven + twelve + six commits on top of the 3.20.0 main), MySheet gained: the bracketed-specifier token and the structured-reference grammar with a canonical writer owned by one file (Phase 4); the `TableReference` node (union tag 327) with ONE resolution primitive over a three-outcome region outcome, error-VALUE failure semantics pinned against mutation, the mini-CSE array-entered arms with two recorded seams left deliberately open (sweep items 32/33/34), the consumer/cache fast paths, bare `=Tabela1` through the name path at evaluation time, and the `<table>` reader that populates the registry at load with every malformed part contained behind one warning (Phases 5 and 6). The oracle (Aspose.Cells 26.6.0 as MEASURED) governed every number: where the engine cannot match — the empty-table `#REF!`, the consumer-error overwrites, the export trap, the `TRUE`-named-table limit — the divergence is registered beside its pins, never sold as parity.

Process record, because it is part of the deliverable: every phase ran brief-driven TDD subagents, a controller review per task, a multi-part final review with every finding verified in the tree before acting, a fix wave, a clean rebase and an ff-only merge. The model plan of record (Opus-class work / mechanical work) was re-decided twice by availability — Fable out of credits, Opus and Sonnet weekly-limited from Sep 11, Kimi monthly-limited — landing on the controller's own subagent instances plus Copilot, with Copilot's findings discounted unless verified (it produced zero actionable findings across three final reviews, fabricating agreement once). Three sessions of controller work; the plan files and the gitignored ledgers carried every restart.

## Deployment Plan

1. `git push origin main` (37 commits ahead; the user or the controller can run it — the epic's rule was that the user cuts releases).
2. Cut **3.21.0** via `.github/workflows/release.yml` `workflow_dispatch` (the user runs it): versionize bumps the version, tags, packs both packages and publishes to NuGet.org via Trusted Publishing; `CHANGELOG.md` is generated (never edited by hand).
3. After the release: `git pull` on main, confirm the tag, and the structured-table work is shipped. The recorded follow-ups live in `plans/structured-table-references-and-aggregate/phase-11-excel-compatibility-sweep.md` items 32-34 (the deliberate seams), plus the deferred minors in each phase's fix-wave file.

## RESUME HERE (2026-09-11, session 3 — Phases 4 AND 5 merged)

**3.20.0 shipped. Phase 4 COMPLETE (`e3dc21f`). Phase 5 COMPLETE and merged (`313a6c8`, core 2552 / 0, Excel 93 / 0, csharpier clean, Release 0 warnings, no trailers).** Phase 6 is rebased onto main and T2 is running. **Phase 6 is the release blocker** — without it a real `.xlsx` cell holding `SUM(Tabela1[Valor])` answers `#NAME?`.

| branch | worktree | head | suite there | state |
| --- | --- | --- | --- | --- |
| `feat/excel-loader` | `MySheet-p6` | `24a1c75` | core 2552 / 0, Excel **123 / 2** | T1 landed + rebased clean. The 2 failures are T2's first work items (both measured consequences of the Phase 4/5 merges — see the phase-6 ledger's session-3 entry). T2 RUNNING on the controller's subagent instance; T3 last (docs, both twins, RE-MEASURE EVERY ANCHOR — plus `docs/pt-BR/performance.md:434` SMALL, handed over by the Phase 5 review) |

**Reviewer lineup (user): every review = one controller subagent instance + Copilot, every Copilot finding verified against the tree before acting. Kimi 403 monthly. No Claude/z.ai CLI dispatches until Monday Sep 14.** Copilot's reviews on both phases so far: gates real, oracle numbers transcribed-not-measured, and on Phase 5 it fabricated agreement on the sweep-32 seam — zero actionable findings both times; the discounts are recorded in both fix-wave files.

Next actions in order: (1) T2 returns → controller review, ledger, then T3 (docs sweep, both twins, re-anchored); (2) Phase 6 review (controller subagent + Copilot) over `main..feat/excel-loader`, de-dup into `phase-6-excel-loader/fix-wave.md`, execute, rebase, ff-merge, verify on main; (3) master plan rows 4/5/6 Complete with counts, Phase C/D Phase Summaries, tell the user **3.21.0 can be cut** (release via `release.yml` `workflow_dispatch`, the user runs it).

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete — releases are manual via `.github/workflows/release.yml` `workflow_dispatch`: versionize bumps and tags, packs both packages, publishes to NuGet.org via Trusted Publishing; the user runs it)_
