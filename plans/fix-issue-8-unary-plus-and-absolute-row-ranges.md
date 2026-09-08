# Fix issue #8 — unary `+` on text/boolean, absolute whole-row references (`$1:$1`)

Fixes the two defects reported in [issue #8](https://github.com/danfma/my-sheet/issues/8) against 3.15.0,
plus a third one found while reproducing the issue's own representative formula (`LET` did not capture
ranges), and the structured `ParseException` the issue asked about in its diagnosability note. Reference
semantics = Aspose.Cells 26.7 as quoted in the issue. Branch: `fix/issue-8-unary-plus-absolute-rows`.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to
`Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with
zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases
are done, fill in **Final Recap** and **Deployment Plan**.

Test runner is **TUnit** (not xUnit): `dotnet test` does not work on .NET 10. Use
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/<Class>/*"`.
Console output from tests only shows with `--output Detailed`. Pre-commit hook runs `csharpier check` and a
Debug build, so format with `dotnet csharpier format .` before committing.

## Root causes (all confirmed by reproduction before fixing)

**Defect 1 — unary `+`** (`Danfma.MySheet/Expressions/UnaryOperation.cs`): `Evaluate` ran
`CoerceToNumber` on the operand for every operator, so `+text` → `#VALUE!` and `+TRUE` → `1`.

**Defect 2 — `$1:$1`** (`Danfma.MySheet/Parsing/Parser.cs`): the Tokenizer lexes `$1` as an `Identifier`
(a `$` starts an identifier so `$A$1` works). `Parser.TryEndpoint` recognised a `NameReference` only as a
column (via `CellAddress.TryParseColumn`, which strips `$` — hence `$A:$A` worked) and had no row
counterpart. Unqualified → `DynamicRange(NameReference,NameReference)` → `#REF!`; sheet-qualified →
`ParseQualifiedReference` threw `Expected a cell reference after '!'`. One missing branch, both symptoms.

**Defect 3 — `LET` and ranges** (`Danfma.MySheet/Expressions/Logical/Let.cs`, found via the issue's real
formula): `Let` bound each name to `value.Evaluate(scope)`. A range node evaluates to `#VALUE!` (a range has
no scalar), so `LET(hdr, Data!$1:$1, MATCH(x, hdr, 0))` matched against `#VALUE!` → `#N/A` — with or without
the `$`. Defined names already handled this (`NamedReferences.EvaluateDefinition` wraps range nodes as a
reference value); `LET` and `CHOOSE` each had their own copy of the rule or lacked it.

## Decisions

- Unary `+` is a **type-preserving no-op** except `Blank → 0`: text, boolean, number, error AND
  reference-typed values (`+OFFSET(...)`) pass through unchanged, so `SUM(+A1:A3)` keeps working.
- Defect 2 fixed in `Parser.TryEndpoint` via a new `CellAddress.TryParseRow` (twin of `TryParseColumn`),
  NOT in the Tokenizer: making `$1` a `Number` token would let `=$1+1` evaluate to `2`, which Excel rejects.
  `=$1+1` stays `#NAME?` (tested). `FormulaWriter` already drops `$` for open ranges, so `$1:$1` → `1:1`.
- Defect 3 fixed with one shared helper `NamedReferences.CaptureValue(expression, context)` used by
  defined names, `LET` and `CHOOSE` (three sites → one rule). An `AnchoredRangeReference` (shared-formula
  master) is resolved to its per-slave rectangle first. A single cell is still bound by value.
- `ParseException` gains `Kind` (`ParseErrorKind` enum, 8 members) and `Token`; `Message` format unchanged.
  The public 2-arg constructor was REPLACED by the 4-arg one (flagged in the PR as an API change; no
  consumer constructs it — every catch site keeps compiling). "Syntax vs semantic" needs no flag: the
  exception type IS the syntax category; documented. `Position` is 0-based into the formula BODY (after
  `=`) — that is why the issue saw "position 13" for `=COUNTA(Other!$1:$1)`; documented.
- Out of scope (pre-existing, unrelated): open ranges inside shared-formula masters are not shifted per
  slave (`1:1` and `$1:$1` both stay fixed).

## Phase 0: Commit the unrelated `code-review-graph install` changes on `main`
Status: Complete

- [x] Clean `.husky/pre-commit` (stray second shebang removed, block commented).
- [x] Commit `chore(tooling): add code-review-graph hooks, skills and project instructions` (`a683bcc`).
- [x] Branch `fix/issue-8-unary-plus-absolute-rows` created from it.

### Verification Plan
- `sh -n .husky/pre-commit` → ok; `git status --short` clean apart from the plan file. ✔

### Phase Summary
Committed `.gitignore`, `.husky/pre-commit`, `.claude/settings.json`, `.claude/skills/*`, `CLAUDE.md`,
`.github/code-review-graph.instruction.md`. `.claude/settings.local.json` and `.claude/worktrees/` are
already ignored (global git ignore / `.git/info/exclude`). Pre-commit hook ran (csharpier + Debug build +
graph update) and passed.

## Phase 1: Unary `+` no-op
Status: Complete

- [x] 10 tests added to `tests/Danfma.MySheet.Tests/Expressions/UnaryOperationTests.cs` (5 RED before fix).
- [x] `UnaryOperation.Evaluate`: `Plus` returns the operand unchanged (blank → 0); `Negate`/`Percent` keep
  `CoerceToNumber`.
- [x] `ArrayEvaluation` has no element-wise unary path (checked) — nothing to mirror.

### Verification Plan
- `--treenode-filter "/*/*/UnaryOperationTests/*"` → 14/14 pass (5 failed before the fix). ✔

### Phase Summary
Issue table reproduced: `+text` → text, `+"literal"` → text, `+TRUE` → Boolean kind, `+42` → 42,
`+blank` → 0, `+#DIV/0!` → `#DIV/0!`, `-text` → `#VALUE!`, `%text` → `#VALUE!`; plus
`SUM(+OFFSET(A1,0,0,3,1))` = 6 (reference pass-through).

## Phase 2: Absolute whole-row endpoints (`$1:$1`, `1:$1`, `$1:1`, `Sheet!$1:$1`)
Status: Complete

- [x] `tests/Danfma.MySheet.Tests/Parsing/AbsoluteRowReferenceTests.cs` (21 cases; 18 RED before fix).
- [x] `CellAddress.TryParseRow` + new `NameReference` arm in `Parser.TryEndpoint`; the qualified path gets
  it for free via `TryEndpointToken` → `TryBuildOpenRange`.
- [x] `DependencyExtractor`/`ReverseDependencyGraph` unchanged: the fix only changes which node is produced
  (`OpenRangeReference`, already handled).

### Verification Plan
- `--treenode-filter "/*/*/AbsoluteRowReferenceTests/*"` → 21/21 pass. ✔ (20/21 after this phase alone;
  the LET-based issue formula needed Phase 2b.)

### Phase Summary
Issue matrix reproduced: unqualified `$1:$1`, `$1:$1000`, `1:$1`, `$1:1` equal the relative form;
qualified `'Other Sheet'!$1:$1` etc. parse and return 3 / 2 / 40. AST equality asserted:
`=$1:$1` ≡ `OpenRangeReference(null,null,1,1,"S")`. Round-trip writes `1:1`.

## Phase 2b: `LET` (and `CHOOSE`, defined names) capture ranges as reference values
Status: Complete

Discovered while running the issue's representative formula after Phase 2: `LET(hdr,'Other Sheet'!$1:$1,
MATCH("h2",hdr,0))` → `#N/A`, also with `1:1` and `A1:C1`. Not caused by `$`; fixed because the issue's
formula cannot work without it.

- [x] 11 tests added to `tests/Danfma.MySheet.Tests/Parsing/LetFunctionTests.cs` (9 RED before fix:
  SUM/MATCH/INDEX/ROWS/COUNTA over bound bounded, whole-row, absolute-row, whole-column and union ranges;
  nested LET). Guards: single cell bound by value; bound range used as scalar → `#VALUE!`.
- [x] `NamedReferences.CaptureValue` shared helper; `EvaluateDefinition`, `Let`, `Choose` use it.

### Verification Plan
- `--treenode-filter "/*/*/LetFunctionTests/*"` → 16/16 pass. ✔
- The issue's LET/MATCH/INDEX header-lookup formula → 40. ✔

### Phase Summary
One rule, three call sites. `DynamicRange` already evaluates to a reference value on its own, so it needs
no arm in the helper.

## Phase 3: Structured `ParseException`
Status: Complete

- [x] `ParseErrorKind` enum + `Kind`/`Token` properties; constructor `(kind, message, position, token)`.
- [x] All 12 throw sites updated (Tokenizer ×3, Parser ×9); `grep -c ParseErrorKind` on throw sites = 12/12.
- [x] `tests/Danfma.MySheet.Tests/Parsing/ParseExceptionTests.cs` — 13 tests, one or two per kind, plus
  `Message` suffix and body-relative `Position` checks.
- [x] Docs: `docs/workbook-and-expressions.md` rule bullet rewritten.

### Verification Plan
- `--treenode-filter "/*/*/ParseExceptionTests|ExpressionParserTests|TokenizerTests/*"` → 63/63. ✔

### Phase Summary
Token text per kind: the unexpected character; the token found where another was expected (empty at end
of input); the function name for a bad arity; the literal's remainder for unterminated string/quoted name.

## Phase 4: Full regression, docs, commits, PR
Status: In progress

- [x] `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` → 0 errors; the single warning
  (TUnitAssertions0015 in a new test) fixed with `.IsTrue()`.
- [x] Core suite: 1176 total, 0 failed. Excel suite: 74 total, 0 failed. No deletions in existing test
  files (`git diff --numstat -- tests/`), so the 55 new cases are pure additions over `main`.
- [x] Docs: operator table (unary `+`), open-range example (`$1:$1`), LET range capture (two places),
  `ParseException` bullet.
- [ ] Commits (English, conventional, no AI attribution), one per defect so versionize lists them
  separately in the CHANGELOG:
  1. `fix(eval): unary + is a type-preserving no-op`
  2. `fix(parser): accept absolute row endpoints ($1:$1) in whole-row ranges`
  3. `fix(eval): LET, CHOOSE and defined names capture ranges as reference values`
  4. `feat(parser): structured ParseException (Kind, Token, Position)`
  5. `docs(plans): record issue #8 fix`
- [ ] Push and open the PR against `main` referencing `#8`; do NOT merge (user reviews).

### Verification Plan
- `git diff --stat main..HEAD` lists only the files above; `dotnet csharpier check .` clean.
- PR body lists the API change (ParseException constructor) explicitly.

### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
Release is automated by `versionize` in `.github/workflows/release.yml`: merging this branch into `main`
and running the workflow bumps the version from the conventional commits (three `fix:` + one `feat:` →
**3.16.0**), regenerates `CHANGELOG.md`, tags and publishes the NuGet package. No manual step beyond
merging the PR and triggering the workflow. Consumers upgrading from 3.15.0 must only touch code that
CONSTRUCTS `ParseException` (none known); catch sites are source-compatible.
