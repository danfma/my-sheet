# ExcelMerge: on-demand values + ignore-sheets overload

Two focused changes to `Danfma.MySheet.Excel/ExcelMerge.cs`: (1) stop materializing every
computed value into an up-front dictionary — evaluate each cell on demand at write time; (2) add
an overload that skips a caller-supplied set of sheet names during the merge. Behavior of the
existing merge output stays identical.

## Scope

- **In:** the two items above, in `ExcelMerge.cs`, plus tests in
  `tests/Danfma.MySheet.Excel.Tests/ExcelMergeTests.cs`.
- **Out (separate plans later):** the feature follow-ups (INDIRECT, OFFSET height/width
  truncation, cross-sheet `:`), and the memory backlog (per-workbook interning, CellStore
  streaming serialization, sparse arrays).

## Decisions (confirmed)

- On-demand: **only eliminate the global `computed` dictionary**. Keep the per-sheet
  `OrderBy(Row).ThenBy(Column)` (it is per-sheet, bounded, and needed for ordered OpenXML
  insertion via `GetOrCreateRow`/`GetOrCreateCell`).
- Skip API: an **overload** `MergeIntoExcel(this Workbook, string path, IReadOnlySet<string> ignoredSheets)`;
  matching is **case-insensitive**, consistent with the sheet-name matching already used in the file.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status
to `Complete` and write its **Phase Summary**; run the phase's **Verification Plan** and record the
result before moving on. When all phases are done, fill in **Final Recap** and **Deployment Plan**.

Test framework is **TUnit** — `dotnet test` does NOT work on .NET 10. Run the Excel suite with:
`dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release -- --treenode-filter "/*/*/ExcelMergeTests/*"`

---

## Phase 1: Evaluate cells on demand (drop the global dictionary)
Status: Complete

Current `MergeIntoExcel` (`ExcelMerge.cs:23-78`) pre-evaluates every cell of every sheet into a
`Dictionary<(string Sheet, string Id), ComputedValue> computed` on a large-stack thread (lines
28-41), then opens the OpenXML doc and reads from that dictionary in the write loop (line 66). The
dictionary is a full second copy of every value (the workbook already memoizes them). Eliminate it
by running the whole merge — doc open + write loop — inside one `RunWithLargeStack` call and
calling `workbook.GetCellValue(sheet.Name, id)` on demand at write time.

- [x] Confirm the `RunWithLargeStack` overload in use: `ExcelMerge.cs:28` uses `Workbook.RunWithLargeStack(Func<T>)`. Check whether an `Action` overload exists (grep `RunWithLargeStack` in `Danfma.MySheet/Workbook.cs`). If only `Func<T>`, return a dummy `0` from the lambda; if an `Action` overload exists, use it.
- [x] Rewrite `MergeIntoExcel(this Workbook workbook, string path)`:
  - Keep `var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();` OUTSIDE the large-stack call (sorting sheets needs no large stack).
  - Remove the `computed` dictionary block (lines 28-41) entirely.
  - Wrap the doc-open + `foreach (sheet in orderedSheets)` write loop in a single `Workbook.RunWithLargeStack(() => { ... })`.
  - Inside the loop, replace `var value = values[(sheet.Name, id)];` with `var value = workbook.GetCellValue(sheet.Name, id);`.
  - Everything else (`FindWorksheet`, `orderedCells` OrderBy, blank skip, `GetOrCreateCell`/`GetOrCreateRow`, `WriteLiteral`) stays byte-identical.

  Target shape:

  ```csharp
  public static void MergeIntoExcel(this Workbook workbook, string path)
  {
      var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

      // Whole merge on one large-stack thread: deep formula chains evaluate safely AND the OpenXML
      // write runs together, so each cell is computed on demand at write time — no up-front value
      // dictionary (the workbook already memoizes each value in its own store).
      Workbook.RunWithLargeStack(() =>
      {
          using var document = SpreadsheetDocument.Open(path, isEditable: true);

          var workbookPart =
              document.WorkbookPart
              ?? throw new InvalidDataException("The document does not contain a workbook part.");

          foreach (var sheet in orderedSheets)
          {
              if (FindWorksheet(workbookPart, sheet.Name) is not { } worksheetPart)
              {
                  continue;
              }

              var worksheet = worksheetPart.Worksheet ??= new Worksheet();
              var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

              var orderedCells = sheet
                  .Select(entry => (entry.Key, Position: CellId.Parse(entry.Key)))
                  .OrderBy(cell => cell.Position.Row)
                  .ThenBy(cell => cell.Position.Column);

              foreach (var (id, position) in orderedCells)
              {
                  var value = workbook.GetCellValue(sheet.Name, id); // on demand, not from a prebuilt dict

                  if (value.Kind == ComputedValueKind.Blank)
                  {
                      continue;
                  }

                  var cell = GetOrCreateCell(GetOrCreateRow(sheetData, position.Row), id, position.Column);
                  WriteLiteral(cell, value);
              }
          }

          return 0; // if only Func<T> exists; use the Action overload instead if present
      });
  }
  ```

- [x] Build the Excel project: `dotnet build Danfma.MySheet.Excel/Danfma.MySheet.Excel.csproj -c Release` — 0 errors, 0 warnings.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release -- --treenode-filter "/*/*/ExcelMergeTests/*"` → the existing `Merge_InjectsLiterals_DropsTargetFormulas_PreservesEverythingElse` and `Merge_TemplateWorkflow_IsCopyThenMerge` still PASS (proves the output is unchanged — same literals, dropped formulas, preserved cells/sheets).
- Grep confirms the dictionary is gone: `grep -c "new Dictionary<(string Sheet, string Id)" Danfma.MySheet.Excel/ExcelMerge.cs` → `0`.
- Grep confirms on-demand eval in the loop: `grep -c "workbook.GetCellValue(sheet.Name, id)" Danfma.MySheet.Excel/ExcelMerge.cs` → `1`.

### Phase Summary
Done. `MergeIntoExcel` now runs the entire merge (OpenXML open + write loop) inside a single
`Workbook.RunWithLargeStack(Func<T>)` call (only a `Func<T>` overload exists → the lambda returns an
unused `0`), and computes each cell with `workbook.GetCellValue(sheet.Name, id)` at write time. The
`Dictionary<(string Sheet, string Id), ComputedValue>` that pre-materialized every value is gone —
removing a full second copy of all cell values (~one entry per non-blank cell, e.g. ~549k at K1
scale). The per-sheet `OrderBy(Row).ThenBy(Column)` is kept (bounded, needed for ordered OpenXML
insertion). **Verification result:** build 0/0; `ExcelMergeTests` 2/2 pass (output unchanged); dict
grep = 0, on-demand grep = 1. No behavior change. Committed on branch `feat/excelmerge-ondemand-skip`.

---

## Phase 2: `MergeIntoExcel` overload that skips a set of sheets
Status: Not started

Add an overload that accepts sheet names to skip. The existing no-arg method delegates to it with
an empty set, so there is one implementation body.

- [ ] Add a private static empty set to delegate to: `private static readonly IReadOnlySet<string> NoIgnoredSheets = new HashSet<string>();`
- [ ] Change the no-arg method to delegate: `public static void MergeIntoExcel(this Workbook workbook, string path) => workbook.MergeIntoExcel(path, NoIgnoredSheets);`
- [ ] Add the overload holding the merge body from Phase 1, with the skip:

  ```csharp
  /// <summary>Merges in place, skipping any sheet whose name is in <paramref name="ignoredSheets"/>
  /// (case-insensitive). Skipped sheets are neither evaluated nor written; the target's copies of
  /// those sheets are left untouched.</summary>
  public static void MergeIntoExcel(this Workbook workbook, string path, IReadOnlySet<string> ignoredSheets)
  {
      // Normalize to a case-insensitive lookup regardless of the caller's set comparer, matching the
      // case-insensitive sheet-name matching used elsewhere in this file. Bounded by sheet count.
      var ignored = ignoredSheets.Count == 0
          ? null
          : new HashSet<string>(ignoredSheets, StringComparer.OrdinalIgnoreCase);

      var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

      Workbook.RunWithLargeStack(() =>
      {
          using var document = SpreadsheetDocument.Open(path, isEditable: true);
          var workbookPart =
              document.WorkbookPart
              ?? throw new InvalidDataException("The document does not contain a workbook part.");

          foreach (var sheet in orderedSheets)
          {
              if (ignored is not null && ignored.Contains(sheet.Name))
              {
                  continue; // caller asked to skip this sheet
              }

              if (FindWorksheet(workbookPart, sheet.Name) is not { } worksheetPart)
              {
                  continue;
              }
              // ... identical body to Phase 1 (worksheet/sheetData, orderedCells, on-demand GetCellValue, WriteLiteral) ...
          }

          return 0;
      });
  }
  ```

  The no-arg method's body becomes the one-line delegation above (do NOT keep a second copy of the
  loop — the overload is the single implementation).

- [ ] Update the `<summary>` XML doc on the class/method to mention the new skip behavior.
- [ ] Build: `dotnet build Danfma.MySheet.Excel/Danfma.MySheet.Excel.csproj -c Release` — 0 warnings.

### Verification Plan
- Add two TUnit tests to `tests/Danfma.MySheet.Excel.Tests/ExcelMergeTests.cs`:
  1. `Merge_WithIgnoredSheet_SkipsIt`: target has sheets `Data` and `Skip` (e.g. `Skip!A1 = "orig"`); workbook has `Data!A2=5` and `Skip!A1=42`; call `wb.MergeIntoExcel(path, new HashSet<string> { "Skip" })`; assert `Data!A2 == 5` (merged) AND `Skip!A1` is still `"orig"` (untouched, not overwritten with 42).
  2. `Merge_IgnoredSheet_IsCaseInsensitive`: same setup, ignore set `{ "skip" }` (lowercase); assert `Skip!A1` is still `"orig"`.
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release -- --treenode-filter "/*/*/ExcelMergeTests/*"` → all ExcelMergeTests (the 2 existing + 2 new) PASS.
- Full Excel suite green: `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → 0 failures.

### Phase Summary
_(write when phase completes)_

---

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
