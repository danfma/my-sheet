# Excel Table fixtures

Real `.xlsx` files carrying an Excel **Table** (`<table>` part, a.k.a. ListObject), authored by the oracle
rather than by ClosedXML so the `<table>` part, the `<tableParts>` relationship and the stored formula text
are exactly what an Excel-style producer writes.

- **Generator:** a throwaway console program (not part of the build) over Aspose.Cells 26.6.0. Recipe:
  `Worksheets[0].Name = "Data"`, `Cells[...].PutValue(...)` for the header and rows, then
  `ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, hasHeaders: true)` (0-based) and
  `DisplayName = "Tabela1"`; the totals variant sets `ShowTotals = true` and
  `ListColumns[1].TotalsCalculation = TotalsCalculation.Sum`; the headerless variant sets
  `ShowHeaderRow = false`. Formulas are assigned to `Cells[...].Formula` (PLAIN entry) and
  `CalculateFormula()` runs before `Save(path, SaveFormat.Xlsx)`, so every formula cell carries the
  oracle's cached `<v>`.
- **Strip:** unlicensed Aspose appends a worksheet named `Evaluation Warning` on save and leaves the
  workbook view on it. After generation the OpenXML SDK removes the `<sheet>` element, its part and the
  calc chain, and resets `workbookView/@activeTab` to 0 (ClosedXML refuses a view past the last sheet), so
  every file below carries only the sheets named here. Nothing else is touched.
- **Loader side:** DocumentFormat.OpenXml 3.5.1.

`Tabela1` = `Data!A1:B3` unless stated: header `Item`/`Valor`, rows `a`/10 and `b`/32.

| File | `<table>` part | Formula cells and the oracle's PLAIN answer |
| --- | --- | --- |
| `f1-plain.xlsx` | `ref="A1:B3" totalsRowShown="0"`, columns `Item`, `Valor`; `D10` = `Valor` (literal) | `D1 SUM(Tabela1[Valor])` 42 · `D2 ROWS(Tabela1[#All])` 3 · `D3 ROWS(Tabela1[#Data])` 2 · `D4 COUNTA(Tabela1[#Headers])` 2 · `D5 SUM(Tabela1[[#Data],[Valor]])` 42 · `D6 SUM(INDIRECT("Tabela1["&D10&"]"))` 42 · `D7 SUM(Tabela1[#Totals])` `#REF!` · `D8 SUM(Tabela1[[#Data],[#Totals]])` 42 · `D9 SUM(Tabela1)` 42 (stored `Tabela1[]`) |
| `f2-totals.xlsx` | `ref="A1:B4" totalsRowCount="1"` (no `totalsRowShown`), `autoFilter ref="A1:B3"`; `B4` = `SUBTOTAL(109,[Valor])` cached 42 | `D1 SUM(Tabela1[Valor])` 42 · `D2 ROWS(Tabela1[#All])` 4 · `D3 SUM(Tabela1[#Totals])` 42 · `D4 ROWS(Tabela1[#Data])` 2 · `D5 SUM(Tabela1[[#Data],[#Totals]])` 84 |
| `f3-noheader.xlsx` | `ref="A2:B3" headerRowCount="0" totalsRowShown="0"`, columns `Item`, `Valor` (row 1 is empty) | `D1 SUM(Tabela1[Valor])` 42 · `D2 SUM(Tabela1[Item])` 0 · `D3 SUM(Tabela1)` 42 · `D4 ROWS(Tabela1[#All])` 2 · `D5 COUNTA(Tabela1[#Headers])` 1 · `D6 ROWS(Tabela1[#Data])` 2 |
| `f4-names.xlsx` | `ref="A1:E3"`, columns `AMOUNT IN USD For Line 1`, `(A) NAME OF PFIC`, `Owner's Share`, `Line_x000a_Break` (typed with a real line break), ` Padded `; rows 10/x/1/100/7 and 32/y/2/200/8 | `G1 SUM(Tabela1[AMOUNT IN USD For Line 1])` 42 · `G2 COUNTA(Tabela1[(A) NAME OF PFIC])` 2 · `G3 SUM(Tabela1[Owner''s Share])` 3 · `G4` typed `Owner's Share`, stored `Owner''s Share`, 3 · `G5 SUM(Tabela1[Line⏎Break])` 300 (a raw newline in `<f>`) · `G6 SUM(Tabela1[[ Padded ]])` 15 |
| `f5-name-My_Table.xlsx` | `name="My\Table" displayName="My\Table" ref="A1:B3"` | `D1 SUM(My\Table[Valor])` 42 |
| `f5-name-TRUE.xlsx` | `name="TRUE" displayName="TRUE" ref="A1:B3"` | `D1 SUM(TRUE[Valor])` 42 |
| `f6-cross-sheet.xlsx` | `Tabela1` on `Data`; a second sheet `Report` (`B2` = 1000, `B3` = 2000) | `Report!D1 SUM(Tabela1[Valor])` 42 · `Report!D2` typed `SUM(Data!Tabela1[Valor])`, stored WITHOUT the qualifier, 42 |
| `f7-header-only.xlsx` | `ref="A1:B1"` — the header row alone, zero data rows | `D1 SUM(Tabela1[Valor])` 0 · `D2 ROWS(Tabela1[#Data])` 0 |
| `f8-two-tables.xlsx` | `Tabela1` (`A1:B3`) and `Tabela2` (`F1:G3`, columns `K`/`V`, rows k1/5, k2/6) on `Data`; `Tabela3` (`A1:A3`, column `X`, rows 1, 2) on `Other`; `C1` = `Dobro` | `D1 SUM(Tabela1[Valor])+SUM(Tabela2[V])+SUM(Tabela3[X])` 56 · `D2` typed `SUM(Tabela1[@Valor])`, stored `SUM(Tabela1[[#This Row],[Valor]])`, 10 · `I2` typed `Tabela1[@Valor]*2`, stored `Tabela1[[#This Row],[Valor]]*2`, 20 |

Every file also writes `name` equal to `displayName`. Measured 2026-09-11.
