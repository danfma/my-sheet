using System.Globalization;
using Danfma.MySheet.Expressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Merges computed values from a MySheet <see cref="Workbook"/> into an EXISTING .xlsx file, in place:
/// every cell we hold is written as its computed literal value — dropping any formula the target cell had —
/// while everything else in the file (styles, other cells, other sheets, shared strings) is left intact.
/// Sheets are matched by name (case-insensitive); sheets missing from the target are skipped, blank values
/// are not written. Text is written as an inline string so the target's shared-string table is untouched.
/// To produce a new report from a pristine template, copy the template first
/// (<c>File.Copy(template, output)</c>) and merge into the copy.
/// </summary>
public static class ExcelMerge
{
    // Shared empty set so the no-argument overload delegates without allocating per call.
    private static readonly IReadOnlySet<string> NoIgnoredSheets = new HashSet<string>();

    /// <summary>Merges in place, editing <paramref name="path"/> directly.</summary>
    public static void MergeIntoExcel(this Workbook workbook, string path) =>
        workbook.MergeIntoExcel(path, NoIgnoredSheets);

    /// <summary>
    /// Merges in place, skipping any sheet whose name is in <paramref name="ignoredSheets"/>
    /// (case-insensitive). A skipped sheet is neither evaluated nor written, so the target's copy of
    /// that sheet is left exactly as it was.
    /// </summary>
    public static void MergeIntoExcel(
        this Workbook workbook,
        string path,
        IReadOnlySet<string> ignoredSheets
    )
    {
        // Normalize to a case-insensitive lookup regardless of the caller's set comparer, matching the
        // case-insensitive sheet-name matching used elsewhere in this file. Bounded by sheet count.
        var ignored =
            ignoredSheets.Count == 0
                ? null
                : new HashSet<string>(ignoredSheets, StringComparer.OrdinalIgnoreCase);

        var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

        // The whole merge runs on one large-stack thread: deep formula chains evaluate safely (see
        // ExcelExport for the rationale) AND the OpenXML write runs together, so each cell is computed
        // ON DEMAND at write time — no up-front dictionary duplicating every value (the workbook already
        // memoizes each value in its own store).
        Workbook.RunWithLargeStack(() =>
        {
            using var document = SpreadsheetDocument.Open(path, isEditable: true);

            var workbookPart =
                document.WorkbookPart
                ?? throw new InvalidDataException("The document does not contain a workbook part.");

            // Merging overrides/drops formula cells, so any calcChain the target carried is now stale.
            // Left in place, Excel reports "Removed records: Formula from /xl/calcChain.xml" and forces a
            // repair on open. Dropping the part lets Excel rebuild the calc chain cleanly and silently.
            if (workbookPart.CalculationChainPart is { } calcChain)
            {
                workbookPart.DeletePart(calcChain);
            }

            foreach (var sheet in orderedSheets)
            {
                if (ignored is not null && ignored.Contains(sheet.Name))
                {
                    continue; // caller asked to skip this sheet: not evaluated, not written
                }

                if (FindWorksheet(workbookPart, sheet.Name) is not { } worksheetPart)
                {
                    continue; // the target has no sheet with this name: skipped by design
                }

                var worksheet = worksheetPart.Worksheet ??= new Worksheet();
                var sheetData =
                    worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

                // Index the existing rows/cells once so each write is an O(1) hash lookup instead of a linear
                // DOM scan — a fully-populated target (the common template case) otherwise turns every write
                // into an O(rows) walk, i.e. O(rows²) overall.
                var index = new SheetDataIndex(sheetData);

                var orderedCells = sheet
                    .Select(entry => (entry.Key, Position: CellId.Parse(entry.Key)))
                    .OrderBy(cell => cell.Position.Row)
                    .ThenBy(cell => cell.Position.Column);

                foreach (var (id, position) in orderedCells)
                {
                    var value = workbook.GetCellValue(sheet.Name, id); // computed on demand, not prebuilt

                    if (value.Kind == ComputedValueKind.Blank)
                    {
                        continue; // blanks are not written, leaving the target cell as it was
                    }

                    var cell = index.GetOrCreateCell(id, position.Row, position.Column);

                    WriteLiteral(cell, value);
                }
            }

            return 0; // RunWithLargeStack exposes only a Func<T> overload; the result is unused
        });
    }

    private static WorksheetPart? FindWorksheet(WorkbookPart workbookPart, string sheetName)
    {
        var sheetElement = workbookPart
            .Workbook?.Sheets?.Elements<XlsxSheet>()
            .FirstOrDefault(sheet =>
                string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)
            );

        return sheetElement?.Id?.Value is { } relationshipId
            ? (WorksheetPart)workbookPart.GetPartById(relationshipId)
            : null;
    }

    // A hash index over one worksheet's <sheetData>, so getting/creating a row or cell is O(1) instead of a
    // linear DOM walk. Rows must stay ordered by RowIndex and cells by column; because writes arrive in
    // ascending (row, column) order, the maxima let the common append land in O(1), and the rare
    // out-of-order insert (a target with gaps we fill low) still finds its successor without touching the
    // hot path.
    private sealed class SheetDataIndex
    {
        private readonly SheetData _sheetData;
        private readonly Dictionary<uint, XlsxRow> _rows = new();
        private readonly Dictionary<XlsxRow, RowCells> _cells = new();
        private uint _maxRow;

        public SheetDataIndex(SheetData sheetData)
        {
            _sheetData = sheetData;

            foreach (var row in sheetData.Elements<XlsxRow>())
            {
                if (row.RowIndex?.Value is { } rowIndex)
                {
                    _rows[rowIndex] = row;
                    _maxRow = Math.Max(_maxRow, rowIndex);
                }
            }
        }

        public Cell GetOrCreateCell(string id, int row, int column)
        {
            var target = GetOrCreateRow((uint)row);

            if (!_cells.TryGetValue(target, out var rowCells))
            {
                rowCells = new RowCells(target);
                _cells[target] = rowCells;
            }

            return rowCells.GetOrCreate(id, (uint)column);
        }

        private XlsxRow GetOrCreateRow(uint rowIndex)
        {
            if (_rows.TryGetValue(rowIndex, out var existing))
            {
                return existing;
            }

            var created = new XlsxRow { RowIndex = rowIndex };

            if (rowIndex > _maxRow)
            {
                _sheetData.AppendChild(created);
                _maxRow = rowIndex;
            }
            else
            {
                var successor = _sheetData.Elements<XlsxRow>().First(row => row.RowIndex?.Value > rowIndex);
                _sheetData.InsertBefore(created, successor);
            }

            _rows[rowIndex] = created;

            return created;
        }
    }

    // The per-row mirror of SheetDataIndex: column → cell, keeping cells ordered by column.
    private sealed class RowCells
    {
        private readonly XlsxRow _row;
        private readonly Dictionary<uint, Cell> _cells = new();
        private uint _maxColumn;

        public RowCells(XlsxRow row)
        {
            _row = row;

            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } reference)
                {
                    var column = (uint)CellId.Parse(reference).Column;
                    _cells[column] = cell;
                    _maxColumn = Math.Max(_maxColumn, column);
                }
            }
        }

        public Cell GetOrCreate(string id, uint column)
        {
            if (_cells.TryGetValue(column, out var existing))
            {
                return existing;
            }

            var created = new Cell { CellReference = id };

            if (column > _maxColumn)
            {
                _row.AppendChild(created);
                _maxColumn = column;
            }
            else
            {
                var successor = _row
                    .Elements<Cell>()
                    .First(cell =>
                        cell.CellReference?.Value is { } reference
                        && CellId.Parse(reference).Column > column
                    );
                _row.InsertBefore(created, successor);
            }

            _cells[column] = created;

            return created;
        }
    }

    // Replaces the cell's CONTENT with a literal (dropping any formula) while leaving its style —
    // formatting lives in the untouched StyleIndex, so a formatted template cell stays formatted.
    private static void WriteLiteral(Cell cell, in ComputedValue value)
    {
        cell.CellFormula = null;
        cell.InlineString = null;
        cell.DataType = null;
        cell.CellValue = null;

        switch (value.Kind)
        {
            case ComputedValueKind.Number:
                cell.CellValue = new CellValue(value.ToDouble().ToString(CultureInfo.InvariantCulture));
                break;

            case ComputedValueKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(value.ToBoolean() ? "1" : "0");
                break;

            case ComputedValueKind.Text:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(XlsxTextFactory.Create(value.ToText()));
                break;

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(error.ToString());
                break;

            case ComputedValueKind.Reference:
                // A bare reference result (e.g. a multi-cell OFFSET) has no single cell value.
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(Error.Value.ToString());
                break;
        }
    }
}
