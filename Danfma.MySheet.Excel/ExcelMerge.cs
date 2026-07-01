using System.Globalization;
using Danfma.MySheet.Expressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

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
    /// <summary>Merges in place, editing <paramref name="path"/> directly.</summary>
    public static void MergeIntoExcel(this Workbook workbook, string path)
    {
        var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

        // Evaluate everything up front on one large-stack thread (see ExcelExport for the rationale).
        var values = Workbook.RunWithLargeStack(() =>
        {
            var computed = new Dictionary<(string Sheet, string Id), ComputedValue>();

            foreach (var sheet in orderedSheets)
            {
                foreach (var id in sheet.Keys)
                {
                    computed[(sheet.Name, id)] = workbook.GetCellValue(sheet.Name, id);
                }
            }

            return computed;
        });

        using var document = SpreadsheetDocument.Open(path, isEditable: true);

        var workbookPart =
            document.WorkbookPart
            ?? throw new InvalidDataException("The document does not contain a workbook part.");

        foreach (var sheet in orderedSheets)
        {
            if (FindWorksheet(workbookPart, sheet.Name) is not { } worksheetPart)
            {
                continue; // the target has no sheet with this name: skipped by design
            }

            var worksheet = worksheetPart.Worksheet ??= new Worksheet();
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var orderedCells = sheet
                .Select(entry => (entry.Key, Position: CellId.Parse(entry.Key)))
                .OrderBy(cell => cell.Position.Row)
                .ThenBy(cell => cell.Position.Column);

            foreach (var (id, position) in orderedCells)
            {
                var value = values[(sheet.Name, id)];

                if (value.Kind == ComputedValueKind.Blank)
                {
                    continue; // blanks are not written, leaving the target cell as it was
                }

                var cell = GetOrCreateCell(GetOrCreateRow(sheetData, position.Row), id, position.Column);

                WriteLiteral(cell, value);
            }
        }
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

    // Rows must stay ordered by RowIndex, so a missing row is inserted before the first higher one.
    private static XlsxRow GetOrCreateRow(SheetData sheetData, int rowIndex)
    {
        foreach (var row in sheetData.Elements<XlsxRow>())
        {
            if (row.RowIndex?.Value == (uint)rowIndex)
            {
                return row;
            }

            if (row.RowIndex?.Value > (uint)rowIndex)
            {
                return sheetData.InsertBefore(new XlsxRow { RowIndex = (uint)rowIndex }, row);
            }
        }

        return sheetData.AppendChild(new XlsxRow { RowIndex = (uint)rowIndex });
    }

    // Cells must stay ordered by column within their row, mirroring GetOrCreateRow.
    private static Cell GetOrCreateCell(XlsxRow row, string id, int column)
    {
        foreach (var cell in row.Elements<Cell>())
        {
            var reference = cell.CellReference?.Value;

            if (string.Equals(reference, id, StringComparison.OrdinalIgnoreCase))
            {
                return cell;
            }

            if (reference is not null && CellId.Parse(reference).Column > column)
            {
                return row.InsertBefore(new Cell { CellReference = id }, cell);
            }
        }

        return row.AppendChild(new Cell { CellReference = id });
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
                cell.InlineString = new InlineString(new XlsxText(value.ToText()));
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
