using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;
using XlsxWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;

namespace Danfma.MySheet.Excel;

/// <summary>How <see cref="ExcelExport.SaveAsExcel(Workbook, string, ExcelExportOptions?)"/> writes formula cells.</summary>
public enum FormulaMode
{
    /// <summary>Write only computed literal values — a flattened snapshot with no formulas.</summary>
    ValuesOnly,

    /// <summary>Write the Excel formula (<c>&lt;f&gt;</c>) plus its computed value (<c>&lt;v&gt;</c>).</summary>
    Formulas,
}

public sealed record ExcelExportOptions
{
    public FormulaMode FormulaMode { get; init; } = FormulaMode.ValuesOnly;
}

/// <summary>
/// Writes a MySheet <see cref="Workbook"/> to an Excel (.xlsx) file — cross-platform via the OpenXML SDK,
/// no Excel installation required. Every cell's value is computed by the MySheet engine (memoized, on a
/// large-stack thread); <see cref="FormulaMode.Formulas"/> also emits the formula text via the core
/// <see cref="FormulaWriter"/> so Excel keeps recalculating the file.
/// </summary>
public static class ExcelExport
{
    /// <summary>Saves the workbook as an .xlsx file.</summary>
    public static void SaveAsExcel(this Workbook workbook, string path, ExcelExportOptions? options = null)
    {
        using var stream = File.Create(path);

        workbook.SaveAsExcel(stream, options);
    }

    /// <summary>Saves the workbook as an .xlsx document into a writable stream.</summary>
    public static void SaveAsExcel(this Workbook workbook, Stream stream, ExcelExportOptions? options = null)
    {
        var mode = (options ?? new ExcelExportOptions()).FormulaMode;
        var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

        // Evaluate everything up front on one large-stack thread, so deep dependency chains cannot
        // overflow while cells are being written. Results are memoized per cell.
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

        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new XlsxWorkbook();

        var sheetList = workbookPart.Workbook.AppendChild(new Sheets());
        var sharedStrings = new SharedStringRegistry();
        var sheetId = 1u;

        foreach (var sheet in orderedSheets)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(BuildSheetData(sheet, values, mode, sharedStrings));

            sheetList.AppendChild(
                new XlsxSheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId++,
                    Name = sheet.Name,
                }
            );
        }

        sharedStrings.WriteTo(workbookPart);
    }

    private static SheetData BuildSheetData(
        Sheet sheet,
        Dictionary<(string Sheet, string Id), ComputedValue> values,
        FormulaMode mode,
        SharedStringRegistry sharedStrings
    )
    {
        var sheetData = new SheetData();
        XlsxRow? currentRow = null;

        // OpenXML requires rows in order and cells in column order within the row.
        var orderedCells = sheet
            .Select(entry => (entry.Key, Position: ParseCellId(entry.Key), Expression: entry.Value))
            .OrderBy(cell => cell.Position.Row)
            .ThenBy(cell => cell.Position.Column);

        foreach (var (id, position, expression) in orderedCells)
        {
            var cell = BuildCell(id, expression, values[(sheet.Name, id)], sheet.Name, mode, sharedStrings);

            if (cell is null)
            {
                continue; // blank: the cell is simply omitted
            }

            if (currentRow is null || currentRow.RowIndex! != (uint)position.Row)
            {
                currentRow = sheetData.AppendChild(new XlsxRow { RowIndex = (uint)position.Row });
            }

            currentRow.AppendChild(cell);
        }

        return sheetData;
    }

    private static Cell? BuildCell(
        string id,
        Expression expression,
        in ComputedValue value,
        string sheetName,
        FormulaMode mode,
        SharedStringRegistry sharedStrings
    )
    {
        // A literal node is written as a literal in both modes; a formula node keeps its formula text
        // (plus the cached value) in Formulas mode and is flattened to its computed value otherwise.
        if (mode == FormulaMode.Formulas && expression is not ValueExpression)
        {
            var cell = new Cell
            {
                CellReference = id,
                CellFormula = new CellFormula(expression.ToFormula(sheetName)),
            };

            ApplyCachedValue(cell, value);

            return cell;
        }

        return BuildLiteralCell(id, value, sharedStrings);
    }

    // The cached value that accompanies a formula: plain number, t="b", t="str" (NOT a shared string —
    // that is the xlsx convention for formula-produced text) or t="e".
    private static void ApplyCachedValue(Cell cell, in ComputedValue value)
    {
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
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(value.ToText());
                break;

            case ComputedValueKind.Error:
                ApplyError(cell, value);
                break;

            case ComputedValueKind.Reference:
                // A bare reference result (e.g. a multi-cell OFFSET) has no single cell value.
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(Error.Value.ToString());
                break;

            // Blank: a formula whose result is blank carries no cached value.
        }
    }

    private static Cell? BuildLiteralCell(string id, in ComputedValue value, SharedStringRegistry sharedStrings)
    {
        var cell = new Cell { CellReference = id };

        switch (value.Kind)
        {
            case ComputedValueKind.Number:
                cell.CellValue = new CellValue(value.ToDouble().ToString(CultureInfo.InvariantCulture));
                return cell;

            case ComputedValueKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(value.ToBoolean() ? "1" : "0");
                return cell;

            case ComputedValueKind.Text:
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(
                    sharedStrings.IndexOf(value.ToText()).ToString(CultureInfo.InvariantCulture)
                );
                return cell;

            case ComputedValueKind.Error:
                ApplyError(cell, value);
                return cell;

            case ComputedValueKind.Reference:
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(Error.Value.ToString());
                return cell;

            default:
                return null; // blank cells are omitted entirely
        }
    }

    private static void ApplyError(Cell cell, in ComputedValue value)
    {
        value.TryGetError(out var error);
        cell.DataType = CellValues.Error;
        cell.CellValue = new CellValue(error.ToString());
    }

    private static (int Row, int Column) ParseCellId(string id)
    {
        var index = 0;
        var column = 0;

        while (index < id.Length && char.IsLetter(id[index]))
        {
            column = column * 26 + (char.ToUpperInvariant(id[index]) - 'A' + 1);
            index++;
        }

        return (int.Parse(id[index..], CultureInfo.InvariantCulture), column);
    }

    /// <summary>Deduplicating shared-string table, written to the workbook only when any text was used.</summary>
    private sealed class SharedStringRegistry
    {
        private readonly Dictionary<string, int> _indexes = new(StringComparer.Ordinal);
        private readonly List<string> _items = [];

        public int IndexOf(string text)
        {
            if (!_indexes.TryGetValue(text, out var index))
            {
                index = _items.Count;
                _items.Add(text);
                _indexes[text] = index;
            }

            return index;
        }

        public void WriteTo(WorkbookPart workbookPart)
        {
            if (_items.Count == 0)
            {
                return;
            }

            var part = workbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable(
                _items.Select(item => new SharedStringItem(new XlsxText(item)))
            );
        }
    }
}
