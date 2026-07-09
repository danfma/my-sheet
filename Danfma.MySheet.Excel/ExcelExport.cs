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

        // The whole write runs on one large-stack thread: each cell value is computed ON DEMAND at write
        // time (deep dependency chains cannot overflow) and streamed straight out via OpenXmlWriter, so we
        // never hold a second dictionary of every value NOR a full worksheet DOM — only the workbook's own
        // memoized store plus the (distinct-string-bounded) shared-string table stay resident.
        Workbook.RunWithLargeStack(() =>
        {
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            var workbookPart = document.AddWorkbookPart();
            var sharedStrings = new SharedStringRegistry();
            var sheetEntries = new List<(string Id, string Name)>(orderedSheets.Length);

            // Stream each worksheet part first: this is where the shared-string registry gets populated,
            // so the shared-string table must be written afterwards.
            foreach (var sheet in orderedSheets)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                WriteWorksheet(worksheetPart, sheet, workbook, mode, sharedStrings);
                sheetEntries.Add((workbookPart.GetIdOfPart(worksheetPart), sheet.Name));
            }

            var xlsxWorkbook = new XlsxWorkbook();
            var sheetList = xlsxWorkbook.AppendChild(new Sheets());
            var sheetId = 1u;

            foreach (var (id, name) in sheetEntries)
            {
                sheetList.AppendChild(new XlsxSheet { Id = id, SheetId = sheetId++, Name = name });
            }

            workbookPart.Workbook = xlsxWorkbook;

            // Defined names go after <sheets> in the workbook element (schema order).
            WriteDefinedNames(workbookPart, workbook);

            sharedStrings.WriteTo(workbookPart);

            return 0; // RunWithLargeStack exposes only a Func<T> overload; the result is unused
        });
    }

    private static void WriteDefinedNames(WorkbookPart workbookPart, Workbook workbook)
    {
        if (workbook.DefinedNames.Count == 0)
        {
            return;
        }

        var definedNames = new DefinedNames();

        foreach (var (name, expression) in workbook.DefinedNames)
        {
            // An empty un-parse context matches no sheet, so every reference is emitted fully qualified
            // (e.g. Data!A1:A3) — required because a workbook-level name has no implicit sheet.
            definedNames.AppendChild(
                new DefinedName(expression.ToFormula(string.Empty)) { Name = name }
            );
        }

        (workbookPart.Workbook ??= new XlsxWorkbook()).AppendChild(definedNames);
    }

    // Streams one worksheet part cell-by-cell via OpenXmlWriter: a transient Cell is built, written, and
    // immediately eligible for GC, so we never materialize the whole <sheetData> tree in memory. Values are
    // computed on demand here (we are already on the large-stack thread from SaveAsExcel).
    private static void WriteWorksheet(
        WorksheetPart part,
        Sheet sheet,
        Workbook workbook,
        FormulaMode mode,
        SharedStringRegistry sharedStrings
    )
    {
        using var writer = OpenXmlWriter.Create(part);

        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());

        // OpenXML requires rows in order and cells in column order within the row.
        var orderedCells = sheet
            .Select(entry => (entry.Key, Position: CellId.Parse(entry.Key), Expression: entry.Value))
            .OrderBy(cell => cell.Position.Row)
            .ThenBy(cell => cell.Position.Column);

        var currentRow = -1;

        foreach (var (id, position, expression) in orderedCells)
        {
            var value = workbook.GetCellValue(sheet.Name, id);
            var cell = BuildCell(id, expression, value, sheet.Name, mode, sharedStrings);

            if (cell is null)
            {
                continue; // blank: the cell is simply omitted
            }

            if (position.Row != currentRow)
            {
                if (currentRow != -1)
                {
                    writer.WriteEndElement(); // close the previous <row>
                }

                writer.WriteStartElement(new XlsxRow { RowIndex = (uint)position.Row });
                currentRow = position.Row;
            }

            writer.WriteElement(cell);
        }

        if (currentRow != -1)
        {
            writer.WriteEndElement(); // close the last <row>
        }

        writer.WriteEndElement(); // </sheetData>
        writer.WriteEndElement(); // </worksheet>
        writer.Close();
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
                _items.Select(item => new SharedStringItem(XlsxTextFactory.Create(item)))
            );
        }
    }
}
