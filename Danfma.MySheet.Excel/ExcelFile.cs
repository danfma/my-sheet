using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Reads Excel (.xlsx) files into a MySheet <see cref="Workbook"/> — cross-platform via the OpenXML SDK,
/// no Excel installation required. Formula cells are parsed into real <c>Expression</c> trees (re-evaluated
/// by the MySheet engine); plain cells become literal values. Dates stay as serial numbers, and a
/// shared-formula cell that carries no formula text (a "slave" of a dragged formula) falls back to its
/// cached literal value.
/// </summary>
public static class ExcelFile
{
    /// <summary>Loads an .xlsx file into a new <see cref="Workbook"/>.</summary>
    public static Workbook Load(string path)
    {
        using var stream = File.OpenRead(path);

        return Load(stream);
    }

    /// <summary>Loads an .xlsx document from a readable stream into a new <see cref="Workbook"/>.</summary>
    public static Workbook Load(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var workbookPart =
            document.WorkbookPart
            ?? throw new InvalidDataException("The document does not contain a workbook part.");

        var sharedStrings = LoadSharedStrings(workbookPart);
        var workbook = new Workbook();

        // Iterating workbook.xml's sheet list in document order makes our Sheet.Index match Excel's tab order.
        foreach (var sheetElement in workbookPart.Workbook?.Sheets?.Elements<XlsxSheet>() ?? [])
        {
            if (sheetElement.Name?.Value is not { } name || sheetElement.Id?.Value is not { } relationshipId)
            {
                continue;
            }

            var sheet = workbook.Sheets.Add(name);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

            // Masters of shared-formula groups, per worksheet: si → (master cell, formula text). Document
            // order guarantees the master (the first cell of the group's ref) is seen before its slaves.
            var sharedFormulas = new Dictionary<uint, (string MasterId, string Formula)>();

            foreach (var cell in worksheetPart.Worksheet?.Descendants<Cell>() ?? [])
            {
                LoadCell(cell, sheet, sharedStrings, sharedFormulas);
            }
        }

        return workbook;
    }

    private static void LoadCell(
        Cell cell,
        Sheet sheet,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, (string MasterId, string Formula)> sharedFormulas
    )
    {
        if (cell.CellReference?.Value is not { } id)
        {
            return;
        }

        // A formula cell becomes a real expression tree, re-evaluated by the MySheet engine (the cached
        // <v> is ignored). A shared-formula master also registers its text for the group.
        if (cell.CellFormula?.Text is { Length: > 0 } formula)
        {
            if (cell.CellFormula.FormulaType?.Value == CellFormulaValues.Shared
                && cell.CellFormula.SharedIndex?.Value is { } masterIndex)
            {
                sharedFormulas[masterIndex] = (id, formula);
            }

            sheet[id] = ExpressionParser.Parse("=" + formula, sheet);

            return;
        }

        // A shared-formula slave carries no text: expand it from its master, shifting relative references
        // by the cell delta — exactly what Excel does. Without a known master it falls back to the cached
        // literal below.
        if (cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared
            && cell.CellFormula.SharedIndex?.Value is { } slaveIndex
            && sharedFormulas.TryGetValue(slaveIndex, out var master))
        {
            var shifted = SharedFormulaShifter.Shift(master.Formula, master.MasterId, id);

            sheet[id] = ExpressionParser.Parse("=" + shifted, sheet);

            return;
        }

        if (LoadLiteral(cell, sharedStrings) is { } literal)
        {
            sheet[id] = literal;
        }
    }

    private static Expression? LoadLiteral(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        var type = cell.DataType?.Value;

        if (type == CellValues.InlineString)
        {
            return new StringValue(cell.InlineString?.InnerText ?? string.Empty);
        }

        if (cell.CellValue?.InnerText is not { } raw)
        {
            // No value and no formula: a style-only/empty cell, which MySheet models as blank.
            return null;
        }

        if (type is null || type == CellValues.Number)
        {
            return new NumberValue(double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture));
        }

        if (type == CellValues.SharedString)
        {
            return new StringValue(sharedStrings[int.Parse(raw, CultureInfo.InvariantCulture)]);
        }

        if (type == CellValues.Boolean)
        {
            return new BooleanValue(raw is "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        if (type == CellValues.Error)
        {
            return new ErrorValue(raw);
        }

        if (type == CellValues.Date)
        {
            // ISO-8601 dates (strict-mode files) are converted to the same serial-number form Excel uses
            // in transitional files, keeping "dates are serial numbers" consistent across both.
            return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
                ? new NumberValue(date.ToOADate())
                : new StringValue(raw);
        }

        // CellValues.String (a formula's cached string) and anything unrecognized read as text.
        return new StringValue(raw);
    }

    private static IReadOnlyList<string> LoadSharedStrings(WorkbookPart workbookPart)
    {
        var table = workbookPart.SharedStringTablePart?.SharedStringTable;

        if (table is null)
        {
            return [];
        }

        // InnerText flattens rich-text runs into the plain concatenated string.
        return [.. table.Elements<SharedStringItem>().Select(item => item.InnerText)];
    }
}
