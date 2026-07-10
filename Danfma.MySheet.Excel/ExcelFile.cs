using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Reads Excel (.xlsx) files into a MySheet <see cref="Workbook"/> — cross-platform via the OpenXML SDK,
/// no Excel installation required. Shared strings and worksheets are STREAMED (forward-only XmlReader over
/// each part), so the OpenXML DOM is never materialized and the only full representation in memory is the
/// MySheet model. Formula cells are parsed into real <c>Expression</c> trees (re-evaluated by the MySheet
/// engine); plain cells become literal values. Dates stay as serial numbers, and a shared-formula cell
/// that carries no formula text (a "slave" of a dragged formula) is expanded from its group master.
/// </summary>
public static class ExcelFile
{
    /// <summary>Loads an .xlsx file into a new <see cref="Workbook"/>.</summary>
    public static Workbook Load(string path)
    {
        using var stream = File.OpenRead(path);

        return Load(stream);
    }

    /// <summary>
    /// Loads an .xlsx document from a stream into a new <see cref="Workbook"/>. The stream must be
    /// readable and seekable (a requirement of the underlying package reader).
    /// </summary>
    public static Workbook Load(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var workbookPart =
            document.WorkbookPart
            ?? throw new InvalidDataException("The document does not contain a workbook part.");

        var sharedStrings = SharedStringsStreamReader.Read(workbookPart.SharedStringTablePart);
        var workbook = new Workbook();

        // Iterating workbook.xml's sheet list in document order makes our Sheet.Index match Excel's tab order.
        foreach (var sheetElement in workbookPart.Workbook?.Sheets?.Elements<XlsxSheet>() ?? [])
        {
            if (
                sheetElement.Name?.Value is not { } name
                || sheetElement.Id?.Value is not { } relationshipId
            )
            {
                continue;
            }

            var sheet = workbook.Sheets.Add(name);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

            WorksheetStreamLoader.Load(worksheetPart, sheet, sharedStrings);
        }

        // Defined names are read after the sheets so their (qualified) references resolve to real sheets.
        LoadDefinedNames(workbookPart, workbook);

        return workbook;
    }

    // A sheet named "" so ExpressionParser can parse a defined name's refersTo; workbook-scoped names are
    // fully qualified, so this context is never actually consulted (parsing only reads the sheet's name).
    private static readonly Sheet DefinedNameContext = new() { Name = string.Empty };

    private static void LoadDefinedNames(WorkbookPart workbookPart, Workbook workbook)
    {
        foreach (
            var definedName in workbookPart.Workbook?.DefinedNames?.Elements<DefinedName>() ?? []
        )
        {
            if (definedName.Name?.Value is not { } name)
            {
                continue;
            }

            // Only workbook-scoped user names: skip the sheet-scoped ones (they carry a localSheetId) and
            // Excel's builtin "_xlnm.*" names (Print_Area, Print_Titles, _FilterDatabase, …).
            if (
                definedName.LocalSheetId is not null
                || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
            )
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(definedName.Text))
            {
                continue;
            }

            try
            {
                // The refersTo (e.g. "Data!$A$1:$A$10") is parsed as a formula; a constant name (e.g. "0.1")
                // parses to a literal.
                var expression = ExpressionParser.Parse("=" + definedName.Text, DefinedNameContext);
                workbook.DefineName(name, expression);
            }
            catch (Exception exception) when (exception is ParseException or ArgumentException)
            {
                // A name we cannot parse, or whose name fails validation, is skipped rather than failing
                // the whole load (a documented interop limitation).
            }
        }
    }
}
