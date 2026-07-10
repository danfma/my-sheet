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
    public static void SaveAsExcel(
        this Workbook workbook,
        string path,
        ExcelExportOptions? options = null
    )
    {
        using var stream = File.Create(path);

        workbook.SaveAsExcel(stream, options);
    }

    /// <summary>Saves the workbook as an .xlsx document into a writable stream.</summary>
    public static void SaveAsExcel(
        this Workbook workbook,
        Stream stream,
        ExcelExportOptions? options = null
    )
    {
        var mode = (options ?? new ExcelExportOptions()).FormulaMode;
        var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

        // The whole write runs on one large-stack thread: each cell value is computed ON DEMAND at write
        // time (deep dependency chains cannot overflow) and streamed straight out via OpenXmlWriter, so we
        // never hold a second dictionary of every value NOR a full worksheet DOM — only the workbook's own
        // memoized store plus the (distinct-string-bounded) shared-string table stay resident.
        Workbook.RunWithLargeStack(() =>
        {
            using var document = SpreadsheetDocument.Create(
                stream,
                SpreadsheetDocumentType.Workbook
            );

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
                sheetList.AppendChild(
                    new XlsxSheet
                    {
                        Id = id,
                        SheetId = sheetId++,
                        Name = name,
                    }
                );
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

    // Streams one worksheet part straight to the package via OpenXmlWriter, writing each <c>/<f>/<v>
    // through reusable template elements + struct attributes so there is NO per-cell heap allocation and
    // never a materialized <sheetData> tree. Values are computed on demand here (we are already on the
    // large-stack thread from SaveAsExcel).
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

        // Reused across the whole sheet: the tag templates carry no attributes of their own (we pass every
        // attribute explicitly) and OpenXmlAttribute is a struct, so the only heap object is this list's
        // backing array, reused via Clear().
        var attributes = new List<OpenXmlAttribute>(2);
        var rowTag = new XlsxRow();
        var cellTag = new Cell();
        var formulaTag = new CellFormula();
        var valueTag = new CellValue();

        // OpenXML requires rows in order and cells in column order within the row. Sorting packed
        // (row << 14 | column) keys against a parallel payload array replaces the LINQ pipeline —
        // no iterator or per-cell tuple boxing on a 600k-cell sheet.
        var sortKeys = new long[sheet.Count];
        var cells = new (string Id, int Row, int Column, Expression Expression)[sheet.Count];
        var next = 0;

        foreach (var entry in sheet)
        {
            var (row, column) = CellId.Parse(entry.Key);

            sortKeys[next] = ((long)row << 14) | (uint)column;
            cells[next] = (entry.Key, row, column, entry.Value);
            next++;
        }

        Array.Sort(sortKeys, cells);

        var currentRow = -1;

        foreach (var (id, row, column, expression) in cells)
        {
            var position = (Row: row, Column: column);
            var value = workbook.GetCellValue(sheet.Name, id);

            // Resolve what this cell serializes to. A formula node (Formulas mode) always writes its <f>,
            // even when the cached value is blank; a literal blank is omitted entirely.
            string? formula;
            CellContent content;

            if (mode == FormulaMode.Formulas && expression is not ValueExpression)
            {
                formula = expression.ToFormula(sheet.Name);
                content = CachedContent(value);
            }
            else
            {
                formula = null;

                if (value.Kind == ComputedValueKind.Blank)
                {
                    continue; // literal blank: the cell is simply omitted
                }

                content = LiteralContent(value, sharedStrings);
            }

            if (position.Row != currentRow)
            {
                if (currentRow != -1)
                {
                    writer.WriteEndElement(); // close the previous <row>
                }

                attributes.Clear();
                attributes.Add(
                    new OpenXmlAttribute(
                        "r",
                        "",
                        ((uint)position.Row).ToString(CultureInfo.InvariantCulture)
                    )
                );
                writer.WriteStartElement(rowTag, attributes);
                currentRow = position.Row;
            }

            attributes.Clear();
            attributes.Add(new OpenXmlAttribute("r", "", id));

            if (content.DataType is { } dataType)
            {
                attributes.Add(new OpenXmlAttribute("t", "", dataType));
            }

            writer.WriteStartElement(cellTag, attributes);

            if (formula is not null)
            {
                writer.WriteStartElement(formulaTag);
                writer.WriteString(formula);
                writer.WriteEndElement();
            }

            if (content.Value is { } text)
            {
                writer.WriteStartElement(valueTag);
                writer.WriteString(text);
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // </c>
        }

        if (currentRow != -1)
        {
            writer.WriteEndElement(); // close the last <row>
        }

        writer.WriteEndElement(); // </sheetData>
        writer.WriteEndElement(); // </worksheet>
        writer.Close();
    }

    // The (t attribute, <v> text) a cell serializes to; a null DataType means no t attribute and a null
    // Value means no <v> child (a formula whose cached result is blank).
    private readonly record struct CellContent(string? DataType, string? Value);

    // A literal cell value. Text goes through the shared-string table (t="s", <v> is the index).
    private static CellContent LiteralContent(
        in ComputedValue value,
        SharedStringRegistry sharedStrings
    ) =>
        value.Kind switch
        {
            ComputedValueKind.Number => new(null, XlsxNumbers.Format(value.ToDouble())),
            ComputedValueKind.Boolean => new("b", value.ToBoolean() ? "1" : "0"),
            ComputedValueKind.Text => new(
                "s",
                XlsxNumbers.Format(sharedStrings.IndexOf(value.ToText()))
            ),
            ComputedValueKind.Error => new("e", ErrorText(value)),
            // A bare reference result (e.g. a multi-cell OFFSET) has no single cell value.
            ComputedValueKind.Reference => new("e", Error.Value.ToString()),
            _ => new(null, null),
        };

    // The cached value that accompanies a formula: plain number, t="b", t="str" (NOT a shared string — that
    // is the xlsx convention for formula-produced text), t="e", or blank (no <v>).
    private static CellContent CachedContent(in ComputedValue value) =>
        value.Kind switch
        {
            ComputedValueKind.Number => new(null, XlsxNumbers.Format(value.ToDouble())),
            ComputedValueKind.Boolean => new("b", value.ToBoolean() ? "1" : "0"),
            ComputedValueKind.Text => new("str", value.ToText()),
            ComputedValueKind.Error => new("e", ErrorText(value)),
            ComputedValueKind.Reference => new("e", Error.Value.ToString()),
            _ => new(null, null),
        };

    private static string ErrorText(in ComputedValue value)
    {
        value.TryGetError(out var error);
        return error.ToString();
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
