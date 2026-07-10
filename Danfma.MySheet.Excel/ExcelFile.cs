using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Options for <see cref="ExcelFile.Load(string, ExcelLoadOptions?)"/> /
/// <see cref="ExcelFile.Load(Stream, ExcelLoadOptions?)"/>. Everything here is optional and additive: the
/// parameterless overloads behave exactly as before (no options means no callback and no behavior change).
/// </summary>
public sealed class ExcelLoadOptions
{
    /// <summary>
    /// Invoked once for every non-fatal issue found while loading — see <see cref="ExcelLoadWarningKind"/>
    /// for the current set. The load never fails because of one of these; the affected item is simply
    /// skipped or degraded, exactly as it was before this option existed. A plain callback (rather than an
    /// accumulated list) so the host decides whether to log, collect, or ignore — and pays nothing when
    /// left <c>null</c> (the default).
    /// </summary>
    public Action<ExcelLoadWarning>? OnWarning { get; init; }
}

/// <summary>
/// One non-fatal issue surfaced via <see cref="ExcelLoadOptions.OnWarning"/> while loading an .xlsx file.
/// </summary>
/// <param name="Kind">What kind of issue this is.</param>
/// <param name="Subject">What the warning is about: the defined name's own name for
/// <see cref="ExcelLoadWarningKind.InvalidDefinedName"/>, or the cell id (e.g. <c>"B7"</c>) for
/// <see cref="ExcelLoadWarningKind.UnparsableDateLiteral"/>.</param>
/// <param name="Detail">A short human-readable detail: the parse exception's message, or the raw literal
/// text that failed to parse.</param>
public readonly record struct ExcelLoadWarning(
    ExcelLoadWarningKind Kind,
    string Subject,
    string Detail
);

/// <summary>The kind of a non-fatal <see cref="ExcelLoadWarning"/> raised while loading an .xlsx file.</summary>
public enum ExcelLoadWarningKind
{
    /// <summary>
    /// A workbook-scoped <c>&lt;definedName&gt;</c> whose <c>refersTo</c> text failed to parse, or whose
    /// name failed validation. The name is skipped (as it always was); the rest of the workbook loads
    /// normally. <see cref="ExcelLoadWarning.Subject"/> is the defined name.
    /// </summary>
    InvalidDefinedName,

    /// <summary>
    /// A <c>t="d"</c> (ISO-8601 date) cell whose literal text failed to parse as a <see cref="DateTime"/>.
    /// The cell falls back to a <c>StringValue</c> holding the raw text, exactly as it always did.
    /// <see cref="ExcelLoadWarning.Subject"/> is the cell id.
    /// </summary>
    UnparsableDateLiteral,
}

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
    public static Workbook Load(string path) => Load(path, options: null);

    /// <summary>
    /// Loads an .xlsx file into a new <see cref="Workbook"/>, reporting non-fatal issues (an invalid
    /// defined name, an unparsable date literal, …) via <paramref name="options"/>'s
    /// <see cref="ExcelLoadOptions.OnWarning"/> instead of letting them pass silently.
    /// </summary>
    public static Workbook Load(string path, ExcelLoadOptions? options)
    {
        using var stream = File.OpenRead(path);

        return Load(stream, options);
    }

    /// <summary>
    /// Loads an .xlsx document from a stream into a new <see cref="Workbook"/>. The stream must be
    /// readable and seekable (a requirement of the underlying package reader).
    /// </summary>
    public static Workbook Load(Stream stream) => Load(stream, options: null);

    /// <summary>
    /// Loads an .xlsx document from a stream into a new <see cref="Workbook"/>, reporting non-fatal issues
    /// via <paramref name="options"/>'s <see cref="ExcelLoadOptions.OnWarning"/>. The stream must be
    /// readable and seekable (a requirement of the underlying package reader).
    /// </summary>
    public static Workbook Load(Stream stream, ExcelLoadOptions? options)
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

            WorksheetStreamLoader.Load(worksheetPart, sheet, sharedStrings, options);
        }

        // Defined names are read after the sheets so their (qualified) references resolve to real sheets.
        LoadDefinedNames(workbookPart, workbook, options);

        return workbook;
    }

    // A sheet named "" so ExpressionParser can parse a defined name's refersTo; workbook-scoped names are
    // fully qualified, so this context is never actually consulted (parsing only reads the sheet's name).
    private static readonly Sheet DefinedNameContext = new() { Name = string.Empty };

    private static void LoadDefinedNames(
        WorkbookPart workbookPart,
        Workbook workbook,
        ExcelLoadOptions? options
    )
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
                var expression = ExpressionParser.ParseFormulaBody(
                    definedName.Text,
                    DefinedNameContext
                );
                workbook.DefineName(name, expression);
            }
            catch (Exception exception) when (exception is ParseException or ArgumentException)
            {
                // A name we cannot parse, or whose name fails validation, is skipped rather than failing
                // the whole load (a documented interop limitation) — surfaced via OnWarning instead of
                // disappearing silently.
                options?.OnWarning?.Invoke(
                    new ExcelLoadWarning(
                        ExcelLoadWarningKind.InvalidDefinedName,
                        name,
                        exception.Message
                    )
                );
            }
        }
    }
}
