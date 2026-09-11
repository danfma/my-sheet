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
/// <see cref="ExcelLoadWarningKind.InvalidDefinedName"/>, the table's <c>displayName</c> (or the sheet
/// name when the part could not be read that far) for
/// <see cref="ExcelLoadWarningKind.InvalidTableDefinition"/>, or the cell id (e.g. <c>"B7"</c>) for every
/// cell-scoped kind (<see cref="ExcelLoadWarningKind.UnparsableDateLiteral"/>,
/// <see cref="ExcelLoadWarningKind.UnparsableFormula"/>,
/// <see cref="ExcelLoadWarningKind.UnparsableCellLiteral"/>).</param>
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

    /// <summary>
    /// A cell whose formula text failed to parse — a syntax MySheet's parser does not accept. Now that
    /// structured references are supported, the remaining common causes are the structured-reference
    /// shapes still out of scope (the this-row form, stored by Excel as
    /// <c>Tabela1[[#This Row],[Valor]]</c> and typed as <c>Tabela1[@Valor]</c>, a column span
    /// <c>Tabela1[[Q1]:[Q3]]</c>, the implicit-table form <c>[Valor]</c>), array literals
    /// (<c>{1;2;3}</c>), and genuinely malformed or otherwise unsupported formula text. The cell falls back
    /// to the cached value Excel stored alongside the formula (blank when the file carries none), so only
    /// that cell degrades — the rest of the workbook loads normally.
    /// <see cref="ExcelLoadWarning.Subject"/> is the cell id; for a shared-formula group it is the MASTER's
    /// cell, reported once for the group (each slave then falls back to its own cached value).
    /// A structured reference whose TABLE is missing or was skipped is NOT this warning — it parses fine
    /// and evaluates to <c>#NAME?</c>; see <see cref="InvalidTableDefinition"/>.
    /// </summary>
    UnparsableFormula,

    /// <summary>
    /// A cell's literal <c>&lt;v&gt;</c> text that could not be decoded as the type its <c>@t</c> claims: a
    /// numeric cell whose text is not a number, or a <c>t="s"</c> cell whose shared-string index is out of
    /// range. The cell degrades to the raw text (numeric case) or to blank (unresolvable index) rather than
    /// failing the load. <see cref="ExcelLoadWarning.Subject"/> is the cell id and
    /// <see cref="ExcelLoadWarning.Detail"/> the raw text. Distinct from
    /// <see cref="UnparsableDateLiteral"/>, which is specifically the <c>t="d"</c> ISO-8601 case.
    /// </summary>
    UnparsableCellLiteral,

    /// <summary>
    /// An Excel <b>Table</b> (<c>&lt;table&gt;</c> part) that could not be registered in
    /// <see cref="Workbook.Tables"/>: a missing or malformed <c>ref</c>, a <c>headerRowCount</c> or
    /// <c>totalsRowCount</c> other than 0 or 1, a column count that disagrees with the <c>ref</c>'s width,
    /// an empty or duplicated column name, a name MySheet's tokenizer could never read (a backslash,
    /// <c>TRUE</c>/<c>FALSE</c> — Excel itself accepts both), a name another table already claimed, or a
    /// part whose XML cannot be read at all (garbage, a wrong root element). The table is skipped and the
    /// rest of the workbook loads normally — its cells are still ordinary cells — but a structured
    /// reference into it then evaluates to <c>#NAME?</c>, which is exactly what Excel shows for a table
    /// that does not exist.
    /// <see cref="ExcelLoadWarning.Subject"/> is the table's <c>displayName</c>, or the SHEET name when
    /// the part could not be read far enough to have one.
    /// </summary>
    InvalidTableDefinition,
}

/// <summary>
/// Reads Excel (.xlsx) files into a MySheet <see cref="Workbook"/> — cross-platform via the OpenXML SDK,
/// no Excel installation required. Shared strings and worksheets are STREAMED (forward-only XmlReader over
/// each part), so the OpenXML DOM is never materialized and the only full representation in memory is the
/// MySheet model. Formula cells are parsed into real <c>Expression</c> trees (re-evaluated by the MySheet
/// engine); plain cells become literal values. Dates stay as serial numbers, and a shared-formula cell
/// that carries no formula text (a "slave" of a dragged formula) is expanded from its group master.
/// Excel <b>Tables</b> are read from each worksheet's <c>&lt;table&gt;</c> parts into
/// <see cref="Workbook.Tables"/> (name, geometry and column names), which is what makes a structured
/// reference such as <c>Tabela1[Valor]</c> resolve; the parts are reached through package relationships,
/// so this does not materialize any worksheet DOM.
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

            // Tables come from a package RELATIONSHIP (TableDefinitionParts), not from the sheet XML — the
            // streaming loader breaks out at </sheetData> and never sees <tableParts>. Read after the cells
            // so this sheet's warnings stay in document order; registration order does not matter because
            // a structured reference resolves at evaluation time, exactly like a defined name.
            TableDefinitionReader.Read(worksheetPart, workbook, sheet.Name, options);
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
