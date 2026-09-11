using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxTable = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Reads a worksheet's Excel <b>Table</b> parts (<c>xl/tables/tableN.xml</c>, a.k.a. ListObjects) into
/// <see cref="Workbook.Tables"/> through
/// <see cref="Workbook.DefineTable(string, string, string, IReadOnlyList{string}, bool, bool)"/>, which
/// owns every geometry, width, name and column-name rule — the reader only maps attributes and contains
/// failures. <c>TableDefinitionParts</c> is a package-RELATIONSHIP enumeration, so this reads the tiny
/// <c>&lt;table&gt;</c> part without materializing the worksheet DOM (measured on a 5000x10 table: 62 KB and
/// 0.23 ms here versus 29.8 MB and 340 ms for <c>WorksheetPart.Worksheet</c>), which is why it does not
/// break <see cref="WorksheetStreamLoader"/>'s streaming contract. Every failure becomes one
/// <see cref="ExcelLoadWarningKind.InvalidTableDefinition"/> and the table is skipped; the load never fails.
/// </summary>
internal static class TableDefinitionReader
{
    public static void Read(
        WorksheetPart part,
        Workbook workbook,
        string sheetName,
        ExcelLoadOptions? options
    )
    {
        // Two-stage containment. Each PART can fail independently (measured on OpenXml 3.5.1: garbage XML
        // → XmlException, a wrong root element → InvalidDataException), so one try around the whole loop
        // would drop a healthy third table because the second is corrupt. The RELATIONSHIP list is guarded
        // separately for a part the package lacks; on 3.5.1 that file already fails at
        // `workbookPart.Workbook` before any sheet is read, so this outer catch is a guard for other SDK
        // versions, not a measured path.
        List<TableDefinitionPart> parts;
        try
        {
            parts = [.. part.TableDefinitionParts];
        }
        catch (Exception exception)
            when (exception is InvalidOperationException or OpenXmlPackageException)
        {
            Warn(options, sheetName, exception.Message);
            return;
        }

        foreach (var tablePart in parts)
        {
            // The warning's Subject: the table's name once the part has been read far enough to have one,
            // the sheet's until then.
            var subject = sheetName;

            try
            {
                var table =
                    tablePart.Table
                    ?? throw new InvalidDataException("The <table> part has no root element.");

                // The schema requires displayName and makes name optional; a producer that wrote only
                // the latter still gets its table registered under it.
                var name = NonBlank(table.DisplayName?.Value) ?? NonBlank(table.Name?.Value);
                if (name is null)
                {
                    Warn(
                        options,
                        sheetName,
                        "The <table> part has neither a displayName nor a name."
                    );
                    continue;
                }

                subject = name;

                // Both counts default per ECMA-376 (a normal table omits headerRowCount and a table without
                // totals omits totalsRowCount; totalsRowShown is a UI flag, not geometry). Excel writes 0
                // or 1 of each, so anything else is skipped rather than guessed at.
                var headerRows = table.HeaderRowCount?.Value ?? 1u;
                var totalsRows = table.TotalsRowCount?.Value ?? 0u;
                if (headerRows > 1 || totalsRows > 1)
                {
                    Warn(
                        options,
                        name,
                        $"Table '{name}': headerRowCount={headerRows} and totalsRowCount={totalsRows}; "
                            + "Excel writes 0 or 1 of each."
                    );
                    continue;
                }

                // Column names in document order, exactly as the part carries them (the SDK has already
                // decoded XML entities); a missing name becomes "" so DefineTable reports it as empty.
                var columnNames = (table.TableColumns?.Elements<TableColumn>() ?? [])
                    .Select(column => DecodeColumnName(column.Name?.Value ?? string.Empty))
                    .ToList();

                // First registration wins. Excel forbids two tables sharing a name, so a duplicate means a
                // corrupt file; refusing here (rather than letting DefineTable's redefine-replaces rule
                // pick the LAST part) keeps a name that formulas on an earlier sheet may already resolve
                // through pointing where it did.
                if (workbook.Tables.ContainsKey(name))
                {
                    Warn(
                        options,
                        name,
                        $"Table '{name}': another table already claimed this name; the first one read is "
                            + "kept."
                    );
                    continue;
                }

                // The ref spans the WHOLE table, header and totals rows included — exactly the record's
                // FirstRow..LastRow with HasHeaderRow/HasTotalsRow, so no arithmetic happens here. An
                // absent ref becomes "" and is rejected by DefineTable like any other non-range.
                workbook.DefineTable(
                    name,
                    sheetName,
                    table.Reference?.Value ?? string.Empty,
                    columnNames,
                    hasHeaderRow: headerRows == 1,
                    hasTotalsRow: totalsRows == 1
                );
            }
            catch (Exception exception)
                when (exception
                        is XmlException
                            or InvalidDataException
                            or InvalidOperationException
                            or ArgumentException
                            or FormatException
                )
            {
                // ArgumentException covers every DefineTable / Table.Validate reject (name, geometry,
                // width, empty or duplicated column, a name a defined name holds); its message already
                // names the table and the offending member. The rest are the SDK's on a corrupt part.
                Warn(options, subject, exception.Message);
            }
        }
    }

    private static string? NonBlank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    /// <summary>
    /// Undoes OOXML's <c>_xHHHH_</c> escaping for CONTROL characters only: exactly four hex digits between
    /// <c>_x</c> and <c>_</c>, decoded when the code point is below 0x20 (a producer is FORCED to escape
    /// those — a raw newline in an attribute would be normalized to a space) or is <c>_x005f_</c> (the
    /// escaped underscore that starts an escape). Every other <c>_x…_</c> run stays literal, so a
    /// user-typed <c>_x0020_</c> survives and a producer that escapes nothing is never mis-decoded. XML
    /// entities (<c>&amp;amp;</c>, <c>&amp;#39;</c>, …) are ALREADY decoded by the SDK and must not be
    /// touched again. Excel's formula-side <c>'</c>-escapes (<c>''</c>, <c>'[</c>, <c>']</c>, <c>'#</c>,
    /// <c>'@</c>) live on the LEXER side: both halves meet at this raw text, compared OrdinalIgnoreCase
    /// by <see cref="Table.TryGetColumnIndex"/>.
    /// </summary>
    internal static string DecodeColumnName(string raw)
    {
        if (raw.IndexOf("_x", StringComparison.Ordinal) < 0)
        {
            return raw;
        }

        var builder = new StringBuilder(raw.Length);
        var i = 0;

        while (i < raw.Length)
        {
            if (TryDecodeEscapeAt(raw, i, out var decoded))
            {
                builder.Append(decoded);
                i += EscapeLength;
            }
            else
            {
                builder.Append(raw[i]);
                i++;
            }
        }

        return builder.ToString();
    }

    private const int EscapeLength = 7; // _xHHHH_

    private static bool TryDecodeEscapeAt(string text, int index, out char decoded)
    {
        decoded = '\0';

        if (
            index + EscapeLength > text.Length
            || text[index] != '_'
            || text[index + 1] != 'x'
            || text[index + EscapeLength - 1] != '_'
            || !int.TryParse(
                text.AsSpan(index + 2, 4),
                NumberStyles.AllowHexSpecifier,
                CultureInfo.InvariantCulture,
                out var code
            )
            || (code >= 0x20 && code != '_')
        )
        {
            return false;
        }

        decoded = (char)code;
        return true;
    }

    private static void Warn(ExcelLoadOptions? options, string subject, string detail) =>
        options?.OnWarning?.Invoke(
            new ExcelLoadWarning(ExcelLoadWarningKind.InvalidTableDefinition, subject, detail)
        );
}
