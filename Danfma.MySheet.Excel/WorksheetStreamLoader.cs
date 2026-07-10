using System.Globalization;
using System.Text;
using System.Xml;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml.Packaging;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Streams one worksheet part's XML forward-only into the sheet — the OpenXML worksheet DOM is never
/// materialized, so the only full representation in memory is the MySheet model itself. Formula cells
/// are parsed into expression trees (the cached <c>&lt;v&gt;</c> is ignored), literals are decoded by
/// their <c>@t</c> type, and shared-formula slaves are expanded from their group master (document
/// order guarantees the master — the first cell of the group's ref — is seen first, ECMA-376
/// §18.3.1.40; an out-of-spec producer that emits a slave earlier is handled by deferring it). Cells
/// without <c>@r</c> (implicit position, allowed by the spec) are placed at the next column of the
/// current row. Matching is by <see cref="XmlReader.LocalName"/> only, so namespace prefixes
/// (<c>x:</c>, strict-mode) don't matter.
/// </summary>
internal static class WorksheetStreamLoader
{
    public static void Load(WorksheetPart part, Sheet sheet, IReadOnlyList<string> sharedStrings)
    {
        // Masters of shared-formula groups: si → (master cell, formula text).
        var sharedFormulas = new Dictionary<uint, (string MasterId, string Formula)>();

        // Out-of-spec slaves seen before their master (lazy; null on the happy path). The cached
        // literal is decoded eagerly because the <v> text is gone once the reader moves on.
        List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? pendingSlaves = null;

        var builder = new StringBuilder();

        using (var source = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(source))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData")
                {
                    if (!reader.IsEmptyElement)
                    {
                        ReadSheetData(
                            reader,
                            sheet,
                            sharedStrings,
                            sharedFormulas,
                            ref pendingSlaves,
                            builder
                        );
                    }

                    break; // nothing after sheetData is load-relevant
                }
            }
        }

        if (pendingSlaves is null)
        {
            return;
        }

        // A deferred slave resolves against its master if one eventually appeared; otherwise it falls
        // back to its cached literal — the same outcome the DOM loader produced for an unknown master.
        foreach (var (id, sharedIndex, cachedLiteral) in pendingSlaves)
        {
            if (sharedFormulas.TryGetValue(sharedIndex, out var master))
            {
                var shifted = SharedFormulaShifter.Shift(master.Formula, master.MasterId, id);

                sheet[id] = ExpressionParser.Parse("=" + shifted, sheet);
            }
            else if (cachedLiteral is not null)
            {
                sheet[id] = cachedLiteral;
            }
        }
    }

    private static void ReadSheetData(
        XmlReader reader,
        Sheet sheet,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, (string MasterId, string Formula)> sharedFormulas,
        ref List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? pendingSlaves,
        StringBuilder builder
    )
    {
        var currentRow = 0;

        reader.Read(); // first child of <sheetData>

        while (
            !reader.EOF
            && !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "sheetData")
        )
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
            {
                // A row without @r is implicitly the one after the previous row.
                currentRow = int.TryParse(
                    reader.GetAttribute("r"),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var rowNumber
                )
                    ? rowNumber
                    : currentRow + 1;

                if (reader.IsEmptyElement)
                {
                    reader.Read();

                    continue;
                }

                ReadRow(
                    reader,
                    sheet,
                    currentRow,
                    sharedStrings,
                    sharedFormulas,
                    ref pendingSlaves,
                    builder
                );

                continue;
            }

            if (reader.NodeType == XmlNodeType.Element)
            {
                reader.Skip();

                continue;
            }

            if (!reader.Read())
            {
                break;
            }
        }
    }

    // Entered ON a non-empty <row>; consumes through </row> and past it.
    private static void ReadRow(
        XmlReader reader,
        Sheet sheet,
        int row,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, (string MasterId, string Formula)> sharedFormulas,
        ref List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? pendingSlaves,
        StringBuilder builder
    )
    {
        var nextColumn = 1;

        reader.Read();

        while (
            !reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
        )
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                nextColumn =
                    ReadCell(
                        reader,
                        sheet,
                        row,
                        nextColumn,
                        sharedStrings,
                        sharedFormulas,
                        ref pendingSlaves,
                        builder
                    ) + 1;

                continue;
            }

            if (reader.NodeType == XmlNodeType.Element)
            {
                reader.Skip();

                continue;
            }

            if (!reader.Read())
            {
                break;
            }
        }

        reader.Read(); // past </row>
    }

    // Entered ON a <c>; consumes through </c> and past it. Returns the cell's 1-based column so the
    // caller can track the implicit position of a following cell that lacks @r.
    private static int ReadCell(
        XmlReader reader,
        Sheet sheet,
        int row,
        int nextColumn,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, (string MasterId, string Formula)> sharedFormulas,
        ref List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? pendingSlaves,
        StringBuilder builder
    )
    {
        var id = reader.GetAttribute("r");
        var type = reader.GetAttribute("t");

        int column;

        if (id is null)
        {
            column = nextColumn;
            id = CellId.Format(row, column);
        }
        else
        {
            column = CellId.Parse(id).Column;
        }

        string? formulaText = null;
        string? raw = null;
        string? inlineText = null;
        var isSharedFormula = false;
        uint sharedIndex = 0;

        if (reader.IsEmptyElement)
        {
            reader.Read();
        }
        else
        {
            reader.Read();

            while (
                !reader.EOF
                && !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
            )
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "f":
                            if (
                                reader.GetAttribute("t") == "shared"
                                && uint.TryParse(
                                    reader.GetAttribute("si"),
                                    NumberStyles.None,
                                    CultureInfo.InvariantCulture,
                                    out var si
                                )
                            )
                            {
                                isSharedFormula = true;
                                sharedIndex = si;
                            }

                            if (reader.IsEmptyElement)
                            {
                                reader.Read();
                            }
                            else
                            {
                                formulaText = reader.ReadElementContentAsString();
                            }

                            continue;
                        case "v":
                            if (reader.IsEmptyElement)
                            {
                                raw = string.Empty;
                                reader.Read();
                            }
                            else
                            {
                                raw = reader.ReadElementContentAsString();
                            }

                            continue;
                        case "is":
                            inlineText = SharedStringsStreamReader.ReadFlattenedText(
                                reader,
                                builder
                            );

                            continue;
                        default:
                            reader.Skip();

                            continue;
                    }
                }

                if (!reader.Read())
                {
                    break;
                }
            }

            reader.Read(); // past </c>
        }

        // A formula cell becomes a real expression tree, re-evaluated by the MySheet engine (the
        // cached <v> is ignored). A shared-formula master also registers its text for the group.
        if (formulaText is { Length: > 0 })
        {
            if (isSharedFormula)
            {
                sharedFormulas[sharedIndex] = (id, formulaText);
            }

            sheet[id] = ExpressionParser.Parse("=" + formulaText, sheet);

            return column;
        }

        // A shared-formula slave carries no text: expand it from its master, shifting relative
        // references by the cell delta — exactly what Excel does.
        if (isSharedFormula)
        {
            if (sharedFormulas.TryGetValue(sharedIndex, out var master))
            {
                var shifted = SharedFormulaShifter.Shift(master.Formula, master.MasterId, id);

                sheet[id] = ExpressionParser.Parse("=" + shifted, sheet);

                return column;
            }

            // Out-of-spec producer: the slave appeared before its master. Defer it; the cached
            // literal is decoded now because the <v> text is unavailable later.
            pendingSlaves ??= [];
            pendingSlaves.Add(
                (id, sharedIndex, DecodeLiteral(type, raw, inlineText, sharedStrings))
            );

            return column;
        }

        if (DecodeLiteral(type, raw, inlineText, sharedStrings) is { } literal)
        {
            sheet[id] = literal;
        }

        return column;
    }

    // Parity port of the DOM loader's LoadLiteral, keyed by the raw @t string.
    private static Expression? DecodeLiteral(
        string? type,
        string? raw,
        string? inlineText,
        IReadOnlyList<string> sharedStrings
    )
    {
        if (type == "inlineStr")
        {
            return new StringValue(inlineText ?? string.Empty);
        }

        if (raw is null)
        {
            // No value and no formula: a style-only/empty cell, which MySheet models as blank.
            return null;
        }

        if (type is null or "n")
        {
            return new NumberValue(
                double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture)
            );
        }

        return type switch
        {
            "s" => new StringValue(sharedStrings[int.Parse(raw, CultureInfo.InvariantCulture)]),
            "b" => new BooleanValue(
                raw is "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase)
            ),
            "e" => new ErrorValue(raw),
            // ISO-8601 dates (strict-mode files) are converted to the same serial-number form Excel
            // uses in transitional files, keeping "dates are serial numbers" consistent across both.
            "d" => DateTime.TryParse(
                raw,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date
            )
                ? new NumberValue(date.ToOADate())
                : new StringValue(raw),
            // "str" (a formula's cached string) and anything unrecognized read as text.
            _ => new StringValue(raw),
        };
    }
}
