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
    // Per-worksheet load state, so the reader helpers stay small.
    private sealed class LoadContext(Sheet sheet, IReadOnlyList<string> sharedStrings)
    {
        public Sheet Sheet { get; } = sheet;
        public IReadOnlyList<string> SharedStrings { get; } = sharedStrings;

        /// <summary>
        /// Masters of shared-formula groups: si → (position, id, text, tokens). The token list is
        /// produced once per master and re-parsed with a per-slave delta; the text survives for the
        /// out-of-spec fallback (negative delta) that still goes through the textual shifter.
        /// </summary>
        public Dictionary<
            uint,
            (int Row, int Column, string MasterId, string Formula, List<Token> Tokens)
        > SharedFormulas { get; } = [];

        /// <summary>
        /// Identical formula text parses to an identical immutable tree, and sharing one instance
        /// across cells is safe because per-cell evaluation state lives in the value store keyed by
        /// (sheet, column, row) — never in the node. Deduplicating saves both the re-parse and the
        /// duplicate tree.
        /// </summary>
        public Dictionary<string, Expression> FormulaCache { get; } = [];

        /// <summary>
        /// Out-of-spec slaves seen before their master (lazy; null on the happy path). The cached
        /// literal is decoded eagerly because the &lt;v&gt; text is gone once the reader moves on.
        /// </summary>
        public List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? PendingSlaves;

        public StringBuilder Builder { get; } = new();
    }

    public static void Load(WorksheetPart part, Sheet sheet, IReadOnlyList<string> sharedStrings)
    {
        var context = new LoadContext(sheet, sharedStrings);

        using (var source = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(source))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData")
                {
                    if (!reader.IsEmptyElement)
                    {
                        ReadSheetData(reader, context);
                    }

                    break; // nothing after sheetData is load-relevant
                }
            }
        }

        if (context.PendingSlaves is null)
        {
            return;
        }

        // A deferred slave resolves against its master if one eventually appeared; otherwise it falls
        // back to its cached literal — the same outcome the DOM loader produced for an unknown master.
        foreach (var (id, sharedIndex, cachedLiteral) in context.PendingSlaves)
        {
            if (context.SharedFormulas.TryGetValue(sharedIndex, out var master))
            {
                var (slaveRow, slaveColumn) = CellId.Parse(id);

                sheet[id] = ExpandSlave(sheet, master, id, slaveRow, slaveColumn);
            }
            else if (cachedLiteral is not null)
            {
                sheet[id] = cachedLiteral;
            }
        }
    }

    // Expands one shared-formula slave from its master. Spec-compliant files always yield
    // non-negative deltas (the master is the FIRST cell of the group's ref), which the token-delta
    // parse handles without any per-slave re-tokenization; a negative delta (out-of-spec si reuse)
    // falls back to the original text shifter, preserving its exact degenerate behavior.
    private static Expression ExpandSlave(
        Sheet sheet,
        (int Row, int Column, string MasterId, string Formula, List<Token> Tokens) master,
        string id,
        int row,
        int column
    )
    {
        var deltaRow = row - master.Row;
        var deltaColumn = column - master.Column;

        if (deltaRow >= 0 && deltaColumn >= 0)
        {
            return ExpressionParser.ParseSharedFormulaBody(
                master.Tokens,
                sheet,
                deltaRow,
                deltaColumn
            );
        }

        var shifted = SharedFormulaShifter.Shift(master.Formula, master.MasterId, id);

        return ExpressionParser.ParseFormulaBody(shifted, sheet);
    }

    private static void ReadSheetData(XmlReader reader, LoadContext context)
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

                ReadRow(reader, context, currentRow);

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
    private static void ReadRow(XmlReader reader, LoadContext context, int row)
    {
        var nextColumn = 1;

        reader.Read();

        while (
            !reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
        )
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                nextColumn = ReadCell(reader, context, row, nextColumn) + 1;

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
    private static int ReadCell(XmlReader reader, LoadContext context, int row, int nextColumn)
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
                                context.Builder
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
                context.SharedFormulas[sharedIndex] = (
                    row,
                    column,
                    id,
                    formulaText,
                    ExpressionParser.TokenizeFormulaBody(formulaText)
                );
            }

            if (!context.FormulaCache.TryGetValue(formulaText, out var expression))
            {
                expression = ExpressionParser.ParseFormulaBody(formulaText, context.Sheet);
                context.FormulaCache[formulaText] = expression;
            }

            context.Sheet[id] = expression;

            return column;
        }

        // A shared-formula slave carries no text: expand it from its master, shifting relative
        // references by the cell delta — exactly what Excel does. The shifted text is unique per
        // cell, so it bypasses the formula cache.
        if (isSharedFormula)
        {
            if (context.SharedFormulas.TryGetValue(sharedIndex, out var master))
            {
                context.Sheet[id] = ExpandSlave(context.Sheet, master, id, row, column);

                return column;
            }

            // Out-of-spec producer: the slave appeared before its master. Defer it; the cached
            // literal is decoded now because the <v> text is unavailable later.
            context.PendingSlaves ??= [];
            context.PendingSlaves.Add(
                (id, sharedIndex, DecodeLiteral(context, type, raw, inlineText))
            );

            return column;
        }

        if (DecodeLiteral(context, type, raw, inlineText) is { } literal)
        {
            context.Sheet[id] = literal;
        }

        return column;
    }

    // Parity port of the DOM loader's LoadLiteral, keyed by the raw @t string.
    private static Expression? DecodeLiteral(
        LoadContext context,
        string? type,
        string? raw,
        string? inlineText
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
            "s" => new StringValue(
                context.SharedStrings[int.Parse(raw, CultureInfo.InvariantCulture)]
            ),
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
