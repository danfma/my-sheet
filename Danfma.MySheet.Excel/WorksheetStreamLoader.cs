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
    private sealed class LoadContext(
        Sheet sheet,
        IReadOnlyList<string> sharedStrings,
        ExcelLoadOptions? options
    )
    {
        public Sheet Sheet { get; } = sheet;
        public IReadOnlyList<string> SharedStrings { get; } = sharedStrings;

        /// <summary>Null unless the caller opted in via <see cref="ExcelFile.Load(string, ExcelLoadOptions?)"/>.</summary>
        public ExcelLoadOptions? Options { get; } = options;

        /// <summary>
        /// Masters of shared-formula groups: si → (position, id, text, tokens, anchored tree). The token
        /// list is produced once per master and re-parsed with a per-slave delta on the LEGACY path; the
        /// text survives for the out-of-spec fallback (negative delta) that still goes through the textual
        /// shifter. <c>AnchoredTree</c>/<c>AnchoredSupported</c> (G3 spike, node-delta shared formulas) are
        /// computed once per group, eagerly, when the master is first seen: the anchored-mode parse of the
        /// SAME token list, shared by every slave via <see cref="SharedFormulaSlave"/> when
        /// <c>AnchoredSupported</c> is true — see <see cref="ExpandSlave"/>.
        /// </summary>
        public Dictionary<
            uint,
            (
                int Row,
                int Column,
                string MasterId,
                string Formula,
                List<Token> Tokens,
                Expression AnchoredTree,
                bool AnchoredSupported
            )
        > SharedFormulas { get; } = [];

        /// <summary>
        /// Identical formula text parses to an identical immutable tree, and sharing one instance
        /// across cells is safe because per-cell evaluation state lives in the value store keyed by
        /// (sheet, column, row) — never in the node. Deduplicating saves both the re-parse and the
        /// duplicate tree.
        /// </summary>
        public Dictionary<string, Expression> FormulaCache { get; } = [];

        /// <summary>
        /// String literals deduped by content (Ordinal): shared strings are already unique instances
        /// per index, but repeated cells referencing the same index (or repeated inline strings) still
        /// allocate a fresh <see cref="StringValue"/> wrapper each time without this cache.
        /// </summary>
        public Dictionary<string, StringValue> StringCache { get; } = new(StringComparer.Ordinal);

        /// <summary>
        /// Error literals not covered by <see cref="ErrorValue"/>'s well-known singletons (e.g. the
        /// rare <c>#NULL!</c> or <c>#GETTING_DATA</c>, or a malformed code) — few distinct values per
        /// file, so a small per-load cache is enough.
        /// </summary>
        public Dictionary<string, ErrorValue> ErrorCache { get; } = new(StringComparer.Ordinal);

        /// <summary>
        /// Out-of-spec slaves seen before their master (lazy; null on the happy path). The cached
        /// literal is decoded eagerly because the &lt;v&gt; text is gone once the reader moves on.
        /// </summary>
        public List<(string Id, uint SharedIndex, Expression? CachedLiteral)>? PendingSlaves;

        public StringBuilder Builder { get; } = new();
    }

    public static void Load(
        WorksheetPart part,
        Sheet sheet,
        IReadOnlyList<string> sharedStrings,
        ExcelLoadOptions? options = null
    )
    {
        var context = new LoadContext(sheet, sharedStrings, options);

        using (var source = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(source))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "dimension")
                {
                    // The dimension ref is a bounding box, not a cell count: presize the dense store
                    // toward it, capped so a sparse sheet with a huge bbox cannot balloon the
                    // reservation (the cap bounds waste to a few MB, reclaimed by the next resize).
                    ApplyDimensionHint(sheet, reader.GetAttribute("ref"));

                    continue;
                }

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

    private const int MaxPresizedCells = 512 * 1024;

    private static void ApplyDimensionHint(Sheet sheet, string? dimensionRef)
    {
        if (dimensionRef is null)
        {
            return;
        }

        var separator = dimensionRef.IndexOf(':');

        if (separator < 0)
        {
            return; // single-cell dimension: nothing worth reserving
        }

        try
        {
            var (fromRow, fromColumn) = CellId.Parse(dimensionRef[..separator]);
            var (toRow, toColumn) = CellId.Parse(dimensionRef[(separator + 1)..]);
            var cells = (long)(toRow - fromRow + 1) * (toColumn - fromColumn + 1);

            if (cells > 0)
            {
                sheet.EnsureCellCapacity((int)long.Min(cells, MaxPresizedCells));
            }
        }
        catch (FormatException)
        {
            // A malformed dimension is a hint we simply don't take.
        }
    }

    // Expands one shared-formula slave from its master. Spec-compliant files always yield
    // non-negative deltas (the master is the FIRST cell of the group's ref), which the token-delta
    // parse handles without any per-slave re-tokenization; a negative delta (out-of-spec si reuse)
    // falls back to the original text shifter, preserving its exact degenerate behavior.
    //
    // G3 spike (node-delta shared formulas): on the spec-compliant (non-negative delta) path, when the
    // group's master parsed fully into anchored nodes (AnchoredSupported), every slave becomes a small
    // SharedFormulaSlave wrapper around the ONE shared AnchoredTree instead of its own independently-parsed,
    // independently-allocated expression tree — the load-time allocation the spike targets. A group whose
    // master contains a shape the anchored Parser mode cannot represent exactly (an open range, a union, a
    // reference-returning endpoint — see AnchoredFormulaSupport) keeps the pre-spike token-delta re-parse
    // per slave, unchanged.
    private static Expression ExpandSlave(
        Sheet sheet,
        (
            int Row,
            int Column,
            string MasterId,
            string Formula,
            List<Token> Tokens,
            Expression AnchoredTree,
            bool AnchoredSupported
        ) master,
        string id,
        int row,
        int column
    )
    {
        var deltaRow = row - master.Row;
        var deltaColumn = column - master.Column;

        if (deltaRow >= 0 && deltaColumn >= 0)
        {
            if (master.AnchoredSupported)
            {
                return new SharedFormulaSlave(master.AnchoredTree, deltaRow, deltaColumn);
            }

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

    // G3 spike (node-delta shared formulas): attempts the anchored-mode parse of a shared-formula group's
    // master ONCE, and validates the result is fully representable by the anchored/delta evaluation model
    // (AnchoredFormulaSupport.IsFullyAnchored). ANY failure — a ParseException from a token shape the
    // anchored Parser mode does not handle gracefully (e.g. a chained cross-sheet range endpoint), or a
    // shape the support check rejects (an open range, a union, a reference-returning endpoint) — is treated
    // as "this group is not anchored-safe": the caller keeps using the legacy per-slave token-delta parse for
    // every slave in the group, exactly as before this spike. Honest fallback, not a guess.
    private static (Expression Tree, bool Supported) TryBuildAnchoredMaster(
        Sheet sheet,
        List<Token> tokens
    )
    {
        try
        {
            var tree = ExpressionParser.ParseAnchoredMasterBody(tokens, sheet);

            return AnchoredFormulaSupport.IsFullyAnchored(tree)
                ? (tree, true)
                : (BlankValue.Instance, false);
        }
        catch (ParseException)
        {
            return (BlankValue.Instance, false);
        }
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
                var tokens = ExpressionParser.TokenizeFormulaBody(formulaText);
                var (anchoredTree, anchoredSupported) = TryBuildAnchoredMaster(
                    context.Sheet,
                    tokens
                );

                context.SharedFormulas[sharedIndex] = (
                    row,
                    column,
                    id,
                    formulaText,
                    tokens,
                    anchoredTree,
                    anchoredSupported
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
                (id, sharedIndex, DecodeLiteral(context, id, type, raw, inlineText))
            );

            return column;
        }

        if (DecodeLiteral(context, id, type, raw, inlineText) is { } literal)
        {
            context.Sheet[id] = literal;
        }

        return column;
    }

    // Parity port of the DOM loader's LoadLiteral, keyed by the raw @t string. `id` is only needed for the
    // "d" case's warning (the cell reference is the useful subject there); every other branch ignores it.
    private static Expression? DecodeLiteral(
        LoadContext context,
        string id,
        string? type,
        string? raw,
        string? inlineText
    )
    {
        if (type == "inlineStr")
        {
            return GetOrAddString(context, inlineText ?? string.Empty);
        }

        if (raw is null)
        {
            // No value and no formula: a style-only/empty cell, which MySheet models as blank.
            return null;
        }

        if (type is null or "n")
        {
            // Not deduped: measured. A Dictionary<double, NumberValue> here was tried and reverted —
            // see the long comment on GetOrAddString for the numbers (BDN on the K1 fixture: number
            // literals are only ~52% duplicate, below the ~54-70% breakeven a dictionary needs to pay
            // for itself against a 24-byte NumberValue, and it made Allocated/Gen1/Gen2 WORSE, not
            // better).
            return new NumberValue(
                double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture)
            );
        }

        return type switch
        {
            "s" => GetOrAddString(
                context,
                context.SharedStrings[int.Parse(raw, CultureInfo.InvariantCulture)]
            ),
            "b" => raw is "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase)
                ? BooleanValue.True
                : BooleanValue.False,
            "e" => GetOrAddError(context, raw),
            // ISO-8601 dates (strict-mode files) are converted to the same serial-number form Excel
            // uses in transitional files, keeping "dates are serial numbers" consistent across both.
            "d" => DecodeDateLiteral(context, id, raw),
            // "str" (a formula's cached string) and anything unrecognized read as text.
            _ => GetOrAddString(context, raw),
        };
    }

    // Split out of the switch expression above only because the failure path has a side effect (the
    // warning callback); the fallback behavior (StringValue of the raw text) is unchanged from before
    // ExcelLoadOptions existed.
    private static Expression DecodeDateLiteral(LoadContext context, string id, string raw)
    {
        if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return new NumberValue(date.ToOADate());
        }

        context.Options?.OnWarning?.Invoke(
            new ExcelLoadWarning(ExcelLoadWarningKind.UnparsableDateLiteral, id, raw)
        );

        return GetOrAddString(context, raw);
    }

    // Caches are per-load (LoadContext dies with the worksheet load), so there is no cap and no
    // cross-worksheet state: duplicate literals within one sheet are exactly the win, bounded by the
    // sheet's own cell count. Only strings and error codes are cached this way — see the comment above
    // the "n"/"d" NumberValue construction sites for why a same-shaped number cache was tried and
    // reverted.
    private static StringValue GetOrAddString(LoadContext context, string text)
    {
        if (!context.StringCache.TryGetValue(text, out var value))
        {
            value = new StringValue(text);
            context.StringCache[text] = value;
        }

        return value;
    }

    // The common Excel error codes already have singletons on ErrorValue; only the long tail (e.g.
    // #NULL!, #GETTING_DATA, or a malformed code) needs the per-load cache.
    private static ErrorValue GetOrAddError(LoadContext context, string code)
    {
        switch (code)
        {
            case "#DIV/0!":
                return ErrorValue.DivByZero;
            case "#VALUE!":
                return ErrorValue.NotValue;
            case "#REF!":
                return ErrorValue.Reference;
            case "#NAME?":
                return ErrorValue.Name;
            case "#NUM!":
                return ErrorValue.Number;
            case "#N/A":
                return ErrorValue.NotAvailable;
        }

        if (!context.ErrorCache.TryGetValue(code, out var value))
        {
            value = new ErrorValue(code);
            context.ErrorCache[code] = value;
        }

        return value;
    }
}
