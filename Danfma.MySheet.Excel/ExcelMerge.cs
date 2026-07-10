using System.Globalization;
using System.Xml;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Merges computed values from a MySheet <see cref="Workbook"/> into an EXISTING .xlsx file, in place:
/// every cell we hold is written as its computed literal value — dropping any formula the target cell had —
/// while everything else in the file (styles, other cells, other sheets) is left intact. Sheets are matched
/// by name (case-insensitive); sheets missing from the target are skipped, blank values are not written. Text
/// is written as a shared string — appended to the target's shared-string table, reusing existing plain-text
/// entries — so a label repeated across many cells costs a single index, not a full inline copy each time.
///
/// Each worksheet is rewritten by STREAMING its XML through an XmlReader→XmlWriter merge-join (existing cells
/// in document order vs our cells in ascending order) instead of materializing the whole worksheet DOM — so
/// the peak memory stays close to the workbook itself even on very large sheets.
///
/// To produce a new report from a pristine template, copy the template first
/// (<c>File.Copy(template, output)</c>) and merge into the copy.
/// </summary>
public static class ExcelMerge
{
    private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Shared empty set so the no-argument overload delegates without allocating per call.
    private static readonly IReadOnlySet<string> NoIgnoredSheets = new HashSet<string>();

    /// <summary>Merges in place, editing <paramref name="path"/> directly.</summary>
    public static void MergeIntoExcel(this Workbook workbook, string path) =>
        workbook.MergeIntoExcel(path, NoIgnoredSheets);

    /// <summary>
    /// Merges in place, skipping any sheet whose name is in <paramref name="ignoredSheets"/>
    /// (case-insensitive). A skipped sheet is neither evaluated nor written, so the target's copy of
    /// that sheet is left exactly as it was.
    /// </summary>
    public static void MergeIntoExcel(
        this Workbook workbook,
        string path,
        IReadOnlySet<string> ignoredSheets
    )
    {
        // Normalize to a case-insensitive lookup regardless of the caller's set comparer, matching the
        // case-insensitive sheet-name matching used elsewhere in this file. Bounded by sheet count.
        var ignored =
            ignoredSheets.Count == 0
                ? null
                : new HashSet<string>(ignoredSheets, StringComparer.OrdinalIgnoreCase);

        var orderedSheets = workbook.Sheets.Values.OrderBy(sheet => sheet.Index).ToArray();

        // The whole merge runs on one large-stack thread: deep formula chains evaluate safely (see
        // ExcelExport for the rationale) AND each cell is computed ON DEMAND at write time — no up-front
        // dictionary duplicating every value (the workbook already memoizes each value in its own store).
        Workbook.RunWithLargeStack(() =>
        {
            using var document = SpreadsheetDocument.Open(path, isEditable: true);

            var workbookPart =
                document.WorkbookPart
                ?? throw new InvalidDataException("The document does not contain a workbook part.");

            // Merging overrides/drops formula cells, so any calcChain the target carried is now stale.
            // Left in place, Excel reports "Removed records: Formula from /xl/calcChain.xml" and forces a
            // repair on open. Dropping the part lets Excel rebuild the calc chain cleanly and silently.
            if (workbookPart.CalculationChainPart is { } calcChain)
            {
                workbookPart.DeletePart(calcChain);
            }

            // One shared-string table for the whole merge: text from every sheet dedups into it, and it is
            // finalized once after all sheets are streamed.
            var sharedStrings = new SharedStrings(workbookPart);

            foreach (var sheet in orderedSheets)
            {
                if (ignored is not null && ignored.Contains(sheet.Name))
                {
                    continue; // caller asked to skip this sheet: not evaluated, not written
                }

                if (FindWorksheet(workbookPart, sheet.Name) is not { } worksheetPart)
                {
                    continue; // the target has no sheet with this name: skipped by design
                }

                StreamMergeWorksheet(worksheetPart, sheet, workbook, sharedStrings);
            }

            sharedStrings.Finish();

            return 0; // RunWithLargeStack exposes only a Func<T> overload; the result is unused
        });
    }

    private static WorksheetPart? FindWorksheet(WorkbookPart workbookPart, string sheetName)
    {
        var sheetElement = workbookPart
            .Workbook?.Sheets?.Elements<XlsxSheet>()
            .FirstOrDefault(sheet =>
                string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)
            );

        return sheetElement?.Id?.Value is { } relationshipId
            ? (WorksheetPart)workbookPart.GetPartById(relationshipId)
            : null;
    }

    // Rewrites the worksheet part by streaming its XML and merging our cells in. Our cells are grouped by
    // row and sorted by column so the merge-join against the existing (already-ordered) rows/cells is linear.
    private static void StreamMergeWorksheet(
        WorksheetPart part,
        Sheet sheet,
        Workbook workbook,
        SharedStrings sharedStrings
    )
    {
        var rows = new SortedDictionary<int, List<(int Column, string Id)>>();

        foreach (var id in sheet.Keys)
        {
            var position = CellId.Parse(id);

            if (!rows.TryGetValue(position.Row, out var cells))
            {
                cells = [];
                rows[position.Row] = cells;
            }

            cells.Add((position.Column, id));
        }

        foreach (var cells in rows.Values)
        {
            cells.Sort((left, right) => left.Column.CompareTo(right.Column));
        }

        using var buffer = new MemoryStream();

        using (var source = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = XmlReader.Create(source))
        using (var writer = XmlWriter.Create(buffer, new XmlWriterSettings { CloseOutput = false }))
        {
            reader.MoveToContent(); // <worksheet>
            MergeWorksheet(reader, writer, workbook, sheet.Name, rows, sharedStrings);
        }

        buffer.Position = 0;
        part.FeedData(buffer);
    }

    // Copies the <worksheet> element and every child verbatim, EXCEPT <sheetData>, which is merged. Anything
    // that is not a cell we own (other cells, cols, mergeCells, validations, …) passes through unchanged.
    private static void MergeWorksheet(
        XmlReader reader,
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        SortedDictionary<int, List<(int Column, string Id)>> rows,
        SharedStrings sharedStrings
    )
    {
        writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        writer.WriteAttributes(reader, defattr: false);

        var sheetDataWritten = false;

        if (reader.IsEmptyElement)
        {
            // A worksheet with no children at all: emit our rows as a fresh <sheetData> if we have any.
            WriteSheetData(writer, workbook, sheetName, rows, sharedStrings);
            writer.WriteEndElement();
            return;
        }

        reader.Read();

        while (!(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "worksheet"))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData")
            {
                MergeSheetData(reader, writer, workbook, sheetName, rows, sharedStrings);
                sheetDataWritten = true;
            }
            else if (reader.NodeType == XmlNodeType.Element)
            {
                writer.WriteNode(reader, defattr: false); // verbatim subtree; advances the reader itself
            }
            else if (!reader.Read())
            {
                break;
            }
        }

        // The target had no <sheetData> but we have values to write: inject one before </worksheet>. (Rare —
        // any authored sheet with cells already carries a sheetData; this keeps parity with the old code path
        // that created it on demand.)
        if (!sheetDataWritten)
        {
            WriteSheetData(writer, workbook, sheetName, rows, sharedStrings);
        }

        writer.WriteEndElement(); // </worksheet>
    }

    // Merge-joins the existing rows (streamed, ascending) with our rows (ascending) inside <sheetData>.
    private static void MergeSheetData(
        XmlReader reader,
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        SortedDictionary<int, List<(int Column, string Id)>> rows,
        SharedStrings sharedStrings
    )
    {
        writer.WriteStartElement(reader.Prefix, "sheetData", reader.NamespaceURI);
        writer.WriteAttributes(reader, defattr: false);

        using var ours = rows.GetEnumerator();
        var hasOurs = ours.MoveNext();

        if (reader.IsEmptyElement)
        {
            while (hasOurs)
            {
                WriteNewRow(
                    writer,
                    workbook,
                    sheetName,
                    ours.Current.Key,
                    ours.Current.Value,
                    sharedStrings
                );
                hasOurs = ours.MoveNext();
            }

            writer.WriteEndElement();
            reader.Read(); // move past the empty <sheetData/>
            return;
        }

        reader.Read();

        while (!(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "sheetData"))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
            {
                var rowNumber = int.Parse(reader.GetAttribute("r")!, CultureInfo.InvariantCulture);

                // Our rows that sort before this existing one are brand-new: emit them first.
                while (hasOurs && ours.Current.Key < rowNumber)
                {
                    WriteNewRow(
                        writer,
                        workbook,
                        sheetName,
                        ours.Current.Key,
                        ours.Current.Value,
                        sharedStrings
                    );
                    hasOurs = ours.MoveNext();
                }

                if (hasOurs && ours.Current.Key == rowNumber)
                {
                    MergeRow(
                        reader,
                        writer,
                        workbook,
                        sheetName,
                        ours.Current.Value,
                        sharedStrings
                    );
                    hasOurs = ours.MoveNext();
                }
                else
                {
                    writer.WriteNode(reader, defattr: false); // no cells here are ours: copy the row verbatim
                }
            }
            else if (reader.NodeType == XmlNodeType.Element)
            {
                writer.WriteNode(reader, defattr: false);
            }
            else if (!reader.Read())
            {
                break;
            }
        }

        // Our rows that sort after every existing row.
        while (hasOurs)
        {
            WriteNewRow(
                writer,
                workbook,
                sheetName,
                ours.Current.Key,
                ours.Current.Value,
                sharedStrings
            );
            hasOurs = ours.MoveNext();
        }

        writer.WriteEndElement(); // </sheetData>
        reader.Read(); // move past </sheetData>
    }

    // Merge-joins one existing <row>'s cells (streamed, ascending) with our cells for that row (ascending),
    // preserving the row's own attributes and every cell we do not own.
    private static void MergeRow(
        XmlReader reader,
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        List<(int Column, string Id)> ourCells,
        SharedStrings sharedStrings
    )
    {
        writer.WriteStartElement(reader.Prefix, "row", reader.NamespaceURI);
        writer.WriteAttributes(reader, defattr: false); // r, spans, s, customFormat, ht, … all preserved

        var index = 0;

        if (reader.IsEmptyElement)
        {
            WriteRemainingCells(writer, workbook, sheetName, ourCells, ref index, sharedStrings);
            writer.WriteEndElement();
            reader.Read(); // move past the empty <row/>
            return;
        }

        reader.Read();

        while (!(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row"))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                var column = CellId.Parse(reader.GetAttribute("r")!).Column;

                // Our cells that sort before this existing one are brand-new.
                while (index < ourCells.Count && ourCells[index].Column < column)
                {
                    WriteOurCell(
                        writer,
                        workbook,
                        sheetName,
                        ourCells[index],
                        style: null,
                        sharedStrings
                    );
                    index++;
                }

                if (index < ourCells.Count && ourCells[index].Column == column)
                {
                    var value = workbook.GetCellValue(sheetName, ourCells[index].Id);

                    if (value.Kind == ComputedValueKind.Blank)
                    {
                        writer.WriteNode(reader, defattr: false); // our blank does not overwrite: keep target
                    }
                    else
                    {
                        // Preserve the target cell's style (formatting lives in the style index), drop its
                        // formula/value by skipping the old subtree and writing ours in its place.
                        var style = reader.GetAttribute("s");
                        WriteCell(writer, ourCells[index].Id, style, value, sharedStrings);
                        reader.Skip();
                    }

                    index++;
                }
                else
                {
                    writer.WriteNode(reader, defattr: false); // not ours: copy the cell verbatim
                }
            }
            else if (reader.NodeType == XmlNodeType.Element)
            {
                writer.WriteNode(reader, defattr: false);
            }
            else if (!reader.Read())
            {
                break;
            }
        }

        WriteRemainingCells(writer, workbook, sheetName, ourCells, ref index, sharedStrings);

        writer.WriteEndElement(); // </row>
        reader.Read(); // move past </row>
    }

    // Writes a brand-new <row> for cells that had no counterpart in the target. Skips the row entirely when
    // every value is blank, matching the "blanks are not written" rule (and never emitting an empty row).
    private static void WriteNewRow(
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        int rowNumber,
        List<(int Column, string Id)> cells,
        SharedStrings sharedStrings
    )
    {
        var nonBlank = new List<(string Id, ComputedValue Value)>(cells.Count);

        foreach (var (_, id) in cells)
        {
            var value = workbook.GetCellValue(sheetName, id);

            if (value.Kind != ComputedValueKind.Blank)
            {
                nonBlank.Add((id, value));
            }
        }

        if (nonBlank.Count == 0)
        {
            return;
        }

        writer.WriteStartElement("row", Ns);
        writer.WriteStartAttribute("r");
        XlsxNumbers.Write(writer, rowNumber);
        writer.WriteEndAttribute();

        foreach (var (id, value) in nonBlank)
        {
            WriteCell(writer, id, style: null, value, sharedStrings);
        }

        writer.WriteEndElement();
    }

    private static void WriteRemainingCells(
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        List<(int Column, string Id)> cells,
        ref int index,
        SharedStrings sharedStrings
    )
    {
        while (index < cells.Count)
        {
            WriteOurCell(writer, workbook, sheetName, cells[index], style: null, sharedStrings);
            index++;
        }
    }

    private static void WriteOurCell(
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        (int Column, string Id) cell,
        string? style,
        SharedStrings sharedStrings
    )
    {
        var value = workbook.GetCellValue(sheetName, cell.Id);

        if (value.Kind != ComputedValueKind.Blank)
        {
            WriteCell(writer, cell.Id, style, value, sharedStrings);
        }
    }

    // Writes a <c> for one computed value: numbers bare, booleans t="b", text as a shared string (t="s", the
    // <v> being the index into the target's shared-string table), errors t="e". A preserved style index is
    // re-emitted so a formatted template cell stays formatted.
    private static void WriteCell(
        XmlWriter writer,
        string reference,
        string? style,
        in ComputedValue value,
        SharedStrings sharedStrings
    )
    {
        writer.WriteStartElement("c", Ns);
        writer.WriteAttributeString("r", reference);

        if (style is not null)
        {
            writer.WriteAttributeString("s", style);
        }

        switch (value.Kind)
        {
            case ComputedValueKind.Number:
                writer.WriteStartElement("v", Ns);
                XlsxNumbers.Write(writer, value.ToDouble());
                writer.WriteEndElement();
                break;

            case ComputedValueKind.Boolean:
                writer.WriteAttributeString("t", "b");
                writer.WriteElementString("v", Ns, value.ToBoolean() ? "1" : "0");
                break;

            case ComputedValueKind.Text:
                writer.WriteAttributeString("t", "s");
                writer.WriteStartElement("v", Ns);
                XlsxNumbers.Write(writer, sharedStrings.IndexOf(value.ToText()));
                writer.WriteEndElement();
                break;

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                writer.WriteAttributeString("t", "e");
                writer.WriteElementString("v", Ns, error.ToString());
                break;

            case ComputedValueKind.Reference:
                // A bare reference result (e.g. a multi-cell OFFSET) has no single cell value.
                writer.WriteAttributeString("t", "e");
                writer.WriteElementString("v", Ns, Error.Value.ToString());
                break;
        }

        writer.WriteEndElement(); // </c>
    }

    // Emits a fresh <sheetData> containing only our rows — used when the target worksheet had none.
    private static void WriteSheetData(
        XmlWriter writer,
        Workbook workbook,
        string sheetName,
        SortedDictionary<int, List<(int Column, string Id)>> rows,
        SharedStrings sharedStrings
    )
    {
        writer.WriteStartElement("sheetData", Ns);

        foreach (var (rowNumber, cells) in rows)
        {
            WriteNewRow(writer, workbook, sheetName, rowNumber, cells, sharedStrings);
        }

        writer.WriteEndElement();
    }

    // Appends our text values to the target's shared-string table and hands back the index to reference, so a
    // repeated label costs one index per cell instead of a full inline copy. Loaded lazily (never touched if
    // the merge writes no text); existing entries keep their positions so passed-through cells stay valid;
    // rich-text entries are left intact and never reused.
    private sealed class SharedStrings
    {
        private readonly WorkbookPart _workbookPart;
        private SharedStringTable? _table;
        private Dictionary<string, int>? _plainIndex;
        private int _count;

        public SharedStrings(WorkbookPart workbookPart) => _workbookPart = workbookPart;

        public int IndexOf(string text)
        {
            EnsureLoaded();

            if (_plainIndex!.TryGetValue(text, out var existing))
            {
                return existing;
            }

            var index = _count++;
            _table!.AppendChild(new SharedStringItem(XlsxTextFactory.Create(text)));
            _plainIndex[text] = index;

            return index;
        }

        // The count/uniqueCount hints are now stale (we appended); drop them so Excel recomputes rather than
        // reporting a mismatch and forcing a repair. No-op if we never touched the table.
        public void Finish()
        {
            if (_table is null)
            {
                return;
            }

            _table.Count = null;
            _table.UniqueCount = null;
        }

        private void EnsureLoaded()
        {
            if (_table is not null)
            {
                return;
            }

            var part =
                _workbookPart.SharedStringTablePart
                ?? _workbookPart.AddNewPart<SharedStringTablePart>();
            _table = part.SharedStringTable ??= new SharedStringTable();
            _plainIndex = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (var item in _table.Elements<SharedStringItem>())
            {
                // Only a plain <si><t>…</t></si> entry (no rich-text runs) is safe to reuse by its text.
                if (item.Text is { } text && !item.Elements<Run>().Any())
                {
                    _plainIndex.TryAdd(text.Text ?? string.Empty, _count);
                }

                _count++;
            }
        }
    }
}
