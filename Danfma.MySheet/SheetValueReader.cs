namespace Danfma.MySheet;

/// <summary>
/// A per-sheet, numeric-address reader over the workbook's memoized values — the bulk-extraction
/// counterpart of <see cref="Workbook.GetCellValue(string,string)"/>. Obtain one via
/// <see cref="Workbook.GetValueReader"/>; it captures the sheet's dense handle once, so each
/// <see cref="GetValue"/> hit is a direct paged-store read with no id string, no A1 parse and no
/// per-cell sheet-name hashing. Reads are thread-safe (seqlock-verified) and misses evaluate
/// on demand exactly like the string path.
/// </summary>
public readonly struct SheetValueReader
{
    private readonly Workbook _workbook;
    private readonly int _handle;
    private readonly string _sheetName;

    internal SheetValueReader(Workbook workbook, int handle, string sheetName)
    {
        _workbook = workbook;
        _handle = handle;
        _sheetName = sheetName;
    }

    /// <summary>The memoized value at the 1-based (column, row) address, evaluating on a miss.</summary>
    public ComputedValue GetValue(int column, int row) =>
        _workbook.GetCellValueDense(_handle, _sheetName, column, row);
}
