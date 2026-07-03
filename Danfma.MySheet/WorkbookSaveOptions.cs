namespace Danfma.MySheet;

/// <summary>
/// Options for <see cref="Workbook.Save(string, WorkbookSaveOptions)"/> and its async counterpart.
/// </summary>
public sealed class WorkbookSaveOptions
{
    /// <summary>
    /// When <c>true</c>, the memoized computed values (the warm cache) are persisted alongside the model in a
    /// self-describing container so a subsequent <see cref="Workbook.Load(string)"/> can skip recomputation
    /// (a "warm start"). When <c>false</c> (the default) the file is byte-identical to
    /// <see cref="Workbook.Save(string)"/> — the raw model with no values. Volatile cells
    /// (<c>NOW</c>/<c>TODAY</c>/<c>RAND</c>/<c>RANDBETWEEN</c>, directly or transitively) and reference-typed
    /// results are never persisted: they recompute lazily on the first read after loading.
    /// </summary>
    public bool IncludeComputedValues { get; init; }

    /// <summary>
    /// Compression applied to the saved payload. Defaults to <see cref="WorkbookCompression.None"/> — a cold,
    /// uncompressed save is byte-identical to <see cref="Workbook.Save(string)"/>. Set to
    /// <see cref="WorkbookCompression.Brotli"/> to shrink the file (typically well under half the raw size on
    /// large workbooks); the compressed bytes are wrapped in the <c>MSWM</c> container so
    /// <see cref="Workbook.Load(string)"/> detects and decompresses them transparently. This is orthogonal to
    /// <see cref="IncludeComputedValues"/>: cold and warm saves can each be compressed. The library never
    /// changes the file name — the caller chooses the path (a <c>.br</c> suffix is the recommended convention
    /// for compressed files).
    /// </summary>
    public WorkbookCompression Compression { get; init; } = WorkbookCompression.None;
}
