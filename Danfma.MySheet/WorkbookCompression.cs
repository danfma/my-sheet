namespace Danfma.MySheet;

/// <summary>
/// Selects the compression applied to a saved workbook. Orthogonal to
/// <see cref="WorkbookSaveOptions.IncludeComputedValues"/>: a cold (model-only) or warm (model + values)
/// save can each be written compressed or uncompressed.
/// </summary>
public enum WorkbookCompression
{
    /// <summary>No compression (the default). A cold, uncompressed save is byte-identical to
    /// <see cref="Workbook.Save(string)"/>.</summary>
    None = 0,

    /// <summary>Compress the payload with Brotli (<see cref="System.IO.Compression.CompressionLevel.Optimal"/>).
    /// The bytes are wrapped in the self-describing <c>MSWM</c> container so a later
    /// <see cref="Workbook.Load(string)"/> detects and transparently decompresses them. Brotli is part of the
    /// BCL, so this adds no third-party dependency.</summary>
    Brotli = 1,
}
