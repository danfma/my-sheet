namespace Danfma.MySheet;

/// <summary>
/// Selects the buffering strategy <see cref="Workbook.Save(string, WorkbookSaveOptions)"/> and
/// <see cref="Workbook.Load(string, WorkbookLoadOptions)"/> (and their async counterparts) use to move bytes
/// between the workbook model and the file. This is a write/read MECHANISM choice only — the bytes on disk
/// are identical regardless of which value is selected; see each option's remarks for the reference source
/// that verified <c>DeserializeAsync</c>'s internal pooling and the container-warm-path spike this trade-off
/// came out of.
/// </summary>
public enum WorkbookIoBuffering
{
    /// <summary>
    /// Pooled, stream-oriented I/O — the default and fastest option. Synchronous paths serialize each part
    /// (model, then values) as one contiguous buffer, reusing MemoryPack's own internal pooling; asynchronous
    /// paths stream through MemoryPack's own segmented, <see cref="System.Buffers.ArrayPool{T}"/>-backed
    /// reader/writer with no full-file <c>byte[]</c> materialization. Peak per-buffer size scales with the
    /// model/value-block size (large workbooks can still reach the Large Object Heap on the synchronous path).
    /// </summary>
    Pooled = 0,

    /// <summary>
    /// LOH-bounded, deterministic-footprint I/O. Writes flow through a single reused ~64KB pooled buffer
    /// (synchronous — <see cref="System.IO.Pipelines.PipeWriter"/> is async-first, so a bounded sync ceiling
    /// needs a hand-rolled writer instead) or a <see cref="System.IO.Pipelines.PipeWriter"/>
    /// (asynchronous); reads are rebuilt from pooled ~64KB segments into a
    /// <see cref="System.Buffers.ReadOnlySequence{T}"/> (synchronous) or read via a
    /// <see cref="System.IO.Pipelines.PipeReader"/> (asynchronous) instead of one contiguous <c>byte[]</c>.
    /// No single allocation reaches the Large Object Heap. This costs roughly 2-3x more time than
    /// <see cref="Pooled"/> (more, smaller I/O calls) — opt in only when LOH pressure, not raw throughput, is
    /// the constraint. A warm-start container's value block is unaffected by this choice on <em>load</em> —
    /// containers are read as one buffer either way (they are the smaller, already-compressed case); this
    /// only changes how the (larger) raw model is read/written.
    /// </summary>
    /// <remarks>
    /// <see cref="WorkbookSaveOptions.Compression"/> set to <see cref="WorkbookCompression.Brotli"/> is a
    /// KNOWN EXCEPTION on write: splitting the same bytes across many small <c>BrotliStream.Write</c> calls
    /// measurably changes the compressed output (verified empirically), so a compressed save always
    /// serializes each part (model, then values) as one whole buffer before handing it to Brotli in the SAME
    /// two calls the default path uses — this keeps every save byte-identical for a given
    /// options/compression combination, at the cost of the LOH-bounded guarantee for that one combination
    /// (compressed + <see cref="Pipelines"/> still allocates one full-size buffer per part). Uncompressed
    /// saves, and every load, get the full LOH-bounded behavior.
    /// </remarks>
    Pipelines = 1,
}
