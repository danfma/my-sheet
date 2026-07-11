using System.Buffers;

namespace Danfma.MySheet;

/// <summary>
/// Forwards <see cref="IBufferWriter{T}"/> calls to <paramref name="inner"/> while tallying the total bytes
/// <see cref="Advance"/>d. Used by the opt-in <see cref="WorkbookIoBuffering.Pipelines"/> container writer
/// (both the synchronous <see cref="StreamBufferWriter"/> and the asynchronous
/// <see cref="System.IO.Pipelines.PipeWriter"/>) to learn the exact (uncompressed) model length as a side
/// effect of the single write pass — the container header is then patched via seek-back.
/// </summary>
internal sealed class CountingBufferWriter<TInner>(TInner inner) : IBufferWriter<byte>
    where TInner : IBufferWriter<byte>
{
    public long TotalBytesWritten { get; private set; }

    public void Advance(int count)
    {
        inner.Advance(count);
        TotalBytesWritten += count;
    }

    public Memory<byte> GetMemory(int sizeHint = 0) => inner.GetMemory(sizeHint);

    public Span<byte> GetSpan(int sizeHint = 0) => inner.GetSpan(sizeHint);
}
