using System.Buffers;

namespace Danfma.MySheet;

/// <summary>
/// Single-buffer, forward-only <see cref="IBufferWriter{T}"/> over a <see cref="Stream"/>: one ~64KB buffer
/// rented once from <see cref="ArrayPool{T}"/> and reused for the whole write, flushed to the stream only
/// when it fills (or a single request needs more room than it has). Backs the opt-in
/// <see cref="WorkbookIoBuffering.Pipelines"/> SYNCHRONOUS write path — a hard, deterministic memory ceiling
/// (no allocation reaches the Large Object Heap) traded for more, smaller <see cref="Stream.Write(byte[],int,int)"/>
/// calls than the default <see cref="WorkbookIoBuffering.Pooled"/> path. <see cref="System.IO.Pipelines.PipeWriter"/>
/// is not used here because its <c>FlushAsync</c> is async-only — blocking on it from a synchronous Save would
/// be sync-over-async.
/// </summary>
/// <remarks>
/// Safe ONLY with a writer that follows the "one outstanding <see cref="GetSpan"/>/<see cref="GetMemory"/>
/// region at a time, committed by <see cref="Advance"/> before the next request" discipline — exactly how
/// MemoryPack's generated <c>MemoryPackWriter&lt;TBufferWriter&gt;</c> drives an <see cref="IBufferWriter{T}"/>
/// (verified against MemoryPack 1.21.4's decompiled source: the real buffer writer's <c>Advance</c> is called
/// only when growing to a new region or at the final flush). Not thread-safe. The caller MUST call
/// <see cref="Dispose"/> to flush the trailing partial buffer and return the rental to the pool.
/// </remarks>
internal sealed class StreamBufferWriter : IBufferWriter<byte>, IDisposable
{
    private const int DefaultBufferSize = 64 * 1024;

    private readonly Stream _destination;
    private byte[] _buffer;
    private int _position;

    public StreamBufferWriter(Stream destination, int bufferSize = DefaultBufferSize)
    {
        _destination = destination;
        _buffer = ArrayPool<byte>.Shared.Rent(bufferSize);
    }

    public void Advance(int count)
    {
        if (count < 0 || _position + count > _buffer.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        _position += count;
    }

    public Memory<byte> GetMemory(int sizeHint = 0) => EnsureCapacity(sizeHint);

    public Span<byte> GetSpan(int sizeHint = 0) => EnsureCapacity(sizeHint).Span;

    // Flushes the current buffer (if it can't satisfy sizeHint) and rents a fresh one, dropping an oversized
    // one-off rental back to the default size so a single huge field (e.g. a long embedded string) does not
    // permanently inflate the steady-state buffer.
    private Memory<byte> EnsureCapacity(int sizeHint)
    {
        var required = Math.Max(sizeHint, 1);

        if (_buffer.Length - _position < required)
        {
            Flush();

            if (_buffer.Length < required || _buffer.Length > DefaultBufferSize)
            {
                ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = ArrayPool<byte>.Shared.Rent(Math.Max(required, DefaultBufferSize));
            }
        }

        return _buffer.AsMemory(_position);
    }

    private void Flush()
    {
        if (_position > 0)
        {
            _destination.Write(_buffer, 0, _position);
            _position = 0;
        }
    }

    public void Dispose()
    {
        Flush();
        ArrayPool<byte>.Shared.Return(_buffer);
    }
}
