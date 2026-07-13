using System.Buffers;
using System.IO.Compression;

namespace Danfma.MySheet;

/// <summary>
/// Write-only <see cref="Stream"/> that re-chunks whatever it receives into EXACTLY 64KB pieces (the final
/// one partial) before handing them to an inner <see cref="BrotliStream"/> — the container format's v3
/// chunking DISCIPLINE (see the format comment atop <c>Workbook.Serialization.cs</c>). Chunk boundaries are a
/// pure function of the cumulative byte count written since construction, never of the CALLER's write-call
/// granularity: feeding it the model+values as two whole-buffer <see cref="Write(byte[],int,int)"/> calls
/// (the default <see cref="WorkbookIoBuffering.Pooled"/> writer) or as many small MemoryPack-driven writes
/// (the streaming <see cref="WorkbookIoBuffering.Pipelines"/>/async writers) produces the byte-IDENTICAL
/// compressed output, because this type buffers into its own fixed-size window regardless of how the
/// incoming spans are sliced. That is what lets every write mechanism share one v3 byte-identity contract —
/// unlike v2's whole-buffer Brotli.Write, which was proven sensitive to caller chunking (see
/// <c>WorkbookIoBufferingTests.BrotliStream_OutputDependsOnWriteCallBoundaries_NotJustContent</c>) and
/// therefore needed every mechanism to materialize the SAME two whole buffers to stay byte-identical.
/// </summary>
/// <remarks>
/// Not thread-safe. The caller MUST call <see cref="Dispose"/> (or <see cref="DisposeAsync"/>) to flush the
/// trailing partial chunk and let the inner <see cref="BrotliStream"/> write its own trailer.
/// </remarks>
internal sealed class BrotliChunkedStream : Stream
{
    internal const int ChunkSize = 64 * 1024;

    private readonly BrotliStream _brotli;
    private readonly byte[] _buffer;
    private int _position;

    public BrotliChunkedStream(Stream destination, CompressionLevel level)
    {
        _brotli = new BrotliStream(destination, level, leaveOpen: true);
        _buffer = ArrayPool<byte>.Shared.Rent(ChunkSize);
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    // Deliberately a NO-OP, not a forward to _brotli.Flush()/FlushAsync(): BrotliStream.Flush is NOT
    // byte-neutral — it forces the encoder to emit whatever it has buffered as its own block, which changes
    // the compressed output (the same write-call-boundary sensitivity the v2 format's whole-buffer fallback
    // existed to avoid — see the format comment atop Workbook.Serialization.cs). MemoryPackSerializer
    // .SerializeAsync(Stream, ...) calls FlushAsync on completion of EACH part; the sync IBufferWriter-driven
    // path (StreamBufferWriter) has no Flush concept at all and can never trigger this. Making Flush a no-op
    // here keeps the two paths byte-identical: our own fixed-64KB chunking (Write/WriteAsync) is the ONLY
    // thing allowed to decide Brotli write boundaries. Nothing is lost — pending bytes still flush correctly
    // on Dispose/DisposeAsync.
    public override void Flush() { }

    public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) =>
        Write(buffer.AsSpan(offset, count));

    public override void Write(ReadOnlySpan<byte> data)
    {
        while (!data.IsEmpty)
        {
            var take = Math.Min(ChunkSize - _position, data.Length);
            data[..take].CopyTo(_buffer.AsSpan(_position, take));
            _position += take;
            data = data[take..];

            if (_position == ChunkSize)
            {
                _brotli.Write(_buffer, 0, ChunkSize);
                _position = 0;
            }
        }
    }

    public override Task WriteAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken
    ) => WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

    public override async ValueTask WriteAsync(
        ReadOnlyMemory<byte> data,
        CancellationToken cancellationToken = default
    )
    {
        while (!data.IsEmpty)
        {
            var take = Math.Min(ChunkSize - _position, data.Length);
            data[..take].CopyTo(_buffer.AsMemory(_position, take));
            _position += take;
            data = data[take..];

            if (_position == ChunkSize)
            {
                await _brotli.WriteAsync(_buffer.AsMemory(0, ChunkSize), cancellationToken);
                _position = 0;
            }
        }
    }

    // Flushes the trailing partial chunk (< ChunkSize bytes — the format's own "the final write may be
    // shorter" allowance) then disposes the inner BrotliStream so it writes its own trailer.
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_position > 0)
            {
                _brotli.Write(_buffer, 0, _position);
                _position = 0;
            }
            ArrayPool<byte>.Shared.Return(_buffer);
            _brotli.Dispose();
        }
        base.Dispose(disposing);
    }

    public override async ValueTask DisposeAsync()
    {
        if (_position > 0)
        {
            await _brotli.WriteAsync(_buffer.AsMemory(0, _position));
            _position = 0;
        }
        ArrayPool<byte>.Shared.Return(_buffer);
        await _brotli.DisposeAsync();
        GC.SuppressFinalize(this);
    }
}
