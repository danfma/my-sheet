namespace Danfma.MySheet;

/// <summary>
/// Forwards writes to <paramref name="inner"/> while tallying the total bytes written. Used by the default
/// (<see cref="WorkbookIoBuffering.Pooled"/>) async container writer to learn the exact (uncompressed) model
/// length as a side effect of the single streaming pass through
/// <c>MemoryPack.MemoryPackSerializer.SerializeAsync</c>, instead of a separate counting or dry-run step —
/// the container header is then patched via seek-back (the destination is always a seekable file). Write-only;
/// every other member throws.
/// </summary>
internal sealed class CountingWriteStream(Stream inner) : Stream
{
    public long TotalBytesWritten { get; private set; }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() => inner.Flush();

    public override Task FlushAsync(CancellationToken cancellationToken) =>
        inner.FlushAsync(cancellationToken);

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        inner.Write(buffer, offset, count);
        TotalBytesWritten += count;
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        inner.Write(buffer);
        TotalBytesWritten += buffer.Length;
    }

    public override async ValueTask WriteAsync(
        ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default
    )
    {
        await inner.WriteAsync(buffer, cancellationToken);
        TotalBytesWritten += buffer.Length;
    }

    public override Task WriteAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken
    ) => WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
}
