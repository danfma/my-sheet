using System.Buffers;

namespace Danfma.MySheet;

/// <summary>
/// A <see cref="ReadOnlySequenceSegment{T}"/> backed by an <see cref="ArrayPool{T}"/> rental. Chains
/// synchronously-read ~64KB chunks into a <see cref="ReadOnlySequence{T}"/> that MemoryPack can deserialize
/// directly, so the opt-in <see cref="WorkbookIoBuffering.Pipelines"/> synchronous load path never
/// materializes the whole file as one contiguous <c>byte[]</c> (the segmented-read mirror of what
/// <c>MemoryPackSerializer.DeserializeAsync</c> already does internally on the async path). Every segment's
/// rental must be returned via <see cref="Return"/> once the deserialize call that consumes the sequence has
/// finished — the payload is intentionally non-contiguous, so every segment has to stay alive until then.
/// </summary>
internal sealed class PooledSequenceSegment : ReadOnlySequenceSegment<byte>
{
    private readonly byte[] _array;

    public PooledSequenceSegment(byte[] array, int length, long runningIndex)
    {
        _array = array;
        Memory = array.AsMemory(0, length);
        RunningIndex = runningIndex;
    }

    public void SetNext(PooledSequenceSegment next) => Next = next;

    public void Return() => ArrayPool<byte>.Shared.Return(_array);
}
