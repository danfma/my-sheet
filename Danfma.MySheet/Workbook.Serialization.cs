using System.Buffers;
using System.Buffers.Binary;
using System.IO.Compression;
using System.IO.Pipelines;
using MemoryPack;

namespace Danfma.MySheet;

public sealed partial class Workbook
{
    // === Save container format ===========================================================================
    // A cold, uncompressed Save writes the RAW MemoryPack of the model, byte-identical to every prior version
    // (a permanent regression contract). Any other combination wraps that model in a self-describing container
    // whose fixed 9-byte header is shared by every version:
    //   magic "MSWM" (4) | version (1) | modelLength (int32 LE, 4) | body
    // `version` selects the body encoding (the header layout and offsets never change — new warm-start files
    // stay v1, so their tests are unaffected):
    //   v1 (uncompressed):     body = model bytes | value-block bytes                    (warm-start, unchanged)
    //   v2 (Brotli, LEGACY):   body = Brotli(model bytes | value-block bytes), written as TWO whole-buffer
    //                          BrotliStream.Write calls. READ-ONLY since M6: nothing in this library writes
    //                          v2 anymore (see the golden fixture in ContainerVersionCompatibilityTests), but
    //                          old v2 files load unchanged forever — the same one-way policy as an append-only
    //                          union tag (e.g. 319-321): a version, once superseded, is never removed from the
    //                          reader.
    //   v3 (Brotli, chunked):  body = Brotli(model bytes | value-block bytes), written as a sequence of
    //                          EXACTLY 64KB BrotliStream.Write calls (the final one shorter) — see
    //                          BrotliChunkedStream. THE DEFAULT for WorkbookCompression.Brotli since M6:
    //                          measured ~13% smaller than v2's whole-buffer write on a large real-world
    //                          payload (Optimal), and — unlike v2 — deterministic and byte-identical across
    //                          every WorkbookIoBuffering write mechanism, because BrotliChunkedStream's chunk
    //                          boundaries depend only on the cumulative byte count, never on the caller's
    //                          write-call granularity (v2 could not make that promise: see
    //                          WorkbookIoBufferingTests.BrotliStream_OutputDependsOnWriteCallBoundaries_NotJustContent,
    //                          which is why every v2 writer had to fall back to materializing the SAME two
    //                          whole buffers). A v3 file is a forward-only artifact: it cannot be opened by a
    //                          version of this library older than the one that introduced v3 (the older
    //                          reader's version switch does not recognize tag 3) — the same one-way boundary
    //                          documented for the shared-formula delta tags.
    // In every version `modelLength` is the UNCOMPRESSED model length, used to slice the (decompressed) body
    // into model vs. values. The value block is the MemoryPack of a List<CachedCellValue> surrogate (empty for
    // a cold compressed save). Load sniffs the 4-byte magic: a match is a container, anything else is a raw
    // (legacy or cold) model — the raw MemoryPack object header is a small member count (Workbook = 0x02),
    // never 'M' (0x4D), so the two are unambiguous.
    //
    // Every container writer below (whole-array default, single-pass streaming, or the opt-in pooled/Pipelines
    // writers) produces the SAME bytes for the same model/values/compression — WorkbookIoBuffering is a
    // write-MECHANISM choice only. The header's modelLength is always known before the model bytes exist in
    // the whole-array path; the streaming paths instead write a placeholder and patch it via seek-back once
    // the real (uncompressed) count is known — safe because Save/SaveAsync only ever target a path, so the
    // destination FileStream is always seekable.
    private static ReadOnlySpan<byte> ContainerMagic => "MSWM"u8;
    private const byte ContainerVersionUncompressed = 1;
    private const byte ContainerVersionBrotli = 2; // legacy: read-only since M6, see the format comment above
    private const byte ContainerVersionBrotliChunked = 3; // default Brotli encoding since M6
    private const int ContainerHeaderLength = 4 + 1 + 4; // magic + version + modelLength

    /// <summary>
    /// Serializes the workbook to a file (MemoryPack). The cell cache and registered custom functions are
    /// not persisted — they are rebuilt/re-registered after loading. Byte-identical across versions.
    /// </summary>
    public void Save(string path) => File.WriteAllBytes(path, MemoryPackSerializer.Serialize(this));

    /// <summary>
    /// Serializes the workbook to a file, honoring <paramref name="options"/>. With
    /// <see cref="WorkbookSaveOptions.IncludeComputedValues"/> the memoized values travel with the model in a
    /// container so a later <see cref="Load(string)"/> starts warm (no recomputation); volatile and
    /// reference-typed cache entries are never persisted. With
    /// <see cref="WorkbookSaveOptions.Compression"/> set to <see cref="WorkbookCompression.Brotli"/> the payload
    /// is Brotli-compressed inside the container. When both are at their defaults
    /// (<see cref="WorkbookCompression.None"/>, no values) the file is byte-identical to
    /// <see cref="Save(string)"/>. Either overload's output is transparently read back by
    /// <see cref="Load(string)"/>. <see cref="WorkbookSaveOptions.IoBuffering"/> only changes how the bytes
    /// are produced (see <see cref="WorkbookIoBuffering"/>), never the bytes themselves.
    /// </summary>
    public void Save(string path, WorkbookSaveOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var compress = options.Compression == WorkbookCompression.Brotli;

        if (!options.IncludeComputedValues && !compress)
        {
            // Cold, uncompressed: the permanent byte-identity contract — raw model, no container/header,
            // whichever write mechanism is chosen (the bytes are identical either way).
            using var coldDestination = File.Create(path);

            if (options.IoBuffering == WorkbookIoBuffering.Pipelines)
            {
                using var writer = new StreamBufferWriter(coldDestination);
                MemoryPackSerializer.Serialize(writer, this);
            }
            else
            {
                coldDestination.Write(MemoryPackSerializer.Serialize(this));
            }

            return;
        }

        var values = options.IncludeComputedValues
            ? SnapshotComputedValues()
            : new List<CachedCellValue>();

        using var destination = File.Create(path);

        if (options.IoBuffering == WorkbookIoBuffering.Pipelines)
        {
            WriteContainerPooled(destination, compress, options.CompressionLevel, values);
        }
        else
        {
            WriteContainerWhole(destination, compress, options.CompressionLevel, values);
        }
    }

    /// <inheritdoc cref="Save(string)"/>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = File.Create(path);
        await MemoryPackSerializer.SerializeAsync(
            stream,
            this,
            cancellationToken: cancellationToken
        );
    }

    /// <inheritdoc cref="Save(string, WorkbookSaveOptions)"/>
    public async Task SaveAsync(
        string path,
        WorkbookSaveOptions options,
        CancellationToken cancellationToken = default
    )
    {
        ArgumentNullException.ThrowIfNull(options);

        var compress = options.Compression == WorkbookCompression.Brotli;

        if (!options.IncludeComputedValues && !compress)
        {
            await using var coldDestination = File.Create(path);

            if (options.IoBuffering == WorkbookIoBuffering.Pipelines)
            {
                // Delegate to the bounded SYNC writer on a pool thread rather than serializing into a
                // PipeWriter: MemoryPack's Serialize(IBufferWriter) is synchronous, and a StreamPipeWriter
                // only drains to the stream during FlushAsync — so a pipe here would accumulate the ENTIRE
                // payload in segments before the single post-serialize flush, defeating the memory ceiling
                // (measured: ~150MB allocated vs ~108MB pooled on a 40MB model). Task.Run keeps the caller
                // unblocked; the file I/O itself is synchronous on that pool thread (acceptable for file
                // writes — it is exactly what the sync Save does). PipeReader stays on the LOAD side, where
                // reading can genuinely interleave with the pipe.
                await Task.Run(
                    () =>
                    {
                        using var writer = new StreamBufferWriter(coldDestination);
                        MemoryPackSerializer.Serialize(writer, this);
                    },
                    cancellationToken
                );
            }
            else
            {
                await MemoryPackSerializer.SerializeAsync(
                    coldDestination,
                    this,
                    cancellationToken: cancellationToken
                );
            }

            return;
        }

        var values = options.IncludeComputedValues
            ? SnapshotComputedValues()
            : new List<CachedCellValue>();

        await using var destination = File.Create(path);

        if (options.IoBuffering == WorkbookIoBuffering.Pipelines)
        {
            // Same PipeWriter-accumulation reasoning as the cold branch above: the bounded sync container
            // writer on a pool thread is strictly better than a pipe fed by a synchronous serializer.
            await Task.Run(
                () => WriteContainerPooled(destination, compress, options.CompressionLevel, values),
                cancellationToken
            );
        }
        else
        {
            await WriteContainerStreamAsync(
                destination,
                compress,
                options.CompressionLevel,
                values,
                cancellationToken
            );
        }
    }

    // Default (WorkbookIoBuffering.Pooled) container writer. modelLength is known up front (both parts are
    // ordinary byte[] — MemoryPack has no synchronous Stream-based Serialize, so this is the cheapest legal
    // synchronous shape: two full-size buffers, no further copies). Brotli, when requested, streams straight
    // from those two buffers into the destination through BrotliChunkedStream (v3's fixed-64KB chunking
    // discipline) — no intermediate MemoryStream.ToArray() copy, and no final "concat everything into one
    // container array" copy (the two buffers this replaces from the historical implementation).
    private void WriteContainerWhole(
        Stream destination,
        bool compress,
        CompressionLevel level,
        List<CachedCellValue> values
    )
    {
        var model = MemoryPackSerializer.Serialize(this);
        var valuesBytes = MemoryPackSerializer.Serialize(values);

        Span<byte> header = stackalloc byte[ContainerHeaderLength];
        ContainerMagic.CopyTo(header);
        header[4] = compress ? ContainerVersionBrotliChunked : ContainerVersionUncompressed;
        BinaryPrimitives.WriteInt32LittleEndian(header.Slice(5, 4), model.Length);
        destination.Write(header);

        if (compress)
        {
            using var brotli = new BrotliChunkedStream(destination, level);
            brotli.Write(model);
            brotli.Write(valuesBytes);
        }
        else
        {
            destination.Write(model);
            destination.Write(valuesBytes);
        }
    }

    // Opt-in (WorkbookIoBuffering.Pipelines) synchronous container writer. UNCOMPRESSED: model/values stream
    // through a single reused ~64KB StreamBufferWriter instead of two full-size byte[] buffers, and the header
    // (a placeholder until the model finishes streaming) is patched via seek-back — the destination is always
    // the file's own seekable FileStream, since Save only ever targets a path. COMPRESSED (v3): model/values
    // stream through that SAME StreamBufferWriter, but now pointed at a BrotliChunkedStream instead of the raw
    // destination — StreamBufferWriter's variable, content-dependent flush sizes are re-aggregated into
    // BrotliChunkedStream's fixed 64KB windows before anything reaches Brotli, so this is fully LOH-bounded
    // (no full-size model/values byte[]) AND byte-identical to the whole-array writer's v3 output (proven by
    // WorkbookIoBufferingTests.Save_PipelinesMatchesPooled_ForEveryOptionCombination) — v3's chunking, unlike
    // v2's, does not depend on write-call boundaries, so there is no whole-buffer fallback to keep here.
    private void WriteContainerPooled(
        Stream destination,
        bool compress,
        CompressionLevel level,
        List<CachedCellValue> values
    )
    {
        Span<byte> header = stackalloc byte[ContainerHeaderLength];
        ContainerMagic.CopyTo(header);
        header[4] = compress ? ContainerVersionBrotliChunked : ContainerVersionUncompressed;
        destination.Write(header);

        long modelLength;

        if (compress)
        {
            using var brotli = new BrotliChunkedStream(destination, level);
            using var raw = new StreamBufferWriter(brotli);
            var counting = new CountingBufferWriter<StreamBufferWriter>(raw);
            MemoryPackSerializer.Serialize(counting, this);
            modelLength = counting.TotalBytesWritten;
            MemoryPackSerializer.Serialize(raw, values);
        }
        else
        {
            using var raw = new StreamBufferWriter(destination);
            var counting = new CountingBufferWriter<StreamBufferWriter>(raw);
            MemoryPackSerializer.Serialize(counting, this);
            modelLength = counting.TotalBytesWritten;
            MemoryPackSerializer.Serialize(raw, values);
        }

        PatchModelLength(destination, modelLength);
    }

    // Default (WorkbookIoBuffering.Pooled) ASYNC container writer. UNCOMPRESSED: genuinely single-pass, no
    // full-size byte[] for the model at all — MemoryPackSerializer.SerializeAsync already streams through its
    // own pooled, segmented buffer writer (verified against MemoryPack 1.21.4's decompiled source);
    // CountingWriteStream taps the exact byte count as those bytes fly by, so the header is patched via
    // seek-back with no separate counting pass. COMPRESSED (v3): the same single-pass SerializeAsync stream,
    // now routed through a BrotliChunkedStream — its fixed 64KB chunking re-aggregates SerializeAsync's many
    // small internal segments into Brotli-facing writes that are byte-identical to every other v3 writer
    // (unlike v2, which needed the whole-buffer fallback this replaces, because its output depended on
    // exactly those write-call boundaries).
    private async Task WriteContainerStreamAsync(
        Stream destination,
        bool compress,
        CompressionLevel level,
        List<CachedCellValue> values,
        CancellationToken cancellationToken
    )
    {
        var header = new byte[ContainerHeaderLength];
        ContainerMagic.CopyTo(header);
        header[4] = compress ? ContainerVersionBrotliChunked : ContainerVersionUncompressed;
        await destination.WriteAsync(header, cancellationToken);

        long modelLength;

        if (compress)
        {
            await using var brotli = new BrotliChunkedStream(destination, level);
            var counting = new CountingWriteStream(brotli);
            await MemoryPackSerializer.SerializeAsync(
                counting,
                this,
                cancellationToken: cancellationToken
            );
            modelLength = counting.TotalBytesWritten;
            await MemoryPackSerializer.SerializeAsync(
                brotli,
                values,
                cancellationToken: cancellationToken
            );
        }
        else
        {
            var counting = new CountingWriteStream(destination);
            await MemoryPackSerializer.SerializeAsync(
                counting,
                this,
                cancellationToken: cancellationToken
            );
            modelLength = counting.TotalBytesWritten;
            await MemoryPackSerializer.SerializeAsync(
                destination,
                values,
                cancellationToken: cancellationToken
            );
        }

        await PatchModelLengthAsync(destination, modelLength, cancellationToken);
    }

    private static void PatchModelLength(Stream destination, long modelLength)
    {
        destination.Position = 5;
        Span<byte> lengthBytes = stackalloc byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(lengthBytes, checked((int)modelLength));
        destination.Write(lengthBytes);
    }

    private static async Task PatchModelLengthAsync(
        Stream destination,
        long modelLength,
        CancellationToken cancellationToken
    )
    {
        destination.Position = 5;
        var lengthBytes = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(lengthBytes, checked((int)modelLength));
        await destination.WriteAsync(lengthBytes, cancellationToken);
    }

    // Snapshots the memoized cache into surrogates, EXCLUDING volatile-tainted entries (they must re-sample on
    // the next read after loading) and reference-typed values (unrepresentable — their cells recompute).
    private List<CachedCellValue> SnapshotComputedValues()
    {
        if (_valueStore is not { } store)
        {
            return new List<CachedCellValue>();
        }

        // Presized off the store's own (cheap, approximate) cell count instead of growing cell by cell —
        // it may slightly overcount (tainted cells the loop below skips), which only costs a few unused
        // slots, never a regrowth.
        var list = new List<CachedCellValue>(store.EstimatedPresentCount());

        // The store already excludes volatile-tainted cells; the surrogate factory drops the unrepresentable
        // Reference kind. Present blank cells ARE carried (an explicitly-empty cached cell round-trips).
        foreach (var (sheetName, id, value) in store.EnumerateNonTainted())
        {
            if (CachedCellValue.TryFrom(sheetName, id, value) is { } surrogate)
            {
                list.Add(surrogate);
            }
        }

        return list;
    }

    // Repopulates the memoized store from a warm value block, reusing the same lazy store creation path as
    // GetCellValue so the field survives MemoryPack's field-initializer bypass and stays consistent.
    private void LoadComputedValues(List<CachedCellValue> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        var store = ValueStore;

        foreach (var entry in values)
        {
            store.LoadEntry(entry.SheetName, entry.CellId, entry.ToComputedValue());
        }
    }

    /// <summary>Loads a workbook from a file written by <see cref="Save(string)"/> and returns the new
    /// instance. Warm-start containers repopulate the value cache; raw (cold/legacy) files load unchanged.</summary>
    public static Workbook Load(string path) => Deserialize(File.ReadAllBytes(path), path);

    /// <summary>
    /// Loads a workbook, honoring <paramref name="options"/>. <see cref="WorkbookLoadOptions.IoBuffering"/>
    /// only changes how the file's bytes are read (see <see cref="WorkbookIoBuffering"/>) — the resulting
    /// workbook is identical to <see cref="Load(string)"/> either way.
    /// </summary>
    public static Workbook Load(string path, WorkbookLoadOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        if (options.IoBuffering == WorkbookIoBuffering.Pooled)
        {
            return Load(path);
        }

        using var stream = new FileStream(
            path,
            new FileStreamOptions
            {
                Mode = FileMode.Open,
                Access = FileAccess.Read,
                Share = FileShare.Read,
                Options = FileOptions.SequentialScan,
            }
        );

        Span<byte> header = stackalloc byte[ContainerHeaderLength];
        var headerRead = stream.ReadAtLeast(
            header,
            ContainerHeaderLength,
            throwOnEndOfStream: false
        );

        if (SniffContainerHeader(header, headerRead))
        {
            // Containers are the smaller, warm-cache case — same byte[] path regardless of IoBuffering.
            stream.Seek(0, SeekOrigin.Begin);
            var bytes = new byte[stream.Length];
            stream.ReadExactly(bytes);
            return DeserializeContainer(bytes, path);
        }

        stream.Seek(0, SeekOrigin.Begin);
        return LoadRawPooledSequence(stream, path);
    }

    // Reads the raw model in ~64KB ArrayPool-backed segments and hands MemoryPack a ReadOnlySequence, instead
    // of one contiguous byte[] covering the whole file — the synchronous mirror of DeserializeAsync's own
    // segmented/pooled internal behavior (System.IO.Pipelines' PipeWriter/PipeReader are async-first, so a
    // synchronous LOH-bounded read needs this hand-built chain instead).
    private static Workbook LoadRawPooledSequence(Stream stream, string path)
    {
        const int SegmentSize = 64 * 1024;

        PooledSequenceSegment? first = null;
        PooledSequenceSegment? last = null;

        try
        {
            long runningIndex = 0;
            int read;

            do
            {
                var buffer = ArrayPool<byte>.Shared.Rent(SegmentSize);
                read = stream.Read(buffer, 0, buffer.Length);

                if (read == 0)
                {
                    ArrayPool<byte>.Shared.Return(buffer);
                    break;
                }

                var segment = new PooledSequenceSegment(buffer, read, runningIndex);
                runningIndex += read;

                if (first is null)
                {
                    first = last = segment;
                }
                else
                {
                    last!.SetNext(segment);
                    last = segment;
                }
            } while (read > 0);

            var sequence = first is null
                ? ReadOnlySequence<byte>.Empty
                : new ReadOnlySequence<byte>(first, 0, last!, last!.Memory.Length);

            return MemoryPackSerializer.Deserialize<Workbook>(sequence)
                ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
        }
        finally
        {
            for (var node = first; node is not null; )
            {
                var next = (PooledSequenceSegment?)node.Next;
                node.Return();
                node = next;
            }
        }
    }

    /// <inheritdoc cref="Load(string)"/>
    public static async Task<Workbook> LoadAsync(
        string path,
        CancellationToken cancellationToken = default
    )
    {
        await using var stream = OpenReadAsync(path);

        var header = new byte[ContainerHeaderLength];
        var headerRead = await stream.ReadAtLeastAsync(
            header,
            ContainerHeaderLength,
            throwOnEndOfStream: false,
            cancellationToken
        );

        if (SniffContainerHeader(header, headerRead))
        {
            // Containers are the smaller, warm-cache case — read the (already-compact) bytes once and reuse
            // the existing synchronous container path unchanged.
            stream.Seek(0, SeekOrigin.Begin);
            var bytes = new byte[stream.Length];
            await stream.ReadExactlyAsync(bytes, cancellationToken);
            return DeserializeContainer(bytes, path);
        }

        // Raw model: the common, larger case. Hand the stream straight to MemoryPack's async deserializer,
        // which reads through pooled, segmented ~64KB chunks built into a ReadOnlySequence instead of
        // materializing the whole file as one contiguous byte[] up front (the historical
        // File.ReadAllBytesAsync + Deserialize(byte[])).
        stream.Seek(0, SeekOrigin.Begin);
        return await MemoryPackSerializer.DeserializeAsync<Workbook>(
                stream,
                cancellationToken: cancellationToken
            ) ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
    }

    /// <inheritdoc cref="Load(string, WorkbookLoadOptions)"/>
    public static async Task<Workbook> LoadAsync(
        string path,
        WorkbookLoadOptions options,
        CancellationToken cancellationToken = default
    )
    {
        ArgumentNullException.ThrowIfNull(options);

        if (options.IoBuffering == WorkbookIoBuffering.Pooled)
        {
            return await LoadAsync(path, cancellationToken);
        }

        await using var stream = OpenReadAsync(path);

        var header = new byte[ContainerHeaderLength];
        var headerRead = await stream.ReadAtLeastAsync(
            header,
            ContainerHeaderLength,
            throwOnEndOfStream: false,
            cancellationToken
        );

        if (SniffContainerHeader(header, headerRead))
        {
            // Containers are the smaller, warm-cache case — same byte[] path regardless of IoBuffering.
            stream.Seek(0, SeekOrigin.Begin);
            var bytes = new byte[stream.Length];
            await stream.ReadExactlyAsync(bytes, cancellationToken);
            return DeserializeContainer(bytes, path);
        }

        stream.Seek(0, SeekOrigin.Begin);

        var reader = PipeReader.Create(stream, new StreamPipeReaderOptions(leaveOpen: true));
        try
        {
            ReadResult result;
            do
            {
                result = await reader.ReadAsync(cancellationToken);

                // PipeReader REQUIRES an AdvanceTo call between reads — without it, the reader has no signal
                // that the buffer was examined and can hang waiting for "new" data that never comes. Consumed
                // stays at Start (nothing processed yet — MemoryPack needs the whole sequence at once);
                // examined moves to End so the next ReadAsync call actually waits for more bytes to arrive.
                if (!result.IsCompleted)
                {
                    reader.AdvanceTo(result.Buffer.Start, result.Buffer.End);
                }
            } while (!result.IsCompleted);

            var workbook =
                MemoryPackSerializer.Deserialize<Workbook>(result.Buffer)
                ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
            reader.AdvanceTo(result.Buffer.End);
            return workbook;
        }
        finally
        {
            await reader.CompleteAsync();
        }
    }

    private static FileStream OpenReadAsync(string path) =>
        new(
            path,
            new FileStreamOptions
            {
                Mode = FileMode.Open,
                Access = FileAccess.Read,
                Share = FileShare.Read,
                Options = FileOptions.Asynchronous | FileOptions.SequentialScan,
            }
        );

    private static bool SniffContainerHeader(ReadOnlySpan<byte> header, int bytesRead) =>
        bytesRead >= ContainerHeaderLength
        && header.Slice(0, ContainerMagic.Length).SequenceEqual(ContainerMagic);

    // Sniffs the 4-byte magic: a container repopulates the warm cache; anything else is a raw model.
    private static Workbook Deserialize(byte[] bytes, string path)
    {
        if (
            bytes.Length >= ContainerHeaderLength
            && bytes.AsSpan(0, ContainerMagic.Length).SequenceEqual(ContainerMagic)
        )
        {
            return DeserializeContainer(bytes, path);
        }

        return MemoryPackSerializer.Deserialize<Workbook>(bytes)
            ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
    }

    // For v1 the body is a zero-copy view over the already-in-memory `bytes` array (the container is the
    // smaller, small-or-already-compressed case, so `bytes` never needs re-chunking). For v2/v3 the body is
    // decompressed into a chain of pooled ~64KB segments instead of one contiguous byte[] — the decompressed
    // payload can be a large multiple of the compressed file size, which is exactly where the historical
    // MemoryStream.ToArray()->ToArray() chain paid its cost. Every rented segment is returned in `finally`.
    private static Workbook DeserializeContainer(byte[] bytes, string path)
    {
        var version = bytes[4];
        var modelLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(5, 4));

        if (modelLength < 0)
        {
            throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
        }

        PooledSequenceSegment? firstSegment = null;

        try
        {
            ReadOnlySequence<byte> body;

            switch (version)
            {
                case ContainerVersionUncompressed:
                    body = new ReadOnlySequence<byte>(bytes.AsMemory(ContainerHeaderLength));
                    break;
                case ContainerVersionBrotli:
                case ContainerVersionBrotliChunked:
                    (body, firstSegment) = BrotliDecompressToSequence(
                        bytes,
                        ContainerHeaderLength,
                        bytes.Length - ContainerHeaderLength,
                        path
                    );
                    break;
                default:
                    throw new InvalidDataException(
                        $"'{path}' is a MySheet container of unsupported version {version}."
                    );
            }

            if (modelLength > body.Length)
            {
                throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
            }

            var workbook =
                MemoryPackSerializer.Deserialize<Workbook>(body.Slice(0, modelLength))
                ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");

            var values =
                MemoryPackSerializer.Deserialize<List<CachedCellValue>>(body.Slice(modelLength))
                ?? new List<CachedCellValue>();

            workbook.LoadComputedValues(values);

            return workbook;
        }
        finally
        {
            for (var node = firstSegment; node is not null; )
            {
                var next = (PooledSequenceSegment?)node.Next;
                node.Return();
                node = next;
            }
        }
    }

    // Decompresses a Brotli-compressed span (v2's whole-buffer body and v3's fixed-64KB-chunked body decode
    // through the exact same BrotliStream reader — chunking is a WRITE-time discipline only, invisible to
    // decompression) into a chain of ArrayPool-backed ~64KB segments chained into a ReadOnlySequence<byte>,
    // the same idiom LoadRawPooledSequence uses for the raw (uncompressed) load path. `source` is read
    // in-place (no upfront ToArray() copy of the compressed bytes — MemoryStream wraps the existing array);
    // only the larger decompressed output is segmented and pooled. On success the caller owns returning every
    // segment (see DeserializeContainer's finally); on failure this method returns its own rentals before
    // rethrowing.
    private static (
        ReadOnlySequence<byte> Body,
        PooledSequenceSegment? First
    ) BrotliDecompressToSequence(byte[] source, int offset, int count, string path)
    {
        const int SegmentSize = 64 * 1024;

        PooledSequenceSegment? first = null;
        PooledSequenceSegment? last = null;

        try
        {
            using var input = new MemoryStream(source, offset, count, writable: false);
            using var brotli = new BrotliStream(input, CompressionMode.Decompress);

            long runningIndex = 0;
            int read;

            do
            {
                var buffer = ArrayPool<byte>.Shared.Rent(SegmentSize);
                read = brotli.Read(buffer, 0, SegmentSize);

                if (read == 0)
                {
                    ArrayPool<byte>.Shared.Return(buffer);
                    break;
                }

                var segment = new PooledSequenceSegment(buffer, read, runningIndex);
                runningIndex += read;

                if (first is null)
                {
                    first = last = segment;
                }
                else
                {
                    last!.SetNext(segment);
                    last = segment;
                }
            } while (read > 0);

            var sequence = first is null
                ? ReadOnlySequence<byte>.Empty
                : new ReadOnlySequence<byte>(first, 0, last!, last!.Memory.Length);

            return (sequence, first);
        }
        catch (Exception inner)
            when (inner is InvalidDataException or IOException or InvalidOperationException)
        {
            for (var node = first; node is not null; )
            {
                var next = (PooledSequenceSegment?)node.Next;
                node.Return();
                node = next;
            }

            throw new InvalidDataException(
                $"'{path}' is a corrupt MySheet container (Brotli payload could not be decompressed).",
                inner
            );
        }
    }
}
