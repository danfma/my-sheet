using System.Buffers.Binary;
using System.IO.Compression;
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
    //   v1 (uncompressed): body = model bytes | value-block bytes            (warm-start, unchanged)
    //   v2 (Brotli):       body = Brotli(model bytes | value-block bytes)    (cold OR warm, compressed)
    // In every version `modelLength` is the UNCOMPRESSED model length, used to slice the (decompressed) body
    // into model vs. values. The value block is the MemoryPack of a List<CachedCellValue> surrogate (empty for
    // a cold compressed save). Load sniffs the 4-byte magic: a match is a container, anything else is a raw
    // (legacy or cold) model — the raw MemoryPack object header is a small member count (Workbook = 0x02),
    // never 'M' (0x4D), so the two are unambiguous.
    //
    // The container writers below (whole-array default, or the single-pass streaming async writer) produce
    // the SAME bytes for the same model/values/compression. The header's modelLength is known before the
    // model bytes exist in the whole-array path; the streaming path instead writes a placeholder and patches
    // it via seek-back once the real (uncompressed) count is known — safe because Save/SaveAsync only ever
    // target a path, so the destination FileStream is always seekable.
    private static ReadOnlySpan<byte> ContainerMagic => "MSWM"u8;
    private const byte ContainerVersionUncompressed = 1;
    private const byte ContainerVersionBrotli = 2;
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
    /// <see cref="Load(string)"/>.
    /// </summary>
    public void Save(string path, WorkbookSaveOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var compress = options.Compression == WorkbookCompression.Brotli;

        if (!options.IncludeComputedValues && !compress)
        {
            // Cold, uncompressed: the permanent byte-identity contract — raw model, no container/header.
            Save(path);
            return;
        }

        var values = options.IncludeComputedValues
            ? SnapshotComputedValues()
            : new List<CachedCellValue>();

        using var destination = File.Create(path);
        WriteContainerWhole(destination, compress, options.CompressionLevel, values);
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
            await SaveAsync(path, cancellationToken);
            return;
        }

        var values = options.IncludeComputedValues
            ? SnapshotComputedValues()
            : new List<CachedCellValue>();

        await using var destination = File.Create(path);
        await WriteContainerStreamAsync(
            destination,
            compress,
            options.CompressionLevel,
            values,
            cancellationToken
        );
    }

    // Container writer. modelLength is known up front (both parts are ordinary byte[] — MemoryPack has no
    // synchronous Stream-based Serialize, so this is the cheapest legal synchronous shape: two full-size
    // buffers, no further copies). Brotli, when requested, streams straight from those two buffers into the
    // destination — no intermediate MemoryStream.ToArray() copy, and no final "concat everything into one
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
        header[4] = compress ? ContainerVersionBrotli : ContainerVersionUncompressed;
        BinaryPrimitives.WriteInt32LittleEndian(header.Slice(5, 4), model.Length);
        destination.Write(header);

        if (compress)
        {
            using var brotli = new BrotliStream(destination, level, leaveOpen: true);
            brotli.Write(model);
            brotli.Write(valuesBytes);
        }
        else
        {
            destination.Write(model);
            destination.Write(valuesBytes);
        }
    }

    // ASYNC container writer. UNCOMPRESSED: genuinely single-pass, no full-size byte[] for the model at all —
    // MemoryPackSerializer.SerializeAsync already streams through its own pooled, segmented buffer writer
    // (verified against MemoryPack 1.21.4's decompiled source); CountingWriteStream taps the exact byte count
    // as those bytes fly by, so the header is patched via seek-back with no separate counting pass.
    // COMPRESSED: falls back to the same two whole-buffer Brotli Write calls as WriteContainerWhole — Brotli's
    // compressed output is sensitive to Write-CALL boundaries, not just content (verified empirically:
    // splitting the same bytes across many small Write calls measurably changes the compressed bytes at
    // Fastest/Optimal levels), and SerializeAsync's internal writer flushes in many small segments, which
    // WOULD break the byte-identity contract if fed straight into Brotli.
    private async Task WriteContainerStreamAsync(
        Stream destination,
        bool compress,
        CompressionLevel level,
        List<CachedCellValue> values,
        CancellationToken cancellationToken
    )
    {
        if (compress)
        {
            WriteContainerWhole(destination, compress: true, level, values);
            return;
        }

        var header = new byte[ContainerHeaderLength];
        ContainerMagic.CopyTo(header);
        header[4] = ContainerVersionUncompressed;
        await destination.WriteAsync(header, cancellationToken);

        var counting = new CountingWriteStream(destination);
        await MemoryPackSerializer.SerializeAsync(
            counting,
            this,
            cancellationToken: cancellationToken
        );
        var modelLength = counting.TotalBytesWritten;
        await MemoryPackSerializer.SerializeAsync(
            destination,
            values,
            cancellationToken: cancellationToken
        );

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

    /// <inheritdoc cref="Load(string)"/>
    public static async Task<Workbook> LoadAsync(
        string path,
        CancellationToken cancellationToken = default
    )
    {
        await using var stream = OpenReadAsync(path);

        // Peek the fixed 9-byte container header (magic + version + modelLength) without a separate stream —
        // on a short/empty file, ReadAtLeastAsync just reports fewer bytes read.
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

    private static Workbook DeserializeContainer(byte[] bytes, string path)
    {
        var version = bytes[4];
        var modelLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(5, 4));

        if (modelLength < 0)
        {
            throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
        }

        // Resolve the (decompressed) body once, then slice model vs. values by modelLength. For v1 the body is
        // the bytes after the header verbatim; for v2 it is the Brotli-decompressed payload.
        var body = version switch
        {
            ContainerVersionUncompressed => bytes.AsMemory(ContainerHeaderLength),
            ContainerVersionBrotli => BrotliDecompress(bytes.AsSpan(ContainerHeaderLength), path),
            _ => throw new InvalidDataException(
                $"'{path}' is a MySheet container of unsupported version {version}."
            ),
        };

        if (modelLength > body.Length)
        {
            throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
        }

        var workbook =
            MemoryPackSerializer.Deserialize<Workbook>(body.Span.Slice(0, modelLength))
            ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");

        var values =
            MemoryPackSerializer.Deserialize<List<CachedCellValue>>(body.Span.Slice(modelLength))
            ?? new List<CachedCellValue>();

        workbook.LoadComputedValues(values);

        return workbook;
    }

    private static byte[] BrotliDecompress(ReadOnlySpan<byte> compressed, string path)
    {
        try
        {
            using var input = new MemoryStream(compressed.ToArray());
            using var brotli = new BrotliStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            brotli.CopyTo(output);
            return output.ToArray();
        }
        catch (Exception inner)
            when (inner is InvalidDataException or IOException or InvalidOperationException)
        {
            throw new InvalidDataException(
                $"'{path}' is a corrupt MySheet container (Brotli payload could not be decompressed).",
                inner
            );
        }
    }
}
