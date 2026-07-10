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
        File.WriteAllBytes(path, SerializeToBytes(options));
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
        await File.WriteAllBytesAsync(path, SerializeToBytes(options), cancellationToken);
    }

    // The bytes a Save writes. A cold, uncompressed save is the raw model, byte-identical to Save(path). Every
    // other combination is a container: warm (uncompressed) stays v1; anything compressed is v2 (Brotli over
    // model||values as ONE stream — a single stream compresses better than two separate blocks).
    private byte[] SerializeToBytes(WorkbookSaveOptions options)
    {
        var model = MemoryPackSerializer.Serialize(this);
        var compress = options.Compression == WorkbookCompression.Brotli;

        // Cold + uncompressed → the historical raw format (permanent byte-identity contract).
        if (!options.IncludeComputedValues && !compress)
        {
            return model;
        }

        // A cold compressed save carries no values; a warm save snapshots the cache.
        var values = MemoryPackSerializer.Serialize(
            options.IncludeComputedValues ? SnapshotComputedValues() : new List<CachedCellValue>()
        );

        return compress
            ? BuildContainer(
                ContainerVersionBrotli,
                model,
                BrotliCompress(model, values, options.CompressionLevel)
            )
            : BuildContainer(ContainerVersionUncompressed, model, Concat(model, values));
    }

    // Prepends the fixed 9-byte header to an already-encoded body. `modelLength` is always the UNCOMPRESSED
    // model length so the reader can slice model vs. values after (optionally) decompressing the body.
    private static byte[] BuildContainer(byte version, byte[] model, byte[] body)
    {
        var buffer = new byte[ContainerHeaderLength + body.Length];
        var span = buffer.AsSpan();

        ContainerMagic.CopyTo(span);
        span[4] = version;
        BinaryPrimitives.WriteInt32LittleEndian(span.Slice(5, 4), model.Length);
        body.CopyTo(span.Slice(ContainerHeaderLength));

        return buffer;
    }

    private static byte[] Concat(byte[] first, byte[] second)
    {
        var buffer = new byte[first.Length + second.Length];
        first.CopyTo(buffer.AsSpan());
        second.CopyTo(buffer.AsSpan(first.Length));
        return buffer;
    }

    // Compresses model||values as a single Brotli stream at the chosen level. One stream compresses better
    // than two independently-compressed blocks because the coder shares context across the whole payload. The
    // level affects only write time and size — Load decompresses any level, so the container is unversioned by it.
    private static byte[] BrotliCompress(byte[] model, byte[] values, CompressionLevel level)
    {
        using var output = new MemoryStream();
        using (var brotli = new BrotliStream(output, level, leaveOpen: true))
        {
            brotli.Write(model);
            brotli.Write(values);
        }

        return output.ToArray();
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
    ) => Deserialize(await File.ReadAllBytesAsync(path, cancellationToken), path);

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
