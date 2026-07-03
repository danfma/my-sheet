using System.IO.Compression;
using MemoryPack;
using MessagePack;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// The set of format variants the spike compares. MemoryPack (production) and MessagePack-indexed are the
// two head-to-head formats; the LZ4 / GZip / Brotli rows exist to answer the honest question the plan raises:
// "does generic compression over the CURRENT format close the gap that a format switch would open?"
internal static class SpikeSerializers
{
    private static readonly MessagePackSerializerOptions Lz4 =
        MessagePackSerializerOptions.Standard.WithCompression(MessagePackCompression.Lz4BlockArray);

    // ── Production: MemoryPack on the real Workbook ──────────────────────────────────────────────────
    public static byte[] MemoryPack(Workbook workbook) => MemoryPackSerializer.Serialize(workbook);

    public static Workbook MemoryPackDeserialize(byte[] bytes) =>
        MemoryPackSerializer.Deserialize<Workbook>(bytes)!;

    // ── Candidate: MessagePack indexed keys on the mirror graph ──────────────────────────────────────
    public static byte[] MessagePackIndexed(MWorkbook mirror) => MessagePackSerializer.Serialize(mirror);

    public static MWorkbook MessagePackIndexedDeserialize(byte[] bytes) =>
        MessagePackSerializer.Deserialize<MWorkbook>(bytes);

    // ── Candidate variant: MessagePack + LZ4 block array ─────────────────────────────────────────────
    public static byte[] MessagePackLz4(MWorkbook mirror) => MessagePackSerializer.Serialize(mirror, Lz4);

    public static MWorkbook MessagePackLz4Deserialize(byte[] bytes) =>
        MessagePackSerializer.Deserialize<MWorkbook>(bytes, Lz4);

    // ── Honest reference: generic compression OVER the current MemoryPack bytes ───────────────────────
    // If GZip/Brotli over MemoryPack matches LZ4-over-MessagePack on size, the disk/network win does not
    // require a format break — it is one GZipStream away on the format we already ship.
    public static byte[] GZip(byte[] payload)
    {
        using var output = new MemoryStream();
        using (var gzip = new GZipStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            gzip.Write(payload, 0, payload.Length);
        }

        return output.ToArray();
    }

    public static byte[] Brotli(byte[] payload)
    {
        using var output = new MemoryStream();
        using (var brotli = new BrotliStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            brotli.Write(payload, 0, payload.Length);
        }

        return output.ToArray();
    }
}
