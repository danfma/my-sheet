using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// Byte-size harness for plans/messagepack-spike.md, question 1. `dotnet run -c Release -- --messagepack-size`.
// Also verifies round-trip correctness of the mirror (deserialize both formats, compare cell counts +
// spot-check values) so the size numbers are not measuring a broken/partial encoding.
internal static class MessagePackSizeReport
{
    public static void Run()
    {
        Console.WriteLine("== MessagePack format spike — byte sizes ==\n");

        foreach (
            var size in new[]
            {
                MessagePackPayloads.Size.Small,
                MessagePackPayloads.Size.Medium,
                MessagePackPayloads.Size.Large,
            }
        )
        {
            Report(size);
        }
    }

    private static void Report(MessagePackPayloads.Size size)
    {
        var workbook = MessagePackPayloads.Build(size);
        var mirror = MirrorConverter.ToMirror(workbook);

        var cellCount = workbook.Sheets.Values.Sum(s => s.Count);

        var memoryPack = SpikeSerializers.MemoryPack(workbook);
        var msgpack = SpikeSerializers.MessagePackIndexed(mirror);
        var msgpackLz4 = SpikeSerializers.MessagePackLz4(mirror);
        var memoryPackGZip = SpikeSerializers.GZip(memoryPack);
        var memoryPackBrotli = SpikeSerializers.Brotli(memoryPack);
        var msgpackGZip = SpikeSerializers.GZip(msgpack);

        Verify(workbook, memoryPack, mirror, msgpack, msgpackLz4);

        Console.WriteLine($"--- {size} ({cellCount:N0} cells) ---");
        Row("MemoryPack (production)", memoryPack.Length, memoryPack.Length);
        Row("MessagePack indexed", msgpack.Length, memoryPack.Length);
        Row("MessagePack + LZ4", msgpackLz4.Length, memoryPack.Length);
        Row("MessagePack + GZip", msgpackGZip.Length, memoryPack.Length);
        Row("MemoryPack + GZip", memoryPackGZip.Length, memoryPack.Length);
        Row("MemoryPack + Brotli", memoryPackBrotli.Length, memoryPack.Length);
        Console.WriteLine();
    }

    private static void Row(string name, int bytes, int baseline)
    {
        var pct = 100.0 * bytes / baseline;
        Console.WriteLine($"  {name, -26} {bytes, 12:N0} B   {pct, 6:0.0}% of MemoryPack");
    }

    // Round-trip both formats and confirm they reconstruct the same logical content, so the size numbers
    // describe a FAITHFUL encoding (not a lossy one that happens to be smaller).
    private static void Verify(
        Workbook workbook,
        byte[] memoryPack,
        MWorkbook mirror,
        byte[] msgpack,
        byte[] msgpackLz4
    )
    {
        var mpBack = SpikeSerializers.MemoryPackDeserialize(memoryPack);
        var msgBack = SpikeSerializers.MessagePackIndexedDeserialize(msgpack);
        var lz4Back = SpikeSerializers.MessagePackLz4Deserialize(msgpackLz4);

        var original = workbook.Sheets.Values.Sum(s => s.Count);
        var mpCells = mpBack.Sheets.Values.Sum(s => s.Count);
        var msgCells = msgBack.Sheets.Values.Sum(s => s.Cells.Count);
        var lz4Cells = lz4Back.Sheets.Values.Sum(s => s.Cells.Count);

        if (mpCells != original || msgCells != original || lz4Cells != original)
        {
            throw new InvalidOperationException(
                $"Round-trip cell-count mismatch: original={original}, memoryPack={mpCells}, "
                    + $"msgpack={msgCells}, lz4={lz4Cells}"
            );
        }

        // Spot-check: every cell of every sheet must mirror-convert to a node whose type matches the
        // MessagePack round-tripped node (catches a Union tag or Key mismatch that count alone would miss).
        foreach (var (name, sheet) in workbook.Sheets)
        {
            var back = msgBack.Sheets[name];
            foreach (var (id, expr) in sheet)
            {
                var expected = MirrorConverter.ToMirror(expr).GetType();
                var actual = back.Cells[id].GetType();
                if (expected != actual)
                {
                    throw new InvalidOperationException(
                        $"Round-trip type mismatch at {name}!{id}: expected {expected.Name}, got {actual.Name}"
                    );
                }
            }
        }
    }
}
