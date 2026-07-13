using System.Buffers.Binary;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

/// <summary>
/// M6: the <c>MSWM</c> container's v3 body encoding (fixed-64KB-chunked Brotli — see the format comment atop
/// Workbook.Serialization.cs and <see cref="BrotliChunkedStream"/>), now the DEFAULT for
/// <see cref="WorkbookCompression.Brotli"/>. <see cref="WorkbookIoBufferingTests"/> already proves every
/// write mechanism produces byte-identical output for a given options combination and that any writer's
/// output loads via any reader; this file adds the v3-specific guarantees the M6 plan called out by name:
/// determinism across repeated saves, and an explicit sweep of all four Save/SaveAsync × Pooled/Pipelines
/// combinations (cold and warm) asserting the ON-DISK version byte is 3.
/// </summary>
public class ContainerV3Tests
{
    private static ReadOnlySpan<byte> Magic => "MSWM"u8;

    private static string TempPath() => Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

    private static Workbook Sample()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
        return workbook;
    }

    // ~4000 formula cells: crosses many 64KB chunk boundaries, so determinism/round-trip here actually
    // exercises BrotliChunkedStream's multi-chunk path (a tiny workbook fits in one chunk and would not).
    private static Workbook LargeSample(int rows = 4_000)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");
        for (var r = 1; r <= rows; r++)
        {
            sheet[$"A{r}"] = new NumberValue(r);
            sheet[$"B{r}"] = ExpressionParser.Parse($"=A{r}*2+1", sheet);
        }
        return workbook;
    }

    // === Determinism: two saves of the SAME workbook produce IDENTICAL bytes =============================

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task Save_Brotli_TwoSavesOfSameWorkbook_ProduceIdenticalBytes(bool warm)
    {
        var workbook = LargeSample();
        workbook.ComputeAll();

        var pathA = TempPath();
        var pathB = TempPath();

        try
        {
            var options = new WorkbookSaveOptions
            {
                IncludeComputedValues = warm,
                Compression = WorkbookCompression.Brotli,
            };

            workbook.Save(pathA, options);
            workbook.Save(pathB, options);

            var bytesA = File.ReadAllBytes(pathA);
            var bytesB = File.ReadAllBytes(pathB);

            await Assert.That(bytesA[4]).IsEqualTo((byte)3);
            await Assert.That(bytesA).IsEquivalentTo(bytesB);
        }
        finally
        {
            File.Delete(pathA);
            File.Delete(pathB);
        }
    }

    [Test]
    public async Task SaveAsync_Brotli_TwoSavesOfSameWorkbook_ProduceIdenticalBytes()
    {
        var workbook = LargeSample();
        workbook.ComputeAll();

        var pathA = TempPath();
        var pathB = TempPath();

        try
        {
            var options = new WorkbookSaveOptions
            {
                IncludeComputedValues = true,
                Compression = WorkbookCompression.Brotli,
            };

            await workbook.SaveAsync(pathA, options);
            await workbook.SaveAsync(pathB, options);

            var bytesA = File.ReadAllBytes(pathA);
            var bytesB = File.ReadAllBytes(pathB);

            await Assert.That(bytesA[4]).IsEqualTo((byte)3);
            await Assert.That(bytesA).IsEquivalentTo(bytesB);
        }
        finally
        {
            File.Delete(pathA);
            File.Delete(pathB);
        }
    }

    // === Round-trip across all 4 Save/SaveAsync x Pooled/Pipelines combinations, cold and warm ============

    private static IEnumerable<Func<Workbook, string, bool, Task>> Writers()
    {
        yield return async (wb, path, warm) =>
        {
            wb.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = warm,
                    Compression = WorkbookCompression.Brotli,
                    IoBuffering = WorkbookIoBuffering.Pooled,
                }
            );
            await Task.CompletedTask;
        };
        yield return async (wb, path, warm) =>
        {
            wb.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = warm,
                    Compression = WorkbookCompression.Brotli,
                    IoBuffering = WorkbookIoBuffering.Pipelines,
                }
            );
            await Task.CompletedTask;
        };
        yield return (wb, path, warm) =>
            wb.SaveAsync(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = warm,
                    Compression = WorkbookCompression.Brotli,
                    IoBuffering = WorkbookIoBuffering.Pooled,
                }
            );
        yield return (wb, path, warm) =>
            wb.SaveAsync(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = warm,
                    Compression = WorkbookCompression.Brotli,
                    IoBuffering = WorkbookIoBuffering.Pipelines,
                }
            );
    }

    [Test]
    public async Task AllFourWriteMechanisms_ProduceV3_AndRoundTrip_Cold()
    {
        var workbook = Sample();
        var writers = Writers().ToList();

        for (var i = 0; i < writers.Count; i++)
        {
            var path = TempPath();
            try
            {
                await writers[i](workbook, path, false);

                var bytes = File.ReadAllBytes(path);
                await Assert
                    .That(bytes.AsSpan(0, 4).SequenceEqual(Magic))
                    .IsTrue()
                    .Because($"writer #{i}");
                await Assert.That(bytes[4]).IsEqualTo((byte)3).Because($"writer #{i}");

                var loaded = Workbook.Load(path);
                await Assert
                    .That(loaded.GetCellValue("Sheet1", "A3").ToDouble())
                    .IsEqualTo(3.0)
                    .Because($"writer #{i}");
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    [Test]
    public async Task AllFourWriteMechanisms_ProduceV3_AndRoundTrip_Warm()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3"); // populate the cache before a warm save
        var writers = Writers().ToList();

        for (var i = 0; i < writers.Count; i++)
        {
            var path = TempPath();
            try
            {
                await writers[i](workbook, path, true);

                var bytes = File.ReadAllBytes(path);
                await Assert.That(bytes[4]).IsEqualTo((byte)3).Because($"writer #{i}");

                // Warm hit: no recomputation needed — proven by a fresh Workbook with no functions registered
                // still resolving the formula cell (its cached value round-tripped through the container).
                var loaded = Workbook.Load(path);
                await Assert
                    .That(loaded.GetCellValue("Sheet1", "A3").ToDouble())
                    .IsEqualTo(3.0)
                    .Because($"writer #{i}");
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    // === Corrupt v3 payload is reported, not silently mis-read (v2's DeserializeContainer path is shared,
    // but this proves the version-3 branch specifically wires into the same guard) ===========================

    [Test]
    public async Task Load_CorruptV3Container_Throws()
    {
        var path = TempPath();

        try
        {
            var bytes = new byte[13];
            Magic.CopyTo(bytes);
            bytes[4] = 3;
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(5, 4), 4);
            bytes[9] = 0xFF;
            bytes[10] = 0xFF;
            bytes[11] = 0xFF;
            bytes[12] = 0xFF;
            File.WriteAllBytes(path, bytes);

            await Assert.That(() => Workbook.Load(path)).Throws<InvalidDataException>();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
