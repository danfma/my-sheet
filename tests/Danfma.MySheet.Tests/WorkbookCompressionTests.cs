using System.Buffers.Binary;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Optional Brotli compression (<see cref="WorkbookSaveOptions.Compression"/>) on save. Compression is
/// orthogonal to <see cref="WorkbookSaveOptions.IncludeComputedValues"/>: cold (model-only) and warm
/// (model + values) saves can each be compressed. Compressed files are wrapped in the <c>MSWM</c> container
/// (version 2) so <see cref="Workbook.Load(string)"/> detects and transparently decompresses them; the cold,
/// uncompressed default remains byte-identical to <see cref="Workbook.Save(string)"/>.
/// </summary>
public class WorkbookCompressionTests
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

    // A workbook big enough that entropy coding clearly beats the raw MemoryPack layout.
    private static Workbook LargeSample(int rows = 5_000)
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

    // === Default stays byte-identical (permanent regression guard, restated for the compression axis) ====

    [Test]
    public async Task Save_CompressionNoneColdModel_IsByteIdenticalToDefault()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3"); // populate cache to prove flags, not emptiness, keep it cold

        var defaultPath = TempPath();
        var optionPath = TempPath();

        try
        {
            workbook.Save(defaultPath);
            workbook.Save(
                optionPath,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = false,
                    Compression = WorkbookCompression.None,
                }
            );

            await Assert
                .That(File.ReadAllBytes(optionPath))
                .IsEquivalentTo(File.ReadAllBytes(defaultPath));
        }
        finally
        {
            File.Delete(defaultPath);
            File.Delete(optionPath);
        }
    }

    // === Cold compressed: container v2, round-trips, and is smaller than raw ==============================

    [Test]
    public async Task ColdCompressed_IsContainerV2_AndRoundTrips()
    {
        var workbook = Sample();
        var path = TempPath();

        try
        {
            workbook.Save(
                path,
                new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli }
            );

            var bytes = File.ReadAllBytes(path);
            await Assert.That(bytes.AsSpan(0, 4).SequenceEqual(Magic)).IsTrue();
            await Assert.That(bytes[4]).IsEqualTo((byte)2); // Brotli version

            var modelLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(5, 4));
            var rawModel = MemoryPackSerializer.Serialize(workbook);
            await Assert.That(modelLength).IsEqualTo(rawModel.Length); // uncompressed model length

            var loaded = Workbook.Load(path);
            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task ColdCompressed_IsSubstantiallySmallerThanRaw()
    {
        var workbook = LargeSample();
        var rawPath = TempPath();
        var compressedPath = TempPath();

        try
        {
            workbook.Save(rawPath);
            workbook.Save(
                compressedPath,
                new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli }
            );

            var raw = new FileInfo(rawPath).Length;
            var compressed = new FileInfo(compressedPath).Length;

            await Assert.That(compressed).IsLessThan(raw);
            // Rough guard: a 5k-cell workbook compresses to well under half its raw size.
            await Assert.That(compressed).IsLessThan(raw / 2);
        }
        finally
        {
            File.Delete(rawPath);
            File.Delete(compressedPath);
        }
    }

    // === Warm compressed: identical values AND zero recomputation ========================================

    [Test]
    public async Task WarmCompressed_RoundTripsEveryKind()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["N1"] = new NumberValue(42);
        sheet["B1"] = ExpressionParser.Parse("=TRUE", sheet);
        sheet["T1"] = ExpressionParser.Parse("=\"hello\"", sheet);
        sheet["E1"] = ExpressionParser.Parse("=1/0", sheet);

        workbook.GetCellValue("Sheet1", "N1");
        workbook.GetCellValue("Sheet1", "B1");
        workbook.GetCellValue("Sheet1", "T1");
        workbook.GetCellValue("Sheet1", "E1");
        workbook.GetCellValue("Sheet1", "K1"); // Blank

        var path = TempPath();

        try
        {
            workbook.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = true,
                    Compression = WorkbookCompression.Brotli,
                }
            );

            var bytes = File.ReadAllBytes(path);
            await Assert.That(bytes[4]).IsEqualTo((byte)2);

            var loaded = Workbook.Load(path);
            await Assert.That(loaded.GetCellValue("Sheet1", "N1").ToDouble()).IsEqualTo(42.0);
            await Assert.That(loaded.GetCellValue("Sheet1", "B1").ToBoolean()).IsTrue();
            await Assert.That(loaded.GetCellValue("Sheet1", "T1").ToText()).IsEqualTo("hello");

            var error = loaded.GetCellValue("Sheet1", "E1");
            await Assert.That(error.Kind).IsEqualTo(ComputedValueKind.Error);
            await Assert.That(error.TryGetError(out var code)).IsTrue();
            await Assert.That(code).IsEqualTo(Error.DivZero);

            await Assert
                .That(loaded.GetCellValue("Sheet1", "K1").Kind)
                .IsEqualTo(ComputedValueKind.Blank);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task WarmCompressed_ServesCachedCell_WithoutRecomputing()
    {
        var ticks = 0;
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=TICK()", sheet);
        workbook.RegisterFunction("TICK", (_, _) => ++ticks);

        workbook.GetCellValue("Sheet1", "A1"); // only A1 cached
        await Assert.That(ticks).IsEqualTo(1);

        var path = TempPath();

        try
        {
            workbook.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = true,
                    Compression = WorkbookCompression.Brotli,
                }
            );

            var afterTicks = 0;
            var loaded = Workbook.Load(path);
            loaded.RegisterFunction("TICK", (_, _) => ++afterTicks);

            // Warm hit: no evaluation, even through the compressed container.
            await Assert.That(loaded.GetCellValue("Sheet1", "A1").ToDouble()).IsEqualTo(1.0);
            await Assert.That(afterTicks).IsEqualTo(0);

            // A2 was never cached → recomputes.
            await Assert.That(loaded.GetCellValue("Sheet1", "A2").ToDouble()).IsEqualTo(1.0);
            await Assert.That(afterTicks).IsEqualTo(1);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task WarmCompressed_ExcludesVolatiles_WhichRecomputeAfterLoad()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["V1"] = ExpressionParser.Parse("=NOW()", sheet);
        sheet["S1"] = new NumberValue(7);

        var early = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Local);
        workbook.TimeProvider = new FixedLocalClock(early);

        var stampBefore = workbook.GetCellValue("Sheet1", "V1").ToDouble();
        workbook.GetCellValue("Sheet1", "S1");

        var path = TempPath();

        try
        {
            workbook.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = true,
                    Compression = WorkbookCompression.Brotli,
                }
            );

            var loaded = Workbook.Load(path);
            var late = new DateTime(2020, 6, 15, 12, 0, 0, DateTimeKind.Local);
            loaded.TimeProvider = new FixedLocalClock(late);

            await Assert.That(loaded.GetCellValue("Sheet1", "S1").ToDouble()).IsEqualTo(7.0);

            var stampAfter = loaded.GetCellValue("Sheet1", "V1").ToDouble();
            await Assert.That(stampAfter).IsGreaterThan(stampBefore);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Async parity ====================================================================================

    [Test]
    public async Task SaveAsyncCompressed_AndLoadAsync_RoundTrips()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3");

        var path = TempPath();

        try
        {
            await workbook.SaveAsync(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = true,
                    Compression = WorkbookCompression.Brotli,
                }
            );

            var loaded = await Workbook.LoadAsync(path);
            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Corrupt container is reported, not silently mis-read ============================================

    [Test]
    public async Task Load_CorruptBrotliContainer_Throws()
    {
        var path = TempPath();

        try
        {
            // Valid header (magic + version 2 + a plausible modelLength) but a garbage Brotli body.
            var bytes = new byte[13];
            Magic.CopyTo(bytes);
            bytes[4] = 2;
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(5, 4), 4);
            bytes[9] = 0xFF;
            bytes[10] = 0xFF;
            bytes[11] = 0xFF;
            bytes[12] = 0xFF;
            File.WriteAllBytes(path, bytes);

            await Assert
                .That(() => Workbook.Load(path))
                .Throws<InvalidDataException>();
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Compression level is a write-time knob; any level round-trips through Load =====================

    [Test]
    [Arguments(System.IO.Compression.CompressionLevel.Fastest)]
    [Arguments(System.IO.Compression.CompressionLevel.Optimal)]
    [Arguments(System.IO.Compression.CompressionLevel.SmallestSize)]
    [Arguments(System.IO.Compression.CompressionLevel.NoCompression)]
    public async Task Save_BrotliAtAnyLevel_RoundTripsThroughLoad(
        System.IO.Compression.CompressionLevel level
    )
    {
        var workbook = LargeSample();
        workbook.ComputeAll();
        var path = TempPath();
        try
        {
            workbook.Save(
                path,
                new WorkbookSaveOptions
                {
                    IncludeComputedValues = true,
                    Compression = WorkbookCompression.Brotli,
                    CompressionLevel = level,
                }
            );

            var loaded = Workbook.Load(path);
            await Assert
                .That(loaded.GetCellValue("Data", "B5000").AsObject())
                .IsEqualTo(workbook.GetCellValue("Data", "B5000").AsObject());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Save_Fastest_IsLargerThanOptimal_ProvingTheLevelIsHonored()
    {
        var workbook = LargeSample();
        var fastest = TempPath();
        var optimal = TempPath();
        try
        {
            workbook.Save(fastest, new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli, CompressionLevel = System.IO.Compression.CompressionLevel.Fastest });
            workbook.Save(optimal, new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli, CompressionLevel = System.IO.Compression.CompressionLevel.Optimal });

            await Assert
                .That(new FileInfo(fastest).Length)
                .IsGreaterThan(new FileInfo(optimal).Length);
        }
        finally
        {
            File.Delete(fastest);
            File.Delete(optimal);
        }
    }

    private sealed class FixedLocalClock(DateTime localNow) : TimeProvider
    {
        private readonly DateTimeOffset _now = new(
            DateTime.SpecifyKind(localNow, DateTimeKind.Unspecified),
            TimeSpan.Zero
        );

        public override DateTimeOffset GetUtcNow() => _now;

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
