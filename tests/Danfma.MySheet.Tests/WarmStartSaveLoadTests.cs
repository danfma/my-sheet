using System.Buffers.Binary;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Warm-start (opt-in <see cref="WorkbookSaveOptions.IncludeComputedValues"/>) save/load contract. The cold
/// path stays byte-identical to the historical raw MemoryPack format (a permanent regression guard); the warm
/// path wraps the SAME model bytes in an <c>MSWM</c> container plus a value block so a load starts warm.
/// </summary>
public class WarmStartSaveLoadTests
{
    private static ReadOnlySpan<byte> Magic => "MSWM"u8;

    private static Workbook Sample()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
        return workbook;
    }

    private static string TempPath() => Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

    // === Cold-path byte-identity (permanent regression contract) =========================================

    [Test]
    public async Task Save_Default_IsByteIdenticalToRawMemoryPack()
    {
        var workbook = Sample();
        var path = TempPath();

        try
        {
            workbook.Save(path);
            var written = File.ReadAllBytes(path);
            var expected = MemoryPackSerializer.Serialize(workbook);

            await Assert.That(written).IsEquivalentTo(expected);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Save_WithIncludeComputedValuesFalse_IsByteIdenticalToDefault()
    {
        var workbook = Sample();
        // Populate the cache to prove the flag — not the empty cache — is what keeps the file cold.
        workbook.GetCellValue("Sheet1", "A3");

        var coldPath = TempPath();
        var optionPath = TempPath();

        try
        {
            workbook.Save(coldPath);
            workbook.Save(optionPath, new WorkbookSaveOptions { IncludeComputedValues = false });

            await Assert
                .That(File.ReadAllBytes(optionPath))
                .IsEquivalentTo(File.ReadAllBytes(coldPath));
        }
        finally
        {
            File.Delete(coldPath);
            File.Delete(optionPath);
        }
    }

    [Test]
    public async Task Save_Warm_ModelBlockEqualsColdBytes()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3");

        var coldPath = TempPath();
        var warmPath = TempPath();

        try
        {
            workbook.Save(coldPath);
            workbook.Save(warmPath, new WorkbookSaveOptions { IncludeComputedValues = true });

            var cold = File.ReadAllBytes(coldPath);
            var warm = File.ReadAllBytes(warmPath);

            // Container layout: magic(4) + version(1) + modelLength(4) + model + values.
            await Assert.That(warm.AsSpan(0, 4).SequenceEqual(Magic)).IsTrue();
            await Assert.That(warm[4]).IsEqualTo((byte)1);

            var modelLength = BinaryPrimitives.ReadInt32LittleEndian(warm.AsSpan(5, 4));
            await Assert.That(modelLength).IsEqualTo(cold.Length);

            var modelBlock = warm.AsSpan(9, modelLength).ToArray();
            await Assert.That(modelBlock).IsEquivalentTo(cold);
        }
        finally
        {
            File.Delete(coldPath);
            File.Delete(warmPath);
        }
    }

    // === Warm round-trip of every persisted kind =========================================================

    [Test]
    public async Task WarmRoundTrip_PreservesEveryKind()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["N1"] = new NumberValue(42);
        sheet["B1"] = ExpressionParser.Parse("=TRUE", sheet);
        sheet["T1"] = ExpressionParser.Parse("=\"hello\"", sheet);
        sheet["E1"] = ExpressionParser.Parse("=1/0", sheet);
        // K1 stays absent → an empty cell caches Blank.

        // Populate the cache (only GetCellValue writes _cache; Evaluate does not).
        workbook.GetCellValue("Sheet1", "N1");
        workbook.GetCellValue("Sheet1", "B1");
        workbook.GetCellValue("Sheet1", "T1");
        workbook.GetCellValue("Sheet1", "E1");
        workbook.GetCellValue("Sheet1", "K1"); // Blank

        var path = TempPath();

        try
        {
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });
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

    // === Zero-recomputation proof (a counting custom function) ===========================================

    [Test]
    public async Task WarmLoad_ServesCachedCell_WithoutRecomputing()
    {
        var ticks = 0;
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=TICK()", sheet);
        workbook.RegisterFunction("TICK", (_, _) => ++ticks);

        // Evaluate ONLY A1 before saving; A2 stays uncached.
        var before = workbook.GetCellValue("Sheet1", "A1").ToDouble();
        await Assert.That(before).IsEqualTo(1.0);
        await Assert.That(ticks).IsEqualTo(1);

        var path = TempPath();

        try
        {
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });

            // Fresh instance: A1 must come from the warm cache without ever invoking TICK.
            var afterTicks = 0;
            var loaded = Workbook.Load(path);
            loaded.RegisterFunction("TICK", (_, _) => ++afterTicks);

            var cached = loaded.GetCellValue("Sheet1", "A1").ToDouble();
            await Assert.That(cached).IsEqualTo(1.0);
            await Assert.That(afterTicks).IsEqualTo(0); // proof: no evaluation on the warm hit

            // A2 was never cached → it recomputes and DOES need the re-registered function.
            var recomputed = loaded.GetCellValue("Sheet1", "A2").ToDouble();
            await Assert.That(recomputed).IsEqualTo(1.0);
            await Assert.That(afterTicks).IsEqualTo(1);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task WarmLoad_UncachedCustomCall_NeedsReRegistration()
    {
        var ticks = 0;
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet);
        sheet["A2"] = ExpressionParser.Parse("=TICK()", sheet);
        workbook.RegisterFunction("TICK", (_, _) => ++ticks);
        workbook.GetCellValue("Sheet1", "A1"); // only A1 cached

        var path = TempPath();

        try
        {
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });

            var loaded = Workbook.Load(path);
            // No re-registration: the uncached A2 resolves to #NAME?.
            var uncached = loaded.GetCellValue("Sheet1", "A2");
            await Assert.That(uncached.Kind).IsEqualTo(ComputedValueKind.Error);
            await Assert.That(uncached.TryGetError(out var code)).IsTrue();
            await Assert.That(code).IsEqualTo(Error.Name);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Volatiles are excluded and recompute post-load ==================================================

    [Test]
    public async Task WarmSave_ExcludesVolatiles_WhichRecomputeAfterLoad()
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
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });

            var loaded = Workbook.Load(path);
            var late = new DateTime(2020, 6, 15, 12, 0, 0, DateTimeKind.Local);
            loaded.TimeProvider = new FixedLocalClock(late);

            // The stable cell came from the warm cache; the volatile one was excluded and re-samples the clock.
            await Assert.That(loaded.GetCellValue("Sheet1", "S1").ToDouble()).IsEqualTo(7.0);

            var stampAfter = loaded.GetCellValue("Sheet1", "V1").ToDouble();
            await Assert.That(stampAfter).IsNotEqualTo(stampBefore);
            await Assert.That(stampAfter).IsGreaterThan(stampBefore);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Warm + edit + InvalidateCache still recomputes ==================================================

    [Test]
    public async Task WarmLoad_ThenEditAndInvalidate_Recomputes()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3"); // caches 3

        var path = TempPath();

        try
        {
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });

            var loaded = Workbook.Load(path);
            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);

            loaded["Sheet1"]["A1"] = new NumberValue(10);
            loaded.InvalidateCache();

            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(12.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Empty cache → warm is a container with an empty value block =====================================

    [Test]
    public async Task WarmSave_WithEmptyCache_IsContainerWithColdModel()
    {
        var workbook = Sample(); // never read → empty cache

        var coldPath = TempPath();
        var warmPath = TempPath();

        try
        {
            workbook.Save(coldPath);
            workbook.Save(warmPath, new WorkbookSaveOptions { IncludeComputedValues = true });

            var cold = File.ReadAllBytes(coldPath);
            var warm = File.ReadAllBytes(warmPath);

            // It IS a container (magic present), not the raw cold file.
            await Assert.That(warm.AsSpan(0, 4).SequenceEqual(Magic)).IsTrue();

            var modelLength = BinaryPrimitives.ReadInt32LittleEndian(warm.AsSpan(5, 4));
            await Assert.That(modelLength).IsEqualTo(cold.Length);
            await Assert.That(warm.AsSpan(9, modelLength).ToArray()).IsEquivalentTo(cold);

            // And it still loads and recomputes correctly.
            var loaded = Workbook.Load(warmPath);
            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(coldPath);
            File.Delete(warmPath);
        }
    }

    // === Cold/legacy raw files still load through the sniff ==============================================

    [Test]
    public async Task Load_ColdFile_StillLoads()
    {
        var workbook = Sample();
        var path = TempPath();

        try
        {
            workbook.Save(path); // raw, no magic
            var loaded = Workbook.Load(path);

            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task SaveAsyncWarm_AndLoadAsync_RoundTrips()
    {
        var workbook = Sample();
        workbook.GetCellValue("Sheet1", "A3");

        var path = TempPath();

        try
        {
            await workbook.SaveAsync(
                path,
                new WorkbookSaveOptions { IncludeComputedValues = true }
            );
            var loaded = await Workbook.LoadAsync(path);

            await Assert.That(loaded.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Surrogate excludes Reference, includes Blank ====================================================

    [Test]
    public async Task Surrogate_ExcludesReference_ButKeepsBlank()
    {
        var reference = ComputedValue.Reference(new CellReference("A1", "Sheet1"));
        await Assert.That(CachedCellValue.TryFrom("Sheet1", "X1", reference)).IsNull();

        var blank = CachedCellValue.TryFrom("Sheet1", "Y1", ComputedValue.Blank);
        await Assert.That(blank).IsNotNull();
        await Assert.That(blank!.ToComputedValue().Kind).IsEqualTo(ComputedValueKind.Blank);
    }

    // A TimeProvider pinned to a fixed local instant, so the NOW() epoch is deterministic across a save/load.
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
