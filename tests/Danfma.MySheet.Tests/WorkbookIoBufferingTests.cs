using System.IO.Compression;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

/// <summary>
/// M5 (audit-fixes-2026-07, "container warm em passe único" + opt-in Pipelines I/O): every write mechanism
/// (default whole-array, default single-pass streaming, opt-in pooled/Pipelines) must produce byte-identical
/// files for the same <see cref="WorkbookSaveOptions"/>, and every read mechanism (default byte[]/segmented,
/// opt-in pooled-sequence/Pipe) must load any of those files identically. Large-ish fixtures are used
/// deliberately so the raw model exceeds a single ~64KB chunk and genuinely exercises the streaming/pooled
/// writers' multi-segment paths (a tiny workbook would fit in one segment and never prove the loop works).
/// </summary>
public class WorkbookIoBufferingTests
{
    private static string TempPath() => Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

    // ~4000 formula cells: comfortably over 64KB of raw MemoryPack, so every writer/reader here must cross
    // multiple StreamBufferWriter/PipeWriter/PooledSequenceSegment chunk boundaries to round-trip correctly.
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

    private static Workbook SmallSample()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
        return workbook;
    }

    private static IEnumerable<(bool Warm, WorkbookCompression Compression)> OptionMatrix()
    {
        yield return (false, WorkbookCompression.None);
        yield return (true, WorkbookCompression.None);
        yield return (false, WorkbookCompression.Brotli);
        yield return (true, WorkbookCompression.Brotli);
    }

    // === M5b/M5c: every save mechanism produces byte-identical files =====================================

    [Test]
    public async Task Save_PipelinesMatchesPooled_ForEveryOptionCombination()
    {
        var workbook = LargeSample();
        workbook.ComputeAll();

        foreach (var (warm, compression) in OptionMatrix())
        {
            var pooledPath = TempPath();
            var pipelinesPath = TempPath();

            try
            {
                workbook.Save(
                    pooledPath,
                    new WorkbookSaveOptions
                    {
                        IncludeComputedValues = warm,
                        Compression = compression,
                        IoBuffering = WorkbookIoBuffering.Pooled,
                    }
                );
                workbook.Save(
                    pipelinesPath,
                    new WorkbookSaveOptions
                    {
                        IncludeComputedValues = warm,
                        Compression = compression,
                        IoBuffering = WorkbookIoBuffering.Pipelines,
                    }
                );

                await Assert
                    .That(File.ReadAllBytes(pipelinesPath))
                    .IsEquivalentTo(File.ReadAllBytes(pooledPath))
                    .Because($"warm={warm} compression={compression}");
            }
            finally
            {
                File.Delete(pooledPath);
                File.Delete(pipelinesPath);
            }
        }
    }

    [Test]
    public async Task SaveAsync_PipelinesMatchesPooled_ForEveryOptionCombination()
    {
        var workbook = LargeSample();
        workbook.ComputeAll();

        foreach (var (warm, compression) in OptionMatrix())
        {
            var pooledPath = TempPath();
            var pipelinesPath = TempPath();

            try
            {
                await workbook.SaveAsync(
                    pooledPath,
                    new WorkbookSaveOptions
                    {
                        IncludeComputedValues = warm,
                        Compression = compression,
                        IoBuffering = WorkbookIoBuffering.Pooled,
                    }
                );
                await workbook.SaveAsync(
                    pipelinesPath,
                    new WorkbookSaveOptions
                    {
                        IncludeComputedValues = warm,
                        Compression = compression,
                        IoBuffering = WorkbookIoBuffering.Pipelines,
                    }
                );

                await Assert
                    .That(File.ReadAllBytes(pipelinesPath))
                    .IsEquivalentTo(File.ReadAllBytes(pooledPath))
                    .Because($"warm={warm} compression={compression}");
            }
            finally
            {
                File.Delete(pooledPath);
                File.Delete(pipelinesPath);
            }
        }
    }

    [Test]
    public async Task SaveAsync_MatchesSync_ForEveryOptionCombination()
    {
        // Save (sync) and SaveAsync must produce the SAME bytes for the same options — the async rewrite
        // (single-pass streaming, M5b) must not diverge from the sync (whole-array) reference.
        var workbook = LargeSample();
        workbook.ComputeAll();

        foreach (var (warm, compression) in OptionMatrix())
        foreach (
            var buffering in new[] { WorkbookIoBuffering.Pooled, WorkbookIoBuffering.Pipelines }
        )
        {
            var syncPath = TempPath();
            var asyncPath = TempPath();

            try
            {
                var options = new WorkbookSaveOptions
                {
                    IncludeComputedValues = warm,
                    Compression = compression,
                    IoBuffering = buffering,
                };

                workbook.Save(syncPath, options);
                await workbook.SaveAsync(asyncPath, options);

                await Assert
                    .That(File.ReadAllBytes(asyncPath))
                    .IsEquivalentTo(File.ReadAllBytes(syncPath))
                    .Because($"warm={warm} compression={compression} buffering={buffering}");
            }
            finally
            {
                File.Delete(syncPath);
                File.Delete(asyncPath);
            }
        }
    }

    [Test]
    public async Task Save_ColdUncompressed_PipelinesIsByteIdenticalToParameterlessSave()
    {
        // The permanent byte-identity contract, restated for the IoBuffering axis.
        var workbook = SmallSample();
        workbook.GetCellValue("Sheet1", "A3"); // populate cache to prove flags (not emptiness) keep it cold

        var referencePath = TempPath();
        var pipelinesPath = TempPath();

        try
        {
            workbook.Save(referencePath);
            workbook.Save(
                pipelinesPath,
                new WorkbookSaveOptions { IoBuffering = WorkbookIoBuffering.Pipelines }
            );

            await Assert
                .That(File.ReadAllBytes(pipelinesPath))
                .IsEquivalentTo(File.ReadAllBytes(referencePath));
        }
        finally
        {
            File.Delete(referencePath);
            File.Delete(pipelinesPath);
        }
    }

    // === M5a/M5c: every load mechanism agrees on content ==================================================

    [Test]
    public async Task LoadAsync_RawFile_MatchesLoad()
    {
        var workbook = LargeSample();
        var path = TempPath();

        try
        {
            workbook.Save(path); // raw, no container — forces LoadAsync's new streaming branch
            var viaLoad = Workbook.Load(path);
            var viaLoadAsync = await Workbook.LoadAsync(path);

            await Assert
                .That(viaLoadAsync.GetCellValue("Data", "B4000").ToDouble())
                .IsEqualTo(viaLoad.GetCellValue("Data", "B4000").ToDouble());
            await Assert
                .That(viaLoadAsync.GetCellValue("Data", "B1").ToDouble())
                .IsEqualTo(viaLoad.GetCellValue("Data", "B1").ToDouble());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_PipelinesRawFile_MatchesPooled()
    {
        var workbook = LargeSample();
        var path = TempPath();

        try
        {
            workbook.Save(path); // raw file

            var pooled = Workbook.Load(path, new WorkbookLoadOptions());
            var pipelines = Workbook.Load(
                path,
                new WorkbookLoadOptions { IoBuffering = WorkbookIoBuffering.Pipelines }
            );

            await Assert
                .That(pipelines.GetCellValue("Data", "B4000").ToDouble())
                .IsEqualTo(pooled.GetCellValue("Data", "B4000").ToDouble());
            await Assert
                .That(pipelines.GetCellValue("Data", "B2500").ToDouble())
                .IsEqualTo(pooled.GetCellValue("Data", "B2500").ToDouble());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task LoadAsync_PipelinesRawFile_MatchesPooled()
    {
        var workbook = LargeSample();
        var path = TempPath();

        try
        {
            workbook.Save(path); // raw file

            var pooled = await Workbook.LoadAsync(path, new WorkbookLoadOptions());
            var pipelines = await Workbook.LoadAsync(
                path,
                new WorkbookLoadOptions { IoBuffering = WorkbookIoBuffering.Pipelines }
            );

            await Assert
                .That(pipelines.GetCellValue("Data", "B4000").ToDouble())
                .IsEqualTo(pooled.GetCellValue("Data", "B4000").ToDouble());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_PipelinesContainerFile_FallsBackAndMatchesPooled()
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
                }
            );

            var pooled = Workbook.Load(path, new WorkbookLoadOptions());
            var pipelines = Workbook.Load(
                path,
                new WorkbookLoadOptions { IoBuffering = WorkbookIoBuffering.Pipelines }
            );

            await Assert
                .That(pipelines.GetCellValue("Data", "B4000").ToDouble())
                .IsEqualTo(pooled.GetCellValue("Data", "B4000").ToDouble());
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Cross product: any writer's output loads correctly via any reader ================================

    [Test]
    public async Task AnySaveMechanism_LoadsCorrectly_ViaAnyLoadMechanism()
    {
        var workbook = LargeSample();
        workbook.ComputeAll();

        foreach (var (warm, compression) in OptionMatrix())
        foreach (
            var saveBuffering in new[] { WorkbookIoBuffering.Pooled, WorkbookIoBuffering.Pipelines }
        )
        foreach (
            var loadBuffering in new[] { WorkbookIoBuffering.Pooled, WorkbookIoBuffering.Pipelines }
        )
        {
            var path = TempPath();

            try
            {
                workbook.Save(
                    path,
                    new WorkbookSaveOptions
                    {
                        IncludeComputedValues = warm,
                        Compression = compression,
                        IoBuffering = saveBuffering,
                    }
                );

                var loaded = Workbook.Load(
                    path,
                    new WorkbookLoadOptions { IoBuffering = loadBuffering }
                );

                await Assert
                    .That(loaded.GetCellValue("Data", "B4000").ToDouble())
                    .IsEqualTo(8001.0)
                    .Because(
                        $"warm={warm} compression={compression} save={saveBuffering} load={loadBuffering}"
                    );
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    // === Brotli chunking spike (documents the finding that shaped M5b/M5c's compressed branches) ==========

    [Test]
    public async Task BrotliStream_OutputDependsOnWriteCallBoundaries_NotJustContent()
    {
        // This is the M5b/M5c "30-minute spike" the plan called for, kept as a permanent regression guard:
        // if a future .NET/BCL change made Brotli's output insensitive to Write-call chunking, the
        // WriteContainerWhole fallback in WriteContainerPooled/StreamAsync/PipeAsync could be simplified —
        // but today, splitting the SAME bytes across many small Write calls measurably changes the
        // compressed output, which is exactly why the compressed branches always fall back to two
        // whole-buffer Write calls instead of streaming through StreamBufferWriter/PipeWriter/SerializeAsync.
        var data = new byte[500_000];
        new Random(42).NextBytes(data);

        byte[] Compress(Action<BrotliStream> write)
        {
            using var output = new MemoryStream();
            using (var brotli = new BrotliStream(output, CompressionLevel.Fastest, leaveOpen: true))
            {
                write(brotli);
            }
            return output.ToArray();
        }

        var oneShot = Compress(b => b.Write(data));
        var chunked = Compress(b =>
        {
            var offset = 0;
            while (offset < data.Length)
            {
                var n = Math.Min(4096, data.Length - offset);
                b.Write(data, offset, n);
                offset += n;
            }
        });

        await Assert.That(oneShot.AsSpan().SequenceEqual(chunked)).IsFalse();
    }
}
