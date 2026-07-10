using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// Speed axis for plans/messagepack-spike.md, question 2 (ShortRun + MemoryDiagnoser).
//   `dotnet run -c Release -- --filter *MessagePackBenchmarks*`
// Serialize + deserialize each payload with each format. The MessagePack path measures the mirror-graph
// (de)serialization only — mirror CONVERSION cost (Expression → MExpr) is excluded from the hot numbers and
// reported separately by the size harness, because a real migration would serialize the production types
// directly (no conversion step). The comparison is therefore format-vs-format, not converter overhead.
[MemoryDiagnoser]
[Config(typeof(Config))]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class MessagePackBenchmarks
{
    private class Config : ManualConfig
    {
        public Config() => AddJob(Job.ShortRun);
    }

    [Params(
        MessagePackPayloads.Size.Small,
        MessagePackPayloads.Size.Medium,
        MessagePackPayloads.Size.Large
    )]
    public MessagePackPayloads.Size Payload { get; set; }

    private Workbook _workbook = null!;
    private MWorkbook _mirror = null!;
    private byte[] _memoryPackBytes = null!;
    private byte[] _msgpackBytes = null!;
    private byte[] _msgpackLz4Bytes = null!;

    [GlobalSetup]
    public void Setup()
    {
        _workbook = MessagePackPayloads.Build(Payload);
        _mirror = MirrorConverter.ToMirror(_workbook);
        _memoryPackBytes = SpikeSerializers.MemoryPack(_workbook);
        _msgpackBytes = SpikeSerializers.MessagePackIndexed(_mirror);
        _msgpackLz4Bytes = SpikeSerializers.MessagePackLz4(_mirror);
    }

    // ── Serialize ───────────────────────────────────────────────────────────────────────────────────
    [Benchmark(Baseline = true), BenchmarkCategory("Serialize")]
    public byte[] MemoryPack_Serialize() => SpikeSerializers.MemoryPack(_workbook);

    [Benchmark, BenchmarkCategory("Serialize")]
    public byte[] MessagePack_Serialize() => SpikeSerializers.MessagePackIndexed(_mirror);

    [Benchmark, BenchmarkCategory("Serialize")]
    public byte[] MessagePackLz4_Serialize() => SpikeSerializers.MessagePackLz4(_mirror);

    // ── Deserialize ─────────────────────────────────────────────────────────────────────────────────
    [Benchmark(Baseline = true), BenchmarkCategory("Deserialize")]
    public Workbook MemoryPack_Deserialize() =>
        SpikeSerializers.MemoryPackDeserialize(_memoryPackBytes);

    [Benchmark, BenchmarkCategory("Deserialize")]
    public MWorkbook MessagePack_Deserialize() =>
        SpikeSerializers.MessagePackIndexedDeserialize(_msgpackBytes);

    [Benchmark, BenchmarkCategory("Deserialize")]
    public MWorkbook MessagePackLz4_Deserialize() =>
        SpikeSerializers.MessagePackLz4Deserialize(_msgpackLz4Bytes);
}
