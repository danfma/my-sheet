using MessagePack;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// MessagePack mirror of Workbook. Mirrors the two SERIALIZED members of the production Workbook
// (Sheets, DefinedNames) — the [MemoryPackIgnore] runtime state (caches, locks, RNG) is not part of the
// wire format either way. The case-insensitive comparers (Sheets/DefinedNames/Sheet.Cells are
// OrdinalIgnoreCase in production) are restored in OnAfterDeserialize, the direct analogue of the
// production [MemoryPackOnDeserialized] RestoreComparers hook — this is the concrete answer to the plan's
// question about IMessagePackSerializationCallbackReceiver + comparer restoration.
[MessagePackObject]
public sealed partial class MWorkbook : IMessagePackSerializationCallbackReceiver
{
    [Key(0)]
    public Dictionary<string, MSheet> Sheets { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [Key(1)]
    public Dictionary<string, MExpr> DefinedNames { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public void OnBeforeSerialize() { }

    public void OnAfterDeserialize()
    {
        Sheets = new Dictionary<string, MSheet>(Sheets, StringComparer.OrdinalIgnoreCase);
        DefinedNames = new Dictionary<string, MExpr>(DefinedNames, StringComparer.OrdinalIgnoreCase);
    }
}

// MessagePack mirror of Sheet. Note: production Sheet.Cells is a case-SENSITIVE Dictionary (cell ids are
// already canonical), so no comparer restore is needed here — mirrored faithfully.
[MessagePackObject]
public sealed partial class MSheet
{
    [Key(0)]
    public string Name { get; set; } = "";

    [Key(1)]
    public int Index { get; set; }

    [Key(2)]
    public Dictionary<string, MExpr> Cells { get; set; } = new();
}
