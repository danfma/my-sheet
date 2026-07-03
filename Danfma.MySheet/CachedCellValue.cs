using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// A serializable surrogate for one memoized cell value, used only inside the warm-start value block (never
/// on <see cref="ComputedValue"/> itself, which stays a non-serializable value type). It carries the cell
/// address plus the flattened <see cref="ComputedValue"/> contents: the kind tag, the numeric payload
/// (Number, or Boolean as 0/1), the text payload, and the Error code.
///
/// <para><see cref="ComputedValueKind.Reference"/> is deliberately unrepresentable here — reference-typed
/// results are rare as a final cell value and cheap to reconstruct, so they are excluded from the snapshot
/// (their cells fall back to recomputation). <see cref="ComputedValueKind.Blank"/> IS carried, so an
/// explicitly-empty cached cell round-trips.</para>
/// </summary>
[MemoryPackable]
internal sealed partial record CachedCellValue(
    string SheetName,
    string CellId,
    ComputedValueKind Kind,
    double Number,
    string? Text,
    int ErrorCode
)
{
    /// <summary>
    /// Builds a surrogate from a cached value, or <c>null</c> when the value must NOT be persisted
    /// (currently only <see cref="ComputedValueKind.Reference"/>).
    /// </summary>
    public static CachedCellValue? TryFrom(string sheetName, string cellId, ComputedValue value)
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Blank:
                return new CachedCellValue(sheetName, cellId, value.Kind, 0d, null, 0);

            case ComputedValueKind.Number:
                return new CachedCellValue(
                    sheetName,
                    cellId,
                    value.Kind,
                    value.ToDouble(),
                    null,
                    0
                );

            case ComputedValueKind.Boolean:
                return new CachedCellValue(
                    sheetName,
                    cellId,
                    value.Kind,
                    value.ToBoolean() ? 1d : 0d,
                    null,
                    0
                );

            case ComputedValueKind.Text:
                return new CachedCellValue(sheetName, cellId, value.Kind, 0d, value.ToText(), 0);

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                return new CachedCellValue(sheetName, cellId, value.Kind, 0d, null, error.Code);

            case ComputedValueKind.Reference:
            default:
                return null;
        }
    }

    /// <summary>Rebuilds the <see cref="ComputedValue"/> this surrogate stands for.</summary>
    public ComputedValue ToComputedValue() =>
        Kind switch
        {
            ComputedValueKind.Number => ComputedValue.Number(Number),
            ComputedValueKind.Boolean => ComputedValue.Boolean(Number != 0d),
            ComputedValueKind.Text => ComputedValue.Text(Text),
            ComputedValueKind.Error => ComputedValue.Error(Error.FromCode(ErrorCode)),
            _ => ComputedValue.Blank,
        };
}
