using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// A serializable surrogate for one memoized cell value, used only inside the warm-start value block (never
/// on <see cref="ComputedValue"/> itself, which stays a non-serializable value type). It carries the cell
/// address plus the flattened <see cref="ComputedValue"/> contents: the kind tag, the numeric payload
/// (Number, or Boolean as 0/1), the text payload, and the Error code.
///
/// <para><see cref="ComputedValueKind.Reference"/> is deliberately unrepresentable here, but it is no longer
/// a LIVE exclusion: since the cell boundary applies Excel's implicit intersection (see
/// <c>Workbook.EvaluateCell</c>), a value the store holds can no longer BE a reference, so no cached cell
/// takes that arm any more. It stays as defence-in-depth over the PUBLIC surface —
/// <see cref="ComputedValue.Reference"/> is public API, so a caller can hand
/// <see cref="TryFrom"/> a reference-kind value directly, and it must still refuse to persist it rather than
/// flatten it into a wrong scalar. <see cref="ComputedValueKind.Blank"/> IS carried, so an explicitly-empty
/// cached cell round-trips.</para>
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
    /// Builds a surrogate from a cached value, or <c>null</c> when the value must NOT be persisted (only
    /// <see cref="ComputedValueKind.Reference"/>, which the memoized store itself can no longer produce —
    /// see the type remarks).
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

            // Unreachable from the snapshot path (Workbook.SnapshotComputedValues feeds this straight from the
            // memoized store, and the cell boundary intersects every reference away before it is stored), but
            // NOT dead: TryFrom takes a public ComputedValue, so a host — or a test, e.g.
            // WarmStartSaveLoadTests.Surrogate_ExcludesReference_ButKeepsBlank — can still construct one.
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
