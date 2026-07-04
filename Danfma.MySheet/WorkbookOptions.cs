namespace Danfma.MySheet;

/// <summary>
/// Immutable runtime options captured by a <see cref="Workbook"/> at construction (see
/// <see cref="Workbook(WorkbookOptions)"/>). Configuration, NOT document state: never serialized, so the
/// wire schema is untouched and a deserialized workbook simply falls back to the defaults. Currently the
/// single tunable section is the dense value store (<see cref="ValueStore"/>); the wrapper exists so future
/// runtime knobs slot in without churning the public constructor.
/// </summary>
public sealed class WorkbookOptions
{
    /// <summary>The default options (every section at its default). Used by the parameterless
    /// <see cref="Workbook()"/> and by a workbook restored from disk.</summary>
    public static WorkbookOptions Default { get; } = new();

    /// <summary>Tuning for the dense paged value store. Defaults to <see cref="ValueStoreOptions.Default"/>.</summary>
    public ValueStoreOptions ValueStore { get; init; } = ValueStoreOptions.Default;
}

/// <summary>
/// Immutable tuning for the dense paged value store (see <c>SheetValueStore</c>). These shape the store's
/// memory/scan trade-off; the defaults are the values the store shipped with and are tuned for the K1
/// compute workload. Runtime configuration only — captured at <see cref="Workbook"/> construction and NEVER
/// serialized. The two sizes MUST be powers of two because the hot path addresses cells with shift/mask (the
/// shift is derived from the size); a non-power-of-two or out-of-range value throws <see cref="ArgumentException"/>
/// at <see cref="Workbook(WorkbookOptions)"/>.
/// </summary>
public sealed class ValueStoreOptions
{
    /// <summary>Default row-page size (rows per page). One page is <c>ComputedValue[1024]</c> ≈ 24.6 KB.</summary>
    public const int DefaultRowPageSize = 1024;

    /// <summary>Default column-group size (columns per group in the two-level column directory).</summary>
    public const int DefaultColumnGroupSize = 64;

    /// <summary>Default sparsity-guard warm-up: pages allocated before density is judged.</summary>
    public const int DefaultSparsityWarmupPages = 64;

    /// <summary>Default sparsity-guard floor: minimum average cells per page to keep allocating pages.</summary>
    public const int DefaultSparsityMinCellsPerPage = 4;

    // Accepted ranges. RowPageSize's floor is 64 because the per-page presence bitmap is 64-bit words (a page
    // must hold at least one full word); its ceiling keeps a single page under ~1.5 MB. ColumnGroupSize's
    // range keeps the group directory a small pointer array. Documented on each property below.
    internal const int MinRowPageSize = 64;
    internal const int MaxRowPageSize = 65_536;
    internal const int MinColumnGroupSize = 8;
    internal const int MaxColumnGroupSize = 4_096;
    internal const int MaxSparsityWarmupPages = 1_000_000;

    /// <summary>The default options. Referenced by <see cref="WorkbookOptions.Default"/> and used whenever a
    /// workbook is created without options or restored from disk.</summary>
    public static ValueStoreOptions Default { get; } = new();

    /// <summary>
    /// Rows per page — the granularity at which the store allocates dense value storage. Must be a power of
    /// two in <c>[64, 65536]</c>. Larger pages scan faster on dense sheets but waste more on small/sparse
    /// ones. Default <see cref="DefaultRowPageSize"/> (1024).
    /// </summary>
    public int RowPageSize { get; init; } = DefaultRowPageSize;

    /// <summary>
    /// Columns per group in the two-level column directory. Must be a power of two in <c>[8, 4096]</c>. A
    /// lone high-index column costs one group (not a giant flat array). Default
    /// <see cref="DefaultColumnGroupSize"/> (64).
    /// </summary>
    public int ColumnGroupSize { get; init; } = DefaultColumnGroupSize;

    /// <summary>
    /// Sparsity guard: how many pages a sheet may allocate before its density is judged. Must be in
    /// <c>[1, 1000000]</c>. Below this the sheet is always allowed to grow dense. Default
    /// <see cref="DefaultSparsityWarmupPages"/> (64).
    /// </summary>
    public int SparsityWarmupPages { get; init; } = DefaultSparsityWarmupPages;

    /// <summary>
    /// Sparsity guard: the minimum average cells-per-page a sheet must sustain (after warm-up) to keep
    /// allocating new pages; below it, further scattered cells are diverted to a per-sheet dictionary instead
    /// of ballooning page memory. Must be in <c>[1, RowPageSize]</c>. Default
    /// <see cref="DefaultSparsityMinCellsPerPage"/> (4).
    /// </summary>
    public int SparsityMinCellsPerPage { get; init; } = DefaultSparsityMinCellsPerPage;

    /// <summary>Validates the option set, throwing <see cref="ArgumentException"/> for a non-power-of-two size
    /// or an out-of-range value. Called at <see cref="Workbook(WorkbookOptions)"/> so bad config fails fast.</summary>
    internal void Validate()
    {
        ValidatePowerOfTwo(RowPageSize, MinRowPageSize, MaxRowPageSize, nameof(RowPageSize));
        ValidatePowerOfTwo(ColumnGroupSize, MinColumnGroupSize, MaxColumnGroupSize, nameof(ColumnGroupSize));

        if (SparsityWarmupPages is < 1 or > MaxSparsityWarmupPages)
        {
            throw new ArgumentException(
                $"{nameof(SparsityWarmupPages)} must be between 1 and {MaxSparsityWarmupPages}; got {SparsityWarmupPages}.",
                nameof(SparsityWarmupPages)
            );
        }

        if (SparsityMinCellsPerPage < 1 || SparsityMinCellsPerPage > RowPageSize)
        {
            throw new ArgumentException(
                $"{nameof(SparsityMinCellsPerPage)} must be between 1 and RowPageSize ({RowPageSize}); got {SparsityMinCellsPerPage}.",
                nameof(SparsityMinCellsPerPage)
            );
        }
    }

    private static void ValidatePowerOfTwo(int value, int min, int max, string name)
    {
        if (value < min || value > max || (value & (value - 1)) != 0)
        {
            throw new ArgumentException(
                $"{name} must be a power of two in [{min}, {max}]; got {value}.",
                name
            );
        }
    }
}
