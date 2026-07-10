using Danfma.MySheet.Expressions;
using MemoryPack;

namespace Danfma.MySheet;

public sealed partial class Workbook
{
    // === Volatile-function epoch model (F1) ===============================================================
    // A volatile cell (NOW/TODAY/RAND/RANDBETWEEN, directly or transitively) is cached WITHIN an epoch and its
    // address is recorded as tainted inside the value store. Recalculate() drops just those cells and re-samples
    // the clock (epoch++); InvalidateCache() drops everything. This keeps NOW()/RAND() coherent within a pass
    // (sampled once) while staying cheap to refresh, all without a dependency graph.

    // "The cell currently being evaluated on this thread touched a volatile." Same thread-local save/reset/
    // propagate pattern as _evaluating: GetCellValue zeroes it before a cell, reads it after (to mark the
    // cell), and ORs it back into the caller's value so volatility propagates up the evaluation stack.
    [ThreadStatic]
    private static bool _volatileTouched;

    // The clock sampled once per epoch (lazily, on the first volatile read), as an Excel serial (local time).
    // null means "not yet sampled this epoch". Guarded by ClockLock so it is sampled exactly once.
    [MemoryPackIgnore]
    private double? _epochNow;

    // Persistent RNG for RAND/RANDBETWEEN, advanced across epochs (never re-seeded per epoch, so epochs differ
    // naturally; a fixed RandomSeed makes the whole run reproducible). Created lazily under RandomLock.
    [MemoryPackIgnore]
    private Random? _random;

    // Guards the once-per-epoch clock sampling (_epochNow). Lazily created (Interlocked) so it survives
    // MemoryPack deserialization, which bypasses field initializers.
    //
    // Split from the RNG's lock (below) on purpose: a multi-threaded workload mixing NOW()/TODAY() with
    // RAND()/RANDBETWEEN() used to serialize every random draw behind the same lock as the clock sample,
    // even though the two have no correctness dependency on each other. Checked before splitting: the only
    // places that reset EITHER field — InvalidateCache and Recalculate, both above — reset `_epochNow` alone;
    // neither touches `_random` (RNG state intentionally persists across epochs, per the comment on
    // `_random`). So there is no "reset both atomically" invariant to preserve — the two locks can be taken
    // independently and in any order, including concurrently from different threads.
    [MemoryPackIgnore]
    private object? _clockLock;

    // Guards the not-thread-safe RNG (_random). See _clockLock for why this is a separate lock.
    [MemoryPackIgnore]
    private object? _randomLock;

    private object ClockLock
    {
        get
        {
            var existing = _clockLock;
            if (existing is not null)
            {
                return existing;
            }

            var created = new object();
            return Interlocked.CompareExchange(ref _clockLock, created, null) ?? created;
        }
    }

    private object RandomLock
    {
        get
        {
            var existing = _randomLock;
            if (existing is not null)
            {
                return existing;
            }

            var created = new object();
            return Interlocked.CompareExchange(ref _randomLock, created, null) ?? created;
        }
    }

    // Backing field for the injectable clock; ignored by MemoryPack (runtime config, not persisted state).
    [MemoryPackIgnore]
    private TimeProvider? _timeProvider;

    /// <summary>
    /// The clock <c>NOW()</c>/<c>TODAY()</c> read (defaults to <see cref="TimeProvider.System"/>). Injectable
    /// so hosts can freeze time for a batch and tests can pin both the instant and the local zone. Excel uses
    /// LOCAL time, so the functions read <see cref="TimeProvider.GetLocalNow"/>. Not serialized.
    /// </summary>
    [MemoryPackIgnore]
    public TimeProvider TimeProvider
    {
        get => _timeProvider ?? TimeProvider.System;
        set => _timeProvider = value;
    }

    /// <summary>
    /// Seed for the <c>RAND</c>/<c>RANDBETWEEN</c> RNG. <c>null</c> (default) seeds it from the clock; a fixed
    /// value makes the whole run's random sequence reproducible. Set it before the first volatile read (the
    /// RNG is created lazily and never re-seeded afterwards). Not serialized (runtime config).
    /// </summary>
    [MemoryPackIgnore]
    public int? RandomSeed { get; set; }

    /// <summary>
    /// Marks the current thread's cell evaluation as having touched a volatile source. Volatile nodes call
    /// this from <c>Evaluate</c>; <see cref="GetCellValue"/> reads the flag to cache-and-mark the cell and to
    /// propagate volatility to dependents. Internal — part of the evaluation contract, not host API.
    /// </summary>
    internal void MarkVolatileTouched() => _volatileTouched = true;

    /// <summary>
    /// The current epoch's clock, sampled once (lazily) as an Excel serial from local time, so every
    /// <c>NOW()</c>/<c>TODAY()</c> in a pass agrees. Thread-safe: the first caller of the epoch samples and
    /// publishes under <see cref="ClockLock"/>, the rest read the published value. Uses a lock separate from
    /// <see cref="RandomLock"/> so NOW()/TODAY() reads never serialize behind RAND()/RANDBETWEEN() draws (or
    /// vice versa) in a multi-threaded evaluation — see the comment on <see cref="_clockLock"/> for why that
    /// split is safe.
    /// </summary>
    internal double EpochNow()
    {
        lock (ClockLock)
        {
            return _epochNow ??= DateSerial.FromDateTime(TimeProvider.GetLocalNow().DateTime);
        }
    }

    /// <summary>
    /// Draws the next value in <c>[0, 1)</c> from the persistent RNG (created lazily from
    /// <see cref="RandomSeed"/>) and marks the evaluation volatile. Thread-safe (<see cref="Random"/> is not).
    /// The RNG is NOT re-seeded per epoch — the sequence continues so successive epochs differ, while the
    /// per-epoch cache keeps a single cell stable within a pass. Uses <see cref="RandomLock"/>, separate from
    /// <see cref="ClockLock"/> — see the comment on <see cref="_clockLock"/>.
    /// </summary>
    internal double NextRandom()
    {
        MarkVolatileTouched();

        lock (RandomLock)
        {
            _random ??= RandomSeed is { } seed ? new Random(seed) : new Random();
            return _random.NextDouble();
        }
    }
}
