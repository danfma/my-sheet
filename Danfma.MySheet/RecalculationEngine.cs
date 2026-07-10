using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// A host-facing reference to a single cell: the sheet name and the A1 id (e.g. <c>"B5"</c>). The public
/// currency of <see cref="RecalculationEngine"/> — you tell it which cells you edited and it answers with the
/// outputs those edits affect, all in this shape.
/// </summary>
public readonly partial record struct CellRef(string Sheet, string Id);

/// <summary>How a <see cref="RecalculationEngine.Recalculate"/> handled an edit.</summary>
public enum RecalculationMode
{
    /// <summary>Only the affected cone was evicted (the cheap path — the common case for input edits).</summary>
    Partial,

    /// <summary>The impact was large (the cone reached a hot source or exceeded the cap, or an edited cell was
    /// outside the dense model), so the whole cache was invalidated and everything recomputes lazily.</summary>
    FullFallback,
}

/// <summary>
/// The verdict of an impact analysis, exposed so the host can decide BEFORE paying a recompute: whether editing
/// a set of cells triggers a FULL recompute (large impact — the cone reaches a hot source or exceeds the cap)
/// or a PARTIAL one (evict-and-pull), the cone size when partial, and a human-readable reason. Producing it is
/// cheap — the predictor bails at the first hot source without walking its huge fan-in — so calling
/// <see cref="RecalculationEngine.EstimateImpact"/> does NOT cost the price of a full recompute.
/// </summary>
public readonly record struct ImpactEstimate(bool RecommendFull, int ConeSize, string Reason)
{
    internal static ImpactEstimate Full(string reason) => new(true, -1, reason);

    internal static ImpactEstimate Partial(int coneSize) =>
        new(false, coneSize, $"cone pequeno ({coneSize} células)");
}

/// <summary>The outcome of a <see cref="RecalculationEngine.Recalculate"/> call.</summary>
/// <param name="Mode">Whether the edit was handled partially (evict-and-pull) or fell back to a full recompute.</param>
/// <param name="DirtyCellCount">The number of cells evicted on a <see cref="RecalculationMode.Partial"/> pass;
/// <c>-1</c> for a <see cref="RecalculationMode.FullFallback"/> (the whole workbook is stale).</param>
/// <param name="AffectedOutputs">The affected output cells (the sinks — "the cells at the top") when
/// <c>collectOutputs</c> was requested and the pass was partial; empty otherwise. These are the cells to read to
/// populate another sheet/PDF without scanning the whole workbook.</param>
/// <param name="StructureRebuilt"><c>true</c> when a formula edit (or a sheet add/remove) had invalidated the
/// dependency graph and this call rebuilt it before serving — the one cost formula edits pay.</param>
/// <param name="Reason">A short human-readable explanation of the chosen path.</param>
public readonly record struct RecalculationResult(
    RecalculationMode Mode,
    int DirtyCellCount,
    IReadOnlyList<CellRef> AffectedOutputs,
    bool StructureRebuilt,
    string Reason
);

/// <summary>
/// Incremental recomputation over a <see cref="Workbook"/> driven by a reverse dependency graph. Given the cells
/// you edited, it computes the affected cone (the edited cells ∪ their transitive dependents ∪ the always-dirty
/// volatiles), evicts exactly those from the memoized cache, and lets the pull-based engine recompute each once
/// on the next read — instead of the whole-workbook <see cref="Workbook.InvalidateCache"/>. It also answers the
/// input→output question: which outputs a set of inputs affects, without scanning the sheet.
///
/// <para><b>Formula edits are safe.</b> The graph is a snapshot, so changing a cell's FORMULA (its dependencies)
/// would make it stale. This engine detects that: <see cref="Sheet.StructuralVersion"/> bumps on every
/// structural edit (a formula added/removed/changed) and on sheet add/remove, and the next call rebuilds the
/// graph before serving (reported via <see cref="RecalculationResult.StructureRebuilt"/>). A pure VALUE edit
/// (changing a literal input) does NOT bump, so it keeps the cheap path. The rebuild is the amortized cost of a
/// formula edit; value edits never pay it.</para>
///
/// <para><b>Contract.</b> Create the engine AFTER the workbook is populated (typically after a first
/// <see cref="Workbook.ComputeAll"/>). You MUST report every cell you edit — a cell whose value/formula changed
/// but that you omit from <see cref="Recalculate"/> keeps its stale cached value. Single-threaded during the
/// edit phase (same contract as the structural index).</para>
/// </summary>
public sealed class RecalculationEngine
{
    private readonly Workbook _workbook;
    private DirtyEngine _engine;
    private Dictionary<Sheet, long> _snapshot;

    internal RecalculationEngine(Workbook workbook)
    {
        _workbook = workbook;
        _engine = DirtyEngine.Build(workbook);
        _snapshot = SnapshotVersions();
    }

    /// <summary>
    /// Analyzes the impact of editing <paramref name="edited"/> WITHOUT recomputing or evicting — the same
    /// full-vs-partial decision <see cref="Recalculate"/> makes internally, surfaced so the host can plan (show
    /// the cost, batch edits, choose a strategy). Rebuilds the graph first if a formula edit made it stale.
    /// </summary>
    public ImpactEstimate EstimateImpact(IReadOnlyCollection<CellRef> edited)
    {
        EnsureFresh();

        if (!TryTranslate(edited, out var deps))
        {
            return ImpactEstimate.Full("célula editada fora do modelo denso (id não-A1)");
        }

        return _engine.EstimateImpact(deps);
    }

    /// <summary>
    /// Marks <paramref name="edited"/> dirty and recomputes selectively (evict-and-pull): computes the affected
    /// cone, evicts only those cells, and returns — reading any cell then yields fresh values (lazily). A large
    /// cone (a hot column) falls back to a full <see cref="Workbook.InvalidateCache"/>. Rebuilds the dependency
    /// graph first if a formula edit (or a sheet add/remove) made it stale. Pass
    /// <paramref name="collectOutputs"/> to also get the affected output cells (the sinks) in the result — the
    /// set to read to populate another sheet/PDF; skipped by default since it costs an extra walk.
    /// </summary>
    public RecalculationResult Recalculate(
        IReadOnlyCollection<CellRef> edited,
        bool collectOutputs = false
    )
    {
        var rebuilt = EnsureFresh();

        if (!TryTranslate(edited, out var deps))
        {
            _workbook.InvalidateCache();
            return new RecalculationResult(
                RecalculationMode.FullFallback,
                -1,
                [],
                rebuilt,
                "célula editada fora do modelo denso (id não-A1) — recompute completo"
            );
        }

        var dirty = _engine.CalculateDirty(deps);

        if (dirty is null)
        {
            return new RecalculationResult(
                RecalculationMode.FullFallback,
                -1,
                [],
                rebuilt,
                "cone grande ou fonte quente — recompute completo"
            );
        }

        IReadOnlyList<CellRef> outputs = [];
        if (collectOutputs)
        {
            outputs = _engine.GetAffectedOutputs(dirty).ConvertAll(ToRef);
        }

        return new RecalculationResult(
            RecalculationMode.Partial,
            dirty.Count,
            outputs,
            rebuilt,
            "cone pequeno — evict-and-pull"
        );
    }

    // Rebuilds the internal graph if the workbook's structure changed (any sheet's version advanced, or a sheet
    // was added/removed) since the last build. Returns whether a rebuild happened.
    private bool EnsureFresh()
    {
        if (!IsStale())
        {
            return false;
        }

        _engine = DirtyEngine.Build(_workbook);
        _snapshot = SnapshotVersions();
        return true;
    }

    // Stale if the set of sheets changed (count/identity) or any sheet's structural version advanced. Sheet
    // identity is stable (the same Sheet object lives for the workbook's life), so a removed sheet drops from
    // the map and a new one is absent from it — both caught here.
    private bool IsStale()
    {
        var sheets = _workbook.Sheets;
        if (sheets.Count != _snapshot.Count)
        {
            return true;
        }

        foreach (var sheet in sheets.Values)
        {
            if (
                !_snapshot.TryGetValue(sheet, out var version)
                || version != sheet.StructuralVersion
            )
            {
                return true;
            }
        }

        return false;
    }

    private Dictionary<Sheet, long> SnapshotVersions()
    {
        var map = new Dictionary<Sheet, long>(_workbook.Sheets.Count);
        foreach (var sheet in _workbook.Sheets.Values)
        {
            map[sheet] = sheet.StructuralVersion;
        }
        return map;
    }

    // Translates host CellRefs to the internal numeric CellDep. Returns false if ANY ref is a non-A1 (overflow)
    // id: those are outside the dense dependency model, so we cannot compute their cone and the caller must fall
    // back to a full recompute (correctness over cleverness).
    private static bool TryTranslate(IReadOnlyCollection<CellRef> refs, out List<CellDep> deps)
    {
        deps = new List<CellDep>(refs.Count);
        foreach (var reference in refs)
        {
            if (!CellAddress.TryGetColumnRow(reference.Id, out var column, out var row))
            {
                return false;
            }
            deps.Add(new CellDep(reference.Sheet, column, row));
        }
        return true;
    }

    private static CellRef ToRef(CellDep dep) =>
        new(dep.Sheet, new CellAddress(dep.Column, dep.Row).ToId());
}
