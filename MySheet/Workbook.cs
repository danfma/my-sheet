using System.Collections;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using MemoryPack;
using MySheet.Expressions;

namespace MySheet;

/// <summary>
/// A user-supplied function implementation. Receives the raw (unevaluated) arguments so it can decide
/// what to evaluate (lazy/short-circuit) and the workbook for context.
/// </summary>
public delegate object? CustomFunction(Expression[] arguments, Workbook workbook);

[MemoryPackable]
public sealed partial class Workbook
{
    // Host-registered custom functions; not serialized (behavior is re-registered after deserialization).
    [MemoryPackIgnore]
    private Dictionary<string, CustomFunction>? _functions;

    // Memoized cell values; not serialized. Invalidation is explicit (see InvalidateCache).
    [MemoryPackIgnore]
    private ConcurrentDictionary<(string Sheet, string Id), object?>? _cache;

    // Cells currently being evaluated on the calling thread, to detect circular references. Thread-local
    // so concurrent (and benign) re-evaluation of the same cell on different threads is not a false cycle.
    [ThreadStatic]
    private static HashSet<(string Sheet, string Id)>? _evaluating;

    public ConcurrentDictionary<string, Sheet> Sheets { get; init; } = new();

    public Sheet this[string key] => Sheets[key];

    /// <summary>
    /// Returns the computed value of a cell, memoizing it. The cache is NOT invalidated automatically on
    /// mutation — call <see cref="InvalidateCache"/> after editing cells.
    /// </summary>
    public object? GetCellValue(string sheetName, string id)
    {
        var cache = _cache ??= new();
        var key = (sheetName, id);

        if (cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var evaluating = _evaluating ??= new();

        if (!evaluating.Add(key))
        {
            // The cell is already on this thread's evaluation stack: a circular reference.
            return ErrorValue.Reference;
        }

        try
        {
            // Compute outside the dictionary (the formula recurses into GetCellValue), then store.
            var value = Sheets[sheetName][id].Compute(new EvaluationContext(this, sheetName, id));
            cache[key] = value;

            return value;
        }
        finally
        {
            evaluating.Remove(key);
        }
    }

    public void InvalidateCache() => _cache?.Clear();

    public void RegisterFunction(string name, CustomFunction function) =>
        (_functions ??= new(StringComparer.OrdinalIgnoreCase))[name] = function;

    public bool TryGetFunction(string name, [MaybeNullWhen(false)] out CustomFunction function)
    {
        if (_functions is null)
        {
            function = null;
            return false;
        }

        return _functions.TryGetValue(name, out function);
    }
}

public static class CollectionExtensions
{
    extension(ConcurrentDictionary<string, Sheet> sheets)
    {
        public Sheet Add(string name) =>
            sheets.GetOrAdd(name, name => new Sheet { Name = name, Index = sheets.Count });
    }
}

[MemoryPackable]
public sealed partial class Sheet : IEnumerable<KeyValuePair<string, Expression>>
{
    public required string Name { get; init; }

    // 0-based insertion order, used by the SHEET function.
    public int Index { get; init; }

    public Dictionary<string, Expression> Cells { get; init; } = new();

    [MemoryPackIgnore]
    public int Count => Cells.Count;

    [MemoryPackIgnore]
    public Expression this[string key]
    {
        get => Cells.TryGetValue(key, out var cell) ? cell : BlankValue.Instance;
        set => Cells[key] = value;
    }

    [MemoryPackIgnore]
    public IEnumerable<string> Keys => Cells.Keys;

    [MemoryPackIgnore]
    public IEnumerable<Expression> Values => Cells.Values;

    public IEnumerator<KeyValuePair<string, Expression>> GetEnumerator()
    {
        return Cells.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)Cells).GetEnumerator();
    }

    public bool ContainsKey(string key)
    {
        return Cells.ContainsKey(key);
    }

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Expression value)
    {
        return Cells.TryGetValue(key, out value);
    }
}
