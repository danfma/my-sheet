using System.Collections;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using MemoryPack;
using MySheet.Expressions;

namespace MySheet;

[MemoryPackable]
public sealed partial class Workbook
{
    public ConcurrentDictionary<string, Sheet> Sheets { get; init; } = new();

    public Sheet this[string key] => Sheets[key];
}

public static class CollectionExtensions
{
    extension(ConcurrentDictionary<string, Sheet> sheets)
    {
        public Sheet Add(string name) => sheets.GetOrAdd(name, name => new Sheet { Name = name });
    }
}

[MemoryPackable]
public sealed partial class Sheet : IEnumerable<KeyValuePair<string, Expression>>
{
    public required string Name { get; init; }

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
