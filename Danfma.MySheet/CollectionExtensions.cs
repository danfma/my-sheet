using System.Collections.Concurrent;

namespace Danfma.MySheet;

public static class CollectionExtensions
{
    extension(ConcurrentDictionary<string, Sheet> sheets)
    {
        public Sheet Add(string name) =>
            sheets.GetOrAdd(name, name => new Sheet { Name = name, Index = sheets.Count });
    }
}
