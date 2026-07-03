# Migrating to 3.0

MySheet 3.0 **encapsulates `Sheet`'s cell store**. The dictionary that used to be a public, mutable,
`init`-settable property is now exposed as a **read-only view**, and all mutation funnels through two
paths: the indexer `set` (unchanged) and a new `Remove`. The break is strictly **compile-time and
narrow** — reading is untouched, writing through `sheet["A1"] = …` is untouched, and **serialized
workbooks are 100% compatible** in both directions (the wire format is byte-identical).

## Why

3.0 turns the cell store into a single write **choke point**. Funnelling every insert, overwrite and
delete through one place is the foundation the later 3.x work builds on — a write-maintained structural
index (kept up to date on the write instead of rebuilt per cache epoch) and, further out, a reverse
dependency graph for incremental recalculation. None of that machinery ships in 3.0; the release is the
encapsulation that makes it possible without touching callers again.

## What changed

`Sheet.Cells` changed type, and two members were added:

| Member | 2.x | 3.0 |
| --- | --- | --- |
| `Sheet.Cells` | `public Dictionary<string, Expression> { get; init; }` | `public IReadOnlyDictionary<string, Expression> { get; }` |
| `Sheet.SetCell(string id, Expression expr)` | — | `internal` — the single write path the indexer `set` delegates to |
| `Sheet.Remove(string id)` | — | `public bool` — removes a cell, returns whether it existed |

Everything else on `Sheet` is unchanged: the read indexer `sheet["A1"]` (blank for a missing cell), the
write indexer `sheet["A1"] = expr`, `Count`, `Keys`, `Values`, `ContainsKey`, `TryGetValue`, and
enumeration (`foreach (var (id, expr) in sheet)`) all keep their exact 2.x behavior.

## Updating your code

### Reading — nothing to do

Every read path is source-compatible. `Cells` is still enumerable and indexable, `Count`/`Keys`/`Values`
still return the same things:

```csharp
foreach (var (id, expr) in sheet) { /* … */ }   // unchanged
var n = sheet.Count;                             // unchanged
if (sheet.ContainsKey("A1")) { /* … */ }         // unchanged
var expr = sheet["A1"];                           // unchanged (blank for a missing cell)
var only = sheet.Cells["A1"];                     // still works — read-only indexer
```

The only reads that break are ones that **mutated through `Cells`**, e.g. `sheet.Cells["A1"] = expr`,
`sheet.Cells.Remove("A1")`, or `sheet.Cells.Clear()`. Those never needed to reach through the property —
use the `Sheet` members below.

### Writing — use the indexer

The write indexer is unchanged and remains the intended write path. If you were mutating the underlying
dictionary directly, switch to it:

```csharp
// 2.x — reaching through the (then-mutable) dictionary
sheet.Cells["A1"] = new NumberValue(10);

// 3.0
sheet["A1"] = new NumberValue(10);   // delegates to the SetCell choke point
```

### Deleting — use `Remove`

Deleting a cell used to mean `sheet.Cells.Remove("A1")`. That is now `Sheet.Remove`:

```csharp
// 2.x
sheet.Cells.Remove("A1");

// 3.0
bool existed = sheet.Remove("A1");   // true if a cell was there, false for a no-op
```

`Remove` shares the **same explicit-invalidation semantics as a write**: it does not clear memoized
values on its own. Removing a cell changes the result of every formula that read it, so — exactly as
after a write — call `workbook.InvalidateCache()` for the change to be observed on the next read.

### Object initializers — build then populate

Because `Cells` lost its `init` accessor, an object initializer that seeded cells no longer compiles.
Construct the sheet, then populate through the indexer:

```csharp
// 2.x
var sheet = new Sheet
{
    Name = "Sheet1",
    Cells = { ["A1"] = new NumberValue(10), ["A2"] = new NumberValue(20) },
};

// 3.0
var sheet = new Sheet { Name = "Sheet1" };
sheet["A1"] = new NumberValue(10);
sheet["A2"] = new NumberValue(20);
```

In practice most code already creates sheets with `workbook.Sheets.Add("Sheet1")` and fills them through
the indexer, so there is nothing to change there.

## Serialization compatibility

**No action required.** The cell dictionary moved from a public property to a private
`[MemoryPackInclude]` field **at the exact same declaration position**. MemoryPack orders members by
declaration, so the wire schema — member #3, a `Dictionary<string, Expression>` — is byte-identical.
Workbooks saved by 2.x (and 1.x) load in 3.0 and vice versa, unchanged. This is guarded by the same
frozen binary fixture that has covered the format since 2.0
(`tests/Danfma.MySheet.Tests/Fixtures/workbook-pre-namespaces.msgpack.bin`), loaded and re-evaluated on
every run.

## Performance — no action required

The write choke point lets 3.0 make the whole-column **structural index write-maintained**. In 2.x that
index was rebuilt per cache epoch (on an open-range read after each `InvalidateCache()`); in 3.0 the
`Sheet` keeps it current as cells are written and deleted, so it is built once per sheet and **survives
`InvalidateCache()`**. The visible effect is that the "load once, then re-read a whole column every
epoch" shape stops paying a per-epoch index rebuild — its cost flattens to track the column, not the
sheet total (see [Performance](performance.md#repeated-whole-column-reads-scale-with-the-column-not-the-sheet)).
This is transparent: same results, same call sequence (still `edit → InvalidateCache() → read`), nothing
to change in your code. The one caller-facing consequence is the encapsulation above — writes and deletes
must go through the indexer and `Remove` (which is how the index stays correct), exactly as documented.

## Behavior

No formula semantics, parsing, evaluation results, or save format changed. The break is confined to the
`Sheet` API surface described above. If a 2.x → 3.0 upgrade changes any computed value or any saved byte,
that is a bug — please report it.
