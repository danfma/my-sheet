# Dynamic Range Endpoints Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the `:` range operator accept reference-returning expressions (`INDEX`, `OFFSET`, `CHOOSE`, parenthesised references) as endpoints, matching Excel, instead of throwing `ParseException`.

**Architecture:** Add a virtual `TryResolveReference` on `Expression` that yields a node's *target reference without dereferencing*. A new `DynamicRange` reference node holds two endpoint expressions and, at evaluation, resolves each to a reference and spans them into a concrete `RangeReference`. `Parser.ParseRange` builds a `DynamicRange` instead of throwing when endpoints are not static.

**Tech Stack:** C# / .NET 10, MemoryPack (serialization), TUnit (tests).

## Global Constraints

- Excel behaviour is the acceptance criterion. For subtle results, derive the expected value with `Aspose.Cells` (already referenced in `external/BenchExpressions`) and bake it as a constant into the test — no runtime Aspose dependency in `Danfma.MySheet.Tests`.
- MemoryPackUnion tags on `Expression` are APPEND-ONLY. Next free tag is **317** (max currently 316 = `OpenRangeReference`). Never renumber/reuse.
- Value-context `Evaluate` output of every function must stay identical (`1×1` → value, multi-cell → reference). `TryResolveReference` is a NEW, separate method.
- Tests run with TUnit — `dotnet test` does NOT work on .NET 10. Run:
  `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/<ClassName>/*"`
- TUnit assertions are async: `await Assert.That(actual).IsEqualTo(expected)`.
- No self-reference lines in commit messages.

---

### Task 1: `TryResolveReference` foundation on `Expression` and static references

**Files:**
- Modify: `Danfma.MySheet/Expressions/Expression.cs` (add virtual, ~line 345 near `Evaluate`)
- Modify: `Danfma.MySheet/Expressions/Reference.cs` (base override on `Reference`)
- Modify: `Danfma.MySheet/Expressions/CellReference.cs`, `RangeReference.cs`, `OpenRangeReference.cs`, `UnionReference.cs`, `NameReference.cs`
- Test: `tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs` (create)

**Interfaces:**
- Produces: `public virtual bool Expression.TryResolveReference(EvaluationContext context, out Reference? reference)` — returns the expression's target reference without dereferencing; `false` when the expression is not reference-producing.

- [ ] **Step 1: Write the failing test**

```csharp
// tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class TryResolveReferenceTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return (workbook, sheet);
    }

    [Test]
    public async Task CellReference_ResolvesToItself()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=A2", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<CellReference>();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }

    [Test]
    public async Task NumberValue_DoesNotResolve()
    {
        var (workbook, _) = Grid();
        var expr = new NumberValue(3);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: FAIL — `Expression` has no member `TryResolveReference`.

- [ ] **Step 3: Add the virtual to `Expression`**

In `Danfma.MySheet/Expressions/Expression.cs`, immediately after `public abstract ComputedValue Evaluate(EvaluationContext context);`:

```csharp
    /// <summary>
    /// Resolves this expression to its target <see cref="Reference"/> WITHOUT dereferencing a single cell
    /// to its value (unlike <see cref="Evaluate"/>). Used by the ':' range operator and reference-context
    /// consumers. Default: not a reference-producing expression.
    /// </summary>
    public virtual bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;
        return false;
    }
```

- [ ] **Step 4: Override on the static reference nodes**

In `Danfma.MySheet/Expressions/Reference.cs`, add to the `Reference` base so every concrete reference resolves to itself:

```csharp
public abstract record Reference : Expression
{
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = this;
        return true;
    }
}
```

In `Danfma.MySheet/Expressions/NameReference.cs`, override to reuse the existing name resolution (LET scope, defined names, cycle guard):

```csharp
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference) =>
        NamedReferences.TryResolveReference(this, context, out reference);
```

(`CellReference`, `RangeReference`, `OpenRangeReference`, `UnionReference` inherit the `Reference` override — no per-file change needed. Confirm each derives from `Reference`.)

- [ ] **Step 4b: Route `NamedReferences` through the virtual (critical wiring)**

Every reference-consuming function (`SUM`, `COUNTIF`, `VLOOKUP` table, `INDEX` base, …) resolves its range argument through `NamedReferences.TryResolveReference` → `TryResolveRaw`. That method currently returns any `Reference` node as-is. A `DynamicRange` (Task 5) IS a `Reference`, so returning it as-is would hand consumers an unresolved node they cannot expand. Route non-name expressions through the virtual so `DynamicRange`/`Index`/`Offset`/`Choose` resolve to their concrete target.

In `Danfma.MySheet/Expressions/NamedReferences.cs`, change `TryResolveRaw` (currently `Parsing/…` lines 86-133) so the `Reference direct` branch delegates to the virtual instead of returning `this` verbatim:

```csharp
    private static bool TryResolveRaw(
        Expression expression,
        EvaluationContext context,
        [NotNullWhen(true)] out Reference? reference
    )
    {
        // Names resolve through LET scope / defined names below. Everything else — a plain reference
        // (resolves to itself), a DynamicRange (resolves to its span), a reference-returning function
        // (INDEX/OFFSET/CHOOSE, resolves to its target) — goes through the virtual.
        if (expression is not NameReference name)
        {
            return expression.TryResolveReference(context, out reference);
        }

        // ... existing NameReference body unchanged (LET scope, defined names, cycle guard, recurse) ...
    }
```

The `NameReference.TryResolveReference` override (Step 4) calls back into `NamedReferences.TryResolveReference`; since `TryResolveRaw` handles `NameReference` directly before delegating, there is no recursion. Behaviour is identical today (references → self, functions → `false`) until Tasks 2-4 add function overrides.

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: PASS (both tests).

- [ ] **Step 6: Commit**

```bash
git add Danfma.MySheet/Expressions/Expression.cs Danfma.MySheet/Expressions/Reference.cs Danfma.MySheet/Expressions/NameReference.cs tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs
git commit -m "feat(expr): add TryResolveReference resolution path on Expression and references"
```

---

### Task 2: `Index.TryResolveReference` (concrete-range path only)

**Files:**
- Modify: `Danfma.MySheet/Expressions/Lookup/Index.cs`
- Test: `tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs` (add)

**Interfaces:**
- Consumes: `Expression.TryResolveReference` (Task 1), `NamedReferences.TryResolveReference`, `CellAddress`.
- Produces: `Index` overrides `TryResolveReference` — concrete-range INDEX yields a `CellReference` (single cell) / `RangeReference`; array-form INDEX returns `false`.

- [ ] **Step 1: Write the failing tests**

```csharp
// add to TryResolveReferenceTests.cs
    [Test]
    public async Task IndexIntoConcreteRange_ResolvesToTargetCell()
    {
        var (workbook, sheet) = Grid();
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        var expr = ExpressionParser.Parse("=INDEX(A1:A3,2)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<CellReference>();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }

    [Test]
    public async Task IndexIntoComputedArray_DoesNotResolve()
    {
        var (workbook, sheet) = Grid();
        // INDEX(ROW($A:$A), 2) indexes a computed vector, not cells: no address, must not resolve.
        var expr = ExpressionParser.Parse("=INDEX(ROW($A:$A),2)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: FAIL — `IndexIntoConcreteRange_ResolvesToTargetCell` (Index has no override, falls to default `false`).

- [ ] **Step 3: Implement `Index.TryResolveReference`**

In `Danfma.MySheet/Expressions/Lookup/Index.cs`, add. It mirrors the concrete-range branch of `Evaluate` (the `NamedReferences.TryResolveReference` + row/column resolution) but returns the ADDRESS as a `CellReference` instead of reading the value. Array/open-row-number forms return `false`.

```csharp
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // Array forms (mini-CSE vector, open-column ROW identity) have no cell address.
        if (Arguments[0] is Row { Arguments: [OpenRangeReference] }
            || (Arguments[0] is not Reference && ArrayEvaluation.IsArrayEligible(Arguments[0])))
        {
            return false;
        }

        if (!NamedReferences.TryResolveReference(Arguments[0], context, out var resolved)
            || resolved is not RangeReference range)
        {
            return false;
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is not null)
        {
            return false;
        }

        double row;
        double column;
        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out column) is not null)
            {
                return false;
            }
            row = first;
        }
        else if (range.RowCount == 1)
        {
            row = 1;
            column = first;
        }
        else
        {
            row = first;
            column = 1;
        }

        if (row < 1 || column < 1 || row > range.RowCount || column > range.ColumnCount)
        {
            return false;
        }

        var start = CellAddress.Parse(range.StartId);
        var target = new CellAddress(start.Column + (int)column - 1, start.Row + (int)row - 1);
        reference = new CellReference(target.ToId(), range.SheetName);
        return true;
    }
```

> Confirm `RangeReference` exposes `RowCount`/`ColumnCount` (used in `Evaluate`). If they are helper methods elsewhere, reuse the same accessor `Evaluate` uses.

- [ ] **Step 4: Run to verify pass**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add Danfma.MySheet/Expressions/Lookup/Index.cs tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs
git commit -m "feat(index): resolve INDEX to a target reference in reference context"
```

---

### Task 3: `Offset.TryResolveReference`

**Files:**
- Modify: `Danfma.MySheet/Expressions/Lookup/Offset.cs`
- Test: `tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs` (add)

**Interfaces:**
- Produces: `Offset` overrides `TryResolveReference` — yields its target cell/range reference (the address it computes, before dereferencing a `1×1` to value).

- [ ] **Step 1: Write the failing test**

```csharp
    [Test]
    public async Task Offset_ResolvesToTargetCell()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=OFFSET(A1,1,0)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: FAIL — Offset has no override.

- [ ] **Step 3: Refactor `Offset` to expose its target reference**

In `Danfma.MySheet/Expressions/Lookup/Offset.cs`, extract the address computation (currently inside `Evaluate`, `Offset.cs:9-66`) into a private helper that returns a `Reference`, then have both `Evaluate` and `TryResolveReference` call it. `Evaluate` keeps its exact output: a `1×1` reference is dereferenced via `GetCellValue`, multi-cell stays `ComputedValue.Reference`.

```csharp
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
        => TryComputeTarget(context, out reference);

    // Computes OFFSET's target reference (a CellReference for 1x1, else a RangeReference) without
    // dereferencing. Returns false on any argument/bounds error.
    private bool TryComputeTarget(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        if (!NamedReferences.TryResolveReference(Arguments[0], context, out var baseReference)
            || !TryBase(baseReference, out var sheetName, out var baseColumn, out var baseRow))
        {
            return false;
        }
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var rows) is not null) return false;
        if (Arguments[2].Evaluate(context).CoerceToNumber(out var columns) is not null) return false;

        double height = 1, width = 1;
        if (Arguments.Length >= 4 && Arguments[3].Evaluate(context).CoerceToNumber(out height) is not null) return false;
        if (Arguments.Length >= 5 && Arguments[4].Evaluate(context).CoerceToNumber(out width) is not null) return false;

        var startColumn = baseColumn + (int)columns;
        var startRow = baseRow + (int)rows;
        if (startColumn < 1 || startRow < 1 || height < 1 || width < 1)
        {
            return false;
        }

        reference = height == 1 && width == 1
            ? new CellReference(new CellAddress(startColumn, startRow).ToId(), sheetName)
            : new RangeReference(
                new CellAddress(startColumn, startRow).ToId(),
                new CellAddress(startColumn + (int)width - 1, startRow + (int)height - 1).ToId(),
                sheetName);
        return true;
    }
```

Then rewrite the tail of `Evaluate` to reuse it (replacing the inline duplicate at `Offset.cs:44-66`):

```csharp
        if (!TryComputeTarget(context, out var target))
        {
            return ComputedValue.Error(Error.Ref);
        }

        return target is CellReference cell
            ? context.Workbook.GetCellValue(cell.SheetName, cell.Id)   // 1x1: dereference
            : ComputedValue.Reference(target!);                        // multi-cell: reference value
```

> Keep the earlier per-argument error branches in `Evaluate` if they return specific errors the tests rely on; `TryComputeTarget` collapses them to `#REF!`, which matches `Evaluate`'s existing `Error.Ref` fallbacks. Verify no OFFSET test asserts a non-`#REF!` error from a bad OFFSET.

- [ ] **Step 4: Run to verify pass (and no OFFSET regression)**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Then run the existing OFFSET suite: `... --treenode-filter "/*/*/*/*Offset*"` (adjust to the actual OFFSET test class name).
Expected: PASS, no regressions.

- [ ] **Step 5: Commit**

```bash
git add Danfma.MySheet/Expressions/Lookup/Offset.cs tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs
git commit -m "feat(offset): resolve OFFSET to a target reference; share compute with Evaluate"
```

---

### Task 4: `Choose.TryResolveReference`

**Files:**
- Modify: `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs` (the `Choose` record, ~line 12)
- Test: `tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs` (add)

**Interfaces:**
- Produces: `Choose` overrides `TryResolveReference` — resolves the chosen argument to a reference.

- [ ] **Step 1: Write the failing test**

```csharp
    [Test]
    public async Task Choose_ResolvesChosenArgumentToReference()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=CHOOSE(2,A1,B5)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("B5");
    }
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: FAIL — Choose has no override.

- [ ] **Step 3: Implement `Choose.TryResolveReference`**

In `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`, inside the `Choose` record:

```csharp
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var index) is not null)
        {
            return false;
        }

        var slot = (int)index;
        if (slot < 1 || slot >= Arguments.Length)
        {
            return false;
        }

        return Arguments[slot].TryResolveReference(context, out reference);
    }
```

- [ ] **Step 4: Run to verify pass**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/TryResolveReferenceTests/*"`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs
git commit -m "feat(choose): resolve CHOOSE to the chosen argument's reference"
```

---

### Task 5: `DynamicRange` node + serialization

**Files:**
- Create: `Danfma.MySheet/Expressions/DynamicRange.cs`
- Modify: `Danfma.MySheet/Expressions/Expression.cs` (add `[MemoryPackUnion(317, typeof(DynamicRange))]`)
- Test: `tests/Danfma.MySheet.Tests/Expressions/DynamicRangeTests.cs` (create)

**Interfaces:**
- Consumes: `Expression.TryResolveReference` (Task 1), `CellAddress`, `RangeReference`.
- Produces: `public sealed partial record DynamicRange(Expression Start, Expression End) : Reference` — resolves its endpoints and spans them into a `RangeReference`.

- [ ] **Step 1: Write the failing test**

```csharp
// tests/Danfma.MySheet.Tests/Expressions/DynamicRangeTests.cs
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Expressions;

public class DynamicRangeTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return (workbook, sheet);
    }

    [Test]
    public async Task Resolves_IndexEndpoint_ToSpanningRange()
    {
        var (workbook, sheet) = Grid();
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        // INDEX(A1:A3,2) -> A2 ; span A2:A3
        var range = new DynamicRange(
            ExpressionParser.Parse("=INDEX(A1:A3,2)", sheet),
            new CellReference("A3", "Sheet1"));

        var ok = range.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        var rr = (RangeReference)reference!;
        await Assert.That(rr.StartId).IsEqualTo("A2");
        await Assert.That(rr.EndId).IsEqualTo("A3");
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesEndpoints()
    {
        var range = new DynamicRange(new CellReference("A1", "Sheet1"), new CellReference("A3", "Sheet1"));
        var bytes = MemoryPackSerializer.Serialize<Expression>(range);
        var back = (DynamicRange)MemoryPackSerializer.Deserialize<Expression>(bytes)!;

        await Assert.That(((CellReference)back.Start).Id).IsEqualTo("A1");
        await Assert.That(((CellReference)back.End).Id).IsEqualTo("A3");
    }
}
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/DynamicRangeTests/*"`
Expected: FAIL — `DynamicRange` does not exist.

- [ ] **Step 3: Create the node**

```csharp
// Danfma.MySheet/Expressions/DynamicRange.cs
using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A ':' range whose endpoints are reference-returning EXPRESSIONS (e.g. INDEX(...):$D1), resolved to a
/// concrete <see cref="RangeReference"/> at evaluation time. The static endpoint forms (A1:B2, $D:$D) are
/// built directly by the parser and never reach this node.
/// </summary>
[MemoryPackable]
public sealed partial record DynamicRange(Expression Start, Expression End) : Reference
{
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        if (!Start.TryResolveReference(context, out var start)
            || !End.TryResolveReference(context, out var end)
            || !TryBox(start!, out var startBox)
            || !TryBox(end!, out var endBox))
        {
            return false;
        }

        var sheet = start is CellReference sc ? sc.SheetName
            : start is RangeReference sr ? sr.SheetName
            : end is CellReference ec ? ec.SheetName
            : ((RangeReference)end!).SheetName;

        var minColumn = Math.Min(startBox.MinColumn, endBox.MinColumn);
        var minRow = Math.Min(startBox.MinRow, endBox.MinRow);
        var maxColumn = Math.Max(startBox.MaxColumn, endBox.MaxColumn);
        var maxRow = Math.Max(startBox.MaxRow, endBox.MaxRow);

        reference = new RangeReference(
            new CellAddress(minColumn, minRow).ToId(),
            new CellAddress(maxColumn, maxRow).ToId(),
            sheet);
        return true;
    }

    // A ':' endpoint evaluated as a scalar mirrors a bare RangeReference: #VALUE! unless a consumer expands
    // it. Delegating to the resolved range inherits that behaviour.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryResolveReference(context, out var reference)
            ? ComputedValue.Reference(reference!)
            : ComputedValue.Error(Error.Ref);

    private readonly record struct Box(int MinColumn, int MinRow, int MaxColumn, int MaxRow);

    // Bounding box of a resolved reference. Only bounded references (cell/rectangle) are supported as a
    // dynamic endpoint; an open reference as a dynamic endpoint is out of scope (returns false -> #REF!).
    private static bool TryBox(Reference reference, out Box box)
    {
        switch (reference)
        {
            case CellReference cell:
                var a = CellAddress.Parse(cell.Id);
                box = new Box(a.Column, a.Row, a.Column, a.Row);
                return true;
            case RangeReference range:
                var s = CellAddress.Parse(range.StartId);
                var e = CellAddress.Parse(range.EndId);
                box = new Box(
                    Math.Min(s.Column, e.Column), Math.Min(s.Row, e.Row),
                    Math.Max(s.Column, e.Column), Math.Max(s.Row, e.Row));
                return true;
            default:
                box = default;
                return false;
        }
    }
}
```

- [ ] **Step 4: Register the union tag**

In `Danfma.MySheet/Expressions/Expression.cs`, after `[MemoryPackUnion(316, typeof(OpenRangeReference))]`:

```csharp
[MemoryPackUnion(317, typeof(DynamicRange))]
```

- [ ] **Step 5: Run to verify pass**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/DynamicRangeTests/*"`
Expected: PASS (both tests).

- [ ] **Step 6: Commit**

```bash
git add Danfma.MySheet/Expressions/DynamicRange.cs Danfma.MySheet/Expressions/Expression.cs tests/Danfma.MySheet.Tests/Expressions/DynamicRangeTests.cs
git commit -m "feat(expr): add DynamicRange node spanning reference-expression endpoints"
```

---

### Task 6: Parser builds `DynamicRange` instead of throwing

**Files:**
- Modify: `Danfma.MySheet/Parsing/Parser.cs` (`ParseRange`, line 466)
- Test: `tests/Danfma.MySheet.Tests/Parsing/DynamicRangeParsingTests.cs` (create)

**Interfaces:**
- Consumes: `DynamicRange` (Task 5).

- [ ] **Step 1: Write the failing tests**

```csharp
// tests/Danfma.MySheet.Tests/Parsing/DynamicRangeParsingTests.cs
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class DynamicRangeParsingTests
{
    private static Sheet NewSheet()
    {
        var workbook = new Workbook();
        return workbook.Sheets.Add("Sheet1");
    }

    [Test]
    public async Task IndexEndpoint_ParsesToDynamicRange()
    {
        var expr = ExpressionParser.Parse("=INDEX($D:$D,2):$D1", NewSheet());
        await Assert.That(expr).IsTypeOf<DynamicRange>();
    }

    [Test]
    public async Task StaticRange_StillParsesToRangeReference()
    {
        var expr = ExpressionParser.Parse("=A1:B2", NewSheet());
        await Assert.That(expr).IsTypeOf<RangeReference>();
    }

    [Test]
    public async Task WholeColumn_StillParsesToOpenRange()
    {
        var expr = ExpressionParser.Parse("=$D:$D", NewSheet());
        await Assert.That(expr).IsTypeOf<OpenRangeReference>();
    }
}
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/DynamicRangeParsingTests/*"`
Expected: FAIL — `IndexEndpoint_ParsesToDynamicRange` throws `ParseException`.

- [ ] **Step 3: Replace the throw**

In `Danfma.MySheet/Parsing/Parser.cs`, `ParseRange`, replace line 466:

```csharp
        throw new ParseException("The ':' range operator requires cell references", colon.Position);
```

with:

```csharp
        // Endpoints that are not statically resolvable (a reference-returning function like INDEX/OFFSET,
        // a parenthesised reference) become a DynamicRange, resolved at evaluation time.
        return new DynamicRange(left, right);
```

- [ ] **Step 4: Run to verify pass**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/DynamicRangeParsingTests/*"`
Expected: PASS (all three).

- [ ] **Step 5: Commit**

```bash
git add Danfma.MySheet/Parsing/Parser.cs tests/Danfma.MySheet.Tests/Parsing/DynamicRangeParsingTests.cs
git commit -m "feat(parser): build DynamicRange for non-static ':' endpoints instead of throwing"
```

---

### Task 7: End-to-end evaluation + Excel-compatibility (Aspose oracle) + regression

**Files:**
- Test: `tests/Danfma.MySheet.Tests/Expressions/DynamicRangeEvaluationTests.cs` (create)

**Interfaces:**
- Consumes: everything above. No production change expected; if a test fails, fix the responsible node's `TryResolveReference` / `DynamicRange` and note it.

**Aspose oracle procedure (run once, offline, to derive expected constants):** in `external/BenchExpressions` add a throwaway console block that builds the same grid in `Aspose.Cells`, sets the formula, calls `CalculateFormula()`, and prints the result. Copy the printed value into the test as a literal. Do NOT reference Aspose from `Danfma.MySheet.Tests`.

- [ ] **Step 1: Write the evaluation + compat + regression tests**

```csharp
// tests/Danfma.MySheet.Tests/Expressions/DynamicRangeEvaluationTests.cs
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class DynamicRangeEvaluationTests
{
    private static (Workbook, Sheet) Grid(params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        foreach (var (id, value) in cells)
            sheet[id] = new NumberValue(value);
        return (workbook, sheet);
    }

    private static object? Calc(Workbook wb, Sheet sheet, string formula) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(new EvaluationContext(wb, sheet.Name)).AsObject();

    // Aspose oracle: SUM over the dynamic range INDEX(A1:A5,2):A4 == A2+A3+A4 == 20+30+40 == 90.
    [Test]
    public async Task Sum_OverDynamicRange_MatchesExcel()
    {
        var (wb, sheet) = Grid(("A1", 10), ("A2", 20), ("A3", 30), ("A4", 40), ("A5", 50));
        await Assert.That(Calc(wb, sheet, "=SUM(INDEX(A1:A5,2):A4)") as double?).IsEqualTo(90.0);
    }

    // Aspose oracle: COUNTIF(INDEX(A1:A5,2):A5, 30) over A2:A5 == 1.
    [Test]
    public async Task CountIf_OverDynamicRange_MatchesExcel()
    {
        var (wb, sheet) = Grid(("A1", 30), ("A2", 10), ("A3", 30), ("A4", 40), ("A5", 30));
        await Assert.That(Calc(wb, sheet, "=COUNTIF(INDEX(A1:A5,2):A5,30)") as double?).IsEqualTo(2.0);
    }

    // Corner rule: reversed endpoints normalise (B2:A1 == A1:B2). Sum of the 2x2 block.
    [Test]
    public async Task ReversedCorners_Normalise()
    {
        var (wb, sheet) = Grid(("A1", 1), ("B1", 2), ("A2", 3), ("B2", 4));
        await Assert.That(Calc(wb, sheet, "=SUM(OFFSET(B2,0,0):A1)") as double?).IsEqualTo(10.0);
    }

    // Array-form INDEX as an endpoint has no address -> #REF! (Excel parity).
    [Test]
    public async Task ArrayIndexEndpoint_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1));
        var result = Calc(wb, sheet, "=SUM(INDEX(ROW($A:$A),2):A1)");
        await Assert.That(result?.ToString()).IsEqualTo("#REF!");
    }

    // Regression: INDEX(ROW(...),k) still returns the row number (value path untouched).
    [Test]
    public async Task IndexIntoRowArray_StillReturnsRowNumber()
    {
        var (wb, sheet) = Grid();
        await Assert.That(Calc(wb, sheet, "=INDEX(ROW($A:$A),3)") as double?).IsEqualTo(3.0);
    }
}
```

> Confirm the `#REF!` string form via `AsObject()` — adjust the assertion to the actual error representation (`ComputedValue.Error` → `AsObject()`), using an existing error-returning test as the pattern.

- [ ] **Step 2: Run to verify (RED where behaviour is missing)**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/DynamicRangeEvaluationTests/*"`
Expected: the compat tests pass if Tasks 1-6 are correct; any failure points at the node to fix.

- [ ] **Step 3: Fix any failing behaviour**

Consumer wiring was done in Task 1 Step 4b (`NamedReferences.TryResolveRaw` routes through the virtual), so `SUM`/`COUNTIF` should already reach `DynamicRange.TryResolveReference`. If a compat test still fails, the defect is in the responsible node's resolution (span corners, INDEX target address, or the `#REF!` representation) — fix it there and re-run.

- [ ] **Step 4: Run the FULL suite (no regressions)**

Run: `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release`
Expected: all tests pass (record the baseline count before Task 1 and compare).

- [ ] **Step 5: Verify the original repros via AllocProbe**

Run: `dotnet run -c Release --project external/AllocProbe -- parse '=INDEX($D:$D,2):$D1' '=OFFSET(A1,1,0):B2' '=A1:B2'`
Expected: the first two now print `OK ... DynamicRange`; `A1:B2` still `RangeReference`.

- [ ] **Step 6: Commit**

```bash
git add tests/Danfma.MySheet.Tests/Expressions/DynamicRangeEvaluationTests.cs
git commit -m "test(expr): end-to-end and Excel-compat coverage for dynamic ranges"
```

---

## Self-Review notes

- **Spec coverage:** Component 1 → Tasks 1-4; Component 2 → Task 5; Component 3 → Task 6; Excel-compat/Aspose oracle → Task 7. INDIRECT is intentionally out of scope (function not implemented; it will light up automatically once implemented with a `TryResolveReference` override — noted for a fast-follow).
- **Consumer wiring:** range-consuming functions (SUM/COUNTIF/VLOOKUP) reach `DynamicRange.TryResolveReference` because Task 1 Step 4b routes `NamedReferences.TryResolveRaw` through the virtual. This also (correctly, per Excel) lets `INDEX`/`OFFSET`/`CHOOSE` be used as a reference argument anywhere `NamedReferences` resolves — a capability gain, verified by the full-suite regression run in Task 7 Step 4.
- **Open question (corner rule):** encoded as the `ReversedCorners_Normalise` test; confirm the expected sum against Aspose if in doubt.
