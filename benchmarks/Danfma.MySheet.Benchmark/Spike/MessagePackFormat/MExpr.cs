using Danfma.MySheet.Expressions;
using MessagePack;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// ── MessagePack mirror of a REPRESENTATIVE SUBSET of the Expression tree ──────────────────────────────
// This is the spike prototype for plans/messagepack-spike.md, question 3 (technical viability). It mirrors
// ~35 of the ~320 production Expression nodes with MessagePack-CSharp attributes:
//   • [MessagePackObject] on every concrete node (positional record + [property: Key(n)] integer keys);
//   • [Union(tag, typeof(...))] on the abstract base, reusing the SAME append-only tags as the MemoryPack
//     [MemoryPackUnion] table in Danfma.MySheet/Expressions/Expression.cs — proving the identical tag scheme
//     is portable (a full migration would script the [Union] list from the MemoryPack tags, same numbers).
// The subset covers exactly the node types the three spike payloads emit (values, references, operations,
// aggregates, logical, lookup, stat-if, text, scalar math, LET, FunctionCall) so the converter never meets
// an unmapped type. Reuses the PRODUCTION BinaryOperator/UnaryOperator enums (public, serialize by value).
[MessagePackObject]
[Union(0, typeof(MStringValue))]
[Union(1, typeof(MNumberValue))]
[Union(2, typeof(MBooleanValue))]
[Union(3, typeof(MBlankValue))]
[Union(4, typeof(MCellReference))]
[Union(5, typeof(MRangeReference))]
[Union(6, typeof(MSum))]
[Union(8, typeof(MBinaryOperation))]
[Union(9, typeof(MUnaryOperation))]
[Union(10, typeof(MAverage))]
[Union(11, typeof(MMin))]
[Union(12, typeof(MMax))]
[Union(13, typeof(MCount))]
[Union(14, typeof(MIf))]
[Union(18, typeof(MIfError))]
[Union(19, typeof(MFunctionCall))]
[Union(21, typeof(MRound))]
[Union(23, typeof(MAbs))]
[Union(27, typeof(MUpper))]
[Union(31, typeof(MLeft))]
[Union(34, typeof(MConcat))]
[Union(39, typeof(MCountIf))]
[Union(41, typeof(MSumIf))]
[Union(45, typeof(MMatch))]
[Union(46, typeof(MIndex))]
[Union(47, typeof(MVLookup))]
[Union(50, typeof(MNameReference))]
[Union(51, typeof(MLet))]
[Union(54, typeof(MUnionReference))]
[Union(167, typeof(MChoose))]
[Union(176, typeof(MAverageIf))]
[Union(191, typeof(MSmall))]
[Union(316, typeof(MOpenRangeReference))]
public abstract partial record MExpr;

// ── Values ────────────────────────────────────────────────────────────────────────────────────────────
[MessagePackObject]
public sealed partial record MStringValue([property: Key(0)] string Value) : MExpr;

[MessagePackObject]
public sealed partial record MNumberValue([property: Key(0)] double Value) : MExpr;

[MessagePackObject]
public sealed partial record MBooleanValue([property: Key(0)] bool Value) : MExpr;

[MessagePackObject]
public sealed partial record MBlankValue : MExpr;

// ── References ──────────────────────────────────────────────────────────────────────────────────────
[MessagePackObject]
public sealed partial record MCellReference(
    [property: Key(0)] string Id,
    [property: Key(1)] string SheetName
) : MExpr;

[MessagePackObject]
public sealed partial record MRangeReference(
    [property: Key(0)] string StartId,
    [property: Key(1)] string EndId,
    [property: Key(2)] string SheetName
) : MExpr;

[MessagePackObject]
public sealed partial record MOpenRangeReference(
    [property: Key(0)] int? ColMin,
    [property: Key(1)] int? ColMax,
    [property: Key(2)] int? RowMin,
    [property: Key(3)] int? RowMax,
    [property: Key(4)] string SheetName
) : MExpr;

[MessagePackObject]
public sealed partial record MUnionReference([property: Key(0)] MExpr[] Areas) : MExpr;

[MessagePackObject]
public sealed partial record MNameReference([property: Key(0)] string Name) : MExpr;

// ── Operations (reuse production enums) ───────────────────────────────────────────────────────────────
[MessagePackObject]
public sealed partial record MBinaryOperation(
    [property: Key(0)] BinaryOperator Operator,
    [property: Key(1)] MExpr Left,
    [property: Key(2)] MExpr Right
) : MExpr;

[MessagePackObject]
public sealed partial record MUnaryOperation(
    [property: Key(0)] UnaryOperator Operator,
    [property: Key(1)] MExpr Operand
) : MExpr;

// ── Aggregates / functions ─────────────────────────────────────────────────────────────────────────
// The vast majority of production function nodes share the exact shape `record Foo(Expression[] Arguments)`,
// so each mirror is a one-line record carrying a single Key(0) MExpr[]. Sum is the lone outlier (its member
// is named `Expressions`), mirrored here with the same wire shape.
[MessagePackObject]
public sealed partial record MSum([property: Key(0)] MExpr[] Expressions) : MExpr;

[MessagePackObject]
public sealed partial record MAverage([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MMin([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MMax([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MCount([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MIf([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MIfError([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MRound([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MAbs([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MUpper([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MLeft([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MConcat([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MCountIf([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MSumIf([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MAverageIf([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MMatch([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MIndex([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MVLookup([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MChoose([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MSmall([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MLet([property: Key(0)] MExpr[] Arguments) : MExpr;

[MessagePackObject]
public sealed partial record MFunctionCall(
    [property: Key(0)] string Name,
    [property: Key(1)] MExpr[] Arguments
) : MExpr;
