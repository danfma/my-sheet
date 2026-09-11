using System.Reflection;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using MemoryPack;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 4 T1 (foundation): the <see cref="TableReference"/> node itself — resolution over the Phase 3
/// registry, the two invariants of <see cref="TableReference.Evaluate"/>, the union tag and the wire. The
/// parser does not emit the node yet (Phase 4 T5), so every tree here is built by hand; the region geometry
/// for the five non-<see cref="TableArea.Data"/> areas is Phase 5 T1's and is pinned here only as the
/// interim <c>#REF!</c> it answers today.
/// </summary>
public class TableReferenceTests
{
    // Data!Tabela1 = A1:B4, header row 1 (Item / Valor), data rows 2..4 = a,10 / b,20 / c,30, no totals row.
    private static Workbook Fixture()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["A2"] = new StringValue("a");
        data["B2"] = new NumberValue(10);
        data["A3"] = new StringValue("b");
        data["B3"] = new NumberValue(20);
        data["A4"] = new StringValue("c");
        data["B4"] = new NumberValue(30);
        workbook.DefineTable("Tabela1", "Data", "A1:B4", ["Item", "Valor"]);
        return workbook;
    }

    private static TableReference Node(
        string table = "Tabela1",
        string? column = "Valor",
        TableArea area = TableArea.Data
    ) => new(table, column, area);

    // === TryResolveRange: the one primitive ==============================================================

    [Test]
    public async Task Data_WithAColumn_ResolvesToTheColumnsDataRows()
    {
        var ok = Node().TryResolveRange(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // T[#Data] and T[] (Aspose rewrites the latter to the bare table name): the whole data body.
    [Test]
    public async Task Data_WithNoColumn_ResolvesToTheWholeDataBody()
    {
        var ok = Node(column: null).TryResolveRange(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("A2", "B4", "Data"));
    }

    // Excel resolves table and column names case-insensitively (Table1[col] -> Table1[Col]).
    [Test]
    public async Task TableAndColumnNames_ResolveCaseInsensitively()
    {
        var ok = Node("TABELA1", "valor").TryResolveRange(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // Unknown table -> #NAME?, the same name space as a defined name (NameReference).
    [Test]
    public async Task UnknownTable_IsName()
    {
        var ok = Node("NoSuch").TryResolveRange(Fixture(), out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Name);
    }

    // Unknown column -> #REF!, Excel's own repair (a deleted column's specifier becomes Table1[#REF!]).
    [Test]
    public async Task UnknownColumn_IsRef()
    {
        var ok = Node(column: "NoSuch").TryResolveRange(Fixture(), out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // Phase 5 ruling R1 (recorded DIVERGENCE, sweep item 33): the oracle treats a header-only table as an
    // EMPTY reference, which this engine has no node for, so Data over zero data rows answers #REF!.
    [Test]
    [Arguments("Valor")]
    [Arguments(null)]
    public async Task HeaderOnlyTable_DataIsRef_ARecordedDivergence(string? column)
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Data");
        workbook.DefineTable("Vazia", "Data", "A1:B1", ["Item", "Valor"]);

        var ok = Node("Vazia", column).TryResolveRange(workbook, out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // INTERIM: Phase 3 shipped Table.TryGetColumnRange for [#Data] only. The other five areas answer #REF!
    // until Phase 5 T1 lands Table.TryGetRegion — this pin goes red on purpose when it does, and the task
    // that lands it replaces this test with the measured geometry (Phase 4 ruling 2's table).
    [Test]
    [Arguments(TableArea.All)]
    [Arguments(TableArea.Headers)]
    [Arguments(TableArea.Totals)]
    [Arguments(TableArea.HeadersAndData)]
    [Arguments(TableArea.DataAndTotals)]
    public async Task NonDataAreas_AreRef_UntilPhase5TryGetRegion(TableArea area)
    {
        var ok = Node(column: null, area: area)
            .TryResolveRange(Fixture(), out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // === Evaluate: the two invariants ====================================================================

    // (a) Evaluate returns the CONCRETE resolved range, never ComputedValue.Reference(this).
    [Test]
    public async Task Evaluate_ReturnsTheConcreteRange_NotItself()
    {
        var value = Node().Evaluate(new EvaluationContext(Fixture()));

        await Assert.That(value.TryGetReference(out var reference)).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // The measured reason for (a): a probe returning Reference(this) made SUM(node) answer 0, because
    // EnumerateValues' catch-all `case Reference` yields the reference back as one non-numeric element.
    [Test]
    public async Task Sum_OverTheNode_ExpandsTheResolvedRange()
    {
        var sum = new Sum([Node()]);

        var value = sum.Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value as double?).IsEqualTo(60.0);
    }

    // (b) Evaluate never answers #VALUE! the way RangeReference.Evaluate does, and (Phase 5 ruling R2) an
    // unresolvable table is an error VALUE that flows to the consumer, never a throw or a short-circuit.
    [Test]
    public async Task Evaluate_UnknownTable_IsTheNameErrorValue()
    {
        var context = new EvaluationContext(Fixture());

        var bare = Node("NoSuch").Evaluate(context).AsObject();
        var summed = new Sum([Node("NoSuch")]).Evaluate(context).AsObject();

        await Assert.That(bare).IsEqualTo(ErrorValue.Name);
        await Assert.That(summed).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task Evaluate_UnknownColumn_IsTheRefErrorValue()
    {
        var value = Node(column: "NoSuch").Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value).IsEqualTo(ErrorValue.Reference);
    }

    // === TryResolveReference ============================================================================

    [Test]
    public async Task TryResolveReference_ResolvesToTheRange()
    {
        var ok = Node().TryResolveReference(new EvaluationContext(Fixture()), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    [Test]
    public async Task TryResolveReference_Unresolvable_IsFalse()
    {
        var ok = Node("NoSuch")
            .TryResolveReference(new EvaluationContext(Fixture()), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }

    // Deliberately NOT overridden: a table's target comes from REGISTERED state, invalidated by the
    // definitions version bump, not by per-pass volatility — overriding it would make DependencyExtractor
    // mark every structured-reference formula always-dirty and throw away its static RangeDep.
    [Test]
    public async Task IsVolatile_IsNotOverridden()
    {
        await Assert.That(Node().IsVolatile).IsFalse();
    }

    // === The enum and the wire ===========================================================================

    // Data = 0 so the overwhelmingly common T[Col] serializes the enum's default byte; the two pairs are
    // appended after the four singletons. Phase 4 ruling 1.
    [Test]
    public async Task TableArea_IsAByteEnum_WithDataAsDefault()
    {
        var members = Enum.GetValues<TableArea>()
            .Select(area => $"{area}={Convert.ToByte(area)}")
            .ToArray();

        await Assert.That(Enum.GetUnderlyingType(typeof(TableArea))).IsEqualTo(typeof(byte));
        await Assert
            .That(members)
            .IsEquivalentTo([
                "Data=0",
                "All=1",
                "Headers=2",
                "Totals=3",
                "HeadersAndData=4",
                "DataAndTotals=5",
            ]);
        await Assert.That(default(TableArea).ToString()).IsEqualTo("Data");
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesEveryMember()
    {
        var node = new TableReference("Tabela1", "Sales Amount", TableArea.HeadersAndData);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = MemoryPackSerializer.Deserialize<Expression>(bytes);

        await Assert.That(back).IsEqualTo(node);
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesANullColumn()
    {
        var node = new TableReference("Tabela1", null, TableArea.All);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = (TableReference)MemoryPackSerializer.Deserialize<Expression>(bytes)!;

        await Assert.That(back.ColumnName).IsNull();
        await Assert.That(back.Area).IsEqualTo(TableArea.All);
    }

    // Union tags are APPEND-ONLY. 322 is Aggregate (Phase 2), 323-326 are Phase 7's producers, so the
    // node took 327 — the number the design wrote (322) was already live. A duplicate tag fails at
    // MemoryPack type initialization, not at compile time, which is why the tag is pinned by value.
    [Test]
    public async Task UnionTag_Is327()
    {
        var tag = typeof(Expression)
            .GetCustomAttributes<MemoryPackUnionAttribute>()
            .Single(attribute => attribute.Type == typeof(TableReference))
            .Tag;

        await Assert.That(tag).IsEqualTo((ushort)327);
    }
}
