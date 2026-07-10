using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Guards the numeric cell-key lever (string-hygiene 3.x, Phase 2): a <see cref="Sheet"/>'s cell store is
/// numeric-keyed <c>(column,row)</c> in memory (with a non-A1 overflow), collapsing the ~35MB of duplicate id
/// strings the old <c>Dictionary&lt;string,Expression&gt;</c> retained at K1 scale — while the wire format and
/// the observable string surface stay untouched. The wire golden proves byte-identity; the round-trip and
/// overflow tests prove behaviour (including non-address keys) survives Save/Load.
/// </summary>
public class CellStoreTests
{
    // Deterministic serialization of a representative workbook (multi-sheet, cross-sheet formulas, multi-letter
    // columns AA/AB, high rows), captured from the build BEFORE _cells became numeric-keyed. The in-memory key
    // changed but the formatter re-emits the historical string→Expression map in the same order, so the bytes
    // must be identical.
    private const string PreChangeCellsWireGolden =
        "AgIAAAD7////BAAAAERhdGED+////wQAAABEYXRhAAAAAAcAAAD9////AgAAAEExAQEAAAAAAAAkQP3///8CAAAARDEAAfr/"
        + "//8FAAAAYXBwbGX9////AgAAAEQyAAH5////BgAAAGJhbmFuYf3///8CAAAARDMAAfn///8GAAAAY2hlcnJ5/f///wIAAABF"
        + "MQEBAAAAAAAA8D/9////AgAAAEUyAQEAAAAAAAAAQP3///8CAAAARTMBAQAAAAAAAAhA+////wQAAABNYWluA/v///8EAAAA"
        + "TWFpbgEAAAAKAAAA/f///wIAAABBMQEBAAAAAAAAJED9////AgAAAEEyAQEAAAAAAAA0QP3///8CAAAAQTMBAQAAAAAAAD5A"
        + "/f///wIAAABCMQgDAQAAAAgDAAAAAAQC/f///wIAAABBMfv///8EAAAATWFpbggDAgAAAAQC/f///wIAAABBMvv///8EAAAA"
        + "TWFpbgEBAAAAAAAAAEAIAwQAAAABAQAAAAAAAAhAAQEAAAAAAAAAQP3///8CAAAAQjIGAQEAAAAFB/3///8CAAAAQTH9////"
        + "AgAAAEEz+////wQAAABNYWluAwAAAAEAAAABAAAAAQAAAP3///8CAAAAQjMKAQEAAAAFB/3///8CAAAAQTH9////AgAAAEEz"
        + "+////wQAAABNYWluAwAAAAEAAAABAAAAAQAAAP3///8CAAAAQjcvAQQAAAAAAfn///8GAAAAYmFuYW5hBQf9////AgAAAEQx"
        + "/f///wIAAABFM/v///8EAAAARGF0YQMAAAACAAAAAQAAAAQAAAABAQAAAAAAAABAAgEA/P///wMAAABCMTEIAwIAAAAEAv3/"
        + "//8CAAAAQTH7////BAAAAERhdGEBAQAAAAAAAAhA+v///wUAAABBQTEwMAEBAAAAAAAARUD8////AwAAAEFCNQEBAAAAAAAA"
        + "HEAAAAAA";

    // The workbook the golden was serialized from — mirrors the exact SetCell sequence.
    private static Workbook BuildWireFixture()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["D1"] = Expression.String("apple");
        data["D2"] = Expression.String("banana");
        data["D3"] = Expression.String("cherry");
        data["E1"] = new NumberValue(1);
        data["E2"] = new NumberValue(2);
        data["E3"] = new NumberValue(3);

        var main = wb.Sheets.Add("Main");
        main["A1"] = new NumberValue(10);
        main["A2"] = new NumberValue(20);
        main["A3"] = new NumberValue(30);
        main["B1"] = ExpressionParser.Parse("=A1+A2*2-3^2", main);
        main["B2"] = ExpressionParser.Parse("=SUM(A1:A3)", main);
        main["B3"] = ExpressionParser.Parse("=AVERAGE(A1:A3)", main);
        main["B7"] = ExpressionParser.Parse("=VLOOKUP(\"banana\", Data!D1:E3, 2, FALSE)", main);
        main["B11"] = ExpressionParser.Parse("=Data!A1*3", main);
        main["AA100"] = new NumberValue(42);
        main["AB5"] = new NumberValue(7);
        return wb;
    }

    [Test]
    public async Task Wire_IsByteIdentical_AfterNumericKeys()
    {
        var bytes = MemoryPackSerializer.Serialize(BuildWireFixture());

        await Assert.That(Convert.ToBase64String(bytes)).IsEqualTo(PreChangeCellsWireGolden);
    }

    [Test]
    public async Task RoundTrip_NewToNew_IsByteStable()
    {
        var first = MemoryPackSerializer.Serialize(BuildWireFixture());
        var restored = MemoryPackSerializer.Deserialize<Workbook>(first)!;
        var second = MemoryPackSerializer.Serialize(restored);

        await Assert.That(Convert.ToBase64String(second)).IsEqualTo(Convert.ToBase64String(first));
    }

    [Test]
    public async Task RoundTrip_PreservesCellsAndEvaluatesIdentically()
    {
        var wb = BuildWireFixture();
        var restored = MemoryPackSerializer.Deserialize<Workbook>(
            MemoryPackSerializer.Serialize(wb)
        )!;

        var main = restored.Sheets["Main"];

        // A1 addresses (including multi-letter columns and high rows) survive with their exact string ids.
        await Assert.That(main.ContainsKey("A1")).IsTrue();
        await Assert.That(main.ContainsKey("AA100")).IsTrue();
        await Assert.That(main.ContainsKey("AB5")).IsTrue();
        await Assert.That(main.Count).IsEqualTo(wb.Sheets["Main"].Count);

        await Assert.That(main["B2"].Evaluate(restored).AsObject() as double?).IsEqualTo(60.0);
        await Assert.That(main["B11"].Evaluate(restored).AsObject() as double?).IsEqualTo(30.0);
        await Assert
            .That(restored.GetCellValue("Main", "AA100").AsObject() as double?)
            .IsEqualTo(42.0);
    }

    [Test]
    public async Task Enumeration_YieldsExactA1Ids()
    {
        var sheet = new Workbook().Sheets.Add("S");
        sheet["A1"] = new NumberValue(1);
        sheet["AA100"] = new NumberValue(2);
        sheet["XFD1048576"] = new NumberValue(3);

        var keys = sheet.Keys.ToHashSet();

        await Assert.That(keys).Contains("A1");
        await Assert.That(keys).Contains("AA100");
        await Assert.That(keys).Contains("XFD1048576");
        await Assert.That(keys.Count).IsEqualTo(3);
    }

    [Test]
    public async Task NonA1Keys_RouteToOverflow_PreservingExactString()
    {
        var wb = new Workbook();
        var sheet = wb.Sheets.Add("S");

        // Not canonical A1: a plain word, a lowercase form, a leading-zero row, a bare row-0. Each must keep its
        // exact key on the public surface (an A1-lenient parse would rewrite them to a canonical id).
        sheet["hello"] = new NumberValue(1);
        sheet["a1"] = new NumberValue(2);
        sheet["A01"] = new NumberValue(3);
        sheet["A0"] = new NumberValue(4);
        sheet["A1"] = new NumberValue(5); // canonical, distinct from "a1"/"A01"

        var restored = MemoryPackSerializer.Deserialize<Workbook>(
            MemoryPackSerializer.Serialize(wb)
        )!;
        var s = restored.Sheets["S"];

        await Assert.That(s.Count).IsEqualTo(5);
        await Assert.That((s["hello"] as NumberValue)?.Value).IsEqualTo(1.0);
        await Assert.That((s["a1"] as NumberValue)?.Value).IsEqualTo(2.0);
        await Assert.That((s["A01"] as NumberValue)?.Value).IsEqualTo(3.0);
        await Assert.That((s["A0"] as NumberValue)?.Value).IsEqualTo(4.0);
        await Assert.That((s["A1"] as NumberValue)?.Value).IsEqualTo(5.0);

        var keys = s.Keys.ToHashSet();
        await Assert.That(keys).Contains("hello");
        await Assert.That(keys).Contains("a1");
        await Assert.That(keys).Contains("A01");
        await Assert.That(keys).Contains("A0");
        await Assert.That(keys).Contains("A1");
    }

    [Test]
    public async Task Remove_DropsCell_ForBothDenseAndOverflow()
    {
        var sheet = new Workbook().Sheets.Add("S");
        sheet["B2"] = new NumberValue(1);
        sheet["note"] = new NumberValue(2);

        await Assert.That(sheet.Remove("B2")).IsTrue();
        await Assert.That(sheet.Remove("note")).IsTrue();
        await Assert.That(sheet.Remove("Z9")).IsFalse();
        await Assert.That(sheet.Count).IsEqualTo(0);
        await Assert.That(sheet.ContainsKey("B2")).IsFalse();
        await Assert.That(sheet.ContainsKey("note")).IsFalse();
    }

    [Test]
    public async Task Overwrite_KeepsSingleCell_AndUpdatesValue()
    {
        var sheet = new Workbook().Sheets.Add("S");
        sheet["C3"] = new NumberValue(1);
        sheet["C3"] = new NumberValue(9);

        await Assert.That(sheet.Count).IsEqualTo(1);
        await Assert.That((sheet["C3"] as NumberValue)?.Value).IsEqualTo(9.0);
    }

    [Test]
    public async Task OpenRange_OverNumericCells_StillAggregates()
    {
        var wb = new Workbook();
        var sheet = wb.Sheets.Add("S");
        for (var row = 1; row <= 5; row++)
        {
            sheet["A" + row] = new NumberValue(row);
        }

        var total = ExpressionParser.Parse("=SUM(A:A)", sheet).Evaluate(wb).AsObject() as double?;

        await Assert.That(total).IsEqualTo(15.0);
    }

    // Guards the CellStoreFormatter.Deserialize read-side literal sharing (perf follow-up to the K1 gcdump
    // finding: a .xlsx load dedupes StringValue via WorksheetStreamLoader, but a Save+Load round trip through
    // .myxl was materializing a fresh instance per cell, undoing that sharing — measured 209 -> 60,008
    // distinct StringValue instances on the K1-synthetic fixture). Repeated text across many cells must come
    // back as the SAME instance after a round trip, and BooleanValue/well-known ErrorValue must come back as
    // the existing singletons.
    [Test]
    public async Task RoundTrip_SharesRepeatedLiteralInstances()
    {
        var wb = new Workbook();
        var sheet = wb.Sheets.Add("S");

        for (var row = 1; row <= 50; row++)
        {
            sheet["A" + row] = Expression.String("repeated-text");
            sheet["B" + row] = row % 2 == 0 ? BooleanValue.True : BooleanValue.False;
            sheet["C" + row] = ErrorValue.DivByZero;
        }

        var restored = MemoryPackSerializer.Deserialize<Workbook>(
            MemoryPackSerializer.Serialize(wb)
        )!;
        var restoredSheet = restored.Sheets["S"];

        var stringInstances = new HashSet<Danfma.MySheet.Expressions.StringValue>(
            ReferenceEqualityComparer.Instance
        );
        var booleanInstances = new HashSet<BooleanValue>(ReferenceEqualityComparer.Instance);
        var errorInstances = new HashSet<ErrorValue>(ReferenceEqualityComparer.Instance);

        for (var row = 1; row <= 50; row++)
        {
            stringInstances.Add((Danfma.MySheet.Expressions.StringValue)restoredSheet["A" + row]);
            booleanInstances.Add((BooleanValue)restoredSheet["B" + row]);
            errorInstances.Add((ErrorValue)restoredSheet["C" + row]);
        }

        await Assert.That(stringInstances.Count).IsEqualTo(1);
        await Assert.That(booleanInstances.Count).IsEqualTo(2); // True and False singletons
        await Assert.That(errorInstances.Count).IsEqualTo(1);

        // The singletons are the actual shared instances, not merely equal values.
        await Assert.That(booleanInstances).Contains(BooleanValue.True);
        await Assert.That(booleanInstances).Contains(BooleanValue.False);
        await Assert.That(errorInstances).Contains(ErrorValue.DivByZero);
    }
}
