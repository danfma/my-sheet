using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Guards the SheetName interning lever (string-hygiene 3.x, Phase 1): a qualified sheet name must resolve to
/// ONE shared instance across every reference that names the same sheet, both after a parse and after a Load,
/// while the wire format and the observable (case-preserving) behaviour stay untouched. At K1 scale the naive
/// per-node copy of the qualifier is ~24MB of duplicate strings; interning collapses it to one instance.
/// </summary>
public class SheetNameInterningTests
{
    // Deterministic pre-change serialization of a workbook exercising all three SheetName-bearing records
    // cross-sheet (CellReference, RangeReference, OpenRangeReference). Captured from the build BEFORE the
    // InternStringFormatter was applied — the interning is read-side only, so the bytes must be identical.
    private const string PreInterningWireGolden =
        "AgIAAAD7////BAAAAERhdGED+////wQAAABEYXRhAAAAAAQAAAD9////AgAAAEExAQEAAAAAAAAkQP3///8CAAAAQTIB"
        + "AQAAAAAAADRA/f///wIAAABBMwEBAAAAAAAAPkD9////AgAAAEIxAQEAAAAAAADwP/v///8EAAAATWFpbgP7////BAAA"
        + "AE1haW4BAAAABAAAAP3///8CAAAAQzEIAwIAAAAEAv3///8CAAAAQTH7////BAAAAERhdGEBAQAAAAAAAAhA/f///wIA"
        + "AABDMgYBAQAAAAUH/f///wIAAABBMf3///8CAAAAQTP7////BAAAAERhdGEDAAAAAQAAAAEAAAABAAAA/f///wIAAABD"
        + "MwYBAQAAAPo8AQUBAAAAAQAAAAEAAAABAAAAAAAAAAAAAAAAAAAAAAAAAPv///8EAAAARGF0Yf3///8CAAAAQzQIAwAA"
        + "AAAIAwAAAAAEAv3///8CAAAAQTH7////BAAAAERhdGEEAv3///8CAAAAQTL7////BAAAAERhdGEEAv3///8CAAAAQTP7"
        + "////BAAAAERhdGEAAAAA";

    // The workbook the golden was serialized from — cross-sheet CellReference, RangeReference and
    // OpenRangeReference, all naming "Data".
    private static Workbook BuildWireFixture()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["A2"] = new NumberValue(20);
        data["A3"] = new NumberValue(30);
        data["B1"] = new NumberValue(1);

        var main = wb.Sheets.Add("Main");
        main["C1"] = ExpressionParser.Parse("=Data!A1*3", main); // CellReference
        main["C2"] = ExpressionParser.Parse("=SUM(Data!A1:A3)", main); // RangeReference
        main["C3"] = ExpressionParser.Parse("=SUM(Data!A:A)", main); // OpenRangeReference
        main["C4"] = ExpressionParser.Parse("=Data!A1+Data!A2+Data!A3", main);
        return wb;
    }

    [Test]
    public async Task Wire_IsByteIdentical_AfterInterning()
    {
        var bytes = MemoryPackSerializer.Serialize(BuildWireFixture());

        await Assert.That(Convert.ToBase64String(bytes)).IsEqualTo(PreInterningWireGolden);
    }

    [Test]
    public async Task Parse_CellReference_SharesSheetNameInstance()
    {
        var sheet = new Workbook().Sheets.Add("S");
        var a = (CellReference)ExpressionParser.Parse("=Data!A1", sheet);
        var b = (CellReference)ExpressionParser.Parse("=Data!Z99", sheet);

        await Assert.That(ReferenceEquals(a.SheetName, b.SheetName)).IsTrue();
    }

    [Test]
    public async Task Parse_RangeReference_SharesSheetNameInstance()
    {
        var sheet = new Workbook().Sheets.Add("S");
        var a = (RangeReference)ExpressionParser.Parse("=Data!A1:A3", sheet);
        var b = (RangeReference)ExpressionParser.Parse("=Data!B2:C4", sheet);

        await Assert.That(ReferenceEquals(a.SheetName, b.SheetName)).IsTrue();
    }

    [Test]
    public async Task Parse_OpenRangeReference_SharesSheetNameInstance()
    {
        var sheet = new Workbook().Sheets.Add("S");
        var a = (OpenRangeReference)ExpressionParser.Parse("=Data!A:A", sheet);
        var b = (OpenRangeReference)ExpressionParser.Parse("=Data!C:D", sheet);

        await Assert.That(ReferenceEquals(a.SheetName, b.SheetName)).IsTrue();
    }

    [Test]
    public async Task Parse_And_Load_ConvergeOnSameInstance()
    {
        var sheet = new Workbook().Sheets.Add("S");
        var parsed = (CellReference)ExpressionParser.Parse("=Data!A1", sheet);

        var wb = new Workbook();
        wb.Sheets.Add("Data")["A1"] = new NumberValue(1);
        var main = wb.Sheets.Add("Main");
        main["C1"] = ExpressionParser.Parse("=Data!A1", main);
        var restored = MemoryPackSerializer.Deserialize<Workbook>(MemoryPackSerializer.Serialize(wb))!;
        var loaded = (CellReference)restored.Sheets["Main"]["C1"];

        // Same intern pool on both paths: a parsed name and a loaded name are the SAME reference.
        await Assert.That(ReferenceEquals(parsed.SheetName, loaded.SheetName)).IsTrue();
    }

    [Test]
    public async Task Load_SharesSheetNameInstance_AcrossManyReferences()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data")["A1"] = new NumberValue(1);
        var main = wb.Sheets.Add("Main");
        for (var i = 1; i <= 50; i++)
        {
            main["C" + i] = ExpressionParser.Parse("=Data!A1", main);
        }

        var restored = MemoryPackSerializer.Deserialize<Workbook>(MemoryPackSerializer.Serialize(wb))!;
        var loadedMain = restored.Sheets["Main"];

        var first = ((CellReference)loadedMain["C1"]).SheetName;
        for (var i = 2; i <= 50; i++)
        {
            var name = ((CellReference)loadedMain["C" + i]).SheetName;
            await Assert.That(ReferenceEquals(first, name)).IsTrue();
        }
    }

    [Test]
    public async Task Parse_PreservesQualifierCasing_Verbatim()
    {
        var sheet = new Workbook().Sheets.Add("S");

        // Ordinal interning keeps the exact token text — no canonicalization to a first-seen casing.
        var lower = (CellReference)ExpressionParser.Parse("=data!A1", sheet);
        var upper = (CellReference)ExpressionParser.Parse("=DATA!A1", sheet);

        await Assert.That(lower.SheetName).IsEqualTo("data");
        await Assert.That(upper.SheetName).IsEqualTo("DATA");
    }

    [Test]
    public async Task Resolution_StaysCaseInsensitive_AfterInterning()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data")["A1"] = new NumberValue(42);
        var main = wb.Sheets.Add("Main");

        // Differently cased qualifier still resolves to the same sheet (lookup is OrdinalIgnoreCase).
        var value = ExpressionParser.Parse("=data!A1", main).Evaluate(wb).AsObject() as double?;

        await Assert.That(value).IsEqualTo(42.0);
    }
}
