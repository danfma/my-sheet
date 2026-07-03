using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

// Covers the 3.0 write-choke-point's second mutation path: Sheet.Remove. Reads/writes stay covered by the
// existing suites; these lock down Remove's return value, its effect on the read surface, and — crucially —
// that it shares the SAME explicit-invalidation semantics as a write (a mutation is not observed until the
// host calls InvalidateCache).
public class SheetRemoveTests
{
    [Test]
    public async Task Remove_ExistingCell_ReturnsTrueAndDisappearsFromReadSurface()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(10);

        await Assert.That(sheet.ContainsKey("A1")).IsTrue();

        var removed = sheet.Remove("A1");

        await Assert.That(removed).IsTrue();
        await Assert.That(sheet.ContainsKey("A1")).IsFalse();
        await Assert.That(sheet.TryGetValue("A1", out _)).IsFalse();
        await Assert.That(sheet.Keys.Contains("A1")).IsFalse();
        await Assert.That(sheet.Count).IsEqualTo(0);
        // The read indexer keeps its blank-for-missing contract.
        await Assert.That(sheet["A1"] is BlankValue).IsTrue();
    }

    [Test]
    public async Task Remove_MissingCell_ReturnsFalse()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(10);

        var removed = sheet.Remove("Z99");

        await Assert.That(removed).IsFalse();
        // The unrelated cell is untouched.
        await Assert.That(sheet.ContainsKey("A1")).IsTrue();
        await Assert.That(sheet.Count).IsEqualTo(1);
    }

    [Test]
    public async Task Remove_DependentFormula_SeesBlankAfterRecalculation()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        sheet["S1"] = ExpressionParser.Parse("=SUM(A1:A3)", sheet);

        await Assert.That(sheet["S1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(60.0);

        var removed = sheet.Remove("A2");
        await Assert.That(removed).IsTrue();

        // Coherent with a write: a mutation is NOT observed until the cache is invalidated, so the dependent
        // formula still serves its memoized value (exactly as sheet["A2"] = ... would before invalidation).
        await Assert.That(sheet["S1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(60.0);

        workbook.InvalidateCache();

        // After recalculation the removed cell is gone: it reads blank and the SUM drops its contribution.
        await Assert
            .That(workbook.GetCellValue("Sheet1", "A2").Kind)
            .IsEqualTo(ComputedValueKind.Blank);
        await Assert.That(sheet["S1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(40.0);
    }
}
