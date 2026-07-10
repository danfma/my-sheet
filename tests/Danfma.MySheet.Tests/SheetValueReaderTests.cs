using Danfma.MySheet.Parsing;
using MySheetString = Danfma.MySheet.Expressions.StringValue;
using NumberValue = Danfma.MySheet.Expressions.NumberValue;

namespace Danfma.MySheet.Tests;

/// <summary>
/// The numeric-address reader must be a faster address form of GetCellValue — identical results for
/// literals, formulas, blanks and misses (on-demand evaluation), and the same manual-invalidation
/// contract.
/// </summary>
public class SheetValueReaderTests
{
    private static Workbook BuildWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");

        sheet["A1"] = new NumberValue(21);
        sheet["A2"] = ExpressionParser.Parse("=A1*2", sheet);
        sheet["B1"] = new MySheetString("hello");

        return workbook;
    }

    [Test]
    public async Task GetValue_MatchesGetCellValue_ForEveryKind()
    {
        var workbook = BuildWorkbook();
        var reader = workbook.GetValueReader("Main");

        // Formula (miss → on-demand evaluation), literal number, literal text and blank all match
        // the string-keyed path.
        await Assert.That(reader.GetValue(1, 2).ToDouble()).IsEqualTo(42.0);
        await Assert
            .That(reader.GetValue(1, 1).ToDouble())
            .IsEqualTo(workbook.GetCellValue("Main", "A1").ToDouble());
        await Assert
            .That(reader.GetValue(2, 1).ToText())
            .IsEqualTo(workbook.GetCellValue("Main", "B1").ToText());
        await Assert
            .That(reader.GetValue(5, 5).Kind)
            .IsEqualTo(workbook.GetCellValue("Main", "E5").Kind);
    }

    [Test]
    public async Task GetValue_ServesLiteralsNeverReadBefore()
    {
        // The pitfall a dense enumerator would have: literals only enter the value store when read.
        // The reader's miss path evaluates on demand, so an untouched literal is still served.
        var workbook = BuildWorkbook();
        var reader = workbook.GetValueReader("Main");

        await Assert.That(reader.GetValue(2, 1).ToText()).IsEqualTo("hello");
    }

    [Test]
    public async Task GetValue_HonorsInvalidateCache()
    {
        var workbook = BuildWorkbook();
        var reader = workbook.GetValueReader("Main");

        await Assert.That(reader.GetValue(1, 2).ToDouble()).IsEqualTo(42.0);

        workbook["Main"]["A1"] = new NumberValue(100);
        workbook.InvalidateCache();

        // The same reader instance stays valid across invalidation (the handle is stable).
        await Assert.That(reader.GetValue(1, 2).ToDouble()).IsEqualTo(200.0);
    }
}
