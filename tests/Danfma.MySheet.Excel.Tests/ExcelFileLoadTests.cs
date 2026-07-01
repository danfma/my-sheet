using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Excel.Tests;

public class ExcelFileLoadTests
{
    /// <summary>Writes an .xlsx fixture with ClosedXML (an independent implementation), loads it with
    /// our reader, runs the assertions and deletes the file.</summary>
    private static async Task WithFixture(Action<XLWorkbook> build, Func<Workbook, Task> assert)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-excel-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                build(fixture);
                fixture.SaveAs(path);
            }

            await assert(ExcelFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_ReadsLiterals_SheetNamesAndOrder()
    {
        await WithFixture(
            fixture =>
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 42.5;
                data.Cell("A2").Value = "hello";
                data.Cell("A3").Value = true;

                fixture.AddWorksheet("Summary").Cell("B2").Value = 7;
            },
            async workbook =>
            {
                await Assert.That(workbook["Data"].Index).IsEqualTo(0);
                await Assert.That(workbook["Summary"].Index).IsEqualTo(1);

                await Assert.That(workbook.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(42.5);
                await Assert.That(workbook.GetCellValue("Data", "A2").ToText()).IsEqualTo("hello");
                await Assert.That(workbook.GetCellValue("Data", "A3").TryGetBoolean(out var flag)).IsTrue();
                await Assert.That(flag).IsTrue();
                await Assert.That(workbook.GetCellValue("Summary", "B2").ToDouble()).IsEqualTo(7.0);
            }
        );
    }

    [Test]
    public async Task Load_ParsesFormulas_AndReevaluatesThem()
    {
        await WithFixture(
            fixture =>
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 2;
                data.Cell("A2").Value = 3;
                data.Cell("A3").FormulaA1 = "A1+A2";
                data.Cell("A4").FormulaA1 = "SUM(A1:A3)";
            },
            async workbook =>
            {
                // Formula cells load as real expression trees, not as their cached values.
                await Assert.That(workbook["Data"]["A3"]).IsTypeOf<BinaryOperation>();
                await Assert.That(workbook["Data"]["A4"]).IsTypeOf<Sum>();

                // And OUR engine re-evaluates them to the expected results.
                await Assert.That(workbook.GetCellValue("Data", "A3").ToDouble()).IsEqualTo(5.0);
                await Assert.That(workbook.GetCellValue("Data", "A4").ToDouble()).IsEqualTo(10.0);
            }
        );
    }

    [Test]
    public async Task Load_CrossSheetFormula()
    {
        await WithFixture(
            fixture =>
            {
                fixture.AddWorksheet("Data").Cell("A1").Value = 10;
                fixture.AddWorksheet("Summary").Cell("A1").FormulaA1 = "Data!A1*2";
            },
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("Summary", "A1").ToDouble()).IsEqualTo(20.0);
            }
        );
    }

    [Test]
    public async Task Load_AbsoluteReferences_AreNormalized()
    {
        await WithFixture(
            fixture =>
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 4;
                data.Cell("A2").Value = 6;
                data.Cell("B1").FormulaA1 = "$A$1+A2";
            },
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("Data", "B1").ToDouble()).IsEqualTo(10.0);
            }
        );
    }

    [Test]
    public async Task Load_ErrorLiteral()
    {
        await WithFixture(
            fixture =>
            {
                fixture.AddWorksheet("Data").Cell("A1").Value = XLError.DivisionByZero;
            },
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("Data", "A1").TryGetError(out var error)).IsTrue();
                await Assert.That(error).IsEqualTo(Error.DivZero);
            }
        );
    }

    [Test]
    public async Task Load_MissingCell_IsBlank()
    {
        await WithFixture(
            fixture =>
            {
                fixture.AddWorksheet("Data").Cell("A1").Value = 1;
            },
            async workbook =>
            {
                await Assert.That(workbook["Data"]["Z9"]).IsTypeOf<BlankValue>();
                await Assert.That(workbook.GetCellValue("Data", "Z9").Kind).IsEqualTo(ComputedValueKind.Blank);
            }
        );
    }
}
