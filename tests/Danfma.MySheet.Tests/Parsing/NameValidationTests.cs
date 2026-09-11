using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// The defined-name validator (<c>NamedReferences.IsValidName</c>) and the table-name validator
/// (<c>Table.IsValidName</c>) share one namespace (Excel's Name Manager rule), so they must also share ONE
/// notion of "looks like a cell reference": Excel reserves a name only when it is a real address into its
/// grid (<c>Parser.IsExcelGridCellReference</c>: 1-3 ASCII letters, a row up to 1,048,576). Before the
/// defined-name validator was pointed at that predicate it used the parser's UNBOUNDED
/// <c>IsCellReference</c>, so Excel's own default table names (<c>Tabela1</c>, <c>Table1</c>, <c>表1</c>)
/// could be registered as a table but not as a name. The pinned property is the agreement itself, over
/// the names that told the two predicates apart, plus one hand-written row per side so the property
/// cannot pass vacuously (two APIs that always throw also "agree").
/// </summary>
public class NameValidationTests
{
    private static bool DefineNameAccepts(string name) =>
        Accepts(() => new Workbook().DefineName(name, new NumberValue(1)));

    private static bool DefineTableAccepts(string name) =>
        Accepts(() => new Workbook().DefineTable(name, "Data", "A1:A2", ["Valor"]));

    private static bool Accepts(Action define)
    {
        try
        {
            define();
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    // Each name is registered on its OWN workbook: the two APIs share a namespace, so a second registration
    // of the same name on one workbook would throw for the collision rather than for the name's shape.
    [Test]
    [Arguments("Tabela1")] // Excel's pt-BR default table name
    [Arguments("Table1")] // Excel's en-US default table name
    [Arguments("表1")] // Excel's CJK default table name: a non-ASCII letter never labels a grid column
    [Arguments("Sales2024")] // same class as Tabela1: a trailing year does not make a name grid-shaped
    [Arguments("Vendas2024")]
    [Arguments("Q1")] // a real cell, so reserved on both sides
    [Arguments("ABC123")]
    [Arguments("$A$1")]
    [Arguments("A01")]
    [Arguments("T1")]
    [Arguments("A1")]
    public async Task DefineName_And_DefineTable_AgreeOnWhetherANameIsGridShaped(string name)
    {
        var asName = DefineNameAccepts(name);
        var asTable = DefineTableAccepts(name);

        await Assert
            .That(asName)
            .IsEqualTo(asTable)
            .Because($"'{name}': DefineName accepts = {asName}, DefineTable accepts = {asTable}");
    }

    // Anti-vacuity guard, accepting side: the property above would also hold if BOTH validators rejected
    // every name, so pin that Excel's default table name is accepted by both.
    [Test]
    public async Task DefineName_ExcelDefaultTableName_IsAccepted()
    {
        var workbook = new Workbook();

        workbook.DefineName("Tabela1", new NumberValue(1));

        await Assert.That(workbook.DefinedNames.ContainsKey("Tabela1")).IsTrue();
        await Assert.That(DefineTableAccepts("Tabela1")).IsTrue();
    }

    // Anti-vacuity guard, rejecting side: a real grid cell — even the shortest, two-character one — is
    // reserved by both validators.
    [Test]
    public async Task DefineName_RealGridCell_Throws()
    {
        var workbook = new Workbook();

        await Assert
            .That(() => workbook.DefineName("T1", new NumberValue(1)))
            .Throws<ArgumentException>();
        await Assert.That(DefineTableAccepts("T1")).IsFalse();
    }
}
