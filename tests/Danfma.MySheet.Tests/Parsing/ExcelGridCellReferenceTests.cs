using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// <c>Parser.IsExcelGridCellReference</c>: the table-name rule's "looks like a cell reference" predicate.
/// Unlike <c>Parser.IsCellReference</c> (unbounded on purpose — MySheet's grid has no ceiling and the
/// parser relies on that), this one is bounded on BOTH the letter run (1-3 letters, so <c>Tabela1</c> and
/// <c>Table1</c>, Excel's own default table names, are never mistaken for cells) and the grid
/// (<c>XFD1048576</c> is the last cell; <c>XFE1</c> and <c>A1048577</c> are outside it).
/// </summary>
public class ExcelGridCellReferenceTests
{
    [Test]
    [Arguments("A1")]
    [Arguments("a1")]
    [Arguments("$A$1")]
    [Arguments("T1")] // two letters, inside the grid: a real cell, so NOT a legal table name
    [Arguments("Q1")]
    [Arguments("XFD1")]
    [Arguments("XFD1048576")] // the last cell of the grid
    [Arguments("A1048576")]
    public async Task InsideTheGrid_IsACellReference(string text)
    {
        await Assert.That(Parser.IsExcelGridCellReference(text)).IsTrue();
    }

    [Test]
    [Arguments("Tabela1")] // Excel's pt-BR default table name: six letters
    [Arguments("Table1")]
    [Arguments("AAAA1")] // four letters: past the letter bound even though the digits are fine
    [Arguments("XFE1")] // column 16,385: past the last column
    [Arguments("ZZZ1")] // three letters but column 18,278
    [Arguments("A1048577")] // past the last row
    [Arguments("A10000000")] // eight digits: past the digit bound
    [Arguments("A0")]
    [Arguments("R1C1")]
    [Arguments("A")]
    [Arguments("1")]
    [Arguments("")]
    [Arguments("$")]
    [Arguments("A1B")]
    [Arguments("A 1")]
    [Arguments("É1")] // a non-ASCII letter cannot label a grid column
    public async Task OutsideTheGrid_OrNotTheShape_IsNot(string text)
    {
        await Assert.That(Parser.IsExcelGridCellReference(text)).IsFalse();
    }

    // The unbounded sibling stays unbounded: this is exactly why the name rule needs its own predicate.
    [Test]
    public async Task IsCellReference_StaysUnbounded()
    {
        await Assert.That(Parser.IsCellReference("Tabela1")).IsTrue();
        await Assert.That(Parser.IsCellReference("XFD1048577")).IsTrue();
    }
}
