using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The A1 parsing helpers on <see cref="CellAddress"/>: the column accumulator's overflow ceiling (the twin
/// of the row ceiling <c>AbsoluteRowReferenceTests</c> pins) and the composed <c>TryParseA1</c> corner
/// parser the table registry's A1 overload builds on.
/// </summary>
public class CellAddressTests
{
    // --- The column accumulator applies the same ceiling as the row accumulator (int.MaxValue). "FXSHRXW"
    // is column 2,147,483,647 exactly; "FXSHRXX" is one past it and must be rejected, never wrapped: before
    // the guard, 24 'A's returned ok=True with column -965696553.

    [Test]
    [Arguments("A", true, 1)]
    [Arguments("$ab", true, 28)]
    [Arguments("XFD", true, 16384)]
    [Arguments("FXSHRXW", true, int.MaxValue)]
    [Arguments("FXSHRXX", false, 0)] // int.MaxValue + 1
    [Arguments("AAAAAAAAAAAAAAAAAAAAAAAA", false, 0)] // 24 letters: wrapped negative before the guard
    [Arguments("", false, 0)]
    [Arguments("$", false, 0)]
    [Arguments("A1", false, 0)]
    public async Task TryParseColumn_RejectsNonLettersAndOverflow(
        string label,
        bool expectedOk,
        int expectedColumn
    )
    {
        var ok = CellAddress.TryParseColumn(label, out var column);

        await Assert.That(ok).IsEqualTo(expectedOk);
        await Assert.That(column).IsEqualTo(expectedColumn);
    }

    // --- TryParseA1 composes the two halves, so it inherits both ceilings, the '$' stripping and the row-0
    // rejection, and adds only the split: at least one letter, then at least one digit, nothing else.

    [Test]
    [Arguments("A1", true, 1, 1)]
    [Arguments("$B$500", true, 2, 500)]
    [Arguments("xfd1048576", true, 16384, 1048576)]
    [Arguments("FXSHRXW2147483647", true, int.MaxValue, int.MaxValue)]
    [Arguments("FXSHRXX1", false, 0, 0)] // column int.MaxValue + 1
    [Arguments("AAAAAAAAAAAAAAAAAAAAAAAA1", false, 0, 0)] // the wrapped-column case M2 measured
    [Arguments("A2147483648", false, 0, 0)] // row int.MaxValue + 1
    [Arguments("A0", false, 0, 0)]
    [Arguments("A", false, 0, 0)]
    [Arguments("1", false, 0, 0)]
    [Arguments("", false, 0, 0)]
    [Arguments("A1B", false, 0, 0)]
    [Arguments("A1:B2", false, 0, 0)]
    [Arguments("A 1", false, 0, 0)]
    public async Task TryParseA1_SplitsAtTheFirstDigit_AndRejectsEitherHalfFailing(
        string text,
        bool expectedOk,
        int expectedColumn,
        int expectedRow
    )
    {
        var ok = CellAddress.TryParseA1(text, out var column, out var row);

        await Assert.That(ok).IsEqualTo(expectedOk);
        await Assert.That(column).IsEqualTo(expectedColumn);
        await Assert.That(row).IsEqualTo(expectedRow);
    }
}
