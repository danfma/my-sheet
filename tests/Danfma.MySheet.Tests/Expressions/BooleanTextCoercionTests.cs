using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MySheetText = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 11b: the text <c>"TRUE"</c> / <c>"FALSE"</c> in a BOOLEAN CONDITION slot. A user bug report of
/// 2026-09-10: <c>=IF("TRUE",1,0)</c> answered <c>#VALUE!</c> here and <c>1</c> in Excel.
/// <para>
/// Oracle: <b>Aspose.Cells 26.6.0</b>, measured 2026-09-10 on this exact fixture, and <b>plain and
/// array-entered (<c>Cell.SetArrayFormula(f, 1, 1)</c>) entry AGREE on every row below</b>, so entry mode is
/// not a factor in this phase and each row is pinned once.
/// </para>
/// <para>
/// <b>The rule, and it is narrow.</b> Exactly the two words <c>TRUE</c> and <c>FALSE</c>, compared
/// case-INSENSITIVELY (<c>OrdinalIgnoreCase</c>), with <b>no trimming</b> and no other accepted spelling.
/// <c>" TRUE "</c>, <c>"yes"</c>, <c>"1"</c>, <c>"0"</c> and <c>""</c> all stay <c>#VALUE!</c>. It applies to
/// a boolean CONDITION slot only: <c>IF</c>'s condition (scalar and array-condition paths), <c>NOT</c>'s
/// argument and <c>IFS</c>' tests. Measured for <c>IFS</c> on 2026-09-10 while answering this phase's open
/// question, which is why <c>IFS</c> is pinned here beside <c>IF</c> and <c>NOT</c>.
/// </para>
/// <para>
/// <b>What must NOT move, and why the guards are half this file.</b> <c>AND</c>, <c>OR</c> and <c>XOR</c>
/// <b>IGNORE</b> a text argument rather than coercing it — measured: <c>AND("FALSE",TRUE)</c> is TRUE, where
/// coercion would give FALSE, and <c>AND("yes",TRUE)</c> is TRUE rather than an error. Their answers here
/// already coincide with that ignore rule, so <b>no other test in the suite would catch a shared-helper fix
/// breaking them</b>: <see cref="And_TextFalse_IsIgnoredNotCoerced"/> is the only thing standing there.
/// Arithmetic (<c>"TRUE"+0</c>), comparison (<c>"TRUE"=TRUE</c>) and <c>SUMPRODUCT(--(A1:A2=TRUE))</c> keep
/// rejecting/ignoring too, and <c>SWITCH</c> compares its expression with <c>=</c> equality rather than
/// coercing it (measured: <c>SWITCH("TRUE",TRUE,1,0)</c> = 0, <c>SWITCH("TRUE","TRUE",1,0)</c> = 1).
/// </para>
/// <para>
/// Deliberately NOT pinned here: <c>COUNTIF(A1:A2,TRUE)</c> answers 1 on this tree against the oracle's 0.
/// That is criteria-family type equality — a different rule and a SILENT wrong number — and it is recorded
/// for the compatibility sweep beside the existing criteria item, not fixed in this slice.
/// </para>
/// </summary>
public class BooleanTextCoercionTests
{
    // A1 = "TRUE", A2 = "FALSE", A3 = "true" (lower case), A4 = " TRUE " (padded), A5 = "yes", A6 = "1" —
    // all TEXT cells, so the cell entry mode of every row above has a fixture. B1:B3 are numbers, and
    // C1:C3 = "TRUE"/"FALSE"/"true" is the text column the ARRAY-condition rows read.
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new MySheetText("TRUE");
        sheet["A2"] = new MySheetText("FALSE");
        sheet["A3"] = new MySheetText("true");
        sheet["A4"] = new MySheetText(" TRUE ");
        sheet["A5"] = new MySheetText("yes");
        sheet["A6"] = new MySheetText("1");

        sheet["B1"] = new NumberValue(10);
        sheet["B2"] = new NumberValue(20);
        sheet["B3"] = new NumberValue(30);

        sheet["C1"] = new MySheetText("TRUE");
        sheet["C2"] = new MySheetText("FALSE");
        sheet["C3"] = new MySheetText("true");

        return (workbook, sheet);
    }

    private static object? OnGrid(string formula)
    {
        var (workbook, sheet) = Grid();
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // ================================================================================================
    // ACCEPTANCE pins — RED before item 3. Each comment names the value this tree produced at branch
    // creation in brackets; every one of them is #VALUE!, so none of these defects is silent.
    // ================================================================================================

    [Test]
    public async Task If_TextTrueLiteral_IsTrue()
    {
        // [#VALUE!] -> oracle 1.
        await Assert.That(OnGrid("=IF(\"TRUE\",1,0)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task If_TextFalseLiteral_IsFalse()
    {
        // [#VALUE!] -> oracle 0.
        await Assert.That(OnGrid("=IF(\"FALSE\",1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task If_CellHoldingTextTrue_IsTrue()
    {
        // [#VALUE!] -> oracle 1. A1 is a TEXT cell: the bug report's second shape.
        await Assert.That(OnGrid("=IF(A1,1,0)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task If_CellHoldingTextFalse_IsFalse()
    {
        // [#VALUE!] -> oracle 0.
        await Assert.That(OnGrid("=IF(A2,1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task If_CellHoldingLowerCaseTrue_IsTrue()
    {
        // [#VALUE!] -> oracle 1. The match is case-INSENSITIVE, which is what makes it an
        // OrdinalIgnoreCase comparison rather than an ordinal one.
        await Assert.That(OnGrid("=IF(A3,1,0)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task If_MixedCaseTrueLiteral_IsTrue()
    {
        // [#VALUE!] -> oracle 1 (measured for "True" and "tRuE" alike).
        await Assert.That(OnGrid("=IF(\"tRuE\",1,0)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task If_MixedCaseFalseLiteral_IsFalse()
    {
        // [#VALUE!] -> oracle 0.
        await Assert.That(OnGrid("=IF(\"FaLsE\",1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task Not_TextTrueLiteral_IsFalse()
    {
        // [#VALUE!] -> oracle FALSE. The same coercion site as IF's condition.
        await Assert.That(OnGrid("=NOT(\"TRUE\")") as bool?).IsFalse();
    }

    [Test]
    public async Task Not_TextFalseLiteral_IsTrue()
    {
        // [#VALUE!] -> oracle TRUE.
        await Assert.That(OnGrid("=NOT(\"FALSE\")") as bool?).IsTrue();
    }

    [Test]
    public async Task Not_CellHoldingTextTrue_IsFalse()
    {
        // [#VALUE!] -> oracle FALSE.
        await Assert.That(OnGrid("=NOT(A1)") as bool?).IsFalse();
    }

    // ------------------------------------------------------------------------------------------------
    // The ARRAY-condition path of IF and the elementwise path of NOT. MySheet routes an ARRAY condition
    // through ArrayOperands.IfOperand rather than If.Evaluate, so it is a SEPARATE call site of the
    // coercion and would otherwise disagree with the scalar path. Both rows below are 2 / 1 on the oracle
    // in BOTH entry modes (plain SUMPRODUCT and SUMPRODUCT array-entered).
    // ------------------------------------------------------------------------------------------------

    [Test]
    public async Task If_OverAnArrayOfTextConditions_CoercesEachElement()
    {
        // [#VALUE!] -> oracle 2: C1 = "TRUE" and C3 = "true" contribute 1 each, C2 = "FALSE" contributes 0.
        await Assert.That(OnGrid("=SUMPRODUCT(IF(C1:C3,1,0))") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Not_OverAnArrayOfTextArguments_CoercesEachElement()
    {
        // [#VALUE!] -> oracle 1: NOT("TRUE") = FALSE = 0, NOT("FALSE") = TRUE = 1.
        await Assert.That(OnGrid("=SUMPRODUCT(--NOT(A1:A2))") as double?).IsEqualTo(1.0);
    }

    // ------------------------------------------------------------------------------------------------
    // IFS — this phase's open question, ANSWERED by measurement on 2026-09-10: its test slot coerces the
    // two words under exactly the same rule as IF's condition, in both entry modes.
    //   IFS("TRUE",1,TRUE,2) = 1   IFS("FALSE",1,TRUE,2) = 2   IFS(A3,1,TRUE,2) = 1
    //   IFS("yes",1,TRUE,2) = IFS(" TRUE ",1,TRUE,2) = IFS("1",1,TRUE,2) = #VALUE!
    // ------------------------------------------------------------------------------------------------

    [Test]
    public async Task Ifs_TextTrueTest_TakesThatBranch()
    {
        // [#VALUE!] -> oracle 1.
        await Assert.That(OnGrid("=IFS(\"TRUE\",1,TRUE,2)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Ifs_TextFalseTest_FallsThroughToTheNextTest()
    {
        // [#VALUE!] -> oracle 2. "FALSE" must coerce to a FALSE test, not error and not match.
        await Assert.That(OnGrid("=IFS(\"FALSE\",1,TRUE,2)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Ifs_CellHoldingLowerCaseTrue_TakesThatBranch()
    {
        // [#VALUE!] -> oracle 1.
        await Assert.That(OnGrid("=IFS(A3,1,TRUE,2)") as double?).IsEqualTo(1.0);
    }

    // ================================================================================================
    // MUST-NOT-MOVE guards — GREEN at branch creation and GREEN after item 3. The ten the phase names,
    // then the boundary rows the measurement added.
    // ================================================================================================

    [Test]
    public async Task If_TextTrueWithSurroundingWhitespace_StaysValueError()
    {
        // GUARD 1 — no trimming. Oracle #VALUE!, so the rule is a case-insensitive EXACT match.
        await Assert.That(OnGrid("=IF(\" TRUE \",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_UnrelatedText_StaysValueError()
    {
        // GUARD 2 — only the two words coerce. Oracle #VALUE!.
        await Assert.That(OnGrid("=IF(\"yes\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_NumericTextOne_StaysValueError()
    {
        // GUARD 3 — numeric TEXT is not a boolean, even though the NUMBER 1 is truthy. Oracle #VALUE!.
        await Assert.That(OnGrid("=IF(\"1\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_CellHoldingNumericTextOne_StaysValueError()
    {
        // GUARD 4 — the same through a cell. A6 = "1" as TEXT. Oracle #VALUE!.
        await Assert.That(OnGrid("=IF(A6,1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task And_TextFalse_IsIgnoredNotCoerced()
    {
        // GUARD 5 — THE row that distinguishes "ignore" from "coerce", and the reason this phase does not
        // touch a shared helper. Oracle TRUE: the text operand is dropped and AND folds TRUE alone.
        // Coercion would make this FALSE.
        await Assert.That(OnGrid("=AND(\"FALSE\",TRUE)") as bool?).IsTrue();
    }

    [Test]
    public async Task Or_TextTrue_IsIgnoredNotCoerced()
    {
        // GUARD 6 — the mirror of guard 5. Oracle FALSE: coercion would make this TRUE.
        await Assert.That(OnGrid("=OR(\"TRUE\",FALSE)") as bool?).IsFalse();
    }

    [Test]
    public async Task Xor_TextTrue_IsIgnoredNotCoerced()
    {
        // GUARD 7 — Oracle FALSE: one evaluable operand, FALSE, so the TRUE parity is even.
        await Assert.That(OnGrid("=XOR(\"TRUE\",FALSE)") as bool?).IsFalse();
    }

    [Test]
    public async Task TextTrue_InArithmetic_StaysValueError()
    {
        // GUARD 8 — arithmetic does NOT coerce the two words. Oracle #VALUE!.
        await Assert.That(OnGrid("=\"TRUE\"+0")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TextTrue_ComparedWithBooleanTrue_IsNotEqual()
    {
        // GUARD 9 — comparison does NOT coerce: different types are never equal. Oracle 0.
        await Assert.That(OnGrid("=IF(\"TRUE\"=TRUE,1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task SumProduct_OverTextComparedWithBooleanTrue_IsZero()
    {
        // GUARD 10 — the array form of guard 9. Oracle 0: neither text cell equals the boolean TRUE.
        await Assert.That(OnGrid("=SUMPRODUCT(--(A1:A2=TRUE))") as double?).IsEqualTo(0.0);
    }

    // ------------------------------------------------------------------------------------------------
    // Boundary rows the measurement added beside the ten.
    // ------------------------------------------------------------------------------------------------

    [Test]
    public async Task Not_UnrelatedText_StaysValueError()
    {
        // Oracle #VALUE! — NOT's argument takes the same narrow rule as IF's condition, no wider.
        await Assert.That(OnGrid("=NOT(\"yes\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_CellHoldingPaddedTextTrue_StaysValueError()
    {
        // Oracle #VALUE! — guard 1 through a cell (A4 = " TRUE ").
        await Assert.That(OnGrid("=IF(A4,1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_NumericTextZero_StaysValueError()
    {
        // Oracle #VALUE! — "0" is no more a boolean than "1" is.
        await Assert.That(OnGrid("=IF(\"0\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_EmptyText_StaysValueError()
    {
        // Oracle #VALUE! — the empty string is NOT the blank cell FALSE.
        await Assert.That(OnGrid("=IF(\"\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task And_UnrelatedText_IsIgnoredNotAnError()
    {
        // Oracle TRUE — ignoring, not erroring: the widened form of guard 5.
        await Assert.That(OnGrid("=AND(\"yes\",TRUE)") as bool?).IsTrue();
    }

    [Test]
    public async Task And_CellHoldingTextFalse_IsIgnoredNotCoerced()
    {
        // Oracle TRUE — the ignore rule reaches text through a CELL as well as a literal.
        await Assert.That(OnGrid("=AND(A2,TRUE)") as bool?).IsTrue();
    }

    [Test]
    public async Task Or_CellHoldingTextTrue_IsIgnoredNotCoerced()
    {
        // Oracle FALSE — the mirror through a cell.
        await Assert.That(OnGrid("=OR(A1,FALSE)") as bool?).IsFalse();
    }

    [Test]
    public async Task Xor_AllTextArguments_StaysValueError()
    {
        // Oracle #VALUE! — every operand ignored leaves nothing evaluable, which is the documented
        // "no logical values -> #VALUE!" remark, NOT a coercion failure.
        await Assert.That(OnGrid("=XOR(\"TRUE\",\"FALSE\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TextTrue_UnaryDoubleNegated_StaysValueError()
    {
        // Oracle #VALUE! — the -- idiom does not open a back door into the rule.
        await Assert.That(OnGrid("=--\"TRUE\"")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Ifs_UnrelatedText_StaysValueError()
    {
        // Oracle #VALUE! — IFS takes the rule, and takes it exactly as narrow as IF does.
        await Assert.That(OnGrid("=IFS(\"yes\",1,TRUE,2)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Ifs_PaddedTextTrue_StaysValueError()
    {
        // Oracle #VALUE! — no trimming in IFS either.
        await Assert.That(OnGrid("=IFS(\" TRUE \",1,TRUE,2)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Ifs_NumericTextOne_StaysValueError()
    {
        // Oracle #VALUE!.
        await Assert.That(OnGrid("=IFS(\"1\",1,TRUE,2)")).IsEqualTo(ErrorValue.NotValue);
    }

    // ------------------------------------------------------------------------------------------------
    // SWITCH — the other half of the open question. Its first argument is NOT a boolean slot: it is
    // compared with '=' equality, and equality does not coerce. Measured 2026-09-10, both entry modes:
    // SWITCH("TRUE",TRUE,1,0) = 0, SWITCH(TRUE,"TRUE",1,0) = 0, SWITCH("TRUE","TRUE",1,0) = 1. These are
    // GREEN today and stay green: SWITCH is NOT a site of this phase's rule.
    // ------------------------------------------------------------------------------------------------

    [Test]
    public async Task Switch_TextTrueAgainstBooleanTrue_DoesNotMatch()
    {
        // Oracle 0 — the default. Text never equals a boolean, in either direction.
        await Assert.That(OnGrid("=SWITCH(\"TRUE\",TRUE,1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task Switch_BooleanTrueAgainstTextTrue_DoesNotMatch()
    {
        // Oracle 0 — the reversed direction of the row above.
        await Assert.That(OnGrid("=SWITCH(TRUE,\"TRUE\",1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task Switch_TextTrueAgainstTextTrue_Matches()
    {
        // Oracle 1 — text-to-text equality, which is what SWITCH actually does here.
        await Assert.That(OnGrid("=SWITCH(\"TRUE\",\"TRUE\",1,0)") as double?).IsEqualTo(1.0);
    }
}
