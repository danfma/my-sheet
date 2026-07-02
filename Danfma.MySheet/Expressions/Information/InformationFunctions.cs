using MemoryPack;

namespace Danfma.MySheet.Expressions.Information;

// Onda 2 — informação: NA, ISERROR/ISERR/ISNA, ISTEXT/ISNONTEXT/ISLOGICAL (checagem de tipo SEM
// coerção, como as IS functions do Excel), ISEVEN/ISODD, ISREF (checagem sintática do nó),
// ISFORMULA (inspeciona a expressão da célula referenciada), N, TYPE, ERROR.TYPE e SHEETS.
// T mora em Expressions.Text — o function-reference (espelhando o catálogo do Excel) o lista em Text.

[MemoryPackable]
public sealed partial record Na(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(Error.NA);
}

[MemoryPackable]
public sealed partial record IsError(Expression[] Arguments) : Function
{
    // TRUE for any error value (#N/A included); errors are inspected, never propagated.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Error);
}

[MemoryPackable]
public sealed partial record IsErr(Expression[] Arguments) : Function
{
    // TRUE for any error value EXCEPT #N/A (per the Excel docs: ISERR(#N/A) = FALSE).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(
            Arguments[0].Evaluate(context).TryGetError(out var error) && error != Error.NA
        );
}

[MemoryPackable]
public sealed partial record IsNa(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(
            Arguments[0].Evaluate(context).TryGetError(out var error) && error == Error.NA
        );
}

[MemoryPackable]
public sealed partial record IsText(Expression[] Arguments) : Function
{
    // No coercion — ISTEXT(19) is FALSE and ISLOGICAL("TRUE") is FALSE, like Excel's IS functions.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Text);
}

[MemoryPackable]
public sealed partial record IsNonText(Expression[] Arguments) : Function
{
    // TRUE for anything that is not text — including blanks and errors (per the Excel docs).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind != ComputedValueKind.Text);
}

[MemoryPackable]
public sealed partial record IsLogical(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Boolean);
}

[MemoryPackable]
public sealed partial record IsEven(Expression[] Arguments) : Function
{
    // The value is truncated before testing (ISEVEN(2.5) = TRUE); nonnumeric -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Boolean(Math.Truncate(number) % 2 == 0);
    }
}

[MemoryPackable]
public sealed partial record IsOdd(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Boolean(Math.Truncate(number) % 2 != 0);
    }
}

[MemoryPackable]
public sealed partial record IsRef(Expression[] Arguments) : Function
{
    // A syntactic check: the argument IS a reference node (cell, range or union), regardless of
    // the referenced value — ISREF(G8) is TRUE even when G8 is empty or holds an error.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0] is Reference);
}

[MemoryPackable]
public sealed partial record IsFormula(Expression[] Arguments) : Function
{
    // TRUE when the referenced cell contains a formula — i.e. its stored expression is not a plain
    // literal (ValueExpression). A formula that evaluates to an error still counts (=3/0 -> TRUE).
    // A range argument tests its top-left cell. Non-reference argument -> #VALUE! (per the docs).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var (sheetName, cellId) = Arguments[0] switch
        {
            CellReference cell => (cell.SheetName, cell.Id),
            RangeReference range => (range.SheetName, range.StartId),
            _ => (null, null),
        };

        if (sheetName is null || cellId is null)
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Boolean(
            context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
            && sheet.TryGetValue(cellId, out var expression)
            && expression is not ValueExpression
        );
    }
}

[MemoryPackable]
public sealed partial record N(Expression[] Arguments) : Function
{
    // Excel's conversion table: number -> itself; TRUE -> 1; FALSE -> 0; an error -> the error;
    // anything else (text — even "7" — and blanks) -> 0.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = Arguments[0].Evaluate(context);

        switch (value.Kind)
        {
            case ComputedValueKind.Number:
                return value;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                return ComputedValue.Number(boolean ? 1 : 0);

            case ComputedValueKind.Error:
                return value;

            default:
                return ComputedValue.Number(0);
        }
    }
}

[MemoryPackable]
public sealed partial record TypeFunction(Expression[] Arguments) : Function
{
    // Excel's code table: number = 1, text = 2, logical = 4, error = 16 (inspected, not
    // propagated), array = 64. A blank evaluates as 0, so its TYPE is 1, matching Excel.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = Arguments[0].Evaluate(context);

        if (value.Kind == ComputedValueKind.Reference)
        {
            // A single-cell reference result reports its value's type; a multi-cell reference is
            // the closest thing the engine has to an array -> 64.
            using var values = value.EnumerateValues(context).GetEnumerator();

            if (!values.MoveNext())
            {
                return ComputedValue.Number(1);
            }

            var first = values.Current;

            if (values.MoveNext())
            {
                return ComputedValue.Number(64);
            }

            value = first;
        }

        return ComputedValue.Number(
            value.Kind switch
            {
                ComputedValueKind.Text => 2,
                ComputedValueKind.Boolean => 4,
                ComputedValueKind.Error => 16,
                _ => 1, // Number and Blank
            }
        );
    }
}

[MemoryPackable]
public sealed partial record ErrorType(Expression[] Arguments) : Function
{
    // #NULL!=1, #DIV/0!=2, #VALUE!=3, #REF!=4, #NAME?=5, #NUM!=6, #N/A=7; a non-error -> #N/A.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (!Arguments[0].Evaluate(context).TryGetError(out var error))
        {
            return ComputedValue.Error(Error.NA);
        }

        if (error == Error.Null)
        {
            return ComputedValue.Number(1);
        }

        if (error == Error.DivZero)
        {
            return ComputedValue.Number(2);
        }

        if (error == Error.Value)
        {
            return ComputedValue.Number(3);
        }

        if (error == Error.Ref)
        {
            return ComputedValue.Number(4);
        }

        if (error == Error.Name)
        {
            return ComputedValue.Number(5);
        }

        if (error == Error.Num)
        {
            return ComputedValue.Number(6);
        }

        return error == Error.NA ? ComputedValue.Number(7) : ComputedValue.Error(Error.NA);
    }
}

[MemoryPackable]
public sealed partial record SheetsCount(Expression[] Arguments) : Function
{
    // SHEETS() — the number of sheets in the workbook (the reference form is not supported: the
    // engine has no 3-D references, so every reference spans exactly one sheet).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(context.Workbook.Sheets.Count);
}
