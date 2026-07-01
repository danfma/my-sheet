using MemoryPack;

namespace Danfma.MySheet.Expressions;

// SHEET([reference]) — the 1-based position of a sheet. With no argument it uses the current sheet.
[MemoryPackable]
public sealed partial record SheetNumber(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var sheetName = Arguments switch
        {
            [] => context.SheetName,
            [CellReference cell] => cell.SheetName,
            [RangeReference range] => range.SheetName,
            [var argument] => argument.Evaluate(context).AsString(),
            _ => null,
        };

        return
            sheetName is not null && context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
            ? ComputedValue.Number(sheet.Index + 1)
            : ComputedValue.Error(Error.Ref);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
