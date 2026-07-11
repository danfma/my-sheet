using MemoryPack;

namespace Danfma.MySheet.Expressions.Information;

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
            // Phase 2 audit (shared-formula delta production): a SHEET(ref) argument inside a shared-formula
            // master is an anchored node; SheetName is a literal component of both anchored node shapes
            // (unaffected by the delta), so no evaluation-time resolution is needed — just read it directly.
            [AnchoredCellReference anchoredCell] => anchoredCell.SheetName,
            [AnchoredRangeReference anchoredRange] => anchoredRange.SheetName,
            [var argument] => argument.Evaluate(context).AsString(),
            _ => null,
        };

        return
            sheetName is not null && context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
            ? ComputedValue.Number(sheet.Index + 1)
            : ComputedValue.Error(Error.Ref);
    }
}
