using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// INDIRECT(ref_text, [a1]) — resolves the A1-style reference named by <c>ref_text</c> at evaluation
/// time. Volatile (Excel parity). MySheet supports A1 style only: <c>a1 = FALSE</c> (R1C1) yields
/// <c>#REF!</c>. Non-text <c>ref_text</c>, malformed reference text, or an unknown sheet also yield
/// <c>#REF!</c>. A single-cell result is dereferenced to its value; a multi-cell result stays a
/// reference (the OFFSET/CHOOSE convention), so range consumers expand it and it works as a ':' endpoint.
/// </summary>
[MemoryPackable]
public sealed partial record Indirect(Expression[] Arguments) : Function
{
    public override bool IsVolatile => true;

    public override ComputedValue Evaluate(EvaluationContext context)
    {
        context.Workbook.MarkVolatileTouched();

        if (!TryResolveReference(context, out var reference))
        {
            return ComputedValue.Error(Error.Ref);
        }

        return reference is CellReference cell
            ? context.Workbook.GetCellValue(cell.SheetName, cell.Id) // single cell: dereference to value
            : ComputedValue.Reference(reference!); // range: stays a reference for consumers to expand
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // a1 flag: default TRUE (A1). Excel treats TRUE/1 as A1 and FALSE/0 as R1C1, which is unsupported.
        if (Arguments.Length >= 2)
        {
            if (Arguments[1].Evaluate(context).CoerceToNumber(out var a1) is not null || a1 == 0)
            {
                return false;
            }
        }

        if (!Arguments[0].Evaluate(context).TryGetText(out var refText))
        {
            return false; // ref_text must be text
        }

        if (
            context.SheetName is not { } sheetName
            || !context.Workbook.Sheets.TryGetValue(sheetName, out var sheet)
        )
        {
            return false; // no current sheet to resolve an unqualified reference against
        }

        Expression parsed;
        try
        {
            parsed = ExpressionParser.ParseFormulaBody(refText, sheet);
        }
        catch (ParseException)
        {
            return false; // malformed reference text → #REF!
        }

        return parsed.TryResolveReference(context, out reference);
    }
}
