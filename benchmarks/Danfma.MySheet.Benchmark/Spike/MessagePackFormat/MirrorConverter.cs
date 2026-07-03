using Danfma.MySheet.Expressions;
using Logical = Danfma.MySheet.Expressions.Logical;
using Lookup = Danfma.MySheet.Expressions.Lookup;
using Math = Danfma.MySheet.Expressions.Mathematics;
using Stat = Danfma.MySheet.Expressions.Statistical;
using Text = Danfma.MySheet.Expressions.Text;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// Converts the PRODUCTION object graph (real Workbook / Expression tree) into the MessagePack mirror graph,
// so the SAME logical workbook is serialized by both formats for an apples-to-apples byte + speed measure.
// Only the representative subset is mapped; an unmapped node throws (loudly, on purpose) so the spike payload
// generators are kept honest — a payload can never silently drift outside the measured subset.
internal static class MirrorConverter
{
    public static MWorkbook ToMirror(Workbook workbook)
    {
        var mirror = new MWorkbook();

        foreach (var (name, sheet) in workbook.Sheets)
        {
            var msheet = new MSheet { Name = sheet.Name, Index = sheet.Index };
            foreach (var (id, expr) in sheet)
            {
                msheet.Cells[id] = ToMirror(expr);
            }

            mirror.Sheets[name] = msheet;
        }

        foreach (var (name, expr) in workbook.DefinedNames)
        {
            mirror.DefinedNames[name] = ToMirror(expr);
        }

        return mirror;
    }

    public static MExpr ToMirror(Expression e) =>
        e switch
        {
            // Values
            StringValue v => new MStringValue(v.Value),
            NumberValue v => new MNumberValue(v.Value),
            BooleanValue v => new MBooleanValue(v.Value),
            BlankValue => new MBlankValue(),

            // References
            CellReference r => new MCellReference(r.Id, r.SheetName),
            RangeReference r => new MRangeReference(r.StartId, r.EndId, r.SheetName),
            OpenRangeReference r => new MOpenRangeReference(r.ColMin, r.ColMax, r.RowMin, r.RowMax, r.SheetName),
            UnionReference r => new MUnionReference(Map(r.Areas)),
            NameReference r => new MNameReference(r.Name),

            // Operations
            BinaryOperation b => new MBinaryOperation(b.Operator, ToMirror(b.Left), ToMirror(b.Right)),
            UnaryOperation u => new MUnaryOperation(u.Operator, ToMirror(u.Operand)),

            // Aggregates / functions (subset)
            Math.Sum f => new MSum(Map(f.Expressions)),
            Stat.Average f => new MAverage(Map(f.Arguments)),
            Stat.Min f => new MMin(Map(f.Arguments)),
            Stat.Max f => new MMax(Map(f.Arguments)),
            Stat.Count f => new MCount(Map(f.Arguments)),
            Logical.If f => new MIf(Map(f.Arguments)),
            Logical.IfError f => new MIfError(Map(f.Arguments)),
            Math.Round f => new MRound(Map(f.Arguments)),
            Math.Abs f => new MAbs(Map(f.Arguments)),
            Text.Upper f => new MUpper(Map(f.Arguments)),
            Text.Left f => new MLeft(Map(f.Arguments)),
            Text.Concat f => new MConcat(Map(f.Arguments)),
            Stat.CountIf f => new MCountIf(Map(f.Arguments)),
            Math.SumIf f => new MSumIf(Map(f.Arguments)),
            Stat.AverageIf f => new MAverageIf(Map(f.Arguments)),
            Lookup.Match f => new MMatch(Map(f.Arguments)),
            Lookup.Index f => new MIndex(Map(f.Arguments)),
            Lookup.VLookup f => new MVLookup(Map(f.Arguments)),
            Lookup.Choose f => new MChoose(Map(f.Arguments)),
            Stat.Small f => new MSmall(Map(f.Arguments)),
            Logical.Let f => new MLet(Map(f.Arguments)),
            FunctionCall f => new MFunctionCall(f.Name, Map(f.Arguments)),

            _ => throw new NotSupportedException(
                $"Spike mirror does not cover {e.GetType().FullName}. Keep payloads within the measured subset."
            ),
        };

    private static MExpr[] Map(Expression[] items)
    {
        var result = new MExpr[items.Length];
        for (var i = 0; i < items.Length; i++)
        {
            result[i] = ToMirror(items[i]);
        }

        return result;
    }
}
