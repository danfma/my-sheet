using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The fixture and harness <c>FilterTests</c>, <c>SortTests</c> and <c>UniqueTests</c> share. The three
/// producers are driven through the AST directly because none is registered yet (registration, the union
/// tags and the classification counts are one later commit), so the formula-level pins in
/// <c>DynamicArrayTests</c> stay red until then and these are the same numbers reached without the parser.
/// Cells are those of <c>DynamicArrayTests</c> plus the ones this task measured (2026-09-10, own probe copy).
/// </summary>
internal static class SelectionProducerFixture
{
    // A1:A3 = 5, 0, 9 | B1:B3 = 1, 2, 3 | C1:C3 = "a", "A", "b" | A5:A8 = 7, <blank>, "t", 7 |
    // E1:E3 = 5, #DIV/0!, 9 | F1:F3 = TRUE, "x", TRUE | G1:G3 = "TRUE", "FALSE", "true" |
    // K1:K3 = " TRUE ", "yes", "1" | L1:L3 = TRUE, " TRUE ", TRUE | M1:M4 = TRUE, "x", 2, FALSE |
    // N1:O4 = (1,"a"), (2,"b"), (1,"c"), (2,"d") | Q1:Q4 = 9, 5, 9, 0 | R1:R4 = 1, "1", TRUE, 1 |
    // S1:S2 = 4, 4 | T1:U3 = (1,"a"), (1,"a"), (1,"A") | V1:V4 = 0, <blank>, "", FALSE |
    // W1:W5 = 2, #DIV/0!, 3, #N/A, #DIV/0! | A9 is never written.
    public static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        void Put(string id, object value) =>
            sheet[id] = value switch
            {
                string s when s.StartsWith('=') => ExpressionParser.Parse(s, sheet),
                string s => new StringValue(s),
                bool b => b ? BooleanValue.True : BooleanValue.False,
                int i => Number(i),
                _ => throw new ArgumentException($"Unsupported cell value: {value}"),
            };

        (string, object)[] cells =
        [
            ("A1", 5),
            ("A2", 0),
            ("A3", 9),
            ("B1", 1),
            ("B2", 2),
            ("B3", 3),
            ("C1", "a"),
            ("C2", "A"),
            ("C3", "b"),
            ("A5", 7),
            ("A7", "t"),
            ("A8", 7),
            ("E1", 5),
            ("E2", "=1/0"),
            ("E3", 9),
            ("F1", true),
            ("F2", "x"),
            ("F3", true),
            ("G1", "TRUE"),
            ("G2", "FALSE"),
            ("G3", "true"),
            ("K1", " TRUE "),
            ("K2", "yes"),
            ("K3", "1"),
            ("L1", true),
            ("L2", " TRUE "),
            ("L3", true),
            ("M1", true),
            ("M2", "x"),
            ("M3", 2),
            ("M4", false),
            ("N1", 1),
            ("O1", "a"),
            ("N2", 2),
            ("O2", "b"),
            ("N3", 1),
            ("O3", "c"),
            ("N4", 2),
            ("O4", "d"),
            ("Q1", 9),
            ("Q2", 5),
            ("Q3", 9),
            ("Q4", 0),
            ("R1", 1),
            ("R2", "1"),
            ("R3", true),
            ("R4", 1),
            ("S1", 4),
            ("S2", 4),
            ("T1", 1),
            ("U1", "a"),
            ("T2", 1),
            ("U2", "a"),
            ("T3", 1),
            ("U3", "A"),
            ("V1", 0),
            ("V3", "=\"\""),
            ("V4", false),
            ("W1", 2),
            ("W2", "=1/0"),
            ("W3", 3),
            ("W4", "=NA()"),
            ("W5", "=1/0"),
        ];

        foreach (var (id, value) in cells)
        {
            Put(id, value);
        }

        return (workbook, sheet);
    }

    /// <summary>Any registered formula fragment as a node — a range, a comparison, a scalar call.</summary>
    public static Expression Ref(string formula, Sheet sheet) =>
        ExpressionParser.Parse("=" + formula, sheet);

    public static object? Eval(Expression node, Workbook workbook) =>
        node.Evaluate(workbook).AsObject();

    public static double Num(object? value) => value is double d ? d : double.NaN;

    public static bool TryBuild(Expression producer, Workbook workbook, out ArrayOperand operand) =>
        ((IArrayProducer)producer).TryBuildArrayOperand(
            new EvaluationContext(workbook),
            out operand
        );

    public static ArrayOperand Build(Expression producer, Workbook workbook)
    {
        TryBuild(producer, workbook, out var operand);
        return operand;
    }

    /// <summary>The element at a row-major index, read at the operand's own extent.</summary>
    public static ComputedValue At(ArrayOperand operand, int index) =>
        operand.At(index, operand.Rows, operand.Columns);

    public static bool Eligible(Expression node, Workbook workbook) =>
        ArrayEvaluation.IsArrayEligible(node, new EvaluationContext(workbook));
}
