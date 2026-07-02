using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — aliases legados (categoria Compatibility do Excel): MODE, STDEV, STDEVP, VAR, VARP,
// RANK, PERCENTILE, PERCENTRANK, QUARTILE, COVAR e FORECAST. Cada alias é um NÓ DISTINTO da forma
// moderna: mesmo comportamento (delegam à mesma lógica), mas o un-parse preserva o nome legado
// que o usuário escreveu — STDEV(...) não pode virar STDEV.S(...) em FORMULATEXT/export.
public class CompatibilityAliasTests
{
    private const string ContextSheet = "Sheet1";

    private static object? Calc(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add(ContextSheet);

        foreach (var (id, value) in cells)
        {
            sheet[id] = value switch
            {
                string s => new Danfma.MySheet.Expressions.StringValue(s),
                double d => new NumberValue(d),
                int i => new NumberValue(i),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static (string, object)[] Column(string column, int firstRow, params double[] values) =>
        [.. values.Select((v, i) => ($"{column}{firstRow + i}", (object)v))];

    // Cada alias tem que avaliar EXATAMENTE como a forma moderna, no mesmo dataset (os goldens
    // das formas modernas estão nos testes por família; aqui o oráculo é a própria equivalência
    // documentada MODE≡MODE.SNGL, STDEV≡STDEV.S, ..., COVAR≡COVARIANCE.P, FORECAST≡FORECAST.LINEAR).
    [Test]
    [Arguments("MODE(A1:A6)", "MODE.SNGL(A1:A6)")]
    [Arguments("STDEV(A1:A6)", "STDEV.S(A1:A6)")]
    [Arguments("STDEVP(A1:A6)", "STDEV.P(A1:A6)")]
    [Arguments("VAR(A1:A6)", "VAR.S(A1:A6)")]
    [Arguments("VARP(A1:A6)", "VAR.P(A1:A6)")]
    [Arguments("RANK(4,A1:A6)", "RANK.EQ(4,A1:A6)")]
    [Arguments("RANK(4,A1:A6,1)", "RANK.EQ(4,A1:A6,1)")]
    [Arguments("PERCENTILE(A1:A6,0.3)", "PERCENTILE.INC(A1:A6,0.3)")]
    [Arguments("PERCENTRANK(A1:A6,4)", "PERCENTRANK.INC(A1:A6,4)")]
    [Arguments("QUARTILE(A1:A6,1)", "QUARTILE.INC(A1:A6,1)")]
    [Arguments("COVAR(A1:A6,B1:B6)", "COVARIANCE.P(A1:A6,B1:B6)")]
    [Arguments("FORECAST(9,A1:A6,B1:B6)", "FORECAST.LINEAR(9,A1:A6,B1:B6)")]
    public async Task Alias_EvaluatesExactlyLikeTheModernForm(string legacy, string modern)
    {
        var cells = Column("A", 1, 5.6, 4, 4, 3, 2, 4)
            .Concat(Column("B", 1, 9, 7, 12, 15, 17, 11))
            .ToArray();

        var legacyResult = Calc("=" + legacy, cells);
        var modernResult = Calc("=" + modern, cells);

        // Guarda anti-vacuidade: os dois lados têm que ser números de verdade (não erro/NaN).
        await Assert.That(double.IsNaN(Num(legacyResult))).IsFalse();
        await Assert.That(Num(legacyResult)).IsEqualTo(Num(modernResult));
    }

    [Test]
    public async Task Mode_MatchesTheModeSnglGoldenExample()
    {
        // O golden do MODE.SNGL ({5.6,4,4,3,2,4} -> 4) vale para o alias.
        await Assert
            .That(Num(Calc("=MODE(A1:A6)", Column("A", 1, 5.6, 4, 4, 3, 2, 4))))
            .IsEqualTo(4.0);
    }

    // O un-parse tem que preservar a grafia legada — um nó distinto por alias garante isso.
    [Test]
    [Arguments("MODE(A1:A2)")]
    [Arguments("STDEV(A1:A2)")]
    [Arguments("STDEVP(A1:A2)")]
    [Arguments("VAR(A1:A2)")]
    [Arguments("VARP(A1:A2)")]
    [Arguments("RANK(1,A1:A2)")]
    [Arguments("PERCENTILE(A1:A2,0.5)")]
    [Arguments("PERCENTRANK(A1:A2,1)")]
    [Arguments("QUARTILE(A1:A2,1)")]
    [Arguments("COVAR(A1:A2,B1:B2)")]
    [Arguments("FORECAST(1,A1:A2,B1:B2)")]
    public async Task Alias_UnParsesWithTheLegacyName(string formula)
    {
        var sheet = new Sheet { Name = ContextSheet };
        var expression = ExpressionParser.Parse("=" + formula, sheet);

        await Assert.That(expression.ToFormula(ContextSheet)).IsEqualTo(formula);
    }

    [Test]
    public async Task AliasAndModernForm_AreDistinctNodes()
    {
        var sheet = new Sheet { Name = ContextSheet };

        var legacy = ExpressionParser.Parse("=STDEV(A1:A2)", sheet);
        var modern = ExpressionParser.Parse("=STDEV.S(A1:A2)", sheet);

        await Assert.That(legacy.GetType()).IsNotEqualTo(modern.GetType());
    }

    [Test]
    public async Task Alias_SurvivesFormulaTextRoundTrip()
    {
        // FORMULATEXT sobre uma célula com STDEV(...) tem que devolver "=STDEV(A1:A2)".
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add(ContextSheet);

        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["B1"] = ExpressionParser.Parse("=STDEV(A1:A2)", sheet);

        var result = ExpressionParser.Parse("=FORMULATEXT(B1)", sheet).Evaluate(workbook);

        await Assert.That(result.AsObject()).IsEqualTo("=STDEV(A1:A2)");
    }
}
