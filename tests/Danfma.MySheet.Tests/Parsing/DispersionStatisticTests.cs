using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — dispersão e momentos: STDEV.S/.P, STDEVA/STDEVPA, VAR.S/.P, VARA/VARPA, AVEDEV,
// DEVSQ, GEOMEAN, HARMEAN, SKEW, SKEW.P, KURT e STANDARDIZE. Golden values das páginas oficiais
// da Microsoft (support.microsoft.com, fetched em 2026-07-02), citadas por função em cada teste.
public class DispersionStatisticTests
{
    private const double Tolerance = 1e-6;

    private static object? Calc(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value switch
            {
                string s => new Danfma.MySheet.Expressions.StringValue(s),
                double d => new NumberValue(d),
                int i => new NumberValue(i),
                bool b => new BooleanValue(b),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static (string, object)[] Column(string column, int firstRow, params double[] values) =>
        [.. values.Select((v, i) => ($"{column}{firstRow + i}", (object)v))];

    // Dataset "Strength" compartilhado pelas 8 páginas STDEV*/VAR*: {1345, 1301, 1368, 1322,
    // 1310, 1370, 1318, 1350, 1303, 1299} em A2:A11.
    private static readonly (string, object)[] Strength = Column(
        "A",
        2,
        1345,
        1301,
        1368,
        1322,
        1310,
        1370,
        1318,
        1350,
        1303,
        1299
    );

    [Test]
    public async Task StDevSAndP_MatchTheGoldenExamples()
    {
        // =STDEV.S(A2:A11) -> 27.46391572 (método "n-1"); =STDEV.P(A2:A11) -> 26.05455814 ("n").
        await Assert
            .That(Num(Calc("=STDEV.S(A2:A11)", Strength)))
            .IsEqualTo(27.46391572)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=STDEV.P(A2:A11)", Strength)))
            .IsEqualTo(26.05455814)
            .Within(Tolerance);
    }

    [Test]
    public async Task VarSAndP_MatchTheGoldenExamples()
    {
        // =VAR.S(A2:A11) -> 754.26667 (dígitos completos impressos na página do VARA);
        // =VAR.P(A2:A11) -> 678.84.
        await Assert.That(Num(Calc("=VAR.S(A2:A11)", Strength))).IsEqualTo(754.26667).Within(1e-4);
        await Assert
            .That(Num(Calc("=VAR.P(A2:A11)", Strength)))
            .IsEqualTo(678.84)
            .Within(Tolerance);
    }

    [Test]
    public async Task AVariants_MatchTheGoldenExamples()
    {
        // As páginas STDEVA/STDEVPA/VARA/VARPA usam o mesmo dataset numérico: 27.46391572 /
        // 26.05455814 / 754.26667 / 678.84.
        await Assert
            .That(Num(Calc("=STDEVA(A2:A11)", Strength)))
            .IsEqualTo(27.46391572)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=STDEVPA(A2:A11)", Strength)))
            .IsEqualTo(26.05455814)
            .Within(Tolerance);
        await Assert.That(Num(Calc("=VARA(A2:A11)", Strength))).IsEqualTo(754.26667).Within(1e-4);
        await Assert
            .That(Num(Calc("=VARPA(A2:A11)", Strength)))
            .IsEqualTo(678.84)
            .Within(Tolerance);
    }

    [Test]
    public async Task AVariants_CountReferencedTextAndLogicals()
    {
        // Regra documentada (STDEVA/VARA): "Arguments that contain TRUE evaluate as 1; arguments
        // that contain text or FALSE evaluate as 0". {1, 2, "x"} conta como {1, 2, 0} -> VARA = 1
        // e STDEVA = 1; {1, 2, TRUE} conta como {1, 2, 1} -> VARA = 1/3 (derivação mecânica da regra).
        (string, object)[] withText = [("A1", 1), ("A2", 2), ("A3", "x")];
        (string, object)[] withTrue = [("A1", 1), ("A2", 2), ("A3", true)];

        await Assert.That(Num(Calc("=VARA(A1:A3)", withText))).IsEqualTo(1.0).Within(Tolerance);
        await Assert.That(Num(Calc("=STDEVA(A1:A3)", withText))).IsEqualTo(1.0).Within(Tolerance);
        await Assert
            .That(Num(Calc("=VARA(A1:A3)", withTrue)))
            .IsEqualTo(1.0 / 3.0)
            .Within(Tolerance);

        // O STDEV.S "não-A" ignora os mesmos valores referenciados: só {1, 2} sobra.
        await Assert
            .That(Num(Calc("=STDEV.S(A1:A3)", withText)))
            .IsEqualTo(Math.Sqrt(0.5))
            .Within(Tolerance);
    }

    [Test]
    public async Task SampleFunctions_WithFewerThanTwoValues_AreDivZero()
    {
        // Método "n-1" indefinido com n < 2 -> #DIV/0! (STDEV.S/VAR.S e variantes A).
        await Assert.That(Calc("=STDEV.S(A1)", ("A1", 5))).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=VAR.S(A1)", ("A1", 5))).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=STDEVA(A1)", ("A1", 5))).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=VARPA(B1:B2)", ("A1", 5))).IsEqualTo(ErrorValue.DivByZero);
    }

    // --- AVEDEV / DEVSQ — golden: páginas oficiais "AVEDEV function" e "DEVSQ function" ---

    [Test]
    public async Task AveDev_MatchesTheGoldenExample()
    {
        // =AVEDEV(A2:A8) com {4,5,6,7,5,4,3} -> 1.020408.
        await Assert
            .That(Num(Calc("=AVEDEV(A2:A8)", Column("A", 2, 4, 5, 6, 7, 5, 4, 3))))
            .IsEqualTo(1.020408)
            .Within(Tolerance);
    }

    [Test]
    public async Task DevSq_MatchesTheGoldenExample()
    {
        // =DEVSQ(A2:A8) com {4,5,8,7,11,4,3} -> 48.
        await Assert
            .That(Num(Calc("=DEVSQ(A2:A8)", Column("A", 2, 4, 5, 8, 7, 11, 4, 3))))
            .IsEqualTo(48.0)
            .Within(Tolerance);
    }

    // --- GEOMEAN / HARMEAN — golden: páginas oficiais "GEOMEAN function" e "HARMEAN function" ---

    [Test]
    public async Task GeoMean_MatchesTheGoldenExample_AndRejectsNonPositives()
    {
        // =GEOMEAN(A2:A8) com {4,5,8,7,11,4,3} -> 5.476987; qualquer ponto <= 0 -> #NUM!.
        var cells = Column("A", 2, 4, 5, 8, 7, 11, 4, 3);

        await Assert
            .That(Num(Calc("=GEOMEAN(A2:A8)", cells)))
            .IsEqualTo(5.476987)
            .Within(Tolerance);
        await Assert.That(Calc("=GEOMEAN(A2:A3,0)", cells)).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task HarMean_MatchesTheGoldenExample_AndRejectsNonPositives()
    {
        // =HARMEAN(A2:A8) com {4,5,8,7,11,4,3} -> 5.028376; qualquer ponto <= 0 -> #NUM!.
        var cells = Column("A", 2, 4, 5, 8, 7, 11, 4, 3);

        await Assert
            .That(Num(Calc("=HARMEAN(A2:A8)", cells)))
            .IsEqualTo(5.028376)
            .Within(Tolerance);
        await Assert.That(Calc("=HARMEAN(A2:A3,-1)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- SKEW / SKEW.P / KURT — golden: páginas oficiais; dataset compartilhado
    // {3,4,5,2,3,4,5,6,4,7} em A2:A11. ---

    private static readonly (string, object)[] SkewData = Column(
        "A",
        2,
        3,
        4,
        5,
        2,
        3,
        4,
        5,
        6,
        4,
        7
    );

    [Test]
    public async Task Skew_MatchesTheGoldenExample()
    {
        // =SKEW(A2:A11) -> 0.359543 (fator amostral n/((n-1)(n-2))).
        await Assert
            .That(Num(Calc("=SKEW(A2:A11)", SkewData)))
            .IsEqualTo(0.359543)
            .Within(Tolerance);
    }

    [Test]
    public async Task SkewP_MatchesTheGoldenExample()
    {
        // =SKEW.P(A2:A11) -> 0.303193 (desvio padrão populacional).
        await Assert
            .That(Num(Calc("=SKEW.P(A2:A11)", SkewData)))
            .IsEqualTo(0.303193)
            .Within(Tolerance);
    }

    [Test]
    public async Task Kurt_MatchesTheGoldenExample()
    {
        // =KURT(A2:A11) -> -0.151799637 (curtose em excesso, fórmula amostral do Excel).
        await Assert
            .That(Num(Calc("=KURT(A2:A11)", SkewData)))
            .IsEqualTo(-0.151799637)
            .Within(Tolerance);
    }

    [Test]
    public async Task SkewAndKurt_WithTooFewPoints_AreDivZero()
    {
        // Regras documentadas: SKEW/SKEW.P com menos de 3 pontos e KURT com menos de 4 -> #DIV/0!
        // (idem desvio padrão zero).
        var two = Column("A", 1, 1, 2);
        var three = Column("A", 1, 1, 2, 3);

        await Assert.That(Calc("=SKEW(A1:A2)", two)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SKEW.P(A1:A2)", two)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=KURT(A1:A3)", three)).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=SKEW(A1:A3)", Column("A", 1, 5, 5, 5)))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    // --- STANDARDIZE — golden: página oficial "STANDARDIZE function" ---

    [Test]
    public async Task Standardize_MatchesTheGoldenExample()
    {
        // =STANDARDIZE(42,40,1.5) -> 1.33333333; standard_dev <= 0 -> #NUM! (regra documentada).
        await Assert
            .That(Num(Calc("=STANDARDIZE(42,40,1.5)")))
            .IsEqualTo(1.33333333)
            .Within(Tolerance);
        await Assert.That(Calc("=STANDARDIZE(42,40,0)")).IsEqualTo(ErrorValue.Number);
    }
}
