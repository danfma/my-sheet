using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — estatísticas bivariadas (CORREL, PEARSON, COVARIANCE.P/.S, RSQ, SLOPE, INTERCEPT,
// STEYX, FORECAST.LINEAR) e escalares (FISHER, FISHERINV, PHI, PERMUT, PERMUTATIONA, PROB).
// Golden values das páginas oficiais da Microsoft (support.microsoft.com, fetched em 2026-07-02),
// citadas por função em cada teste.
public class BivariateStatisticTests
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
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static (string, object)[] Columns(
        int firstRow,
        double[] columnA,
        double[] columnB
    ) =>
        [
            .. columnA.Select((v, i) => ($"A{firstRow + i}", (object)v)),
            .. columnB.Select((v, i) => ($"B{firstRow + i}", (object)v)),
        ];

    // Dataset compartilhado por CORREL (screenshot oficial) e COVARIANCE.P:
    // Data1 {3,2,4,5,6}, Data2 {9,7,12,15,17}.
    private static readonly (string, object)[] CorrelData =
        Columns(2, [3, 2, 4, 5, 6], [9, 7, 12, 15, 17]);

    // Dataset compartilhado por RSQ e STEYX: known_y {2,3,9,1,8,7,5}, known_x {6,5,11,7,5,4,4}
    // (a página do SLOPE usa o MESMO known_y impresso como datas seriais 2,3,9,1,8,7,5).
    private static readonly (string, object)[] RegressionData =
        Columns(3, [2, 3, 9, 1, 8, 7, 5], [6, 5, 11, 7, 5, 4, 4]);

    // --- CORREL / PEARSON ---

    [Test]
    public async Task Correl_MatchesTheGoldenExample()
    {
        // =CORREL(A2:A6,B2:B6) -> 0.997054486 (página oficial "CORREL function",
        // 995dcef7-0c0a-4bed-a3fb-239d7b68ca92; exemplo publicado como screenshot, transcrito).
        await Assert
            .That(Num(Calc("=CORREL(A2:A6,B2:B6)", CorrelData)))
            .IsEqualTo(0.997054486)
            .Within(Tolerance);
    }

    [Test]
    public async Task Pearson_MatchesTheGoldenExample()
    {
        // =PEARSON(A3:A7,B3:B7) com {9,7,5,3,1}/{10,6,1,5,3} -> 0.699379 (página oficial
        // "PEARSON function", 0c3e30fc-e5af-49c4-808a-3ef66e034c18).
        await Assert
            .That(Num(Calc("=PEARSON(A3:A7,B3:B7)", Columns(3, [9, 7, 5, 3, 1], [10, 6, 1, 5, 3]))))
            .IsEqualTo(0.699379)
            .Within(Tolerance);
    }

    [Test]
    public async Task Correl_ShapeMismatch_IsNA_AndZeroVariance_IsDivZero()
    {
        // Regras documentadas: números de pontos diferentes -> #N/A; desvio padrão zero -> #DIV/0!.
        await Assert
            .That(Calc("=CORREL(A2:A5,B2:B6)", CorrelData))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Calc("=CORREL(A1:A3,B1:B3)", Columns(1, [5, 5, 5], [1, 2, 3])))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Correl_IgnoresPairsWithANonNumericSide()
    {
        // Regra documentada: texto/lógicos/vazios em array são ignorados — o PAR inteiro cai.
        // Com o par extra (texto, 99) o resultado tem que ser o MESMO do dataset golden.
        (string, object)[] cells =
        [
            .. CorrelData, ("A7", "x"), ("B7", 99),
        ];

        await Assert
            .That(Num(Calc("=CORREL(A2:A7,B2:B7)", cells)))
            .IsEqualTo(0.997054486)
            .Within(Tolerance);
    }

    // --- COVARIANCE.P / COVARIANCE.S ---

    [Test]
    public async Task CovarianceP_MatchesTheGoldenExample()
    {
        // =COVARIANCE.P(A2:A6,B2:B6) -> 5.2 (página oficial "COVARIANCE.P function",
        // 6f0e1e6d-956d-4e4b-9943-cfef0bf9edfc).
        await Assert
            .That(Num(Calc("=COVARIANCE.P(A2:A6,B2:B6)", CorrelData)))
            .IsEqualTo(5.2)
            .Within(Tolerance);
    }

    [Test]
    public async Task CovarianceS_MatchesTheGoldenExample()
    {
        // =COVARIANCE.S(A3:A5,B3:B5) com {2,4,8}/{5,11,12} -> 9.666666667 (página oficial
        // "COVARIANCE.S function", 0a539b74-7371-42aa-a18f-1f5320314977); 1 ponto -> #DIV/0!.
        var cells = Columns(3, [2, 4, 8], [5, 11, 12]);

        await Assert
            .That(Num(Calc("=COVARIANCE.S(A3:A5,B3:B5)", cells)))
            .IsEqualTo(9.666666667)
            .Within(Tolerance);
        await Assert
            .That(Calc("=COVARIANCE.S(A3,B3)", cells))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    // --- RSQ / SLOPE / INTERCEPT / STEYX / FORECAST.LINEAR ---

    [Test]
    public async Task Rsq_MatchesTheGoldenExample()
    {
        // =RSQ(A3:A9,B3:B9) -> 0.05795 (página oficial "RSQ function",
        // d7161715-250d-4a01-b80d-a8364f2be08f).
        await Assert
            .That(Num(Calc("=RSQ(A3:A9,B3:B9)", RegressionData)))
            .IsEqualTo(0.05795)
            .Within(1e-5);
    }

    [Test]
    public async Task Slope_MatchesTheGoldenExample()
    {
        // =SLOPE(A3:A9,B3:B9) -> 0.305556 (página oficial "SLOPE function",
        // 11fb8f97-3117-4813-98aa-61d7e01276b9; known_y impresso como datas seriais 2,3,9,1,8,7,5).
        await Assert
            .That(Num(Calc("=SLOPE(A3:A9,B3:B9)", RegressionData)))
            .IsEqualTo(0.305556)
            .Within(Tolerance);
    }

    [Test]
    public async Task Intercept_MatchesTheGoldenExample()
    {
        // =INTERCEPT(A2:A6,B2:B6) com y {2,3,9,1,8} / x {6,5,11,7,5} -> 0.0483871 (página oficial
        // "INTERCEPT function", 2a9b74e2-9d47-4772-b663-3bca70bf63ef).
        await Assert
            .That(Num(Calc("=INTERCEPT(A2:A6,B2:B6)", Columns(2, [2, 3, 9, 1, 8], [6, 5, 11, 7, 5]))))
            .IsEqualTo(0.0483871)
            .Within(Tolerance);
    }

    [Test]
    public async Task Steyx_MatchesTheGoldenExample()
    {
        // =STEYX(A3:A9,B3:B9) -> 3.305719 (página oficial "STEYX function",
        // 6ce74b2c-449d-4a6e-b9ac-f9cef5ba48ab); menos de 3 pontos -> #DIV/0! (regra documentada).
        await Assert
            .That(Num(Calc("=STEYX(A3:A9,B3:B9)", RegressionData)))
            .IsEqualTo(3.305719)
            .Within(Tolerance);
        await Assert
            .That(Calc("=STEYX(A3:A4,B3:B4)", RegressionData))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task ForecastLinear_MatchesTheGoldenExample()
    {
        // =FORECAST.LINEAR(30,A2:A6,B2:B6) com known_y {6,7,9,15,21} / known_x {20,28,31,38,40}
        // -> 10.607253 (página oficial "FORECAST and FORECAST.LINEAR functions",
        // 50ca49c9-7b40-4892-94e4-7ad38bbeda99). Atenção à ordem: x novo, DEPOIS ys, DEPOIS xs.
        var cells = Columns(2, [6, 7, 9, 15, 21], [20, 28, 31, 38, 40]);

        await Assert
            .That(Num(Calc("=FORECAST.LINEAR(30,A2:A6,B2:B6)", cells)))
            .IsEqualTo(10.607253)
            .Within(Tolerance);

        // Regras documentadas: x não numérico -> #VALUE!; var(known_x) = 0 -> #DIV/0!;
        // tamanhos diferentes -> #N/A.
        await Assert
            .That(Calc("=FORECAST.LINEAR(\"a\",A2:A6,B2:B6)", cells))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=FORECAST.LINEAR(30,A2:A4,B2:B4)", Columns(2, [1, 2, 3], [7, 7, 7])))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=FORECAST.LINEAR(30,A2:A5,B2:B6)", cells))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    // --- FISHER / FISHERINV / PHI — golden: páginas oficiais ---

    [Test]
    public async Task Fisher_MatchesTheGoldenExample()
    {
        // =FISHER(0.75) -> 0.9729551; |x| >= 1 -> #NUM! (regras documentadas).
        await Assert.That(Num(Calc("=FISHER(0.75)"))).IsEqualTo(0.9729551).Within(Tolerance);
        await Assert.That(Calc("=FISHER(1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=FISHER(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task FisherInv_MatchesTheGoldenExample()
    {
        // =FISHERINV(0.972955) -> 0.75.
        await Assert.That(Num(Calc("=FISHERINV(0.972955)"))).IsEqualTo(0.75).Within(Tolerance);
    }

    [Test]
    public async Task Phi_MatchesTheGoldenExample()
    {
        // =PHI(0.75) -> 0.301137432 (densidade da normal padrão).
        await Assert.That(Num(Calc("=PHI(0.75)"))).IsEqualTo(0.301137432).Within(Tolerance);
    }

    // --- PERMUT / PERMUTATIONA — golden: páginas oficiais ---

    [Test]
    public async Task Permut_MatchesTheGoldenExamples()
    {
        // =PERMUT(100,3) -> 970200; =PERMUT(3,2) -> 6; n <= 0, k < 0 ou n < k -> #NUM!
        // (regras documentadas; argumentos truncados).
        await Assert.That(Num(Calc("=PERMUT(100,3)"))).IsEqualTo(970200.0);
        await Assert.That(Num(Calc("=PERMUT(3,2)"))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=PERMUT(3.9,2.9)"))).IsEqualTo(6.0);
        await Assert.That(Calc("=PERMUT(0,0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERMUT(3,-1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERMUT(2,3)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task PermutationA_MatchesTheGoldenExamples()
    {
        // =PERMUTATIONA(3,2) -> 9; =PERMUTATIONA(2,2) -> 4 (n^k, com repetição).
        await Assert.That(Num(Calc("=PERMUTATIONA(3,2)"))).IsEqualTo(9.0);
        await Assert.That(Num(Calc("=PERMUTATIONA(2,2)"))).IsEqualTo(4.0);
        await Assert.That(Calc("=PERMUTATIONA(-1,2)")).IsEqualTo(ErrorValue.Number);
    }

    // --- PROB — golden: página oficial "PROB function" (9ac30561-c81c-4259-8253-34f0a238fc49):
    // x {0,1,2,3} com probabilidades {0.2,0.3,0.1,0.4}. ---

    private static readonly (string, object)[] ProbData =
        Columns(3, [0, 1, 2, 3], [0.2, 0.3, 0.1, 0.4]);

    [Test]
    public async Task Prob_MatchesTheGoldenExamples()
    {
        // =PROB(A3:A6,B3:B6,2) -> 0.1 (upper omitido: x igual a lower);
        // =PROB(A3:A6,B3:B6,1,3) -> 0.8.
        await Assert
            .That(Num(Calc("=PROB(A3:A6,B3:B6,2)", ProbData)))
            .IsEqualTo(0.1)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PROB(A3:A6,B3:B6,1,3)", ProbData)))
            .IsEqualTo(0.8)
            .Within(Tolerance);
    }

    [Test]
    public async Task Prob_InvalidProbabilities_AreNum_AndShapeMismatch_IsNA()
    {
        // Regras documentadas: prob <= 0 ou > 1 -> #NUM!; soma != 1 -> #NUM!;
        // tamanhos diferentes -> #N/A.
        await Assert
            .That(Calc("=PROB(A1:A2,B1:B2,1)", Columns(1, [1, 2], [0.5, 1.5])))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Calc("=PROB(A1:A2,B1:B2,1)", Columns(1, [1, 2], [0.5, 0.4])))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Calc("=PROB(A3:A6,B3:B5,2)", ProbData))
            .IsEqualTo(ErrorValue.NotAvailable);
    }
}
