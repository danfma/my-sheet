using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — estatísticas de ordem/posição: MEDIAN, MODE.SNGL, LARGE, SMALL, RANK.EQ, RANK.AVG,
// PERCENTILE.INC/.EXC, PERCENTRANK.INC/.EXC, QUARTILE.INC/.EXC e TRIMMEAN. Golden values das
// páginas oficiais da Microsoft (support.microsoft.com, fetched em 2026-07-02), citadas por
// função em cada teste.
public class OrderStatisticTests
{
    private const double Tolerance = 1e-9;

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

    private static (string, object)[] Column(string column, int firstRow, params double[] values) =>
        [.. values.Select((v, i) => ($"{column}{firstRow + i}", (object)v))];

    // --- MEDIAN — golden: página oficial "MEDIAN function" (d0916313-4753-414c-8537-ce85bdd967d2) ---

    [Test]
    public async Task Median_MatchesTheGoldenExamples()
    {
        // Dados 1..6: =MEDIAN(A2:A6) -> 3 (5 valores); =MEDIAN(A2:A7) -> 3.5 (média dos dois do meio).
        var cells = Column("A", 2, 1, 2, 3, 4, 5, 6);

        await Assert.That(Num(Calc("=MEDIAN(A2:A6)", cells))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=MEDIAN(A2:A7)", cells))).IsEqualTo(3.5);
    }

    [Test]
    public async Task Median_Empty_IsNumError()
    {
        await Assert.That(Calc("=MEDIAN(A1:A3)")).IsEqualTo(ErrorValue.Number);
    }

    // --- MODE.SNGL — golden: página oficial "MODE.SNGL function" ---

    [Test]
    public async Task ModeSngl_MatchesTheGoldenExample()
    {
        // =MODE.SNGL(A2:A7) com 5.6/4/4/3/2/4 -> 4 ("most frequently occurring number").
        await Assert
            .That(Num(Calc("=MODE.SNGL(A2:A7)", Column("A", 2, 5.6, 4, 4, 3, 2, 4))))
            .IsEqualTo(4.0);
    }

    [Test]
    public async Task ModeSngl_NoDuplicates_IsNAError()
    {
        // Regra documentada: "If the data set contains no duplicate data points, MODE.SNGL
        // returns the #N/A error value."
        await Assert
            .That(Calc("=MODE.SNGL(A2:A5)", Column("A", 2, 1, 2, 3, 4)))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task ModeSngl_Tie_TakesTheFirstValueInScanOrder()
    {
        // Semântica de desempate definida para o engine: contagens iguais -> o PRIMEIRO valor
        // encontrado na varredura vence (2 e depois 7, ambos duplas -> 2).
        await Assert
            .That(Num(Calc("=MODE.SNGL(A2:A6)", Column("A", 2, 2, 7, 2, 7, 1))))
            .IsEqualTo(2.0);
    }

    // --- LARGE / SMALL — golden: páginas oficiais "LARGE function" e "SMALL function" ---

    [Test]
    public async Task Large_MatchesTheGoldenExamples()
    {
        // Dados A2:B6 = {3,5,3,5,4} / {4,2,4,6,7}: LARGE(...,3) -> 5; LARGE(...,7) -> 4.
        (string, object)[] cells =
        [
            ("A2", 3),
            ("A3", 5),
            ("A4", 3),
            ("A5", 5),
            ("A6", 4),
            ("B2", 4),
            ("B3", 2),
            ("B4", 4),
            ("B5", 6),
            ("B6", 7),
        ];

        await Assert.That(Num(Calc("=LARGE(A2:B6,3)", cells))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=LARGE(A2:B6,7)", cells))).IsEqualTo(4.0);
    }

    [Test]
    public async Task Small_MatchesTheGoldenExamples()
    {
        // Dados 1: {3,4,5,2,3,4,6,4,7}; Dados 2: {1,4,8,3,7,12,54,8,23}:
        // SMALL(A2:A10,4) -> 4; SMALL(B2:B10,2) -> 3.
        var cells = Column("A", 2, 3, 4, 5, 2, 3, 4, 6, 4, 7)
            .Concat(Column("B", 2, 1, 4, 8, 3, 7, 12, 54, 8, 23))
            .ToArray();

        await Assert.That(Num(Calc("=SMALL(A2:A10,4)", cells))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=SMALL(B2:B10,2)", cells))).IsEqualTo(3.0);
    }

    [Test]
    public async Task LargeAndSmall_KOutOfRange_IsNumError()
    {
        // Regra documentada: k <= 0 ou k maior que o número de pontos -> #NUM! (e array vazio -> #NUM!).
        var cells = Column("A", 1, 1, 2, 3);

        await Assert.That(Calc("=LARGE(A1:A3,0)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LARGE(A1:A3,4)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=SMALL(A1:A3,0)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=SMALL(B1:B3,1)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- RANK.EQ / RANK.AVG — golden: páginas oficiais "RANK.EQ function" e "RANK.AVG function" ---

    private static readonly (string, object)[] RankData = Column("A", 2, 7, 3.5, 3.5, 1, 2);

    [Test]
    public async Task RankEq_MatchesTheGoldenExamples()
    {
        // Dados {7, 3.5, 3.5, 1, 2}: RANK.EQ(A2,...,1) -> 5; RANK.EQ(A6,...) -> 4 (desc default);
        // RANK.EQ(A3,...,1) -> 3 (empate recebe o mesmo rank superior).
        await Assert.That(Num(Calc("=RANK.EQ(A2,A2:A6,1)", RankData))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=RANK.EQ(A6,A2:A6)", RankData))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=RANK.EQ(A3,A2:A6,1)", RankData))).IsEqualTo(3.0);
    }

    [Test]
    public async Task RankEq_NumberAbsent_IsNAError()
    {
        await Assert.That(Calc("=RANK.EQ(9,A2:A6)", RankData)).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task RankAvg_MatchesTheGoldenExample_AndAveragesTies()
    {
        // =RANK.AVG(94,B2:B8) com {89,88,92,101,94,97,95} -> 4 (exemplo da página).
        await Assert
            .That(Num(Calc("=RANK.AVG(94,B2:B8)", Column("B", 2, 89, 88, 92, 101, 94, 97, 95))))
            .IsEqualTo(4.0);

        // Empate: a página do RANK.EQ documenta o fator de correção — para 3.5 (dupla) em
        // {7,3.5,3.5,1,2} asc, o rank médio é 3 + 0.5 = 3.5.
        await Assert.That(Num(Calc("=RANK.AVG(A3,A2:A6,1)", RankData))).IsEqualTo(3.5);
    }

    [Test]
    public async Task RankAvg_NumberAbsent_IsNAError()
    {
        await Assert.That(Calc("=RANK.AVG(9,A2:A6)", RankData)).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task RankEq_OrderZeroAndOmitted_AreEquivalent_Descending()
    {
        // order=0 e order omitido resolvem para o mesmo rank descendente.
        await Assert
            .That(Num(Calc("=RANK.EQ(A6,A2:A6,0)", RankData)))
            .IsEqualTo(Num(Calc("=RANK.EQ(A6,A2:A6)", RankData)));
        await Assert
            .That(Num(Calc("=RANK.AVG(A3,A2:A6,0)", RankData)))
            .IsEqualTo(Num(Calc("=RANK.AVG(A3,A2:A6)", RankData)));
    }

    [Test]
    public async Task RankEq_ThreeWayTie_WholeGroupSharesTheTopRank_AndAvgIsTheMean()
    {
        // {10,10,10,7,3} descendente: os três 10 ocupam os ranks 1,2,3 -> RANK.EQ = 1 (topo do
        // grupo) para qualquer um deles; RANK.AVG = média(1,2,3) = 2.
        var cells = Column("A", 2, 10, 10, 10, 7, 3);

        await Assert.That(Num(Calc("=RANK.EQ(A2,A2:A6)", cells))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=RANK.EQ(A4,A2:A6)", cells))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=RANK.AVG(A2,A2:A6)", cells))).IsEqualTo(2.0);

        // Ascendente: o grupo de 10 ocupa os ranks 3,4,5 (os dois menores, 3 e 7, vêm antes) ->
        // RANK.EQ = 3 (topo do grupo em ordem ascendente); RANK.AVG = média(3,4,5) = 4.
        await Assert.That(Num(Calc("=RANK.EQ(A2,A2:A6,1)", cells))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=RANK.AVG(A2,A2:A6,1)", cells))).IsEqualTo(4.0);
    }

    [Test]
    public async Task RankEq_NonNumericValuesInRef_AreIgnored()
    {
        // Texto dentro do ref não participa da contagem — visão numérica efetiva é {5,3,1}.
        (string, object)[] cells =
        [
            ("A2", 5),
            ("A3", "text"),
            ("A4", 3),
            ("A5", "more text"),
            ("A6", 1),
        ];

        await Assert.That(Num(Calc("=RANK.EQ(A2,A2:A6)", cells))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=RANK.EQ(A2,A2:A6,1)", cells))).IsEqualTo(3.0);
    }

    // --- RANK.EQ / RANK.AVG over a SNAPSHOT-ADMITTED range (>= the 256-cell cache threshold,
    // Workbook.RangeCacheMinimumCells) — the binary-search fast path must match the small-range
    // linear-scan path bit for bit. 300 cells B1:B300; every 50th cell is non-numeric ("skip", ignored
    // exactly like any other referenced text) and the rest cycle 1..10, so the numeric view is 294
    // values: 1..9 appear 30x each, 10 appears 24x (its 6 missing occurrences are the "skip" rows,
    // which all land on the value-10 residue by construction). 11 never appears (the #N/A case). ---

    private static (Workbook Workbook, Sheet Sheet) BuildBigRankSheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Rank");

        for (var i = 1; i <= 300; i++)
        {
            sheet[$"B{i}"] =
                i % 50 == 0
                    ? new Danfma.MySheet.Expressions.StringValue("skip")
                    : new NumberValue((i - 1) % 10 + 1);
        }

        return (workbook, sheet);
    }

    // Evaluates `formula` three times against the SAME workbook: once with the range cache bypassed
    // (always the linear path — the parity oracle) and twice with the cache enabled to drive the range
    // past second-use admission (read 1 only marks the range and stays on the linear path; read 2 builds
    // the RangeSnapshot and is served by it). Returns (bypass, admitted) so callers can assert both equal
    // each other AND the hand-derived golden value.
    private static (object? Bypass, object? Admitted) CalcAdmitted(
        Workbook workbook,
        Sheet sheet,
        string formula
    )
    {
        workbook.RangeCacheDisabled = true;
        workbook.InvalidateCache();
        var bypass = ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        workbook.RangeCacheDisabled = false;
        workbook.InvalidateCache();
        _ = ExpressionParser.Parse(formula, sheet).Evaluate(workbook); // read 1 -> marker only
        var admitted = ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject(); // read 2 -> snapshot

        return (bypass, admitted);
    }

    [Test]
    public async Task RankEq_BigSnapshotAdmittedRange_MatchesLinearBypass_Descending()
    {
        var (workbook, sheet) = BuildBigRankSheet();

        // countGreater(5) = 30*4 (6,7,8,9) + 24 (10) = 144 -> rank = 145.
        var (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.EQ(5,B1:B300)");
        await Assert.That(Num(bypass)).IsEqualTo(145.0);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));

        // The max value (10, 24-way tie): countGreater = 0 -> rank = 1 (top of the tied group).
        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.EQ(10,B1:B300)");
        await Assert.That(Num(bypass)).IsEqualTo(1.0);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));
    }

    [Test]
    public async Task RankEq_BigSnapshotAdmittedRange_MatchesLinearBypass_Ascending()
    {
        var (workbook, sheet) = BuildBigRankSheet();

        // countLess(5) = 30*4 (1,2,3,4) = 120 -> rank = 121.
        var (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.EQ(5,B1:B300,1)");
        await Assert.That(Num(bypass)).IsEqualTo(121.0);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));

        // countLess(10) = 270 -> rank = 271 (top of the tied group in ascending order).
        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.EQ(10,B1:B300,1)");
        await Assert.That(Num(bypass)).IsEqualTo(271.0);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));
    }

    [Test]
    public async Task RankAvg_BigSnapshotAdmittedRange_MatchesLinearBypass()
    {
        var (workbook, sheet) = BuildBigRankSheet();

        // Descending: 144 + (30+1)/2 = 159.5.
        var (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.AVG(5,B1:B300)");
        await Assert.That(Num(bypass)).IsEqualTo(159.5);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));

        // Ascending: 120 + (30+1)/2 = 135.5.
        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.AVG(5,B1:B300,1)");
        await Assert.That(Num(bypass)).IsEqualTo(135.5);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));

        // The 24-way tie at the max value: descending 0+(24+1)/2 = 12.5; ascending 270+12.5 = 282.5.
        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.AVG(10,B1:B300)");
        await Assert.That(Num(bypass)).IsEqualTo(12.5);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));

        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.AVG(10,B1:B300,1)");
        await Assert.That(Num(bypass)).IsEqualTo(282.5);
        await Assert.That(Num(admitted)).IsEqualTo(Num(bypass));
    }

    [Test]
    public async Task RankEqAndAvg_BigSnapshotAdmittedRange_AbsentValue_IsNAError()
    {
        var (workbook, sheet) = BuildBigRankSheet();

        var (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.EQ(11,B1:B300)");
        await Assert.That(bypass).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(admitted).IsEqualTo(ErrorValue.NotAvailable);

        (bypass, admitted) = CalcAdmitted(workbook, sheet, "=RANK.AVG(11,B1:B300)");
        await Assert.That(bypass).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(admitted).IsEqualTo(ErrorValue.NotAvailable);
    }

    // --- PERCENTILE.INC — golden: página oficial "PERCENTILE.INC function" ---

    [Test]
    public async Task PercentileInc_MatchesTheGoldenExample()
    {
        // =PERCENTILE.INC(A2:A5,0.3) com {1,3,2,4} -> 1.9 (interpolação em k(n-1)).
        await Assert
            .That(Num(Calc("=PERCENTILE.INC(A2:A5,0.3)", Column("A", 2, 1, 3, 2, 4))))
            .IsEqualTo(1.9)
            .Within(Tolerance);
    }

    [Test]
    public async Task PercentileInc_KOutOfRange_IsNumError()
    {
        // Regra documentada: k < 0 ou k > 1 -> #NUM! (e array vazio -> #NUM!).
        var cells = Column("A", 2, 1, 3, 2, 4);

        await Assert.That(Calc("=PERCENTILE.INC(A2:A5,-0.1)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERCENTILE.INC(A2:A5,1.1)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERCENTILE.INC(B2:B5,0.5)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- PERCENTILE.EXC — golden: página oficial "PERCENTILE.EXC function" (exemplo oficial em
    // imagem, transcrito): dados {1..9}. ---

    [Test]
    public async Task PercentileExc_MatchesTheGoldenExamples()
    {
        // =PERCENTILE.EXC(B2:B10,0.25) -> 2.5 (posição k(n+1)); k=0, 0.01 e 2 -> #NUM!.
        var cells = Column("B", 2, 1, 2, 3, 4, 5, 6, 7, 8, 9);

        await Assert
            .That(Num(Calc("=PERCENTILE.EXC(B2:B10,0.25)", cells)))
            .IsEqualTo(2.5)
            .Within(Tolerance);
        await Assert.That(Calc("=PERCENTILE.EXC(B2:B10,0)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERCENTILE.EXC(B2:B10,0.01)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=PERCENTILE.EXC(B2:B10,2)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- PERCENTRANK.INC — golden: página oficial "PERCENTRANK.INC function":
    // dados {13,12,11,8,4,3,2,1,1,1}. ---

    [Test]
    public async Task PercentRankInc_MatchesTheGoldenExamples()
    {
        // (2) -> 0.333 (3/(3+6)); (4) -> 0.555; (8) -> 0.666; (5) -> 0.583 (interpolado, "um
        // quarto do caminho" entre os ranks de 4 e 8). Resultado TRUNCADO em 3 dígitos por default.
        var cells = Column("A", 2, 13, 12, 11, 8, 4, 3, 2, 1, 1, 1);

        await Assert
            .That(Num(Calc("=PERCENTRANK.INC(A2:A11,2)", cells)))
            .IsEqualTo(0.333)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PERCENTRANK.INC(A2:A11,4)", cells)))
            .IsEqualTo(0.555)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PERCENTRANK.INC(A2:A11,8)", cells)))
            .IsEqualTo(0.666)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PERCENTRANK.INC(A2:A11,5)", cells)))
            .IsEqualTo(0.583)
            .Within(Tolerance);
    }

    [Test]
    public async Task PercentRankInc_OutOfRangeOrBadSignificance_IsError()
    {
        var cells = Column("A", 2, 13, 12, 11, 8, 4, 3, 2, 1, 1, 1);

        // x fora do alcance dos dados -> #N/A; significance < 1 -> #NUM! (regras documentadas).
        await Assert
            .That(Calc("=PERCENTRANK.INC(A2:A11,99)", cells))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=PERCENTRANK.INC(A2:A11,2,0)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- PERCENTRANK.EXC — golden: página oficial "PERCENTRANK.EXC function":
    // dados {1,2,3,6,6,6,7,8,9}. ---

    [Test]
    public async Task PercentRankExc_MatchesTheGoldenExamples()
    {
        // (7) -> 0.7; (5.43) -> 0.381; (5.43, 1) -> 0.3 (1 dígito significativo).
        var cells = Column("A", 2, 1, 2, 3, 6, 6, 6, 7, 8, 9);

        await Assert
            .That(Num(Calc("=PERCENTRANK.EXC(A2:A10,7)", cells)))
            .IsEqualTo(0.7)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PERCENTRANK.EXC(A2:A10,5.43)", cells)))
            .IsEqualTo(0.381)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=PERCENTRANK.EXC(A2:A10,5.43,1)", cells)))
            .IsEqualTo(0.3)
            .Within(Tolerance);
    }

    // --- QUARTILE.INC — golden: página oficial "QUARTILE.INC function" ---

    [Test]
    public async Task QuartileInc_MatchesTheGoldenExample()
    {
        // =QUARTILE.INC(A2:A9,1) com {1,2,4,7,8,9,10,12} -> 3.5.
        var cells = Column("A", 2, 1, 2, 4, 7, 8, 9, 10, 12);

        await Assert
            .That(Num(Calc("=QUARTILE.INC(A2:A9,1)", cells)))
            .IsEqualTo(3.5)
            .Within(Tolerance);

        // Regras documentadas: quart 0 -> mínimo, 4 -> máximo, quart fora de 0-4 -> #NUM!.
        await Assert.That(Num(Calc("=QUARTILE.INC(A2:A9,0)", cells))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=QUARTILE.INC(A2:A9,4)", cells))).IsEqualTo(12.0);
        await Assert.That(Calc("=QUARTILE.INC(A2:A9,5)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- QUARTILE.EXC — golden: página oficial "QUARTILE.EXC function" ---

    [Test]
    public async Task QuartileExc_MatchesTheGoldenExamples()
    {
        // Dados {6,7,15,36,39,40,41,42,43,47,49}: quart 1 -> 15; quart 3 -> 43;
        // quart <= 0 ou >= 4 -> #NUM! (regra documentada).
        var cells = Column("A", 2, 6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49);

        await Assert
            .That(Num(Calc("=QUARTILE.EXC(A2:A12,1)", cells)))
            .IsEqualTo(15.0)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=QUARTILE.EXC(A2:A12,3)", cells)))
            .IsEqualTo(43.0)
            .Within(Tolerance);
        await Assert.That(Calc("=QUARTILE.EXC(A2:A12,0)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=QUARTILE.EXC(A2:A12,4)", cells)).IsEqualTo(ErrorValue.Number);
    }

    // --- TRIMMEAN — golden: página oficial "TRIMMEAN function" ---

    [Test]
    public async Task TrimMean_MatchesTheGoldenExample()
    {
        // =TRIMMEAN(A2:A12,0.2) com {4,5,6,7,2,3,4,5,1,2,3} -> 3.778 (INT(11·0.2/2)=1 cortado de
        // cada ponta; 34/9 = 3.777…).
        await Assert
            .That(
                Num(Calc("=TRIMMEAN(A2:A12,0.2)", Column("A", 2, 4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3)))
            )
            .IsEqualTo(34.0 / 9.0)
            .Within(Tolerance);
    }

    [Test]
    public async Task TrimMean_PercentOutOfRange_IsNumError()
    {
        // Regra documentada: percent < 0 ou > 1 -> #NUM! (percent = 1 também rejeitado: cortaria
        // o conjunto inteiro — contrato percent ∈ [0, 1) declarado no reference).
        var cells = Column("A", 2, 1, 2, 3, 4);

        await Assert.That(Calc("=TRIMMEAN(A2:A5,-0.1)", cells)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=TRIMMEAN(A2:A5,1)", cells)).IsEqualTo(ErrorValue.Number);
    }
}
