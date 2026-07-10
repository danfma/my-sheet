using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — agregações matemáticas: SUMPRODUCT, SUMX2MY2, SUMX2PY2, SUMXMY2 e SUBTOTAL. Golden
// values das páginas oficiais da Microsoft (support.microsoft.com, fetched em 2026-07-02),
// citadas por função em cada teste.
public class MathAggregateTests
{
    private static object? Calc(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value switch
            {
                string s when s.StartsWith('=') => ExpressionParser.Parse(s, sheet),
                string s => new Danfma.MySheet.Expressions.StringValue(s),
                double d => new NumberValue(d),
                int i => new NumberValue(i),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- SUMPRODUCT — golden: página oficial "SUMPRODUCT function"
    // (16753e75-9f68-4874-94ac-4d2145a2fd2e). O dataset do exemplo 1 só existe como imagem na
    // página; a EQUIVALÊNCIA documentada é o oráculo: "=SUMPRODUCT(C2:C5,D2:D5)" tem o mesmo
    // resultado da forma longa "=C2*D2+C3*D3+C4*D4+C5*D5". ---

    [Test]
    public async Task SumProduct_EqualsTheDocumentedLongForm()
    {
        (string, object)[] groceries =
        [
            ("C2", 2.5),
            ("C3", 4.25),
            ("C4", 8.0),
            ("C5", 3.11),
            ("D2", 4),
            ("D3", 3),
            ("D4", 1),
            ("D5", 5),
        ];

        var product = Num(Calc("=SUMPRODUCT(C2:C5,D2:D5)", groceries));
        var longForm = Num(Calc("=C2*D2+C3*D3+C4*D4+C5*D5", groceries));

        await Assert.That(product).IsEqualTo(longForm);
    }

    [Test]
    public async Task SumProduct_TreatsNonNumericEntriesAsZero()
    {
        // Regra documentada: "SUMPRODUCT treats non-numeric array entries as if they were zeros."
        (string, object)[] cells =
        [
            ("A1", 2),
            ("A2", "x"),
            ("A3", 4),
            ("B1", 10),
            ("B2", 20),
            ("B3", 30),
        ];

        await Assert.That(Num(Calc("=SUMPRODUCT(A1:A3,B1:B3)", cells))).IsEqualTo(140.0);

        // Um único array: SUMPRODUCT soma os valores (não-numéricos contam 0).
        await Assert.That(Num(Calc("=SUMPRODUCT(A1:A3)", cells))).IsEqualTo(6.0);
    }

    [Test]
    public async Task SumProduct_ShapeMismatch_IsValueError()
    {
        // Regra documentada: "The array arguments must have the same dimensions. If they do not,
        // SUMPRODUCT returns the #VALUE! error value."
        (string, object)[] cells = [("A1", 1), ("A2", 2), ("B1", 3)];

        await Assert.That(Calc("=SUMPRODUCT(A1:A2,B1:B3)", cells)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task SumProduct_PropagatesACellError()
    {
        // A cell error inside any array propagates as the function result (before it could be
        // treated as a non-numeric zero). The scan is position-major then array-major, so the
        // first error in that order wins.
        (string, object)[] cells =
        [
            ("A1", 2),
            ("A2", "=1/0"),
            ("A3", 4),
            ("B1", 10),
            ("B2", 20),
            ("B3", 30),
        ];

        await Assert.That(Calc("=SUMPRODUCT(A1:A3,B1:B3)", cells)).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task SumProduct_ShapeMismatchIsReportedBeforeACellError()
    {
        // Dimension validation runs ahead of the value scan: a length mismatch is #VALUE! even when
        // a cell error is also present in one of the arrays.
        (string, object)[] cells = [("A1", 1), ("A2", "=1/0"), ("B1", 3), ("B2", 4), ("B3", 5)];

        await Assert.That(Calc("=SUMPRODUCT(A1:A2,B1:B3)", cells)).IsEqualTo(ErrorValue.NotValue);
    }

    // --- SUMX2MY2 / SUMX2PY2 / SUMXMY2 — golden: páginas oficiais (9e599cc5, 826b60b4,
    // 9d144ac1): array_x {2,3,9,1,8,7,5}, array_y {6,5,11,7,5,4,4}. ---

    private static readonly (string, object)[] PairData =
    [
        ("A2", 2),
        ("A3", 3),
        ("A4", 9),
        ("A5", 1),
        ("A6", 8),
        ("A7", 7),
        ("A8", 5),
        ("B2", 6),
        ("B3", 5),
        ("B4", 11),
        ("B5", 7),
        ("B6", 5),
        ("B7", 4),
        ("B8", 4),
    ];

    [Test]
    public async Task SumXFamily_MatchesTheGoldenExamples()
    {
        // =SUMX2MY2(...) -> -55; =SUMX2PY2(...) -> 521; =SUMXMY2(...) -> 79.
        await Assert.That(Num(Calc("=SUMX2MY2(A2:A8,B2:B8)", PairData))).IsEqualTo(-55.0);
        await Assert.That(Num(Calc("=SUMX2PY2(A2:A8,B2:B8)", PairData))).IsEqualTo(521.0);
        await Assert.That(Num(Calc("=SUMXMY2(A2:A8,B2:B8)", PairData))).IsEqualTo(79.0);
    }

    [Test]
    public async Task SumXFamily_LengthMismatch_IsNA()
    {
        // Regra documentada: número de valores diferente entre array_x e array_y -> #N/A.
        await Assert
            .That(Calc("=SUMX2MY2(A2:A7,B2:B8)", PairData))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task SumXFamily_IgnoresPairsWithANonNumericSide()
    {
        // Regra documentada: texto/lógicos/vazios são ignorados — o PAR inteiro cai
        // (diferente do SUMPRODUCT, que zera a entrada).
        (string, object)[] cells =
        [
            ("A1", 3),
            ("A2", "x"),
            ("A3", 5),
            ("B1", 1),
            ("B2", 100),
            ("B3", 2),
        ];

        // Só os pares (3,1) e (5,2): (9-1) + (25-4) = 29.
        await Assert.That(Num(Calc("=SUMX2MY2(A1:A3,B1:B3)", cells))).IsEqualTo(29.0);
    }

    // --- SUBTOTAL — golden: página oficial "SUBTOTAL function"
    // (7b027003-f060-4ade-9040-e478765b9939): A2 = 120, A3 = 10, A4 = 150, A5 = 23. ---

    private static readonly (string, object)[] SubtotalData =
    [
        ("A2", 120),
        ("A3", 10),
        ("A4", 150),
        ("A5", 23),
    ];

    [Test]
    public async Task Subtotal_MatchesTheGoldenExamples()
    {
        // =SUBTOTAL(9,A2:A5) -> 303; =SUBTOTAL(1,A2:A5) -> 75.75.
        await Assert.That(Num(Calc("=SUBTOTAL(9,A2:A5)", SubtotalData))).IsEqualTo(303.0);
        await Assert.That(Num(Calc("=SUBTOTAL(1,A2:A5)", SubtotalData))).IsEqualTo(75.75);
    }

    [Test]
    public async Task Subtotal_HiddenVariants_BehaveTheSame()
    {
        // 101-111 = mesma agregação; o MySheet não tem linhas ocultas (limite de modelo
        // documentado no function-reference, §A5 do plano).
        await Assert.That(Num(Calc("=SUBTOTAL(109,A2:A5)", SubtotalData))).IsEqualTo(303.0);
        await Assert.That(Num(Calc("=SUBTOTAL(101,A2:A5)", SubtotalData))).IsEqualTo(75.75);
    }

    [Test]
    public async Task Subtotal_MapsEveryDocumentedCode()
    {
        // Tabela documentada de function_num: 2 COUNT, 3 COUNTA, 4 MAX, 5 MIN, 6 PRODUCT,
        // 7 STDEV, 8 STDEVP, 10 VAR, 11 VARP — derivados mecanicamente do dataset da página.
        (string, object)[] cells = [.. SubtotalData, ("A6", "text")];

        await Assert.That(Num(Calc("=SUBTOTAL(2,A2:A6)", cells))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=SUBTOTAL(3,A2:A6)", cells))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=SUBTOTAL(4,A2:A5)", cells))).IsEqualTo(150.0);
        await Assert.That(Num(Calc("=SUBTOTAL(5,A2:A5)", cells))).IsEqualTo(10.0);
        await Assert.That(Num(Calc("=SUBTOTAL(6,A2:A5)", cells))).IsEqualTo(120.0 * 10 * 150 * 23);

        // 7/8/10/11 têm que casar com as funções homônimas avaliadas no mesmo range
        // (guarda anti-vacuidade: o lado SUBTOTAL tem que ser um número de verdade).
        await Assert.That(double.IsNaN(Num(Calc("=SUBTOTAL(7,A2:A5)", cells)))).IsFalse();
        await Assert
            .That(Num(Calc("=SUBTOTAL(7,A2:A5)", cells)))
            .IsEqualTo(Num(Calc("=STDEV.S(A2:A5)", cells)));
        await Assert
            .That(Num(Calc("=SUBTOTAL(8,A2:A5)", cells)))
            .IsEqualTo(Num(Calc("=STDEV.P(A2:A5)", cells)));
        await Assert
            .That(Num(Calc("=SUBTOTAL(10,A2:A5)", cells)))
            .IsEqualTo(Num(Calc("=VAR.S(A2:A5)", cells)));
        await Assert
            .That(Num(Calc("=SUBTOTAL(11,A2:A5)", cells)))
            .IsEqualTo(Num(Calc("=VAR.P(A2:A5)", cells)));
    }

    [Test]
    public async Task Subtotal_IgnoresNestedSubtotals()
    {
        // Regra documentada: "If there are other subtotals within ref1, ref2,… (or nested
        // subtotals), these nested subtotals are ignored to avoid double counting." B1 e B3 são
        // SUBTOTALs (130 e 173); só B2 = 5 conta no SUBTOTAL externo.
        (string, object)[] cells =
        [
            .. SubtotalData,
            ("B1", "=SUBTOTAL(9,A2:A3)"),
            ("B2", 5),
            ("B3", "=SUBTOTAL(9,A4:A5)"),
        ];

        await Assert.That(Num(Calc("=SUBTOTAL(9,B1:B3)", cells))).IsEqualTo(5.0);

        // Os SUBTOTALs internos continuam avaliando normalmente quando referenciados direto.
        await Assert.That(Num(Calc("=B1+B3", cells))).IsEqualTo(303.0);

        // COUNTA aninhado também pula as células SUBTOTAL.
        await Assert.That(Num(Calc("=SUBTOTAL(3,B1:B3)", cells))).IsEqualTo(1.0);
    }

    [Test]
    public async Task Subtotal_InvalidCode_IsValueError()
    {
        // function_num fora de 1-11/101-111 -> #VALUE!.
        await Assert.That(Calc("=SUBTOTAL(0,A2:A5)", SubtotalData)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SUBTOTAL(12,A2:A5)", SubtotalData)).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=SUBTOTAL(100,A2:A5)", SubtotalData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=SUBTOTAL(112,A2:A5)", SubtotalData))
            .IsEqualTo(ErrorValue.NotValue);
    }
}
