using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 4 — condicionais estatísticas (AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS) e variantes A
// (AVERAGEA, MAXA, MINA). Golden values das páginas oficiais da Microsoft
// (support.microsoft.com, fetched em 2026-07-02), citadas por função em cada teste.
public class StatisticalConditionalTests
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
                bool b => new BooleanValue(b),
                _ => throw new ArgumentException($"Unsupported cell value: {value.GetType()}"),
            };
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- AVERAGEIF — golden: página oficial "AVERAGEIF function"
    // (faec8e2e-0dec-4308-af69-f5576d8ac642): A2:A5 = 100000/200000/300000/400000 (property),
    // B2:B5 = 7000/14000/21000/28000 (commission). ---

    private static readonly (string, object)[] Commissions =
    [
        ("A2", 100000),
        ("A3", 200000),
        ("A4", 300000),
        ("A5", 400000),
        ("B2", 7000),
        ("B3", 14000),
        ("B4", 21000),
        ("B5", 28000),
    ];

    [Test]
    public async Task AverageIf_MatchesTheGoldenExamples()
    {
        // =AVERAGEIF(B2:B5,"<23000") -> 14000; =AVERAGEIF(A2:A5,"<250000") -> 150000;
        // =AVERAGEIF(A2:A5,">250000",B2:B5) -> 24500 (exemplos da página).
        await Assert
            .That(Num(Calc("=AVERAGEIF(B2:B5,\"<23000\")", Commissions)))
            .IsEqualTo(14000.0);
        await Assert
            .That(Num(Calc("=AVERAGEIF(A2:A5,\"<250000\")", Commissions)))
            .IsEqualTo(150000.0);
        await Assert
            .That(Num(Calc("=AVERAGEIF(A2:A5,\">250000\",B2:B5)", Commissions)))
            .IsEqualTo(24500.0);
    }

    [Test]
    public async Task AverageIf_NoMatch_IsDivZero()
    {
        // =AVERAGEIF(A2:A5,"<95000") -> #DIV/0! ("there are 0 property values that meet this
        // condition", exemplo da página).
        await Assert
            .That(Calc("=AVERAGEIF(A2:A5,\"<95000\")", Commissions))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task AverageIf_Wildcards_MatchTheGoldenExample2()
    {
        // Exemplo 2 da página: regiões East/West/North/South (New Office)/MidWest com lucros
        // 45678/23789/-4789/0/9678; "=*West" -> 16733.5 e "<>*(New Office)" -> 18589.
        (string, object)[] regions =
        [
            ("A2", "East"),
            ("A3", "West"),
            ("A4", "North"),
            ("A5", "South (New Office)"),
            ("A6", "MidWest"),
            ("B2", 45678),
            ("B3", 23789),
            ("B4", -4789),
            ("B5", 0),
            ("B6", 9678),
        ];

        await Assert
            .That(Num(Calc("=AVERAGEIF(A2:A6,\"=*West\",B2:B6)", regions)))
            .IsEqualTo(16733.5);
        await Assert
            .That(Num(Calc("=AVERAGEIF(A2:A6,\"<>*(New Office)\",B2:B6)", regions)))
            .IsEqualTo(18589.0);
    }

    // --- AVERAGEIFS — golden: página oficial "AVERAGEIFS function"
    // (48910c45-1fc0-4389-a028-f7c5c3001690), exemplo 2 (imóveis): B = preço, C = cidade,
    // D = quartos, E = garagem. (O exemplo 1 da página imprime um Result inconsistente com a
    // própria descrição — 75 vs 80.5 — e foi evitado; o exemplo 2 é consistente.) ---

    private static readonly (string, object)[] Homes =
    [
        ("B2", 230000),
        ("B3", 197000),
        ("B4", 345678),
        ("B5", 321900),
        ("B6", 450000),
        ("B7", 395000),
        ("C2", "Issaquah"),
        ("C3", "Bellevue"),
        ("C4", "Bellevue"),
        ("C5", "Issaquah"),
        ("C6", "Bellevue"),
        ("C7", "Bellevue"),
        ("D2", 3),
        ("D3", 2),
        ("D4", 4),
        ("D5", 2),
        ("D6", 5),
        ("D7", 4),
        ("E2", "No"),
        ("E3", "Yes"),
        ("E4", "Yes"),
        ("E5", "Yes"),
        ("E6", "Yes"),
        ("E7", "No"),
    ];

    [Test]
    public async Task AverageIfs_MatchesTheGoldenExamples()
    {
        // =AVERAGEIFS(B2:B7,C2:C7,"Bellevue",D2:D7,">2",E2:E7,"Yes") -> 397839;
        // =AVERAGEIFS(B2:B7,C2:C7,"Issaquah",D2:D7,"<=3",E2:E7,"No") -> 230000.
        await Assert
            .That(
                Num(Calc("=AVERAGEIFS(B2:B7,C2:C7,\"Bellevue\",D2:D7,\">2\",E2:E7,\"Yes\")", Homes))
            )
            .IsEqualTo(397839.0);
        await Assert
            .That(
                Num(Calc("=AVERAGEIFS(B2:B7,C2:C7,\"Issaquah\",D2:D7,\"<=3\",E2:E7,\"No\")", Homes))
            )
            .IsEqualTo(230000.0);
    }

    [Test]
    public async Task AverageIfs_NoMatch_IsDivZero()
    {
        // Exemplo 1 da página: =AVERAGEIFS(C2:C5,C2:C5,">95") -> #DIV/0! (nenhuma nota > 95).
        await Assert
            .That(
                Calc(
                    "=AVERAGEIFS(C2:C5,C2:C5,\">95\")",
                    ("C2", 85),
                    ("C3", 80),
                    ("C4", 93),
                    ("C5", 75)
                )
            )
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task AverageIfs_ShapeMismatch_IsValueError()
    {
        // Regra documentada: "Each criteria_range must be the same size and shape as average_range."
        await Assert
            .That(Calc("=AVERAGEIFS(B2:B4,C2:C7,\"x\")", Homes))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // --- MAXIFS / MINIFS — golden: páginas oficiais "MAXIFS function"
    // (dfd611e6-da2c-488a-919b-9b6376b28883) e "MINIFS function"
    // (6ca1ddaa-079b-4e74-80cc-72eef32e6599). ---

    [Test]
    public async Task MaxIfs_MatchesTheGoldenExamples()
    {
        // Exemplo 1: grades 89/93/96/85/91/88 com pesos 1/2/2/3/1/1 -> MAXIFS(A2:A7,B2:B7,1) = 91.
        (string, object)[] grades =
        [
            ("A2", 89),
            ("A3", 93),
            ("A4", 96),
            ("A5", 85),
            ("A6", 91),
            ("A7", 88),
            ("B2", 1),
            ("B3", 2),
            ("B4", 2),
            ("B5", 3),
            ("B6", 1),
            ("B7", 1),
        ];

        await Assert.That(Num(Calc("=MAXIFS(A2:A7,B2:B7,1)", grades))).IsEqualTo(91.0);

        // Exemplo 3: pesos 10/1/100/1/1/50, grades b/a/a/b/a/b, levels 100/100/200/300/100/400
        // -> MAXIFS(A2:A7,B2:B7,"b",D2:D7,">100") = 50.
        (string, object)[] levels =
        [
            ("A2", 10),
            ("A3", 1),
            ("A4", 100),
            ("A5", 1),
            ("A6", 1),
            ("A7", 50),
            ("B2", "b"),
            ("B3", "a"),
            ("B4", "a"),
            ("B5", "b"),
            ("B6", "a"),
            ("B7", "b"),
            ("D2", 100),
            ("D3", 100),
            ("D4", 200),
            ("D5", 300),
            ("D6", 100),
            ("D7", 400),
        ];

        await Assert
            .That(Num(Calc("=MAXIFS(A2:A7,B2:B7,\"b\",D2:D7,\">100\")", levels)))
            .IsEqualTo(50.0);
    }

    [Test]
    public async Task MaxIfs_NoMatch_IsZero_AndShapeMismatch_IsValueError()
    {
        // Exemplo 6 (sem match -> 0) e exemplo 5 (shapes diferentes -> #VALUE!) da página.
        (string, object)[] cells =
        [
            ("A2", 10),
            ("A3", 1),
            ("A4", 100),
            ("A5", 1),
            ("A6", 1),
            ("B2", "b"),
            ("B3", "a"),
            ("B4", "a"),
            ("B5", "b"),
            ("B6", "a"),
            ("D2", 100),
            ("D3", 100),
            ("D4", 200),
            ("D5", 300),
            ("D6", 100),
        ];

        await Assert
            .That(Num(Calc("=MAXIFS(A2:A6,B2:B6,\"a\",D2:D6,\">200\")", cells)))
            .IsEqualTo(0.0);
        await Assert.That(Calc("=MAXIFS(A2:A5,B2:C6,\"a\")", cells)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task MinIfs_MatchesTheGoldenExamples()
    {
        // Exemplo 1: grades 89/93/96/85/91/88 com pesos 1/2/2/3/1/1 -> MINIFS(A2:A7,B2:B7,1) = 88.
        (string, object)[] grades =
        [
            ("A2", 89),
            ("A3", 93),
            ("A4", 96),
            ("A5", 85),
            ("A6", 91),
            ("A7", 88),
            ("B2", 1),
            ("B3", 2),
            ("B4", 2),
            ("B5", 3),
            ("B6", 1),
            ("B7", 1),
        ];

        await Assert.That(Num(Calc("=MINIFS(A2:A7,B2:B7,1)", grades))).IsEqualTo(88.0);

        // Exemplo 3: pesos 10..15, grades b/a/a/b/b/b, levels 100/100/200/300/300/400
        // -> MINIFS(A2:A7,B2:B7,"b",D2:D7,">100") = 13.
        (string, object)[] levels =
        [
            ("A2", 10),
            ("A3", 11),
            ("A4", 12),
            ("A5", 13),
            ("A6", 14),
            ("A7", 15),
            ("B2", "b"),
            ("B3", "a"),
            ("B4", "a"),
            ("B5", "b"),
            ("B6", "b"),
            ("B7", "b"),
            ("D2", 100),
            ("D3", 100),
            ("D4", 200),
            ("D5", 300),
            ("D6", 300),
            ("D7", 400),
        ];

        await Assert
            .That(Num(Calc("=MINIFS(A2:A7,B2:B7,\"b\",D2:D7,\">100\")", levels)))
            .IsEqualTo(13.0);
    }

    // --- Variantes A — golden: páginas oficiais "AVERAGEA function"
    // (f5f84098-d453-4f4c-bbba-3d2c66356091), "MAXA function" (814bda1e-3840-4bff-9365-2f59ac2ee62d)
    // e "MINA function" (245a6f46-7ca5-4dc7-ab49-805341bc31d3). ---

    [Test]
    public async Task AverageA_TextCountsAsZero()
    {
        // =AVERAGEA(A2:A6) com 10/7/9/2/"Not available" -> 5.6 (o texto conta como 0).
        await Assert
            .That(
                Num(
                    Calc(
                        "=AVERAGEA(A2:A6)",
                        ("A2", 10),
                        ("A3", 7),
                        ("A4", 9),
                        ("A5", 2),
                        ("A6", "Not available")
                    )
                )
            )
            .IsEqualTo(5.6);
    }

    [Test]
    public async Task MaxA_TrueCountsAsOne()
    {
        // =MAXA(A2:A6) com 0/0.2/0.5/0.4/TRUE -> 1 (TRUE avalia como 1).
        await Assert
            .That(
                Num(
                    Calc(
                        "=MAXA(A2:A6)",
                        ("A2", 0),
                        ("A3", 0.2),
                        ("A4", 0.5),
                        ("A5", 0.4),
                        ("A6", true)
                    )
                )
            )
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task MinA_FalseCountsAsZero()
    {
        // =MINA(A2:A6) com FALSE/0.2/0.5/0.4/0.8 -> 0 (FALSE avalia como 0).
        await Assert
            .That(
                Num(
                    Calc(
                        "=MINA(A2:A6)",
                        ("A2", false),
                        ("A3", 0.2),
                        ("A4", 0.5),
                        ("A5", 0.4),
                        ("A6", 0.8)
                    )
                )
            )
            .IsEqualTo(0.0);
    }
}
