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

    // --- SUMPRODUCT sobre um array COMPUTADO — mesma página golden (16753e75). As duas regras
    // documentadas em jogo: "SUMPRODUCT treats non-numeric array entries as if they were zeros" e
    // "The array arguments must have the same dimensions. If they do not, SUMPRODUCT returns the
    // #VALUE! error value." A2 = 0 é a mentira deliberada do fixture — é a ÚNICA célula cujo sinal
    // (<>0) é FALSE, então um resultado 3 provaria que os sinais nunca chegaram à soma. ---

    private static readonly (string, object)[] FlagData =
    [
        ("A1", 5),
        ("A2", 0),
        ("A3", 9),
        ("B1", 2),
        ("B2", 4),
        ("B3", 6),
    ];

    [Test]
    public async Task SumProduct_ConsumesAComputedArrayArgument()
    {
        // (A1:A3<>0)*1 is [1,0,1] — a computed array, not a range. SUMPRODUCT reads its arguments
        // through PositionalRange, which gained an array backing for exactly this: before it, a
        // computed-array argument was evaluated ONCE as a scalar (#VALUE! for the operation forms,
        // and the leftmost row number for ROW(range) — a wrong NUMBER, not an error).
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:A3<>0)*1)", FlagData))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:A3<>0)*1,B1:B3)", FlagData))).IsEqualTo(8.0);
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:A3>0)*(B1:B3>3))", FlagData))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=SUMPRODUCT(ROW(A1:A3))", FlagData))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=SUMPRODUCT(IF(A1:A3>0,1,0))", FlagData))).IsEqualTo(2.0);
    }

    [Test]
    public async Task CriteriaFamily_StillRefusesAComputedArray()
    {
        // The negative control for the arm above: SUMPRODUCT is the one member of the positional-scan family
        // that opted IN to computed arrays (PositionalRange.OpenArrayOrRange), and its siblings keep the
        // plain PositionalRange.Open, where a non-reference argument falls to ArgumentFlattening's `default`
        // arm and is EVALUATED as a scalar — (A1:A3)*1 is then a one-element sequence holding the #VALUE! of
        // a range in an arithmetic operation. "Refuses" therefore has two SHAPES, and both are pinned
        // because only one of them is an error.
        //
        // PAIRED forms — a real criteria range beside the collapsed argument — see 1 element against 3 and
        // raise the scan's up-front length mismatch. Excel says #VALUE! here too, so this half is parity.
        // All four are listed because the docs name them as a group: a family member drifting off this
        // shared mismatch check would otherwise leave the docs' "four shapes" sentence quietly wrong.
        await Assert
            .That(Calc("=SUMIFS((A1:A3)*1,A1:A3,\">0\")", FlagData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AVERAGEIFS((A1:A3)*1,A1:A3,\">0\")", FlagData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=MAXIFS((A1:A3)*1,A1:A3,\">0\")", FlagData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=MINIFS((A1:A3)*1,A1:A3,\">0\")", FlagData))
            .IsEqualTo(ErrorValue.NotValue);

        // SINGLE-criteria forms have nothing to mismatch against: the lone #VALUE! element matches no
        // criterion, so SUMIF/COUNTIF report an empty scan (0) and AVERAGEIF divides by a zero count. These
        // are SILENT answers, not errors, which is exactly why they are pinned — a reader who assumes the
        // whole family errors would be wrong, and any move of these consumers onto OpenArrayOrRange has to
        // change this line deliberately.
        await Assert.That(Calc("=SUMIF((A1:A3)*1,\">0\")", FlagData)).IsEqualTo(0.0);
        await Assert.That(Calc("=COUNTIF((A1:A3)*1,\">0\")", FlagData)).IsEqualTo(0.0);
        // COUNTIFS belongs to the SILENT half despite its plural name: its first argument IS the criteria
        // range, so a single pair has no second length to disagree with. The docs state the four shapes as
        // pinned, and this is the line that makes COUNTIFS part of that claim.
        await Assert.That(Calc("=COUNTIFS((A1:A3)*1,\">0\")", FlagData)).IsEqualTo(0.0);
        await Assert
            .That(Calc("=AVERAGEIF((A1:A3)*1,\">0\")", FlagData))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task SumProduct_DoesNotCoerceTheLogicalsOfAComputedArray()
    {
        // The "non-numeric entries count as zero" rule covers the TRUE/FALSE of a bare comparison too:
        // (A1:A3<>0) is [TRUE,FALSE,TRUE] and sums to 0 — which is precisely why the Excel idiom
        // multiplies the flags by 1.
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:A3<>0))", FlagData))).IsEqualTo(0.0);
    }

    [Test]
    public async Task SumProduct_OrientationMismatch_IsValueError()
    {
        // "The array arguments must have the same DIMENSIONS" — not the same cell count: a 3x1 column
        // and a 1x3 row both hold 3 cells and are still #VALUE! in Excel. A count-only check accepted
        // this pair and answered 25 (5*5 + 0*0 + 9*0, pairing column A against row 1).
        await Assert
            .That(Calc("=SUMPRODUCT(A1:A3,A1:C1)", FlagData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=SUMPRODUCT((A1:A3<>0)*1,A1:D1)", FlagData))
            .IsEqualTo(ErrorValue.NotValue);

        // The same rule over a COMPUTED array whose count matches: (A1:A3<>0)*1 is 3x1 and A1:C1 is 1x3,
        // three cells each, so only the SHAPE can reject it — the count check never fires here (the A1:D1
        // case above is caught by the count, 3 vs 4, before the shape is even consulted).
        await Assert
            .That(Calc("=SUMPRODUCT((A1:A3<>0)*1,A1:C1)", FlagData))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task SumProduct_OrientationMismatch_IsValueError_BehindAShapelessArgument()
    {
        // A SHAPELESS argument must never become the pivot the other arguments are judged against. MyName
        // is a defined name — it takes the materialized fallback, which knows no rectangle (0/0) — so a
        // check that always compared against argument 1 would short-circuit on it and never compare
        // A1:A3 (3x1) with A1:C1 (1x3) to EACH OTHER, answering 125 where Excel answers #VALUE!. All
        // three arguments hold 3 cells, so the count check cannot see the mismatch either.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);
        workbook.DefineName("MyName", "Sheet1!$A$1:$A$3");

        await Assert
            .That(
                ExpressionParser
                    .Parse("=SUMPRODUCT(MyName,A1:A3,A1:C1)", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(ErrorValue.NotValue);

        // …and the tolerance itself is preserved: a shapeless argument is judged by its COUNT alone, so
        // pairing the same name with a 3x1 computed array stays legal — 5*1 + 0*0 + 9*1 = 14.
        await Assert
            .That(
                ExpressionParser
                    .Parse("=SUMPRODUCT(MyName,(A1:A3<>0)*1)", sheet)
                    .Evaluate(workbook)
                    .AsObject()
            )
            .IsEqualTo(14.0);
    }

    [Test]
    public async Task SumProduct_OrientationMismatch_IsValueError_EvenFromTheRangeSnapshot()
    {
        // The dimension rule must not depend on WHICH backing answered. The shared per-epoch
        // RangeSnapshot is admitted only on a range's SECOND read of the epoch and only above
        // RangeCacheMinimumCells (256) populated cells, and it hands the values over as a flat list —
        // so a snapshot-served range has to carry its rectangle's shape along, or two cells holding
        // the IDENTICAL formula would disagree (measured with the shape dropped: ZZ1 #VALUE!, ZZ2
        // 9045054). A1:A300 is 300x1 and A5:KN5 is 1x300 — 300 cells each, above the threshold.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        for (var row = 1; row <= 300; row++)
        {
            sheet["A" + row] = new NumberValue(row);
        }

        for (var column = 2; column <= 300; column++) // A5 is already populated by the column above
        {
            sheet[new CellAddress(column, 5).ToId()] = new NumberValue(column);
        }

        var rowEnd = new CellAddress(300, 5).ToId(); // KN5 — the 300th column
        var formula = $"=SUMPRODUCT(A1:A300,A5:{rowEnd})";
        sheet["ZZ1"] = ExpressionParser.Parse(formula, sheet);
        sheet["ZZ2"] = ExpressionParser.Parse(formula, sheet);

        // ZZ1 reads both ranges cold (streaming cursor); ZZ2 is the second read, served by the snapshot.
        await Assert
            .That(workbook.GetCellValue("Sheet1", "ZZ1").AsObject())
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(workbook.GetCellValue("Sheet1", "ZZ2").AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task SumProduct_PropagatesAnErrorElementOfAComputedArray()
    {
        // An error INSIDE the computed array propagates as the function result (Excel agrees); it is
        // never absorbed by the "non-numeric counts as zero" rule. E2/E3 are blank, so only E1 errs.
        (string, object)[] cells = [("A1", 5), ("A2", 0), ("A3", 9), ("E1", "=1/0")];

        await Assert
            .That(Calc("=SUMPRODUCT((E1:E3)*1,A1:A3)", cells))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task SumProduct_PairsRangesAndArraysInTheSamePositionOrder()
    {
        // The transpose quartet: a range cursor walks COLUMN-major while the element-wise array stream
        // is ROW-major, so one side must be transposed on read. All four forms are the same sum of
        // squares, 1*1 + 2*2 + 3*3 + 4*4 = 30; without the transpose the two MIXED forms silently
        // return 29, because they pair A2 with B1.
        (string, object)[] grid = [("A1", 1), ("A2", 2), ("B1", 3), ("B2", 4)];

        await Assert.That(Num(Calc("=SUMPRODUCT(A1:B2,A1:B2)", grid))).IsEqualTo(30.0);
        await Assert.That(Num(Calc("=SUMPRODUCT(A1:B2,(A1:B2)*1)", grid))).IsEqualTo(30.0);
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:B2)*1,A1:B2)", grid))).IsEqualTo(30.0);
        await Assert.That(Num(Calc("=SUMPRODUCT((A1:B2)*1,(A1:B2)*1)", grid))).IsEqualTo(30.0);
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
