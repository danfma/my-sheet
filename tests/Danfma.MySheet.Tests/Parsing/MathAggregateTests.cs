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
    public async Task Subtotal_IgnoresNestedSubtotals_ThroughAUnionAndAnOpenRange()
    {
        // A exclusão de SUBTOTALs aninhados tem que valer nas CINCO formas de referência do scan, não só
        // no range fechado: A3 é um SUBTOTAL (= 3), então tanto a união quanto a coluna inteira somam
        // apenas A1 + A2 = 3. (Sem a exclusão o resultado seria 6.)
        (string, object)[] cells = [("A1", 1), ("A2", 2), ("A3", "=SUBTOTAL(9,A1:A2)")];

        await Assert.That(Num(Calc("=SUBTOTAL(9,(A1:A2,A3:A3))", cells))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=SUBTOTAL(9,A:A)", cells))).IsEqualTo(3.0);
    }

    // --- SUBTOTAL / AGGREGATE 1-13 sobre um argumento que é um ARRAY calculado (não uma referência) ---
    // A página oficial define ref1 como "the first named range or reference", e o oráculo do plano
    // (Aspose.Cells 26.6.0, medido em 2026-09-09) mostra que a definição é LITERAL: os `ref` do SUBTOTAL e da
    // forma-referência do AGGREGATE (function_num 1-13) têm que ser REFERÊNCIAS, e um array calculado nessa
    // posição é #VALUE! — inclusive a constante `=SUBTOTAL(9,{1,2,3})`, e inclusive com a fórmula entrada
    // como CSE. É exatamente por isso que a forma-array (14-19) existe: só ELA aceita um array, e continua
    // presa mais abaixo (`=AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/…,1)` = 1 no mesmo oráculo).
    //
    // REVERSÃO do 11f5eaf, que dobrava o array como o SUM dobra. Aquela premissa era uma INFERÊNCIA por
    // analogia com o SUM — nunca medida — e está errada. Medido no oráculo:
    //   =SUBTOTAL(9,ROW(A1:A3))      -> #VALUE!      (o 11f5eaf pinava 6)
    //   =SUBTOTAL(9,(A1:A3<>0)*1)    -> #VALUE!      (pinava 2)
    //   =SUBTOTAL(2,(A1:A3<>0)*1)    -> #VALUE!      (pinava 3)
    //   =SUBTOTAL(3,(A1:A3<>0)*1)    -> #VALUE!      (pinava 3)
    //   =AGGREGATE(9,4,ROW(A1:A3))   -> #VALUE!      (pinava 6)
    //   =AGGREGATE(2,4,(A1:A3<>0)*1) -> #VALUE!      (pinava 3)
    //   =AGGREGATE(9,6,(A1:A3<>0)*1) -> #VALUE!      (pinava 2)
    private static readonly (string, object)[] SubtotalArrayData =
    [
        ("A1", 5),
        ("A2", 0),
        ("A3", 9),
    ];

    [Test]
    public async Task Subtotal_Sum_RejectsARowNumberArray()
    {
        await Assert
            .That(Calc("=SUBTOTAL(9,ROW(A1:A3))", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);

        // ANTI-VACUIDADE: o #VALUE! é a REJEIÇÃO da posição `ref`, não um mini-CSE quebrado — o mesmo array,
        // no mesmo fixture, continua dobrando dentro de um consumidor que aceita array. (O Excel só dá 6 aqui
        // sob CSE; sem CSE ele intersecta implicitamente e dá 1. O MySheet avalia sempre como array — modelo
        // documentado em workbook-and-expressions.md, "Implicit array arguments" — e isso não muda aqui.)
        await Assert.That(Num(Calc("=SUM(ROW(A1:A3))", SubtotalArrayData))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Subtotal_Sum_RejectsAComputedBooleanArray()
    {
        await Assert
            .That(Calc("=SUBTOTAL(9,(A1:A3<>0)*1)", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num(Calc("=SUM((A1:A3<>0)*1)", SubtotalArrayData))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Subtotal_Count_RejectsAComputedArray()
    {
        // COUNT e COUNTA nunca propagam erro de CÉLULA (o acumulador só tallia), então o #VALUE! aqui prova
        // que a rejeição acontece ANTES da varredura, no roteamento do argumento — e não vem do acumulador.
        await Assert
            .That(Calc("=SUBTOTAL(2,(A1:A3<>0)*1)", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Subtotal_CountA_RejectsAComputedArray()
    {
        await Assert
            .That(Calc("=SUBTOTAL(3,(A1:A3<>0)*1)", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Aggregate_ReferenceForm_RejectsAComputedArray()
    {
        // O GÊMEO AGGREGATE dos quatro testes acima, no MESMO fixture: a forma-referência (1-13) passa pelo
        // AggregateCodes.Feed exatamente como o SUBTOTAL, então a rejeição vale para os dois. As três options
        // cobrem os dois estados do bit de erro (4 = "ignore nothing", 6 = "ignore error values"): nem uma
        // nem outra transforma o argumento-array em população.
        await Assert
            .That(Calc("=AGGREGATE(9,4,ROW(A1:A3))", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(2,4,(A1:A3<>0)*1)", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(9,6,(A1:A3<>0)*1)", SubtotalArrayData))
            .IsEqualTo(ErrorValue.NotValue);

        // A forma-ARRAY (14-19) sobre o MESMO array continua sendo o caminho que ACEITA — é o discriminador
        // que prova que a rejeição é da posição `ref`, não do array: medido 1 e 3 no oráculo.
        await Assert
            .That(
                Num(
                    Calc(
                        "=AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>\"\")*(A1:A3<>0)),1)",
                        SubtotalArrayData
                    )
                )
            )
            .IsEqualTo(1.0);
        await Assert
            .That(
                Num(
                    Calc(
                        "=AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>\"\")*(A1:A3<>0)),2)",
                        SubtotalArrayData
                    )
                )
            )
            .IsEqualTo(3.0);
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

    // --- AGGREGATE — golden: página oficial "AGGREGATE function"
    // (43b9278e-6aa7-4f17-92b6-e19993fa26df, fetched em 2026-09-09). A página fornece DUAS sintaxes —
    // "AGGREGATE(function_num, options, ref1, [ref2], …)" (forma-referência) e
    // "AGGREGATE(function_num, options, array, [k])" (forma-array) — a tabela function_num 1-19 e a
    // tabela options 0-7, todas transcritas nos testes abaixo. ---

    // Fixture da tabela function_num: seis números com um valor REPETIDO (5), porque MODE.SNGL precisa de
    // uma moda única, mais uma célula de texto para que COUNT (6) e COUNTA (7) não coincidam.
    private static readonly (string, object)[] AggregateData =
    [
        ("A1", 5),
        ("A2", 0),
        ("A3", 9),
        ("A4", 5),
        ("A5", 3),
        ("A6", 8),
        ("A7", "text"),
    ];

    // Compara um AGGREGATE com a função homônima sobre a MESMA população, com guarda anti-vacuidade: o
    // lado AGGREGATE tem que ser um número de verdade (um #VALUE! dos dois lados passaria calado).
    private static async Task AssertSameAsNamesake(string aggregate, string namesake)
    {
        var actual = Num(Calc(aggregate, AggregateData));

        await Assert.That(double.IsNaN(actual)).IsFalse();
        await Assert.That(actual).IsEqualTo(Num(Calc(namesake, AggregateData)));
    }

    [Test]
    public async Task Aggregate_MapsEveryDocumentedFunctionNum()
    {
        // A tabela function_num inteira, 1-19, contra a função homônima. options = 4 ("ignore nothing") é
        // exatamente a semântica da homônima, então a igualdade é a própria definição da tabela.
        await AssertSameAsNamesake("=AGGREGATE(1,4,A1:A7)", "=AVERAGE(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(2,4,A1:A7)", "=COUNT(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(3,4,A1:A7)", "=COUNTA(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(4,4,A1:A7)", "=MAX(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(5,4,A1:A7)", "=MIN(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(6,4,A1:A7)", "=PRODUCT(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(7,4,A1:A7)", "=STDEV.S(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(8,4,A1:A7)", "=STDEV.P(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(9,4,A1:A7)", "=SUM(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(10,4,A1:A7)", "=VAR.S(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(11,4,A1:A7)", "=VAR.P(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(12,4,A1:A7)", "=MEDIAN(A1:A7)");
        await AssertSameAsNamesake("=AGGREGATE(13,4,A1:A7)", "=MODE.SNGL(A1:A7)");

        // 14-19 são a forma-array: o quarto argumento (k / quart) é obrigatório e vai para a homônima.
        await AssertSameAsNamesake("=AGGREGATE(14,4,A1:A7,2)", "=LARGE(A1:A7,2)");
        await AssertSameAsNamesake("=AGGREGATE(15,4,A1:A7,2)", "=SMALL(A1:A7,2)");
        await AssertSameAsNamesake("=AGGREGATE(16,4,A1:A7,0.25)", "=PERCENTILE.INC(A1:A7,0.25)");
        await AssertSameAsNamesake("=AGGREGATE(17,4,A1:A7,1)", "=QUARTILE.INC(A1:A7,1)");
        await AssertSameAsNamesake("=AGGREGATE(18,4,A1:A7,0.25)", "=PERCENTILE.EXC(A1:A7,0.25)");
        await AssertSameAsNamesake("=AGGREGATE(19,4,A1:A7,1)", "=QUARTILE.EXC(A1:A7,1)");
    }

    // Fixture das options: A2 é um #DIV/0! entre dois números. Mentira deliberada — 14 (a soma dos
    // sobreviventes) não coincide com nenhum valor de célula nem com a contagem.
    private static readonly (string, object)[] AggregateErrorData =
    [
        ("A1", 5),
        ("A2", "=1/0"),
        ("A3", 9),
    ];

    [Test]
    public async Task Aggregate_IgnoreErrorsBit_SelectsTheSurvivingPopulation()
    {
        // Tabela options: 4 = "ignore nothing" (o erro propaga, como em SUM) e 6 = "ignore error values"
        // (a célula de erro sai da população). Este par é o bit 1 inteiro.
        await Assert
            .That(Calc("=AGGREGATE(9,4,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Num(Calc("=AGGREGATE(9,6,A1:A3)", AggregateErrorData))).IsEqualTo(14.0);

        // COUNTA é a ÚNICA diferença comportamental real do bit: sem ele a célula de erro conta, com ele
        // sai da contagem. A tabela de options diz "ignore error values" mas NÃO diz o que o COUNTA passa a
        // contar — os dois valores abaixo, porém, estão MEDIDOS no oráculo do plano (Aspose.Cells 26.6.0,
        // em 2026-09-09): =AGGREGATE(3,4,E1:E3) = 3 e =AGGREGATE(3,6,E1:E3) = 2. Não é mais uma dedução.
        await Assert.That(Num(Calc("=AGGREGATE(3,4,A1:A3)", AggregateErrorData))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=AGGREGATE(3,6,A1:A3)", AggregateErrorData))).IsEqualTo(2.0);

        // COUNT não precisa de regra nenhuma: só conta números, e um erro nunca foi um.
        await Assert.That(Num(Calc("=AGGREGATE(2,4,A1:A3)", AggregateErrorData))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=AGGREGATE(2,6,A1:A3)", AggregateErrorData))).IsEqualTo(2.0);

        // "0 or omitted": um argumento options omitido chega como BlankValue e coage para 0.
        await Assert
            .That(Calc("=AGGREGATE(9,,A1:A3)", AggregateErrorData))
            .IsEqualTo(Calc("=AGGREGATE(9,0,A1:A3)", AggregateErrorData));
        await Assert
            .That(Calc("=AGGREGATE(9,,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Aggregate_WholeArgumentError_PropagatesThroughTheIgnoreErrorsBit()
    {
        // O bit 1 ("ignore error values") vale para as CÉLULAS alcançadas através de uma referência e para
        // os ELEMENTOS de um array — NÃO para um argumento inteiro que É um erro. Medido no oráculo do plano
        // (Aspose.Cells 26.6.0, em 2026-09-09), com o MySheet de antes entre parênteses:
        //   =AGGREGATE(9,6,1/0)       -> #DIV/0!   (devolvia 0)
        //   =AGGREGATE(2,6,1/0)       -> #DIV/0!   (devolvia 0)
        //   =AGGREGATE(9,6,A1:A3,1/0) -> #DIV/0!   (devolvia 14 — o ref válido somava e o erro sumia)
        await Assert
            .That(Calc("=AGGREGATE(9,6,1/0)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=AGGREGATE(2,6,1/0)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(Calc("=AGGREGATE(9,6,A1:A3,1/0)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);

        // ANTI-VACUIDADE do caso de dois argumentos: sem o erro ao lado, o MESMO primeiro ref responde 14.
        await Assert.That(Num(Calc("=AGGREGATE(9,6,A1:A3)", AggregateErrorData))).IsEqualTo(14.0);

        // CONTROLE: o SUBTOTAL já propagava (ele nunca ignora erros), e o oráculo concorda —
        // =SUBTOTAL(9,A1:A3,1/0) -> #DIV/0!. É o mesmo caminho de código com o bit desligado.
        await Assert
            .That(Calc("=SUBTOTAL(9,A1:A3,1/0)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);

        // CONTRASTE que prova que o bit continua valendo onde deve: a mesma divisão por zero, agora dentro
        // de uma CÉLULA alcançada por referência (A2), continua sendo ignorada. 0 nos dois lados, medido.
        await Assert.That(Num(Calc("=AGGREGATE(9,6,A2)", AggregateErrorData))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=AGGREGATE(3,6,A2)", AggregateErrorData))).IsEqualTo(0.0);

        // A forma-array sobre o MESMO escalar-erro: o erro passa a propagar em vez de ser engolido (antes a
        // população ficava vazia e o SMALL respondia #NUM!).
        //
        // DIVERGÊNCIA CONHECIDA, registrada como decisão em aberto no plano mestre: o oráculo devolve
        // #VALUE! aqui — mas devolve 7 para =AGGREGATE(15,6,7,1), ou seja, ele ACEITA um escalar comum na
        // posição `array` e só rejeita um escalar de ERRO. Essa assimetria é a mesma de =SUBTOTAL(9,7) = 7
        // no MySheet contra #VALUE! no oráculo (um literal numa posição de referência), e não é resolvida
        // aqui. Propagar o erro de verdade é a resposta principiada e é o que fica pinado.
        await Assert
            .That(Calc("=AGGREGATE(15,6,1/0,1)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Aggregate_HiddenRowBit_IsANoOp()
    {
        // §S6 do plano — CAVEAT DE MODELO: o MySheet não tem linhas ocultas, então o bit 0 ("ignore hidden
        // rows") não tem nada para ignorar e 1/3/5/7 são idênticos a 0/2/4/6. As quatro igualdades abaixo
        // são o registro executável desse limite; num engine com filtro elas poderiam divergir.
        foreach (var (hidden, plain) in new[] { (1, 0), (3, 2), (5, 4), (7, 6) })
        {
            await Assert
                .That(Calc($"=AGGREGATE(9,{hidden},A1:A3)", AggregateErrorData))
                .IsEqualTo(Calc($"=AGGREGATE(9,{plain},A1:A3)", AggregateErrorData));
        }

        // Anti-vacuidade: os dois LADOS do par têm que ser respostas de verdade, e os pares entre si têm
        // que DIFERIR — senão a igualdade acima passaria com tudo devolvendo o mesmo erro.
        await Assert
            .That(Calc("=AGGREGATE(9,1,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Num(Calc("=AGGREGATE(9,3,A1:A3)", AggregateErrorData))).IsEqualTo(14.0);
        await Assert
            .That(Calc("=AGGREGATE(9,5,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Num(Calc("=AGGREGATE(9,7,A1:A3)", AggregateErrorData))).IsEqualTo(14.0);
    }

    // Fixture do skip aninhado: B1 é um SUBTOTAL (3) e B2 um AGGREGATE (3), B3 = 5 é o único valor
    // "normal". 5 (tudo pulado) e 11 (nada pulado) não coincidem com nenhum valor de célula.
    private static readonly (string, object)[] AggregateNestedData =
    [
        ("A1", 1),
        ("A2", 2),
        ("B1", "=SUBTOTAL(9,A1:A2)"),
        ("B2", "=AGGREGATE(9,0,A1:A2)"),
        ("B3", 5),
    ];

    [Test]
    public async Task Aggregate_NestedBit_SkipsBothSubtotalAndAggregateCells()
    {
        // Tabela options: 0-3 dizem "Ignore nested SUBTOTAL and AGGREGATE functions"; 4-7 não dizem
        // ("Ignore nothing" / hidden rows / error values). O bit 2 é, portanto, o INVERSO do skip.
        foreach (var options in new[] { 0, 1, 2, 3 })
        {
            await Assert
                .That(Num(Calc($"=AGGREGATE(9,{options},B1:B3)", AggregateNestedData)))
                .IsEqualTo(5.0);
        }

        foreach (var options in new[] { 4, 5, 6, 7 })
        {
            await Assert
                .That(Num(Calc($"=AGGREGATE(9,{options},B1:B3)", AggregateNestedData)))
                .IsEqualTo(11.0);
        }

        // COUNTA enxerga o mesmo skip: 1 célula sobrevivente contra 3.
        await Assert.That(Num(Calc("=AGGREGATE(3,0,B1:B3)", AggregateNestedData))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=AGGREGATE(3,4,B1:B3)", AggregateNestedData))).IsEqualTo(3.0);

        // Discriminador dos dois arms do enum: a regra do SUBTOTAL é ESTREITA de propósito (a página do
        // SUBTOTAL só documenta "nested subtotals are ignored"; a redação "SUBTOTAL and AGGREGATE" só
        // aparece na tabela de options do AGGREGATE), então B2 — um AGGREGATE — CONTA aqui: 3 + 5 = 8.
        // Se as duas regras fossem a mesma, o teste acima passaria com o predicado errado.
        // MEDIDO (não mais inferido) no oráculo do plano, Aspose.Cells 26.6.0 em 2026-09-09: com A1=1, A2=2
        // e A3="=AGGREGATE(9,0,A1:A2)"=3, o =SUBTOTAL(9,A1:A3) responde 6 — o AGGREGATE aninhado conta.
        //
        // DIVERGÊNCIA REGISTRADA no bloco acima: o MESMO oráculo também CONTA um AGGREGATE aninhado sob as
        // options 0-3 (daria 8, não 5), contrariando a própria tabela de options da Microsoft ("Ignore
        // nested SUBTOTAL and AGGREGATE functions"). Aqui a página documentada vence o oráculo — o código
        // e os 5.0 acima ficam como estão —, e a divergência está anotada no plano da fase.
        await Assert.That(Num(Calc("=SUBTOTAL(9,B1:B3)", AggregateNestedData))).IsEqualTo(8.0);
    }

    [Test]
    public async Task Aggregate_InvalidFunctionNumOrOptions_IsValueError()
    {
        // function_num fora de 1-19 e options fora de 0-7 → #VALUE!.
        await Assert
            .That(Calc("=AGGREGATE(20,0,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(0,0,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(9,8,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(9,-1,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Aggregate_ArrayForm_RequiresItsFourthArgument()
    {
        // "If a second ref argument is necessary but not provided, AGGREGATE returns a #VALUE! error":
        // 14-19 precisam do k, e três argumentos não o fornecem.
        await Assert
            .That(Calc("=AGGREGATE(15,6,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=AGGREGATE(19,6,A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.NotValue);

        // Menos de 3 argumentos é erro de ARIDADE, rejeitado já no parse (como o Excel rejeita na entrada).
        await Assert
            .That(() => Calc("=AGGREGATE(15,6)", AggregateErrorData))
            .Throws<ParseException>();
    }

    [Test]
    public async Task Aggregate_KOutsideThePostSkipPopulation_IsNumError()
    {
        // Option 6 tira o #DIV/0! da população, que fica com DOIS números — então k = 3 já passou do fim,
        // mesmo havendo três células. É a fronteira que prova que k é medido DEPOIS do skip. MEDIDO no
        // oráculo (Aspose.Cells 26.6.0, 2026-09-09): =AGGREGATE(15,6,E1:E3,2) = 9 e ...,3) = #NUM!.
        await Assert.That(Num(Calc("=AGGREGATE(15,6,A1:A3,2)", AggregateErrorData))).IsEqualTo(9.0);
        await Assert
            .That(Calc("=AGGREGATE(15,6,A1:A3,3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.Number);

        // k = 4 (o valor do brief) passa tanto do fim da população pós-skip quanto do número de CÉLULAS,
        // então fixa a mesma regra por um caminho que nem precisa do skip para falhar.
        await Assert
            .That(Calc("=AGGREGATE(15,6,A1:A3,4)", AggregateErrorData))
            .IsEqualTo(ErrorValue.Number);

        // População inteiramente de erros sob a option 6 → população vazia → #NUM! (o mesmo que o SMALL já
        // responde para um range vazio). Também MEDIDO no oráculo, e não adotado por analogia: sobre três
        // células de erro, =AGGREGATE(15,6,A1:A3,1), =AGGREGATE(14,6,A1:A3,1) e =AGGREGATE(16,6,A1:A3,0.5)
        // dão todos #NUM! (enquanto =AGGREGATE(9,6,A1:A3) dá 0 e =AGGREGATE(3,6,A1:A3) dá 0).
        (string, object)[] allErrors = [("A1", "=1/0"), ("A2", "=1/0"), ("A3", "=1/0")];

        await Assert.That(Calc("=AGGREGATE(15,6,A1:A3,1)", allErrors)).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Aggregate_ArrayForm_OverAPlainRange_HonoursTheNestedSkip()
    {
        // A tabela de options é declarada uma vez para a função INTEIRA, então o skip aninhado das options
        // 0-3 também vale na forma-array — e sobre um range simples ela passa pelo mesmo Gather, que é quem
        // carrega o skip. Fixture DISCRIMINANTE (o valor aninhado é o MAIOR do range, senão LARGE daria o
        // mesmo dos dois lados): A1=1, A2=2, A3=SUBTOTAL(9,A1:A2)=3, logo LARGE(…,1) é 2 com o skip e 3 sem.
        // MEDIDO no oráculo (Aspose.Cells 26.6.0, 2026-09-09): 2 e 3.
        (string, object)[] cells = [("A1", 1), ("A2", 2), ("A3", "=SUBTOTAL(9,A1:A2)")];

        await Assert.That(Num(Calc("=AGGREGATE(14,0,A1:A3,1)", cells))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=AGGREGATE(14,4,A1:A3,1)", cells))).IsEqualTo(3.0);
    }

    [Test]
    public async Task Aggregate_ArrayForm_StreamsAComputedArrayAndDropsItsErrorElements()
    {
        // A forma-array sobre um array COMPUTADO cujos ELEMENTOS incluem um erro, sob a option 6 — o
        // caminho 16-19, que coleta o stream inteiro pelo acumulador em vez do heap do 14/15.
        // 1/A1:A3 = {1/5; #DIV/0!; 1/9}: o erro sai da população e o PERCENTILE.INC em 0.5 dos dois
        // sobreviventes é a média deles. MEDIDO no oráculo: 0.15555555555555556.
        await Assert
            .That(Num(Calc("=AGGREGATE(16,6,1/A1:A3,0.5)", SubtotalArrayData)))
            .IsEqualTo(((1.0 / 9.0) + 0.2) / 2);

        // O mesmo código sobre o RANGE simples: SMALL(…,1) de {5,0,9} é 0 — o literal do relatório, igual
        // no oráculo. Um 0 é fácil de sair por acidente (população vazia responderia #NUM!, mas um SUM
        // vazio responderia 0), então o par com k = 2 é o que prova que a população foi mesmo varrida.
        await Assert.That(Num(Calc("=AGGREGATE(15,6,A1:A3,1)", SubtotalArrayData))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=AGGREGATE(15,6,A1:A3,2)", SubtotalArrayData))).IsEqualTo(5.0);
    }

    [Test]
    public async Task SubtotalAndAggregate_KeepEveryReferenceProducingShapeOffTheArrayPath()
    {
        // PIN DE REGRESSÃO da rejeição de array-computado (AggregateCodes.Feed): as formas que PRODUZEM uma
        // referência sem serem um RangeReference escrito à mão têm que continuar caindo no Gather, e não no
        // gate de array — senão a rejeição comeria argumentos legítimos. As três que existem hoje: um NOME
        // definido, um DynamicRange (o ':' com extremidade que retorna referência) e uma FUNÇÃO que devolve
        // referência (OFFSET). Nenhuma delas muda de resposta entre a base e o head; as nove respostas estão
        // MEDIDAS no oráculo (Aspose.Cells 26.6.0, 2026-09-09) e são as mesmas: 14 / 14 / 0.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);
        workbook.DefineName("Rng", "Sheet1!$A$1:$A$3");

        object? Eval(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        foreach (var shape in new[] { "Rng", "INDEX(A1:A3,1,1):A3", "OFFSET(A1,0,0,3,1)" })
        {
            // Forma-referência, com e sem o bit de skip aninhado: a soma inteira do range, 14.
            await Assert.That(Num(Eval($"=SUBTOTAL(9,{shape})"))).IsEqualTo(14.0);
            await Assert.That(Num(Eval($"=AGGREGATE(9,4,{shape})"))).IsEqualTo(14.0);

            // Forma-array sobre a MESMA referência: SMALL de {5,0,9} em k=1 é 0. Se o argumento tivesse
            // sido desviado para o caminho de array — ou para o escalar —, a população não seria essa e a
            // resposta viria #VALUE! ou #NUM!, não 0.
            await Assert.That(Num(Eval($"=AGGREGATE(15,6,{shape},1)"))).IsEqualTo(0.0);
        }
    }

    [Test]
    public async Task Subtotal_WholeArgumentError_PropagatesEvenThroughTheCountingCodes()
    {
        // PIN DE REGRESSÃO de um efeito COLATERAL da correção do argumento-erro inteiro: os códigos 2
        // (COUNT) e 3 (COUNTA) nunca propagam erro de CÉLULA — o acumulador só tallia —, então antes da
        // correção um argumento que É um erro sumia neles: =SUBTOTAL(2,1/0) dava 0 e =SUBTOTAL(3,1/0) dava
        // 1 (o erro contava como "não vazio"). Agora o erro volta antes de chegar ao acumulador, nos três
        // códigos. MEDIDO no oráculo: #DIV/0! nos três. Já passa — não há RED a pinar, o comportamento veio
        // junto com a correção; o pino existe para que ele não volte a sumir.
        await Assert.That(Calc("=SUBTOTAL(2,1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SUBTOTAL(3,1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SUBTOTAL(9,1/0)")).IsEqualTo(ErrorValue.DivByZero);

        // O terceiro colateral, no AGGREGATE: uma FUNÇÃO de referência que falha (OFFSET para fora da
        // planilha) devolve um #REF! que a option 6 engolia — dava 0. Agora propaga, como o oráculo.
        await Assert
            .That(Calc("=AGGREGATE(9,6,OFFSET(A1,-5,0))", AggregateErrorData))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=AGGREGATE(9,4,OFFSET(A1,-5,0))", AggregateErrorData))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Aggregate_MissingSheetReference_IsRefError()
    {
        // Guarda estrutural ANTES do scan: sem ela a option 6 (que engole erros) veria um range vazio e
        // responderia 0 em vez de #REF!.
        await Assert
            .That(Calc("=AGGREGATE(9,6,Ghost!A1:A3)", AggregateErrorData))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=AGGREGATE(15,6,Ghost!A1:A3,1)", AggregateErrorData))
            .IsEqualTo(ErrorValue.Reference);
    }

    // Fixture das duas formas: mesma FORMA de chamada (quatro argumentos), significados diferentes.
    private static readonly (string, object)[] AggregateFormsData =
    [
        ("A1", 5),
        ("A2", 0),
        ("A3", 9),
        ("B1", 1),
        ("B2", 2),
        ("B3", 3),
    ];

    [Test]
    public async Task Aggregate_FormIsChosenByFunctionNumAlone()
    {
        // function_num 9 ≤ 13 → forma-referência: B1:B3 é um SEGUNDO ref, não um k. 14 + 6 = 20 (e
        // SUM(A1:A3) sozinho é 14, então o segundo ref realmente entra).
        await Assert
            .That(Num(Calc("=AGGREGATE(9,6,A1:A3,B1:B3)", AggregateFormsData)))
            .IsEqualTo(20.0);
        await Assert.That(Num(Calc("=SUM(A1:A3)", AggregateFormsData))).IsEqualTo(14.0);

        // function_num 15 ≥ 14 → forma-array: o 2 é o k do SMALL sobre {0,5,9}, logo 5. Se o 2 fosse lido
        // como um segundo ref (a forma-referência), a resposta seria 14 + 2 = 16.
        await Assert.That(Num(Calc("=AGGREGATE(15,6,A1:A3,2)", AggregateFormsData))).IsEqualTo(5.0);
    }
}
