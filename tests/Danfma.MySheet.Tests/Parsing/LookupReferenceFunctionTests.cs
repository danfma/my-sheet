using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Onda 3 — Lookup & Reference escalar: CHOOSE, HLOOKUP, LOOKUP (vetor + array), COLUMN, COLUMNS,
// XMATCH, ADDRESS, AREAS e FORMULATEXT. Golden values das páginas oficiais da Microsoft
// (support.microsoft.com, fetched em 2026-07-02), citadas por função em cada teste.
public class LookupReferenceFunctionTests
{
    // Cells aceita string/double/int literais; um string começando com '=' é parseado como fórmula.
    private static object? Calc(string formula, params (string Id, object Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        Fill(sheet, cells);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static void Fill(Sheet sheet, (string Id, object Value)[] cells)
    {
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
    }

    // --- CHOOSE — golden: página oficial "CHOOSE function" (fc5c184f-cb62-4ec7-a46e-38653b98f5bc) ---

    [Test]
    public async Task Choose_ReturnsTheNthValue()
    {
        // =CHOOSE(2,A2,A3,A4,A5) -> "2nd"; =CHOOSE(4,B2,B3,B4,B5) -> "Bolts" (exemplos da página).
        (string, object)[] cells =
        [
            ("A2", "1st"), ("A3", "2nd"), ("A4", "3rd"), ("A5", "Finished"),
            ("B2", "Nails"), ("B3", "Screws"), ("B4", "Nuts"), ("B5", "Bolts"),
        ];

        await Assert.That(Calc("=CHOOSE(2,A2,A3,A4,A5)", cells)).IsEqualTo("2nd");
        await Assert.That(Calc("=CHOOSE(4,B2,B3,B4,B5)", cells)).IsEqualTo("Bolts");

        // =CHOOSE(3,"Wide",115,"world",8) -> "world" (exemplo da página).
        await Assert.That(Calc("=CHOOSE(3,\"Wide\",115,\"world\",8)")).IsEqualTo("world");
    }

    [Test]
    public async Task Choose_TruncatesAFractionalIndex()
    {
        // Regra documentada: "If index_num is a fraction, it is truncated to the lowest integer
        // before being used."
        await Assert.That(Calc("=CHOOSE(2.9,\"a\",\"b\",\"c\")")).IsEqualTo("b");
    }

    [Test]
    public async Task Choose_OutOfRange_IsValueError()
    {
        // Regra documentada: index_num < 1 ou > número de valores -> #VALUE!.
        await Assert.That(Calc("=CHOOSE(0,\"a\",\"b\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=CHOOSE(4,\"a\",\"b\",\"c\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Choose_OnlyEvaluatesTheChosenArgument()
    {
        // Lazy como IF: o argumento NÃO escolhido não pode ser avaliado (a função custom lançaria).
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        workbook.RegisterFunction(
            "BOOM",
            (_, _) => throw new InvalidOperationException("The unchosen argument was evaluated.")
        );

        var result = ExpressionParser.Parse("=CHOOSE(1,10,BOOM())", sheet).Evaluate(workbook);

        await Assert.That(result.AsObject() as double?).IsEqualTo(10.0);
    }

    [Test]
    public async Task Choose_RangeArgument_FlowsIntoRangeAwareFunctions()
    {
        // A página documenta value como range ("The value arguments to CHOOSE can be range
        // references as well as single values"); o range escolhido expande no consumidor.
        await Assert
            .That(
                Calc(
                    "=SUM(CHOOSE(2,A1:A2,B1:B2))",
                    ("A1", 1), ("A2", 2), ("B1", 10), ("B2", 20)
                ) as double?
            )
            .IsEqualTo(30.0);
    }

    // --- HLOOKUP — golden: página oficial "HLOOKUP function" (a3034eec-b719-4ba3-bb65-e1ad662ed95f):
    // A1:C1 = Axles/Bearings/Bolts; A2:C2 = 4/4/9; A3:C3 = 5/7/10; A4:C4 = 6/8/11. ---

    private static readonly (string, object)[] HLookupTable =
    [
        ("A1", "Axles"), ("B1", "Bearings"), ("C1", "Bolts"),
        ("A2", 4), ("B2", 4), ("C2", 9),
        ("A3", 5), ("B3", 7), ("C3", 10),
        ("A4", 6), ("B4", 8), ("C4", 11),
    ];

    [Test]
    public async Task HLookup_MatchesTheGoldenExamples()
    {
        // =HLOOKUP("Axles",A1:C4,2,TRUE) -> 4; =HLOOKUP("Bearings",A1:C4,3,FALSE) -> 7;
        // =HLOOKUP("Bolts",A1:C4,4) -> 11 (exemplos da página).
        await Assert
            .That(Calc("=HLOOKUP(\"Axles\",A1:C4,2,TRUE)", HLookupTable) as double?)
            .IsEqualTo(4.0);
        await Assert
            .That(Calc("=HLOOKUP(\"Bearings\",A1:C4,3,FALSE)", HLookupTable) as double?)
            .IsEqualTo(7.0);
        await Assert
            .That(Calc("=HLOOKUP(\"Bolts\",A1:C4,4)", HLookupTable) as double?)
            .IsEqualTo(11.0);
    }

    [Test]
    public async Task HLookup_ApproximateTextKey()
    {
        // =HLOOKUP("B",A1:C4,3,TRUE) -> 5: "an exact match for 'B' is not found, the largest value
        // in row 1 that is less than 'B' is used: 'Axles'" (exemplo da página). Regressão do bug
        // 3460bb3: chave TEXTO com match aproximado exige ordenação cross-type.
        await Assert
            .That(Calc("=HLOOKUP(\"B\",A1:C4,3,TRUE)", HLookupTable) as double?)
            .IsEqualTo(5.0);

        // A própria chave do canto (primeira coluna) em modo aproximado.
        await Assert
            .That(Calc("=HLOOKUP(\"Axles\",A1:C4,4,TRUE)", HLookupTable) as double?)
            .IsEqualTo(6.0);
    }

    [Test]
    public async Task HLookup_ErrorContract()
    {
        // Regras documentadas: row_index_num < 1 -> #VALUE!; > linhas da tabela -> #REF!;
        // lookup abaixo da menor chave (aproximado) -> #N/A.
        await Assert
            .That(Calc("=HLOOKUP(\"Axles\",A1:C4,0)", HLookupTable))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=HLOOKUP(\"Axles\",A1:C4,5)", HLookupTable))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=HLOOKUP(\"A\",A1:C4,2,TRUE)", HLookupTable))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    // --- LOOKUP — golden: página oficial "LOOKUP function" (446d94af-663b-451d-8251-369d5e3864cb):
    // A2:A6 = 4.14/4.19/5.17/5.77/6.39; B2:B6 = red/orange/yellow/green/blue. ---

    private static readonly (string, object)[] LookupVectors =
    [
        ("A2", 4.14), ("A3", 4.19), ("A4", 5.17), ("A5", 5.77), ("A6", 6.39),
        ("B2", "red"), ("B3", "orange"), ("B4", "yellow"), ("B5", "green"), ("B6", "blue"),
    ];

    [Test]
    public async Task Lookup_VectorForm_MatchesTheGoldenExamples()
    {
        // =LOOKUP(4.19,A2:A6,B2:B6) -> "orange"; =LOOKUP(5.75,...) -> "yellow";
        // =LOOKUP(7.66,...) -> "blue" (exemplos da página).
        await Assert.That(Calc("=LOOKUP(4.19,A2:A6,B2:B6)", LookupVectors)).IsEqualTo("orange");
        await Assert.That(Calc("=LOOKUP(5.75,A2:A6,B2:B6)", LookupVectors)).IsEqualTo("yellow");
        await Assert.That(Calc("=LOOKUP(7.66,A2:A6,B2:B6)", LookupVectors)).IsEqualTo("blue");
    }

    [Test]
    public async Task Lookup_BelowTheSmallestValue_IsNotAvailable()
    {
        // =LOOKUP(0,A2:A6,B2:B6) -> #N/A: "If lookup_value is smaller than the smallest value in
        // lookup_vector, LOOKUP returns the #N/A error value" (exemplo + regra da página).
        await Assert.That(Calc("=LOOKUP(0,A2:A6,B2:B6)", LookupVectors))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Lookup_WithoutResultVector_ReturnsTheMatchedKey()
    {
        // Sem result_vector o resultado vem do próprio lookup_vector: o maior valor <= 5.75 é 5.17.
        await Assert.That(Calc("=LOOKUP(5.75,A2:A6)", LookupVectors) as double?).IsEqualTo(5.17);
    }

    [Test]
    public async Task Lookup_ArrayForm_WideArray_SearchesFirstRowReturnsLastRow()
    {
        // Regra documentada: "If array covers an area that is wider than it is tall (more columns
        // than rows), LOOKUP searches for the value of lookup_value in the first row" e "LOOKUP
        // always selects the last value in the row or column". 4 colunas x 2 linhas: chaves
        // a/b/c/d na linha 1, valores 1/2/3/4 na linha 2 -> LOOKUP("c") casa "c" e retorna 3.
        (string, object)[] wide =
        [
            ("A1", "a"), ("B1", "b"), ("C1", "c"), ("D1", "d"),
            ("A2", 1), ("B2", 2), ("C2", 3), ("D2", 4),
        ];

        await Assert.That(Calc("=LOOKUP(\"c\",A1:D2)", wide) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Lookup_ArrayForm_TallArray_SearchesFirstColumnReturnsLastColumn()
    {
        // Regra documentada: "If an array is square or is taller than it is wide (more rows than
        // columns), LOOKUP searches in the first column". 2 colunas x 3 linhas: chaves a/b/c,
        // valores 1/2/3 -> "bump" não é exato; o maior valor <= "bump" é "b" -> 2.
        (string, object)[] tall =
        [
            ("A1", "a"), ("B1", 1),
            ("A2", "b"), ("B2", 2),
            ("A3", "c"), ("B3", 3),
        ];

        await Assert.That(Calc("=LOOKUP(\"bump\",A1:B3)", tall) as double?).IsEqualTo(2.0);
    }

    // --- COLUMN / COLUMNS — golden: páginas oficiais "COLUMN function"
    // (44e8c754-711c-4df3-9da4-47a55042554b) e "COLUMNS function" (4e8e7b4e-e603-43e8-b177-956088fa48ca) ---

    [Test]
    public async Task Column_OfAReference()
    {
        // =COLUMN(D10) -> 4 (exemplo da página); um range usa "the number of the leftmost column".
        await Assert.That(Calc("=COLUMN(D10)") as double?).IsEqualTo(4.0);
        await Assert.That(Calc("=COLUMN(B2:E4)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Column_WithoutArgument_UsesTheCurrentCell()
    {
        // Regra documentada: sem argumento, "it is assumed to be the reference of the cell in
        // which the COLUMN function appears".
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["C5"] = ExpressionParser.Parse("=COLUMN()", sheet);

        await Assert.That(workbook.GetCellValue("Sheet1", "C5").AsDouble()).IsEqualTo(3.0);
    }

    [Test]
    public async Task Columns_CountsColumnsInRange()
    {
        // =COLUMNS(C1:E4) -> 3 (exemplo da página); uma célula única conta 1 (espelho de ROWS).
        await Assert.That(Calc("=COLUMNS(C1:E4)") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=COLUMNS(A5)") as double?).IsEqualTo(1.0);
    }

    // --- XMATCH — golden: página oficial "XMATCH function" (d966da31-7a6b-4a13-a1c6-5a33ed6a0312) ---

    [Test]
    public async Task XMatch_ExactMatch_ReturnsThePosition()
    {
        // =XMATCH(4,{5,4,3,2,1}) -> 2 (exemplo da página; o array constante vira células aqui).
        (string, object)[] countdown = [("A1", 5), ("A2", 4), ("A3", 3), ("A4", 2), ("A5", 1)];

        await Assert.That(Calc("=XMATCH(4,A1:A5)", countdown) as double?).IsEqualTo(2.0);
        await Assert.That(Calc("=XMATCH(99,A1:A5)", countdown)).IsEqualTo(ErrorValue.NotAvailable);
    }

    private static readonly (string, object)[] Fruits =
    [
        ("C3", "Apple"), ("C4", "Grape"), ("C5", "Pear"), ("C6", "Banana"), ("C7", "Cherry"),
    ];

    [Test]
    public async Task XMatch_NextLargerTextKey()
    {
        // =XMATCH("Gra",C3:C7,1) -> 2 (exemplo da página: "Gra" com match_mode 1 encontra "Grape"
        // na posição 2). Chave TEXTO com match aproximado — regressão do bug 3460bb3.
        await Assert.That(Calc("=XMATCH(\"Gra\",C3:C7,1)", Fruits) as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task XMatch_NextSmallerTextKey()
    {
        // match_mode -1 (exact-or-next-smaller) com chave texto: mesmo caso coberto no XLOOKUP
        // ("Cz" -> "Cedar Falls"), aqui como POSIÇÃO 1-based (regressão 3460bb3).
        (string, object)[] towns =
        [
            ("A1", "Bradbury Creek"), ("A2", "Cedar Falls"), ("A3", "Dunmore"),
        ];

        await Assert.That(Calc("=XMATCH(\"Cz\",A1:A3,-1)", towns) as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task XMatch_WildcardMode()
    {
        // match_mode 2: curingas * ? ~ (mesma semântica do XLOOKUP).
        await Assert.That(Calc("=XMATCH(\"Gr?pe\",C3:C7,2)", Fruits) as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task XMatch_ReverseSearch_FindsTheLastOccurrence()
    {
        // search_mode -1 (last-to-first): com duplicatas, a posição reportada é a da ÚLTIMA.
        (string, object)[] duplicated = [("A1", 1), ("A2", 2), ("A3", 2), ("A4", 3)];

        await Assert.That(Calc("=XMATCH(2,A1:A4,0,-1)", duplicated) as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=XMATCH(2,A1:A4)", duplicated) as double?).IsEqualTo(2.0);
    }

    // --- ADDRESS — golden: página oficial "ADDRESS function" (d0c26c0d-3991-446b-8de4-ab46431d4f89) ---

    [Test]
    public async Task Address_AbsNumForms()
    {
        // =ADDRESS(2,3) -> "$C$2"; =ADDRESS(2,3,2) -> "C$2" (exemplos da página);
        // abs_num 3 = "Relative row; absolute column" -> "$C2"; 4 = "Relative" -> "C2" (tabela).
        await Assert.That(Calc("=ADDRESS(2,3)")).IsEqualTo("$C$2");
        await Assert.That(Calc("=ADDRESS(2,3,2)")).IsEqualTo("C$2");
        await Assert.That(Calc("=ADDRESS(2,3,3)")).IsEqualTo("$C2");
        await Assert.That(Calc("=ADDRESS(2,3,4)")).IsEqualTo("C2");
    }

    [Test]
    public async Task Address_R1C1AbsoluteForm()
    {
        // Derivado dos exemplos =ADDRESS(2,3,1,FALSE,...) -> ...R2C3: a forma absoluta R1C1.
        await Assert.That(Calc("=ADDRESS(2,3,1,FALSE)")).IsEqualTo("R2C3");

        // Limite declarado: as formas R1C1 RELATIVAS (ex.: R2C[3] do exemplo =ADDRESS(2,3,2,FALSE))
        // não são modeladas -> #VALUE!.
        await Assert.That(Calc("=ADDRESS(2,3,2,FALSE)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Address_WithSheetText()
    {
        // =ADDRESS(2,3,1,FALSE,"[Book1]Sheet1") -> "'[Book1]Sheet1'!R2C3";
        // =ADDRESS(2,3,1,FALSE,"EXCEL SHEET") -> "'EXCEL SHEET'!R2C3" (exemplos da página);
        // nome simples fica sem aspas: a página descreve "returns Sheet2!$A$1".
        await Assert
            .That(Calc("=ADDRESS(2,3,1,FALSE,\"[Book1]Sheet1\")"))
            .IsEqualTo("'[Book1]Sheet1'!R2C3");
        await Assert
            .That(Calc("=ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")"))
            .IsEqualTo("'EXCEL SHEET'!R2C3");
        await Assert.That(Calc("=ADDRESS(1,1,1,TRUE,\"Sheet2\")")).IsEqualTo("Sheet2!$A$1");
    }

    [Test]
    public async Task Address_InvalidInputs_AreValueErrors()
    {
        await Assert.That(Calc("=ADDRESS(0,1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ADDRESS(1,0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ADDRESS(1,1,5)")).IsEqualTo(ErrorValue.NotValue);
    }

    // --- AREAS — golden: página oficial "AREAS function" (8392ba32-7a41-43b3-96b0-3695d2ec6152) ---

    [Test]
    public async Task Areas_CountsTheAreasOfAReference()
    {
        // =AREAS(B2:D4) -> 1; =AREAS((B2:D4,E5,F6:I9)) -> 3 (exemplos da página).
        await Assert.That(Calc("=AREAS(B2:D4)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=AREAS((B2:D4,E5,F6:I9))") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=AREAS(A1)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Areas_NonReference_IsValueError()
    {
        await Assert.That(Calc("=AREAS(5)")).IsEqualTo(ErrorValue.NotValue);
    }

    // --- FORMULATEXT — golden: página oficial "FORMULATEXT function"
    // (0a786771-54fd-4ae2-96ee-09cda35439c8) ---

    [Test]
    public async Task FormulaText_ReturnsTheFormulaAsText()
    {
        // Exemplo da página: A2 contém =TODAY() e =FORMULATEXT(A2) -> "=TODAY()". (TODAY ainda não
        // é built-in — parseia como chamada custom, o que não muda o un-parse do texto.)
        await Assert.That(Calc("=FORMULATEXT(A2)", ("A2", "=TODAY()"))).IsEqualTo("=TODAY()");
        await Assert.That(Calc("=FORMULATEXT(A2)", ("A2", "=SUM(B1:B3)"))).IsEqualTo("=SUM(B1:B3)");
    }

    [Test]
    public async Task FormulaText_UsesTheReferencedCellsSheetContext()
    {
        // O contexto de sheet do un-parse é o da célula REFERENCIADA: a referência local B1 da
        // fórmula em Sheet2 volta sem qualificação, mesmo lida a partir de Sheet1.
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        var sheet2 = workbook.Sheets.Add("Sheet2");

        sheet2["A1"] = ExpressionParser.Parse("=B1+1", sheet2);
        sheet1["A1"] = ExpressionParser.Parse("=FORMULATEXT(Sheet2!A1)", sheet1);

        await Assert.That(workbook.GetCellValue("Sheet1", "A1").AsString()).IsEqualTo("=B1+1");
    }

    [Test]
    public async Task FormulaText_NonFormulaCells_AreNotAvailable()
    {
        // Regra documentada: "The cell used as the Reference argument does not contain a formula"
        // -> #N/A. Vale para literal e para célula vazia.
        await Assert.That(Calc("=FORMULATEXT(A2)", ("A2", 7))).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=FORMULATEXT(A2)")).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task FormulaText_NonReference_IsValueError()
    {
        // Regra documentada: "Invalid data types used as inputs will produce a #VALUE! error value."
        await Assert.That(Calc("=FORMULATEXT(5)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task FormulaText_RangeArgument_ReadsTheTopLeftCell()
    {
        await Assert
            .That(Calc("=FORMULATEXT(A2:A4)", ("A2", "=1+1"), ("A3", 5)))
            .IsEqualTo("=1+1");
    }
}
