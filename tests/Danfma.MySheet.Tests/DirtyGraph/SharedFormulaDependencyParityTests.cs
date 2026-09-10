using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 3 (shared-formula-delta): DependencyExtractor deixou de marcar SharedFormulaSlave/AnchoredCellReference/
// AnchoredRangeReference como AlwaysDirty (aproximação conservadora do spike) e passou a extrair as dependências
// EFETIVAS, aplicando o (DeltaRow, DeltaColumn) DA PRÓPRIA escrava aos componentes relativos da árvore ancorada
// do master (os $-ancorados ficam parados).
//
// A PROVA de paridade tem duas camadas:
//  (1) por-fórmula (ScanBoth): compara, para o MESMO texto de master + delta, o DependencyScan do caminho
//      ANCORADO (SharedFormulaSlave em cima da árvore de ExpressionParser.ParseAnchoredMasterBody — o caminho
//      de produção) contra o caminho LEGADO (ExpressionParser.ParseSharedFormulaBody — o parse por-slave
//      token-delta que existia ANTES do spike, preservado intacto para o fallback "honesto"). São dois
//      pipelines de parse independentes; se produzem o MESMO DependencyScan para toda forma suportada pelo
//      modo ancorado, a extração está correta.
//  (2) no grafo (AnchoredWorkbook_vs_LegacyWorkbook_...): dois workbooks completos — um com células
//      SharedFormulaSlave de verdade, outro com a expansão legada nas MESMAS células — comparados via
//      ReverseDependencyGraph.Build (Diagnostics() + GetAllDependents para um conjunto de células de amostra).
public class SharedFormulaDependencyParityTests
{
    private static Sheet NewSheet(string name = "S") => new() { Name = name };

    // Constrói os dois DependencyScans (ancorado vs legado) para o MESMO texto de master + delta de escrava.
    private static (DependencyScan Anchored, DependencyScan Legacy) ScanBoth(
        string masterBody,
        int deltaRow,
        int deltaColumn,
        Workbook? workbook = null
    )
    {
        var sheet = NewSheet();
        var tokens = ExpressionParser.TokenizeFormulaBody(masterBody);

        var anchoredMaster = ExpressionParser.ParseAnchoredMasterBody(tokens, sheet);
        var slave = new SharedFormulaSlave(anchoredMaster, deltaRow, deltaColumn);
        var anchoredScan = DependencyExtractor.Extract(slave, workbook);

        var legacyTree = ExpressionParser.ParseSharedFormulaBody(
            tokens,
            sheet,
            deltaRow,
            deltaColumn
        );
        var legacyScan = DependencyExtractor.Extract(legacyTree, workbook);

        return (anchoredScan, legacyScan);
    }

    private static async Task AssertParity(DependencyScan anchored, DependencyScan legacy)
    {
        await Assert.That(anchored.AlwaysDirty).IsEqualTo(legacy.AlwaysDirty);
        await Assert.That(anchored.Cells.SetEquals(legacy.Cells)).IsTrue();
        await Assert.That(new HashSet<RangeDep>(anchored.Ranges).SetEquals(legacy.Ranges)).IsTrue();
    }

    // --- (1) Por-fórmula: cada forma que AnchoredFormulaSupport aceita no master, para vários deltas ---

    [Test]
    [Arguments(0, 0)]
    [Arguments(1, 0)]
    [Arguments(0, 1)]
    [Arguments(3, 2)]
    public async Task MixedAnchors_AllFourDollarCombinations(int deltaRow, int deltaColumn)
    {
        var (anchored, legacy) = ScanBoth("$A$1+A$1+$A2+A2", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(0, 1)]
    [Arguments(2, 3)]
    public async Task BoundedRange_ShiftsBothEndpoints(int deltaRow, int deltaColumn)
    {
        var (anchored, legacy) = ScanBoth("SUM(B2:D4)", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(4, 0)]
    public async Task RangeWithAbsoluteAnchorOnOneCorner_OnlyTheRelativeCornerMoves(
        int deltaRow,
        int deltaColumn
    )
    {
        var (anchored, legacy) = ScanBoth("SUM($B$2:D4)", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(2, 0)]
    public async Task CrossSheetReference_CarriesTheSheetNameUnshifted(
        int deltaRow,
        int deltaColumn
    )
    {
        var (anchored, legacy) = ScanBoth("Data!A1*3", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(3, 1)]
    public async Task IfBothBranches_CollectAsDependencies(int deltaRow, int deltaColumn)
    {
        var (anchored, legacy) = ScanBoth("IF(A1>5,B1,C1)", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(2, 0)]
    public async Task CustomFunctionOverAnchoredArgument_IsAlwaysDirty_ButStillCollectsIt(
        int deltaRow,
        int deltaColumn
    )
    {
        // MYFUNC não é built-in (vira FunctionCall) — always-dirty nos dois caminhos, mas a dependência do
        // argumento ancorado ainda deve ser coletada (super-aproximação seguindo o mesmo padrão de OFFSET).
        var (anchored, legacy) = ScanBoth("MYFUNC(A1)", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    [Test]
    [Arguments(1)]
    public async Task DefinedName_ResolvesIdenticallyRegardlessOfSlaveDelta(int deltaRow)
    {
        var workbook = new Workbook();
        workbook.DefineName("Sales", "Data!A1:A3");

        var (anchored, legacy) = ScanBoth("SUM(Sales)+A1", deltaRow, 0, workbook);

        await AssertParity(anchored, legacy);
        // A dependência do nome não se move com o delta (posição-independente); só A1 desloca.
        await Assert.That(anchored.Ranges).Contains(new RangeDep("Data", 1, 1, 1, 3));
    }

    [Test]
    [Arguments(1, 0)]
    [Arguments(2, 1)]
    public async Task NestedFunctionsOverRangesAndCells_MixedShape(int deltaRow, int deltaColumn)
    {
        var (anchored, legacy) = ScanBoth("SUM(A1:A3)+MAX(B1,C1)*$D$1", deltaRow, deltaColumn);

        await AssertParity(anchored, legacy);
    }

    // Fase 7 (arrays dinâmicos). Um master com produtor de array continua ELEGÍVEL ao modelo ancorado, e
    // isso não é gratuito: `AnchoredFormulaSupport.IsFullyAnchored` tem um `_ => false` no fim (LET,
    // construções só-de-array, tipos futuros), então um nó novo cairia lá e o GRUPO INTEIRO voltaria ao
    // fallback legado por escrava — mais lento, e sem nada avisando. As quatro passam porque esta fase não
    // acrescenta tipo de nó nenhum: elas são Function comuns, e o arm `Function function =>` aceita uma
    // função quando todo argumento é aceito, chegando aos argumentos pelo mesmo acessor
    // (FormulaWriter.Call) que o writer e o DependencyExtractor usam.
    [Test]
    [Arguments(0, 0)]
    [Arguments(1, 0)]
    [Arguments(2, 3)]
    public async Task DynamicArrayProducerMaster_StaysAnchored_AndKeepsDependencyParity(
        int deltaRow,
        int deltaColumn
    )
    {
        // O caso que o item 19 pede pelo nome, mais os outros três produtores. A paridade é a asserção
        // forte: os DOIS pipelines de parse (ancorado + delta da escrava, e o parse legado por escrava)
        // precisam chegar ao MESMO DependencyScan, incluindo o deslocamento dos componentes relativos.
        foreach (
            var body in new[]
            {
                "SUM(FILTER($A$1:$A$3, $B$1:$B$3>0))",
                "SUM(FILTER(A1:A3,B1:B3>0))",
                "SUM(SORT(A1:A3))",
                "SUM(UNIQUE(A1:A3))",
                "SUM(SEQUENCE(3))",
                "SUM(SEQUENCE(2,2,A1,B1))",
            }
        )
        {
            var (anchored, legacy) = ScanBoth(body, deltaRow, deltaColumn);

            await AssertParity(anchored, legacy);
            await Assert.That(anchored.AlwaysDirty).IsFalse();
        }

        // E o veredito de elegibilidade em si, direto no AnchoredFormulaSupport: as quatro são aceitas com
        // os argumentos ancorados …
        foreach (
            var body in new[]
            {
                "SUM(FILTER($A$1:$A$3, $B$1:$B$3>0))",
                "SUM(SORT($A$1:$A$3))",
                "SUM(UNIQUE($A$1:$A$3))",
                "SUM(SEQUENCE(3))",
            }
        )
        {
            var master = ExpressionParser.ParseAnchoredMasterBody(
                ExpressionParser.TokenizeFormulaBody(body),
                NewSheet()
            );

            await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(master)).IsTrue();
        }

        // … e o controle que mostra que o veredito acima não é um "true" universal: um range ABERTO dentro
        // do produtor continua sendo gatilho de fallback, porque o arm de função só aceita a função quando
        // TODO argumento é aceito.
        var openRange = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody("SUM(FILTER($A:$A, $A:$A>0))"),
            NewSheet()
        );

        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(openRange)).IsFalse();
    }

    // --- (2) No grafo: workbook ancorado (SharedFormulaSlave real) vs workbook legado (mesma célula, árvore
    // expandida) — Diagnostics() + GetAllDependents idênticos para um conjunto de células de amostra ---

    [Test]
    public async Task AnchoredWorkbook_vs_LegacyWorkbook_SameGraphStructure()
    {
        var anchoredWorkbook = new Workbook();
        var legacyWorkbook = new Workbook();
        var anchoredSheet = anchoredWorkbook.Sheets.Add("S");
        var legacySheet = legacyWorkbook.Sheets.Add("S");

        // Grupo escalar: master B2 = A{row}*2, escravas B3..B6 (delta 1..4). Cada escrava depende de UMA
        // célula de input distinta (A3..A6).
        var scalarTokens = ExpressionParser.TokenizeFormulaBody("A2*2");
        var scalarAnchoredMaster = ExpressionParser.ParseAnchoredMasterBody(
            scalarTokens,
            anchoredSheet
        );

        for (var row = 2; row <= 6; row++)
        {
            anchoredSheet[$"A{row}"] = new NumberValue(row * 10);
            legacySheet[$"A{row}"] = new NumberValue(row * 10);
        }

        anchoredSheet["B2"] = ExpressionParser.ParseFormulaBody("A2*2", anchoredSheet); // master (ordinary tree)
        legacySheet["B2"] = ExpressionParser.ParseFormulaBody("A2*2", legacySheet);

        for (var row = 3; row <= 6; row++)
        {
            var delta = row - 2;
            anchoredSheet[$"B{row}"] = new SharedFormulaSlave(scalarAnchoredMaster, delta, 0);
            legacySheet[$"B{row}"] = ExpressionParser.ParseSharedFormulaBody(
                scalarTokens,
                legacySheet,
                delta,
                0
            );
        }

        // Grupo de range: master G2 = SUM(D2:F2), escravas G3,G4 (delta 1,2).
        var rangeTokens = ExpressionParser.TokenizeFormulaBody("SUM(D2:F2)");
        var rangeAnchoredMaster = ExpressionParser.ParseAnchoredMasterBody(
            rangeTokens,
            anchoredSheet
        );

        foreach (var col in new[] { "D", "E", "F" })
        {
            for (var row = 2; row <= 4; row++)
            {
                anchoredSheet[$"{col}{row}"] = new NumberValue(row);
                legacySheet[$"{col}{row}"] = new NumberValue(row);
            }
        }

        anchoredSheet["G2"] = ExpressionParser.ParseFormulaBody("SUM(D2:F2)", anchoredSheet);
        legacySheet["G2"] = ExpressionParser.ParseFormulaBody("SUM(D2:F2)", legacySheet);

        for (var row = 3; row <= 4; row++)
        {
            var delta = row - 2;
            anchoredSheet[$"G{row}"] = new SharedFormulaSlave(rangeAnchoredMaster, delta, 0);
            legacySheet[$"G{row}"] = ExpressionParser.ParseSharedFormulaBody(
                rangeTokens,
                legacySheet,
                delta,
                0
            );
        }

        var anchoredGraph = ReverseDependencyGraph.Build(anchoredWorkbook);
        var legacyGraph = ReverseDependencyGraph.Build(legacyWorkbook);

        // Footprint idêntico: mesmo número de arestas de célula/range, buckets, sempre-dirty, etc.
        await Assert.That(anchoredGraph.Diagnostics()).IsEqualTo(legacyGraph.Diagnostics());

        // GetAllDependents idêntico para: uma dependência escalar (A4 -> só B4), uma dependência de range
        // (E3 -> só G3), e uma célula que ninguém lê (D4 é lido por G4 também, mas F2 só por G2).
        CellDep Cell(string id) =>
            CellAddress.TryGetColumnRow(id, out var col, out var row2)
                ? new CellDep("S", col, row2)
                : default;

        foreach (var source in new[] { "A4", "E3", "F2", "D2" })
        {
            var anchoredDeps = anchoredGraph.GetAllDependents([Cell(source)]);
            var legacyDeps = legacyGraph.GetAllDependents([Cell(source)]);

            await Assert.That(anchoredDeps.SetEquals(legacyDeps)).IsTrue();
        }
    }
}
