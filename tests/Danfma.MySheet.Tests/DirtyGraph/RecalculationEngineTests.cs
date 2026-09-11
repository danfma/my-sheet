using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 7 — API pública do RecalculationEngine + suporte a edição de FÓRMULA (lazy-rebuild) e casos de borda.
public class RecalculationEngineTests
{
    private static CellRef Ref(string id) => new("Sheet1", id);

    // ============ O caso definitivo: uma NOVA dependência após editar a fórmula ============================

    // Editar a FÓRMULA de uma célula muda suas deps. Sem rebuild do grafo, uma edição de VALOR posterior na
    // NOVA dependência não marcaria a célula dirty (o grafo velho aponta para a dep ANTIGA) → valor stale.
    // Este é o defeito que a Fase 7 corrige.
    [Test]
    public async Task FormulaEdit_RewiresDependencies_SoLaterValueEditsFollowTheNewSource()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = new NumberValue(2);
        s["B1"] = ExpressionParser.Parse("=A1", s); // B1 lê A1

        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(1.0); // aquece

        // Reescreve B1 para ler A2 (estrutural). O grafo deve reconstruir.
        s["B1"] = ExpressionParser.Parse("=A2", s);
        var afterFormula = engine.Recalculate([Ref("B1")]);
        await Assert.That(afterFormula.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(2.0);

        // Editar A2 (a NOVA dep) — B1 deve seguir. Edição de VALOR: NÃO reconstrói.
        s["A2"] = new NumberValue(50);
        var afterNewDep = engine.Recalculate([Ref("A2")]);
        await Assert.That(afterNewDep.StructureRebuilt).IsFalse();
        await Assert.That(wb.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(50.0);

        // Editar A1 (a dep ANTIGA) — B1 NÃO deve mais reagir.
        s["A1"] = new NumberValue(999);
        engine.Recalculate([Ref("A1")]);
        await Assert.That(wb.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(50.0);
    }

    // ============ Bump da versão estrutural (o gatilho do rebuild) =========================================

    [Test]
    public async Task StructuralVersion_BumpsOnFormulaShapeEdits_NotOnValueEdits()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);

        // literal → literal (edição de valor): NÃO bumpa.
        var v0 = s.StructuralVersion;
        s["A1"] = new NumberValue(2);
        await Assert.That(s.StructuralVersion).IsEqualTo(v0);

        // literal → fórmula: bumpa.
        s["A1"] = ExpressionParser.Parse("=1+1", s);
        await Assert.That(s.StructuralVersion).IsGreaterThan(v0);

        // fórmula → fórmula: bumpa.
        var v1 = s.StructuralVersion;
        s["A1"] = ExpressionParser.Parse("=2+2", s);
        await Assert.That(s.StructuralVersion).IsGreaterThan(v1);

        // fórmula → literal: bumpa.
        var v2 = s.StructuralVersion;
        s["A1"] = new NumberValue(3);
        await Assert.That(s.StructuralVersion).IsGreaterThan(v2);

        // Remover um literal: NÃO bumpa. Remover uma fórmula: bumpa.
        s["A2"] = new NumberValue(9);
        var v3 = s.StructuralVersion;
        s.Remove("A2");
        await Assert.That(s.StructuralVersion).IsEqualTo(v3);

        s["A3"] = ExpressionParser.Parse("=A1", s);
        var v4 = s.StructuralVersion;
        s.Remove("A3");
        await Assert.That(s.StructuralVersion).IsGreaterThan(v4);
    }

    // ============ Casos de borda ==========================================================================

    [Test]
    public async Task CrossSheetEdit_MarksTheDependentOnTheOtherSheet()
    {
        var wb = new Workbook();
        var s1 = wb.Sheets.Add("Sheet1");
        var s2 = wb.Sheets.Add("Sheet2");
        s1["A1"] = new NumberValue(10);
        s2["B1"] = ExpressionParser.Parse("=Sheet1!A1+1", s2);

        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Sheet2", "B1").ToDouble()).IsEqualTo(11.0);

        s1["A1"] = new NumberValue(100);
        engine.Recalculate([new CellRef("Sheet1", "A1")]);
        await Assert.That(wb.GetCellValue("Sheet2", "B1").ToDouble()).IsEqualTo(101.0);
    }

    [Test]
    public async Task UnionReferenceDependent_IsMarkedByAnEditToAnyArea()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = new NumberValue(2);
        s["C1"] = new NumberValue(3);
        s["D1"] = ExpressionParser.Parse("=SUM((A1:A2,C1))", s); // união de duas áreas

        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Sheet1", "D1").ToDouble()).IsEqualTo(6.0);

        s["C1"] = new NumberValue(30);
        engine.Recalculate([Ref("C1")]);
        await Assert.That(wb.GetCellValue("Sheet1", "D1").ToDouble()).IsEqualTo(33.0);
    }

    [Test]
    public async Task Cycle_DoesNotHang_AndRecalculateTerminates()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = ExpressionParser.Parse("=B1+1", s);
        s["B1"] = ExpressionParser.Parse("=A1+1", s);

        var engine = wb.CreateRecalculationEngine(); // build não deve travar no ciclo

        s["A1"] = ExpressionParser.Parse("=B1+2", s);
        var result = engine.Recalculate([Ref("A1")]); // não deve travar

        // O ciclo resolve para #REF! (guarda de recursão); só afirmamos que não trava e responde.
        await Assert.That(wb.GetCellValue("Sheet1", "A1").Kind).IsEqualTo(ComputedValueKind.Error);
        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.Partial);
    }

    [Test]
    public async Task SheetAdded_TriggersRebuild()
    {
        var wb = new Workbook();
        var s1 = wb.Sheets.Add("Sheet1");
        s1["A1"] = new NumberValue(1);

        var engine = wb.CreateRecalculationEngine();

        // Adiciona uma sheet nova com uma fórmula; o engine deve detectar a mudança de conjunto de sheets.
        var s2 = wb.Sheets.Add("Sheet2");
        s2["B1"] = ExpressionParser.Parse("=Sheet1!A1*2", s2);

        var result = engine.Recalculate([new CellRef("Sheet1", "A1")]);
        await Assert.That(result.StructureRebuilt).IsTrue();
    }

    [Test]
    public async Task SheetRemoved_TriggersRebuild_WithoutCrashing()
    {
        var wb = new Workbook();
        var s1 = wb.Sheets.Add("Sheet1");
        var s2 = wb.Sheets.Add("Sheet2");
        s1["A1"] = new NumberValue(1);
        s2["B1"] = ExpressionParser.Parse("=Sheet1!A1*2", s2);

        var engine = wb.CreateRecalculationEngine();
        wb.GetCellValue("Sheet2", "B1");

        wb.Sheets.TryRemove("Sheet2", out _);

        s1["A1"] = new NumberValue(5);
        var result = engine.Recalculate([new CellRef("Sheet1", "A1")]);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Sheet1", "A1").ToDouble()).IsEqualTo(5.0);
    }

    [Test]
    public async Task NonA1EditedCell_FallsBackToFullRecompute()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["B1"] = ExpressionParser.Parse("=A1+1", s);

        var engine = wb.CreateRecalculationEngine();
        wb.GetCellValue("Sheet1", "B1");

        // Um id que não é um endereço A1 (só letras) → fora do modelo denso → fallback full.
        var result = engine.Recalculate([new CellRef("Sheet1", "notacell")]);
        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
    }

    [Test]
    public async Task CollectOutputs_ReturnsTheAffectedSinks()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s); // sink
        s["B1"] = ExpressionParser.Parse("=A1*2", s); // sink

        var engine = wb.CreateRecalculationEngine();

        s["A1"] = new NumberValue(10);
        var result = engine.Recalculate([Ref("A1")], collectOutputs: true);

        var outputs = new HashSet<string>();
        foreach (var output in result.AffectedOutputs)
        {
            outputs.Add(output.Id);
        }

        await Assert.That(outputs.SetEquals(["A3", "B1"])).IsTrue();
    }

    [Test]
    public async Task ValueEdit_DoesNotRebuild_AndPropagates()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s);

        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0); // aquece

        s["A1"] = new NumberValue(10);
        var result = engine.Recalculate([Ref("A1")]);

        await Assert.That(result.StructureRebuilt).IsFalse();
        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.Partial);
        await Assert.That(wb.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(12.0);
    }

    // ============ Redefinição de nome (nenhuma Sheet muda — só o mapa de nomes do Workbook) ================

    // Um nome é resolvido para sua definição ATUAL na construção do grafo (DependencyExtractor.ResolveName).
    // Redefinir um nome — repontá-lo para outra referência — muda a estrutura de dependências exatamente como
    // editar uma fórmula, mas não toca em nenhuma Sheet, então nenhum Sheet.StructuralVersion bumpa. Sem
    // rastrear isso separadamente (Workbook.DefinitionsVersion), o grafo antigo (nome→alvo velho) sobrevive à
    // redefinição, e uma edição na dependente NOVA não marca dirty a fórmula que usa o nome → valor stale.
    [Test]
    public async Task NameRedefinition_MarksTheGraphStale_SoDependentsFollowTheNewTarget()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new NumberValue(10);
        data["A2"] = new NumberValue(99);
        wb.DefineName("N", "Data!A1");
        main["B1"] = ExpressionParser.Parse("=SUM(N)", main);

        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(10.0); // aquece com a def. antiga

        // Repontar N para A2 — estrutural, mas nenhuma Sheet foi tocada.
        wb.DefineName("N", "Data!A2");

        // Fluxo normal: reporta a NOVA dependência (A2) como editada. Sem o fix, o grafo continua apontando
        // N→A1, então A2 não tem dependentes nele e B1 nunca é marcado dirty (fica servindo 10, stale).
        var result = engine.Recalculate([new CellRef("Data", "A2")]);

        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(99.0);
    }

    // ============ Redefinição = evento de invalidação TOTAL (não só rebuild do grafo) ======================

    // Bumpar a versão e reconstruir o grafo NÃO basta: os valores já memoizados sobrevivem à redefinição. Medido
    // hoje, com Rng = Data!A1:A2 alargado para Data!A1:A4, o engine responde `mode=Partial dirty=0 rebuilt=True`
    // e Main!B1 continua servindo 3 (quer 15) — só InvalidateCache() limpa. É um defeito PRÉ-EXISTENTE do
    // DefineName, que a redefinição de tabela herdaria: uma mudança de definição tem de ser tratada como
    // invalidação completa.
    [Test]
    public async Task RedefiningADefinedName_ForcesAFullRecompute()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new NumberValue(1);
        data["A2"] = new NumberValue(2);
        data["A3"] = new NumberValue(4);
        data["A4"] = new NumberValue(8);
        wb.DefineName("Rng", "Data!A1:A2");
        main["B1"] = ExpressionParser.Parse("=SUM(Rng)", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece

        wb.DefineName("Rng", "Data!A1:A4");
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.DirtyCellCount).IsEqualTo(-1);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0);
    }

    // O par documentado planejar-depois-agir consome o sinal. EnsureFresh atualiza o snapshot como efeito
    // colateral, então quem rodar PRIMEIRO come a única evidência de que a definição mudou: medido hoje,
    // EstimateImpact([]) devolve `RecommendFull=False ConeSize=0` e o Recalculate([]) seguinte devolve
    // `mode=Partial dirty=0 rebuilt=False` com B1 = 3. A invalidação precisa ser PEGAJOSA — lida por
    // EstimateImpact sem limpar, limpa só no braço que chama InvalidateCache.
    [Test]
    public async Task RedefiningADefinedName_ForcesAFullRecompute_EvenAfterEstimateImpact()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new NumberValue(1);
        data["A2"] = new NumberValue(2);
        data["A3"] = new NumberValue(4);
        data["A4"] = new NumberValue(8);
        wb.DefineName("Rng", "Data!A1:A2");
        main["B1"] = ExpressionParser.Parse("=SUM(Rng)", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece

        wb.DefineName("Rng", "Data!A1:A4");
        var estimate = engine.EstimateImpact([]);
        var result = engine.Recalculate([]);

        await Assert.That(estimate.RecommendFull).IsTrue();
        await Assert.That(estimate.ConeSize).IsEqualTo(-1);
        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.DirtyCellCount).IsEqualTo(-1);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0);
    }

#if MYSHEET_TABLES
    // O gêmeo de tabela: redefinir uma TABELA é o MESMO evento de invalidação total, porque o contador é único
    // (DefinitionsVersion). Quando estes dois foram escritos (Fase 3) não havia nó de fórmula que REFERENCIASSE
    // uma tabela, então o valor stale é montado do único jeito que havia: uma edição de célula que NUNCA é
    // reportada ao engine. Literal→literal não bumpa Sheet.StructuralVersion, logo o único sinal de staleness é
    // a redefinição. O gêmeo que usa o nó de verdade é StructuredReferenceGraph_Rebuilds_AfterDefineTable,
    // abaixo (Fase 5).
    [Test]
    public async Task RedefiningATable_ForcesAFullRecompute()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new NumberValue(1);
        data["A2"] = new NumberValue(2);
        data["A3"] = new NumberValue(4);
        data["A4"] = new NumberValue(8);
        wb.DefineTable("Vendas", "Data", "A1:A2", ["Produto"]);
        main["B1"] = ExpressionParser.Parse("=SUM(Data!A1:A2)", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece

        data["A2"] = new NumberValue(14); // edição NÃO reportada: sozinha, não gera nenhum sinal
        wb.DefineTable("Vendas", "Data", "A1:A4", ["Produto"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.DirtyCellCount).IsEqualTo(-1);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0);
    }

    [Test]
    public async Task RedefiningATable_ForcesAFullRecompute_EvenAfterEstimateImpact()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new NumberValue(1);
        data["A2"] = new NumberValue(2);
        data["A3"] = new NumberValue(4);
        data["A4"] = new NumberValue(8);
        wb.DefineTable("Vendas", "Data", "A1:A2", ["Produto"]);
        main["B1"] = ExpressionParser.Parse("=SUM(Data!A1:A2)", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece

        data["A2"] = new NumberValue(14); // edição NÃO reportada
        wb.DefineTable("Vendas", "Data", "A1:A4", ["Produto"]);
        var estimate = engine.EstimateImpact([]);
        var result = engine.Recalculate([]);

        await Assert.That(estimate.RecommendFull).IsTrue();
        await Assert.That(estimate.ConeSize).IsEqualTo(-1);
        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.DirtyCellCount).IsEqualTo(-1);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0);
    }

    // Fase 5: o gêmeo dos dois acima com uma fórmula que REFERENCIA a tabela de verdade, o que torna o valor
    // stale real em vez de montado. Sem o bump de DefinitionsVersion + InvalidateCache, alargar a tabela
    // deixaria B1 servindo 3 (a soma da geometria ANTIGA) para sempre: nenhuma Sheet foi tocada, então
    // Sheet.StructuralVersion não bumpa, e o único sinal é a redefinição. O parser ainda não emite o nó nesta
    // branch (Fase 4 T5), então a árvore é montada à mão.
    [Test]
    public async Task StructuredReferenceGraph_Rebuilds_AfterDefineTable()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Valor");
        data["A2"] = new NumberValue(1);
        data["A3"] = new NumberValue(2);
        data["A4"] = new NumberValue(4);
        data["A5"] = new NumberValue(8);
        wb.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);
        main["B1"] = new Sum([new TableReference("Tabela1", "Valor", TableArea.Data)]);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece: A2:A3

        // O controle que dá dentes ao resto: A4 está FORA da geometria antiga, então editá-lo e reportá-lo
        // não pode mover B1. Se movesse, o 15 abaixo não provaria nada sobre a redefinição.
        data["A4"] = new NumberValue(1000);
        engine.Recalculate([new CellRef("Data", "A4")]);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0);
        data["A4"] = new NumberValue(4);
        engine.Recalculate([new CellRef("Data", "A4")]);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0);

        // Alarga a tabela: [Valor] passa a ser A2:A5. Nenhuma célula mudou — só a definição.
        wb.DefineTable("Tabela1", "Data", "A1:A5", ["Valor"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.DirtyCellCount).IsEqualTo(-1);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0); // 1+2+4+8

        // E o grafo RECONSTRUÍDO aponta para o retângulo novo: A5 agora é dependência de verdade, e segui-la
        // é uma edição de VALOR (não reconstrói). É este par — o controle de A4 acima e esta linha — que
        // distingue "recalculou tudo uma vez" de "o grafo aprendeu a geometria nova".
        data["A5"] = new NumberValue(90);
        var afterValueEdit = engine.Recalculate([new CellRef("Data", "A5")]);

        await Assert.That(afterValueEdit.StructureRebuilt).IsFalse();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(97.0); // 1+2+4+90

        // O segundo controle, e o que prova que o 97 vem do cone dirty e não de "recomputa tudo sempre": uma
        // célula fora da geometria NOVA também não move B1.
        data["A9"] = new NumberValue(5000);
        engine.Recalculate([new CellRef("Data", "A9")]);

        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(97.0);
    }

    // O espelho do caso irresolúvel: registrar a tabela DEPOIS é uma mudança de definição como qualquer outra,
    // então o #NAME? memoizado não sobrevive a ela. (Que o extractor marque a fórmula always-dirty enquanto a
    // tabela não existe é asserido em DependencyExtractorTests, não aqui.)
    [Test]
    public async Task AnUnresolvableStructuredReference_FollowsALaterDefineTable()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Valor");
        data["A2"] = new NumberValue(1);
        data["A3"] = new NumberValue(2);
        main["B1"] = new Sum([new TableReference("Tabela1", "Valor", TableArea.Data)]);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").AsObject()).IsEqualTo(ErrorValue.Name);

        wb.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0);
    }

    // Item 24's second half, unblocked when Phase 4 T5's parse arms merged: the name is defined through the
    // STRING overload, so the definition goes through the real parser — which now emits a TableReference —
    // and through HasUnqualifiedReference, whose `_ => false` arm is what admits a structured reference in
    // a definition (a table name is workbook-scoped; there is no sheet to qualify). The Expression overload
    // the twin above used bypasses both halves, which is why this shape was unreachable before the parser:
    // the definition below is a STRING, and the formula that reads the name is a PARSED formula.
    [Test]
    public async Task ADefinedNameOverAStructuredReference_Rebuilds_AfterDefineTable()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Valor");
        data["A2"] = new NumberValue(1);
        data["A3"] = new NumberValue(2);
        data["A4"] = new NumberValue(4);
        data["A5"] = new NumberValue(8);
        wb.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);
        wb.DefineName("Nome", "SUM(Tabela1[Valor])");
        main["B1"] = ExpressionParser.Parse("=Nome", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece: A2:A3

        // O mesmo controle do gêmeo de cima: A4 está FORA da geometria antiga, então editá-lo (e reportar)
        // não pode mover B1 — senão o 15 abaixo não provaria nada sobre a redefinição.
        data["A4"] = new NumberValue(1000);
        engine.Recalculate([new CellRef("Data", "A4")]);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0);
        data["A4"] = new NumberValue(4);
        engine.Recalculate([new CellRef("Data", "A4")]);
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0);

        // Alarga a tabela: a DEFINIÇÃO do nome passa a denotar A2:A5. Nenhuma célula mudou — só a definição.
        wb.DefineTable("Tabela1", "Data", "A1:A5", ["Valor"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0); // 1+2+4+8

        // E o grafo RECONSTRUÍDO aponta para o retângulo novo, através da definição do nome: A5 agora é
        // dependência de verdade, e segui-la é uma edição de VALOR (não reconstrói).
        data["A5"] = new NumberValue(90);
        var afterValueEdit = engine.Recalculate([new CellRef("Data", "A5")]);

        await Assert.That(afterValueEdit.StructureRebuilt).IsFalse();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(97.0); // 1+2+4+90
    }

    // O gêmeo do NOME NU da tabela (`=SUM(Tabela1)`): o parser o entrega como NameReference, o extractor o
    // marca conservadoramente AlwaysDirty (pinado em DependencyExtractorTests) — então não há retângulo
    // estático a aprender, e o valor segue a redefinição pela recomputação de toda passada. O par abaixo é
    // o que prova que a marca conservadora faz o trabalho dela: 3 antes, 15 depois do alargamento.
    [Test]
    public async Task ABareTableName_FollowsALaterDefineTable()
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var main = wb.Sheets.Add("Main");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Valor");
        data["A2"] = new NumberValue(1);
        data["A3"] = new NumberValue(2);
        data["A4"] = new NumberValue(4);
        data["A5"] = new NumberValue(8);
        wb.DefineTable("Tabela1", "Data", "A1:A3", ["Valor"]);
        main["B1"] = ExpressionParser.Parse("=SUM(Tabela1)", main);

        wb.ComputeAll();
        var engine = wb.CreateRecalculationEngine();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(3.0); // aquece: A2:A3

        wb.DefineTable("Tabela1", "Data", "A1:A5", ["Valor"]);
        var result = engine.Recalculate([]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.FullFallback);
        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(wb.GetCellValue("Main", "B1").ToDouble()).IsEqualTo(15.0); // 1+2+4+8
    }
#endif
}
