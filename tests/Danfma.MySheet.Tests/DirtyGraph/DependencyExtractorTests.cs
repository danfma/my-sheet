using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 1 do spike do grafo de dependências: extração forward de deps do AST + classificação always-dirty.
public class DependencyExtractorTests
{
    private static DependencyScan Scan(string formula, Workbook? workbook = null)
    {
        var sheet = new Sheet { Name = "Sheet1" };
        return DependencyExtractor.Extract(ExpressionParser.Parse(formula, sheet), workbook);
    }

    private static CellDep Cell(int col, int row, string sheet = "Sheet1") => new(sheet, col, row);

    [Test]
    public async Task SimpleArithmetic_CollectsBothCells()
    {
        var scan = Scan("=A1+B2");

        await Assert.That(scan.Cells.Count).IsEqualTo(2);
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Cells.Contains(Cell(2, 2))).IsTrue(); // B2
        await Assert.That(scan.Ranges.Count).IsEqualTo(0);
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task BoundedRange_IsASingleRangeDep()
    {
        var scan = Scan("=SUM(A1:C3)");

        await Assert.That(scan.Cells.Count).IsEqualTo(0);
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 3, 1, 3));
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task WholeColumn_IsAnOpenRangeDep()
    {
        var scan = Scan("=SUM(A:A)");

        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, null, null));
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task Union_CollectsEachArea()
    {
        var scan = Scan("=SUM((A1:A2,C1))");

        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 2)); // A1:A2
        await Assert.That(scan.Cells.Contains(Cell(3, 1))).IsTrue(); // C1
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task Offset_IsAlwaysDirty_ButStillCollectsTheBaseCell()
    {
        var scan = Scan("=OFFSET(A1,1,0)");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1 (a base é uma dep real)
    }

    [Test]
    public async Task Indirect_IsAlwaysDirty()
    {
        var scan = Scan("=B1&INDIRECT(B2)");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(2, 1))).IsTrue(); // B1
        await Assert.That(scan.Cells.Contains(Cell(2, 2))).IsTrue(); // B2
    }

    [Test]
    public async Task Index_IsNotDirty_TheWholeRangeIsTheDependency()
    {
        // INDEX varre o range inteiro (row/col computados), então a dep é o range — enumerável, não dirty.
        var scan = Scan("=INDEX(A1:C10,2,3)");

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 3, 1, 10));
    }

    [Test]
    public async Task If_CollectsConditionAndBothBranches()
    {
        // Super-aproximação: ambos os ramos são deps (qualquer um pode ser o resultado).
        var scan = Scan("=IF(A1>5,B1,C1)");

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Cells.Contains(Cell(2, 1))).IsTrue(); // B1
        await Assert.That(scan.Cells.Contains(Cell(3, 1))).IsTrue(); // C1
    }

    [Test]
    public async Task Now_IsAlwaysDirty_WithNoCellDeps()
    {
        var scan = Scan("=NOW()");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Count).IsEqualTo(0);
        await Assert.That(scan.Ranges.Count).IsEqualTo(0);
    }

    [Test]
    public async Task CrossSheetReference_CarriesTheSheetName()
    {
        var scan = Scan("=Sheet2!A1*3");

        await Assert.That(scan.Cells.Contains(Cell(1, 1, "Sheet2"))).IsTrue();
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task DynamicRangeEndpoint_IsAlwaysDirty()
    {
        // A1:INDEX(B1:B9,3) — endpoint reference-returning → DynamicRange → always-dirty, mas coleta o que dá.
        var scan = Scan("=SUM(A1:INDEX(B1:B9,3))");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 2, 2, 1, 9)); // B1:B9
    }

    [Test]
    public async Task CustomFunctionCall_IsAlwaysDirty()
    {
        var scan = Scan("=MYFUNC(A1)"); // não é built-in → FunctionCall

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
    }

    [Test]
    public async Task DefinedName_ResolvesToItsReference()
    {
        var workbook = new Workbook();
        workbook.DefineName("Sales", "Data!A1:A3");

        var scan = Scan("=SUM(Sales)", workbook);

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Ranges).Contains(new RangeDep("Data", 1, 1, 1, 3));
    }

    [Test]
    public async Task UnresolvableName_IsAlwaysDirty()
    {
        var scan = Scan("=SUM(Unknown)"); // sem workbook → nome irresolúvel

        await Assert.That(scan.AlwaysDirty).IsTrue();
    }

    // Fase 8 (lifting elementwise): a fórmula lifted não muda NADA aqui, e é isso que o pin registra. O arm
    // `case Function function:` visita os argumentos via FormulaWriter.Call — o mesmo acessor do registry que
    // o writer e o AnchoredFormulaSupport usam — e esta fase não acrescenta nó nenhum à árvore, então o range
    // dentro do LEN continua sendo uma RangeDep normal. Se não fosse, a célula ficaria fora do cone dirty de
    // uma edição em A1:A3 e serviria valor stale.
    [Test]
    public async Task LiftedFunctionArgument_IsStillARangeDep()
    {
        var scan = Scan("=SUM(LEN(A1:A3))");

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 3));

        // O mesmo para a metade unária do lifting, e para dois lifts empilhados.
        await Assert
            .That(Scan("=SUM(-(A1:A3>1))").Ranges)
            .Contains(new RangeDep("Sheet1", 1, 1, 1, 3));
        await Assert
            .That(Scan("=SUM(LEN(TRIM(A1:A3)))").Ranges)
            .Contains(new RangeDep("Sheet1", 1, 1, 1, 3));
    }

    // ------------------------------------------------------------------------------------------------
    // Fase 7 (arrays dinâmicos). Os quatro produtores — FILTER, SORT, UNIQUE e SEQUENCE — NÃO exigiram
    // mudança nenhuma neste arquivo, e o objetivo destes pins é justamente TRAVAR essa afirmação em vez
    // de deixá-la implícita. Ela é load-bearing e não é óbvia:
    //
    //   * quem produz as deps é o arm genérico `case Function function:` (:207) via `VisitArguments`
    //     (:223-247), que resolve os argumentos pelo MESMO acessor de registry que o FormulaWriter e o
    //     AnchoredFormulaSupport usam. Registrar as quatro (item 11) é o que torna os ranges delas
    //     enumeráveis — nada além disso;
    //   * o `default: return;` de :216-217 é uma armadilha de dependência PERDIDA em silêncio para
    //     qualquer nó que o registry não cubra, então "nada a fazer aqui" precisa de prova;
    //   * o COMPRIMENTO do resultado do FILTER depende dos dados, mas o CONJUNTO de dependências não —
    //     o range de origem é lido inteiro de qualquer forma. É exatamente a superaproximação que este
    //     arquivo documenta em :40-45, e é por isso que marcar AlwaysDirty seria uma pessimização
    //     gratuita: uma fórmula always-dirty recomputa em TODA passada de recálculo, para sempre.
    //
    // Uma classificação errada aqui não quebra nenhum valor — ela só custa desempenho em silêncio, o que
    // é a razão de o pin existir antes de alguém "consertar" a laziness do FILTER marcando-o volátil.
    // ------------------------------------------------------------------------------------------------

    [Test]
    public async Task DynamicArrayProducer_ContributesItsRanges_AndIsNeverAlwaysDirty()
    {
        // FILTER lê DOIS ranges (a origem e o vetor include) e nenhum outro: exatamente dois RangeDeps,
        // nenhuma CellDep, e AlwaysDirty falso.
        var filter = Scan("=SUM(FILTER(A1:A3,B1:B3>0))");

        await Assert.That(filter.AlwaysDirty).IsFalse();
        await Assert.That(filter.Ranges.Count).IsEqualTo(2);
        await Assert.That(filter.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 3)); // A1:A3
        await Assert.That(filter.Ranges).Contains(new RangeDep("Sheet1", 2, 2, 1, 3)); // B1:B3
        await Assert.That(filter.Cells.Count).IsEqualTo(0);

        // SORT e UNIQUE: um range cada, e os argumentos escalares opcionais (sort_index, sort_order,
        // by_col, by_row, exactly_once) não acrescentam nem removem nada.
        foreach (
            var formula in new[]
            {
                "=SUM(SORT(A1:A3))",
                "=SUM(SORT(A1:A3,2,-1,TRUE))",
                "=SUM(UNIQUE(A1:A3))",
                "=SUM(UNIQUE(A1:A3,FALSE,TRUE))",
            }
        )
        {
            var scan = Scan(formula);

            await Assert.That(scan.AlwaysDirty).IsFalse();
            await Assert.That(scan.Ranges.Count).IsEqualTo(1);
            await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 3));
        }

        // SEQUENCE não lê célula nenhuma: nenhum range, nenhuma célula — e MESMO ASSIM não é always-dirty.
        // É o contraste direto com Indirect_IsAlwaysDirty (:68) e Now_IsAlwaysDirty acima: "não tem
        // dependência" não é o mesmo que "a dependência não é enumerável". SEQUENCE(3) é uma constante
        // estrutural; INDIRECT(B2) e NOW() não são.
        var sequence = Scan("=SUM(SEQUENCE(5))");

        await Assert.That(sequence.AlwaysDirty).IsFalse();
        await Assert.That(sequence.Ranges.Count).IsEqualTo(0);
        await Assert.That(sequence.Cells.Count).IsEqualTo(0);

        // Mas um argumento de SEQUENCE que aponta para célula É dependência de verdade — se não fosse,
        // editar A1 não recalcularia a fórmula e ela serviria a sequência antiga.
        var seeded = Scan("=SUM(SEQUENCE(2,3,A1,B1))");

        await Assert.That(seeded.AlwaysDirty).IsFalse();
        await Assert.That(seeded.Cells.Contains(Cell(1, 1))).IsTrue(); // A1 (start)
        await Assert.That(seeded.Cells.Contains(Cell(2, 1))).IsTrue(); // B1 (step)
    }

    [Test]
    public async Task DynamicArrayProducer_OverANameOrAnotherProducer_StillEnumeratesEveryRange()
    {
        // Um NOME definido dentro do argumento do produtor resolve para o range que ele denota, na folha
        // que ele nomeia — o mesmo caminho de DefinedName_ResolvesToItsReference, agora sob o produtor.
        var workbook = new Workbook();
        workbook.DefineName("Sales", "Data!A1:A3");

        foreach (var formula in new[] { "=SUM(FILTER(Sales,Sales>0))", "=SUM(SORT(Sales))" })
        {
            var scan = Scan(formula, workbook);

            await Assert.That(scan.AlwaysDirty).IsFalse();
            await Assert.That(scan.Ranges).Contains(new RangeDep("Data", 1, 1, 1, 3));
        }

        // Um produtor DENTRO de outro (em qualquer um dos dois slots do FILTER) continua entregando os
        // dois ranges: a recursão de VisitArguments não para no primeiro nó de função.
        var nested = Scan("=SUM(FILTER(SORT(A1:A3),UNIQUE(B1:B3)>0))");

        await Assert.That(nested.AlwaysDirty).IsFalse();
        await Assert.That(nested.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 3)); // A1:A3
        await Assert.That(nested.Ranges).Contains(new RangeDep("Sheet1", 2, 2, 1, 3)); // B1:B3

        // Produtor sob um operador, cruzando folha, e sobre uma união de áreas — três formas em que o
        // range poderia se perder no caminho.
        var mixed = Scan("=SUM(FILTER(A1:A3,B1:B3>0)*C1:C3)");

        await Assert.That(mixed.AlwaysDirty).IsFalse();
        await Assert.That(mixed.Ranges.Count).IsEqualTo(3);
        await Assert.That(mixed.Ranges).Contains(new RangeDep("Sheet1", 3, 3, 1, 3)); // C1:C3

        await Assert
            .That(Scan("=SUM(FILTER(Sheet2!A1:A3,Sheet2!B1:B3>0))").Ranges)
            .Contains(new RangeDep("Sheet2", 1, 1, 1, 3));

        var union = Scan("=SUM(FILTER((A1:A2,C1),B1:B2>0))");

        await Assert.That(union.AlwaysDirty).IsFalse();
        await Assert.That(union.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 2)); // A1:A2
        await Assert.That(union.Cells.Contains(Cell(3, 1))).IsTrue(); // C1

        // Um range ABERTO dentro do produtor continua sendo um RangeDep aberto e NÃO always-dirty: a
        // recusa do mini-CSE em transmitir um range aberto (a divergência declarada da fase, pinada em
        // MiniCseConsumerTests) é uma decisão de AVALIAÇÃO e não muda o grafo.
        var open = Scan("=SUM(FILTER(A:A,A:A>0))");

        await Assert.That(open.AlwaysDirty).IsFalse();
        await Assert.That(open.Ranges).Contains(new RangeDep("Sheet1", 1, 1, null, null));
    }

    [Test]
    public async Task TheNotAlwaysDirtyVerdict_HasTeeth_AnUnregisteredTwinIsAlwaysDirty()
    {
        // O QUE FAZ DOS PINS ACIMA UMA PROVA E NÃO UMA TAUTOLOGIA. `AlwaysDirty` é falso por padrão, então
        // asserir "é falso" não distingue "o extractor reconheceu a função" de "o extractor não fez nada".
        // O controle é o gêmeo NÃO REGISTRADO: mesma forma, mesmos ranges, nome inexistente — vira
        // FunctionCall e cai no arm de função custom (comportamento desconhecido ⇒ always-dirty), como em
        // CustomFunctionCall_IsAlwaysDirty. Se FILTER algum dia sair do registry, ou se as quatro forem
        // reclassificadas, este par deixa de ser um par e o teste falha nomeando o motivo.
        await Assert.That(Scan("=SUM(FILTER(A1:A3,B1:B3>0))").AlwaysDirty).IsFalse();
        await Assert.That(Scan("=SUM(FILTERX(A1:A3,B1:B3>0))").AlwaysDirty).IsTrue();

        await Assert.That(Scan("=SUM(SEQUENCE(5))").AlwaysDirty).IsFalse();
        await Assert.That(Scan("=SUM(SEQUENCEX(5))").AlwaysDirty).IsTrue();

        // E os dois gêmeos coletam os MESMOS ranges: a diferença está só no veredito, o que confirma que
        // o always-dirty do não registrado vem da classificação e não de uma falha em enxergar o range.
        await Assert
            .That(Scan("=SUM(FILTERX(A1:A3,B1:B3>0))").Ranges)
            .Contains(new RangeDep("Sheet1", 1, 1, 1, 3));

        // Um nó realmente não enumerável DENTRO de um produtor continua contaminando o veredito, como
        // deve: o produtor não é uma barreira.
        await Assert.That(Scan("=SUM(SORT(OFFSET(A1,1,0)))").AlwaysDirty).IsTrue();
        await Assert.That(Scan("=SUM(UNIQUE(INDIRECT(\"A1:A3\")))").AlwaysDirty).IsTrue();
    }
}
