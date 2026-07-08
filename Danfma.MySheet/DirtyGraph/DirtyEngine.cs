namespace Danfma.MySheet.DirtyGraph;

/// <summary>
/// O motor de recomputação incremental (evict-and-pull) do spike. Sobre um <see cref="Workbook"/> e seu
/// <see cref="ReverseDependencyGraph"/>: dado um conjunto de células editadas, computa o cone dirty (as
/// editadas ∪ seus dependentes transitivos ∪ as sempre-dirty), EVICTA exatamente essas do cache memoizado,
/// e deixa o motor pull-based recomputar cada uma 1× na próxima leitura. Nunca chama
/// <see cref="Workbook.InvalidateCache"/> — só descarta o que a edição realmente afeta.
///
/// <para><b>Contrato.</b> O grafo é construído uma vez e assume estrutura ESTÁVEL: edições de VALOR (trocar o
/// literal de uma célula de input) mantêm o grafo válido; editar a FÓRMULA de uma célula muda suas
/// dependências e exigiria reconstruir/manter o grafo (Fase 3 incremental, fora do spike atual). Uso na fase
/// de edição, single-thread (mesmo contrato do índice estrutural).</para>
/// </summary>
internal sealed class DirtyEngine
{
    private readonly Workbook _workbook;
    private readonly ReverseDependencyGraph _graph;

    private DirtyEngine(Workbook workbook, ReverseDependencyGraph graph)
    {
        _workbook = workbook;
        _graph = graph;
    }

    /// <summary>Constrói o motor (e o grafo reverso) a partir do estado atual do workbook.</summary>
    public static DirtyEngine Build(Workbook workbook) =>
        new(workbook, ReverseDependencyGraph.Build(workbook));

    public ReverseDependencyGraph Graph => _graph;

    // Acima deste tamanho de cone, o impacto é grande demais (≈ alcançou a coluna quente / meio workbook):
    // full-recompute é o certo e a enumeração por-célula degradaria. Casa com a distribuição bimodal do K1.
    private const int LargeConeCap = 50_000;

    /// <summary>
    /// Marca <paramref name="edited"/> como dirty e recomputa seletivamente:
    /// <list type="bullet">
    /// <item>Cone PEQUENO (o caso comum): computa o cone transitivo (output-sensitive) ∪ editadas ∪
    /// sempre-dirty, EVICTA só essas do cache, e retorna o conjunto — ler qualquer célula produz valores
    /// frescos (pull). Recomputa só o que a edição afeta.</item>
    /// <item>Cone GRANDE (> <see cref="LargeConeCap"/>, a cauda da coluna quente): retorna <c>null</c> após
    /// um <see cref="Workbook.InvalidateCache"/> — ali o impacto é ~metade do workbook e o full-recompute é o
    /// caminho certo.</item>
    /// </list>
    /// </summary>
    public HashSet<CellDep>? CalculateDirty(IReadOnlyCollection<CellDep> edited)
    {
        var dirty = _graph.GetAllDependentsBounded(edited, LargeConeCap);

        if (dirty is null)
        {
            _workbook.InvalidateCache(); // cone grande → full-recompute (lazy no próximo read)
            return null;
        }

        foreach (var cell in edited)
        {
            dirty.Add(cell); // o cache da própria célula editada está stale
        }

        foreach (var cell in _graph.AlwaysDirty)
        {
            dirty.Add(cell); // voláteis/dinâmicos recomputam sempre
        }

        foreach (var cell in dirty)
        {
            _workbook.EvictDense(cell.Sheet, cell.Column, cell.Row);
        }

        return dirty;
    }

    /// <summary>
    /// Recomputa o cone dirty imediatamente (pull) lendo cada célula dirty via o cache — cada uma é avaliada
    /// no máximo 1×. Conveniência para medir o custo de recompute; na prática o host lê só os outputs que
    /// precisa. Roda em thread de pilha grande (cadeias profundas).
    /// </summary>
    public void PullAll(IReadOnlyCollection<CellDep> dirty) =>
        Workbook.RunWithLargeStack(() =>
        {
            foreach (var cell in dirty)
            {
                _workbook.GetCellValueDense(
                    _workbook.ResolveDenseHandle(cell.Sheet),
                    cell.Sheet,
                    cell.Column,
                    cell.Row
                );
            }
            return 0;
        });

    /// <summary>
    /// Os OUTPUTS afetados pela edição: as células do cone dirty que nenhuma OUTRA célula dirty lê (os
    /// <em>sinks</em> do sub-DAG sujo — "as células no topo"). É o conjunto que, dado um input, você
    /// extrairia para popular outra planilha/PDF sem varrer o workbook inteiro. O(dirty × dependentes), bom
    /// para o caso comum (cone pequeno); um cone gigante (coluna quente) torna isto caro — limitação anotada.
    /// </summary>
    public List<CellDep> GetAffectedOutputs(IReadOnlyCollection<CellDep> dirty)
    {
        var dirtySet = dirty as HashSet<CellDep> ?? [.. dirty];
        var outputs = new List<CellDep>();
        var dependents = new List<CellDep>();

        foreach (var cell in dirtySet)
        {
            dependents.Clear();
            _graph.CollectDirectDependents(cell, dependents);

            var hasDirtyDependent = false;
            foreach (var dependent in dependents)
            {
                if (dirtySet.Contains(dependent))
                {
                    hasDirtyDependent = true;
                    break;
                }
            }

            if (!hasDirtyDependent)
            {
                outputs.Add(cell); // nada dirty depende dele → é um output/sink
            }
        }

        return outputs;
    }
}
