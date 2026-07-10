using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.DirtyGraph;

// See the structural-layers invariant note atop SheetStructuralIndex.cs: this graph is layer 2 (a built
// snapshot, staled by Sheet.StructuralVersion / Workbook.NamesVersion), distinct from the write-maintained
// SheetStructuralIndex (layer 1) and the epoch-scoped SheetValueStore pages (layer 3).

/// <summary>
/// O grafo de dependências REVERSA de um workbook: dado uma célula editada, quais fórmulas dependem dela
/// (direta e transitivamente). Construído varrendo cada fórmula uma vez com o
/// <see cref="DependencyExtractor"/> e invertendo as arestas.
///
/// <para><b>A mitigação de memória (o risco da Fase 2).</b> Dependências de RANGE NÃO são expandidas por
/// célula — isso explodiria (<c>=SUM(A:A)</c> geraria 100k arestas). Cada fórmula-range guarda UMA aresta
/// <c>(retângulo/coluna-aberta, fórmula)</c>. Logo o total de arestas de range é O(nº de fórmulas-range),
/// não O(células cobertas).</para>
///
/// <para><b>Índice de contenção bucketizado por coluna (a consulta).</b> "Quem depende de (col,row)?" não
/// pode varrer todas as arestas-de-range (centenas de milhares) por ponto. Uma aresta-de-range com eixo de
/// coluna FECHADO <c>[cMin,cMax]</c> é indexada em cada bucket de coluna do span (largura pequena — nunca a
/// altura), e a consulta testa só a contenção de LINHA nesse bucket. Ranges com eixo de coluna ABERTO
/// (linha inteira <c>1:5</c>) vão para um fallback por-sheet varrido por consulta (raros). Um span de coluna
/// patologicamente largo (> <see cref="WideColumnCap"/>) também cai no fallback, evitando blowup de buckets.</para>
///
/// <para>Fórmulas <see cref="DependencyScan.AlwaysDirty"/> (OFFSET/INDIRECT/voláteis/custom) ficam num
/// conjunto à parte que a marcação dirty (Fase 3) sempre inclui. Não thread-safe: usado na fase de edição.</para>
/// </summary>
internal sealed class ReverseDependencyGraph
{
    // Acima deste span de colunas, uma aresta-de-range vai para o fallback em vez de indexar N buckets.
    private const int WideColumnCap = 256;

    // fonte (célula) → fórmulas que a leem diretamente.
    private readonly Dictionary<CellDep, List<CellDep>> _cellDependents = new();

    // (sheet, coluna) → arestas-de-range de eixo-coluna FECHADO que cobrem essa coluna (testa linha na query).
    private readonly Dictionary<(string Sheet, int Column), List<RangeEdge>> _columnBuckets = new();

    // sheet → arestas-de-range de eixo-coluna ABERTO (ou span largo demais): varridas por consulta.
    private readonly Dictionary<string, List<RangeEdge>> _openColumnRanges = new(
        StringComparer.Ordinal
    );

    // Fórmulas cujas dependências não são enumeráveis (sempre recomputam).
    private readonly HashSet<CellDep> _alwaysDirty = [];

    // Contadores para diagnóstico.
    private int _rangeDepCount; // arestas-de-range lógicas (uma por fórmula-range)
    private int _bucketEntries; // entradas físicas nos buckets + fallback (o que ocupa memória)
    private int _wholeColumnDeps; // deps de range com AMBOS os limites de linha abertos (coluna inteira)

    // Fontes "QUENTES" (fan-in direto acima do limiar): editar/alcançar uma explode o cone. Precomputadas no
    // build para o PREDITOR barato — a análise desiste ao tocar numa SEM varrer seu bucket gigante.
    private const int HotFanInThreshold = 10_000;
    private readonly HashSet<(string Sheet, int Column)> _hotColumns = new();

    private readonly record struct RangeEdge(RangeDep Range, CellDep Formula);

    /// <summary>Constrói o grafo reverso varrendo todas as fórmulas do workbook.</summary>
    public static ReverseDependencyGraph Build(Workbook workbook)
    {
        var graph = new ReverseDependencyGraph();

        foreach (var sheet in workbook.Sheets.Values)
        {
            var sheetName = sheet.Name;
            foreach (var (id, expression) in sheet)
            {
                if (expression is ValueExpression)
                {
                    continue; // literais não leem nada
                }

                if (!CellAddress.TryGetColumnRow(id, out var column, out var row))
                {
                    continue; // id não-A1 (host edge case)
                }

                var formula = new CellDep(sheetName, column, row);
                graph.AddFormula(formula, DependencyExtractor.Extract(expression, workbook));
            }
        }

        graph.ComputeHotSources();
        return graph;
    }

    // Marca como quente cada coluna cujo bucket de arestas-de-range excede o limiar: editar uma célula ali
    // dispara todas essas fórmulas de uma vez (fan-in enorme). Uma passada barata sobre os buckets.
    private void ComputeHotSources()
    {
        foreach (var (key, bucket) in _columnBuckets)
        {
            if (bucket.Count > HotFanInThreshold)
            {
                _hotColumns.Add(key);
            }
        }
    }

    // Uma célula é "fonte quente" se está numa coluna quente (muitas fórmulas-range a leem) OU tem muitas
    // arestas de célula diretas. Editá-la, ou o cone alcançá-la, garante um impacto enorme.
    private bool IsHotSource(CellDep cell) =>
        _hotColumns.Contains((cell.Sheet, cell.Column))
        || (_cellDependents.TryGetValue(cell, out var deps) && deps.Count > HotFanInThreshold);

    private void AddFormula(CellDep formula, DependencyScan scan)
    {
        foreach (var source in scan.Cells)
        {
            if (!_cellDependents.TryGetValue(source, out var list))
            {
                _cellDependents[source] = list = [];
            }
            list.Add(formula);
        }

        foreach (var range in scan.Ranges)
        {
            _rangeDepCount++;
            if (range.RowMin is null && range.RowMax is null)
            {
                _wholeColumnDeps++; // coluna inteira: qualquer linha da coluna casa (não podável por linha)
            }
            var edge = new RangeEdge(range, formula);

            // Eixo de coluna fechado e não-largo: indexa por cada coluna do span (testa linha na consulta).
            if (range.ColMin is { } cMin && range.ColMax is { } cMax && cMax - cMin < WideColumnCap)
            {
                for (var col = cMin; col <= cMax; col++)
                {
                    var key = (range.Sheet, col);
                    if (!_columnBuckets.TryGetValue(key, out var bucket))
                    {
                        _columnBuckets[key] = bucket = [];
                    }
                    bucket.Add(edge);
                    _bucketEntries++;
                }
            }
            else
            {
                // Eixo de coluna aberto (linha inteira) ou span largo demais → fallback por-sheet.
                if (!_openColumnRanges.TryGetValue(range.Sheet, out var open))
                {
                    _openColumnRanges[range.Sheet] = open = [];
                }
                open.Add(edge);
                _bucketEntries++;
            }
        }

        if (scan.AlwaysDirty)
        {
            _alwaysDirty.Add(formula);
        }
    }

    /// <summary>As fórmulas sempre-dirty (OFFSET/INDIRECT/voláteis/custom) — recomputam em toda passada.</summary>
    public IReadOnlyCollection<CellDep> AlwaysDirty => _alwaysDirty;

    /// <summary>Adiciona a <paramref name="into"/> as fórmulas que dependem DIRETAMENTE de
    /// <paramref name="source"/> (arestas de célula + fórmulas-range cujo range contém a célula).</summary>
    public void CollectDirectDependents(CellDep source, ICollection<CellDep> into)
    {
        if (_cellDependents.TryGetValue(source, out var direct))
        {
            foreach (var dependent in direct)
            {
                into.Add(dependent);
            }
        }

        // Bucket da coluna do ponto: a coluna já casa, então só testa a linha.
        if (_columnBuckets.TryGetValue((source.Sheet, source.Column), out var bucket))
        {
            foreach (var edge in bucket)
            {
                if (RowContains(edge.Range, source.Row))
                {
                    into.Add(edge.Formula);
                }
            }
        }

        // Fallback de eixo-coluna aberto: testa a contenção completa.
        if (_openColumnRanges.TryGetValue(source.Sheet, out var open))
        {
            foreach (var edge in open)
            {
                if (Contains(edge.Range, source.Column, source.Row))
                {
                    into.Add(edge.Formula);
                }
            }
        }
    }

    /// <summary>
    /// O fecho transitivo dos dependentes de <paramref name="sources"/> — todas as fórmulas afetadas por
    /// editar aquelas células. NÃO inclui as fontes nem as sempre-dirty (a Fase 3 une as duas). Guarda ciclos.
    ///
    /// <para><b>Fixpoint, não BFS por-célula.</b> Uma BFS que consulta os buckets por célula da fronteira fica
    /// quadrática numa coluna com fan-in gigante (editar 1 célula lá dispara centenas de milhares de
    /// fórmulas-range, cada uma re-varrendo o bucket). Em vez disso: propaga as arestas de CÉLULA por BFS
    /// (barato, lookup de dicionário) e varre as arestas de RANGE UMA VEZ por passe externo, disparando uma
    /// fórmula-range se seu range contém QUALQUER célula dirty — via um índice de linhas-dirty por coluna
    /// (<see cref="SortedSet{T}"/> com <c>GetViewBetween</c>). Custo O(arestas × profundidade), limitado.</para>
    /// </summary>
    public HashSet<CellDep> GetAllDependents(IEnumerable<CellDep> sources)
    {
        var affected = new HashSet<CellDep>();
        var dirtyRows = new Dictionary<(string Sheet, int Column), SortedSet<int>>();
        var frontier = new Queue<CellDep>();

        void MarkDirty(CellDep cell)
        {
            var key = (cell.Sheet, cell.Column);
            if (!dirtyRows.TryGetValue(key, out var rows))
            {
                dirtyRows[key] = rows = [];
            }
            if (rows.Add(cell.Row))
            {
                frontier.Enqueue(cell); // célula genuinamente nova → propaga
            }
        }

        foreach (var source in sources)
        {
            MarkDirty(source);
        }

        var rangeChanged = true;
        while (rangeChanged)
        {
            // (A) Propaga as arestas de CÉLULA por BFS (cada dependente é um lookup O(1)).
            while (frontier.Count > 0)
            {
                var current = frontier.Dequeue();
                if (_cellDependents.TryGetValue(current, out var deps))
                {
                    foreach (var dependent in deps)
                    {
                        if (affected.Add(dependent))
                        {
                            MarkDirty(dependent);
                        }
                    }
                }
            }

            // (B) Varre as arestas de RANGE UMA VEZ: dispara as que contêm alguma célula dirty. Fórmulas
            // recém-disparadas voltam para a fronteira (A) do próximo passe. Converge em passes = profundidade
            // das cadeias que passam por ranges.
            rangeChanged = false;
            foreach (var bucket in _columnBuckets.Values)
            {
                rangeChanged |= FireRanges(bucket, dirtyRows, affected, MarkDirty);
            }
            foreach (var open in _openColumnRanges.Values)
            {
                rangeChanged |= FireRanges(open, dirtyRows, affected, MarkDirty);
            }
        }

        return affected;
    }

    private static bool FireRanges(
        List<RangeEdge> edges,
        Dictionary<(string Sheet, int Column), SortedSet<int>> dirtyRows,
        HashSet<CellDep> affected,
        Action<CellDep> markDirty
    )
    {
        var fired = false;
        foreach (var edge in edges)
        {
            if (!affected.Contains(edge.Formula) && RangeHasDirty(edge.Range, dirtyRows))
            {
                affected.Add(edge.Formula);
                markDirty(edge.Formula);
                fired = true;
            }
        }
        return fired;
    }

    // O range contém alguma célula dirty? Usa o índice de linhas-dirty por coluna.
    private static bool RangeHasDirty(
        in RangeDep range,
        Dictionary<(string Sheet, int Column), SortedSet<int>> dirtyRows
    )
    {
        // Eixo de coluna fechado: itera só as colunas do span (largura pequena).
        if (range.ColMin is { } cMin && range.ColMax is { } cMax)
        {
            for (var col = cMin; col <= cMax; col++)
            {
                if (dirtyRows.TryGetValue((range.Sheet, col), out var rows) && RowsHit(rows, range))
                {
                    return true;
                }
            }
            return false;
        }

        // Eixo de coluna aberto (linha inteira / span largo): varre as colunas dirty da sheet (raro).
        foreach (var ((sheet, col), rows) in dirtyRows)
        {
            if (
                string.Equals(sheet, range.Sheet, StringComparison.OrdinalIgnoreCase)
                && (range.ColMin is not { } lo || col >= lo)
                && (range.ColMax is not { } hi || col <= hi)
                && RowsHit(rows, range)
            )
            {
                return true;
            }
        }
        return false;
    }

    private static bool RowsHit(SortedSet<int> dirtyRows, in RangeDep range)
    {
        if (range.RowMin is null && range.RowMax is null)
        {
            return dirtyRows.Count > 0; // coluna inteira: qualquer linha dirty da coluna casa
        }
        return dirtyRows
                .GetViewBetween(range.RowMin ?? int.MinValue, range.RowMax ?? int.MaxValue)
                .Count > 0;
    }

    /// <summary>
    /// PREDITOR + enumeração num só passo, para o evict-and-pull. BFS por-célula via os buckets, output-
    /// sensitive (custo acompanha o cone — rápido no caso comum). Decide FULL de forma BARATA por duas vias:
    /// (1) ao alcançar uma <em>fonte quente</em> precomputada, desiste NA HORA sem varrer seu bucket gigante;
    /// (2) se o cone acumulado excede <paramref name="cap"/>. Retorna a <see cref="ImpactEstimate"/> (o
    /// veredito + tamanho + razão) e, quando parcial, o CONE (para o caller evictar) — senão <c>null</c>.
    /// </summary>
    public (ImpactEstimate Estimate, HashSet<CellDep>? Cone) Analyze(
        IEnumerable<CellDep> sources,
        int cap
    )
    {
        var affected = new HashSet<CellDep>();
        var frontier = new Queue<CellDep>();
        foreach (var source in sources)
        {
            frontier.Enqueue(source);
        }

        var next = new List<CellDep>();
        while (frontier.Count > 0)
        {
            var current = frontier.Dequeue();

            // Barato: se a célula da fronteira é uma fonte quente, o impacto é enorme → full, SEM varrer o
            // bucket dela (o ganho vs. o preditor reativo, que varria as ~490k arestas antes de desistir).
            if (IsHotSource(current))
            {
                return (
                    ImpactEstimate.Full(
                        $"o cone alcança uma fonte quente ({current.Sheet}!col{current.Column})"
                    ),
                    null
                );
            }

            next.Clear();
            CollectDirectDependents(current, next);

            foreach (var dependent in next)
            {
                if (affected.Add(dependent))
                {
                    if (affected.Count > cap)
                    {
                        return (ImpactEstimate.Full($"cone acima do cap ({cap} células)"), null);
                    }
                    frontier.Enqueue(dependent);
                }
            }
        }

        return (ImpactEstimate.Partial(affected.Count), affected);
    }

    private static bool RowContains(in RangeDep range, int row) =>
        (range.RowMin is not { } rMin || row >= rMin)
        && (range.RowMax is not { } rMax || row <= rMax);

    private static bool Contains(in RangeDep range, int column, int row) =>
        (range.ColMin is not { } cMin || column >= cMin)
        && (range.ColMax is not { } cMax || column <= cMax)
        && RowContains(range, row);

    // === Diagnóstico (footprint) ========================================================================

    /// <summary>Números de footprint do grafo, para o probe de memória da Fase 2.</summary>
    public GraphDiagnostics Diagnostics()
    {
        var cellEdges = 0;
        foreach (var list in _cellDependents.Values)
        {
            cellEdges += list.Count;
        }

        var maxBucket = 0;
        foreach (var bucket in _columnBuckets.Values)
        {
            if (bucket.Count > maxBucket)
            {
                maxBucket = bucket.Count;
            }
        }

        // Estimativa grosseira (piso indicativo, não medição de heap): ~40B por aresta física.
        var bytes =
            (long)cellEdges * 40 + (long)_bucketEntries * 40 + (long)_cellDependents.Count * 32;

        return new GraphDiagnostics(
            _cellDependents.Count,
            cellEdges,
            _rangeDepCount,
            _bucketEntries,
            _columnBuckets.Count,
            maxBucket,
            _wholeColumnDeps,
            _alwaysDirty.Count,
            bytes
        );
    }
}

/// <summary>Footprint do grafo reverso.</summary>
internal readonly record struct GraphDiagnostics(
    int DistinctSourceCells,
    int CellEdges,
    int RangeDeps,
    int RangeBucketEntries,
    int ColumnBuckets,
    int MaxColumnBucket,
    int WholeColumnDeps,
    int AlwaysDirtyFormulas,
    long EstimatedBytes
);
