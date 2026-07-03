# Spike: referências de coluna inteira (`A:A`) / linha inteira (`1:1`) — estratégia de armazenamento

Benchmark isolado (sem tocar código de produção) para decidir como enumerar "todas as células
POPULADAS da coluna A" dado que `Sheet.Cells` é `Dictionary<string, Expression>` esparso e a
avaliação é memoizada por época (`Workbook.GetCellValue` + `InvalidateCache`/`Recalculate`).
Rodado na branch `experiment/whole-column-benchmark`.

## Setup

- Código: `benchmarks/Danfma.MySheet.Benchmark/Spike/WholeColumn/` (protótipos standalone;
  `CellAddress.Parse`/`CellId.Parse` são `internal`, então a lógica foi espelhada byte a byte em
  `KeyParse.cs` — inclusive a alocação do substring no parse da linha).
- Run: `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --filter '*WholeColumn*'`
- Ambiente: .NET 10.0.9, Arm64 RyuJIT, macOS. **`ShortRunJob` (3 iterações)** — médias com delta
  < ~15% devem ser tratadas com desconfiança; `Allocated` é determinístico mesmo em run curto.
  Linhas com StdDev alto estão sinalizadas na análise.
- Shapes (semente fixa 12345 no esparso):
  - **DenseColumn**: A1:A10000 + colunas B..F com 1.000 células cada (15.000 células).
  - **Sparse**: 10.000 células em 100 colunas × linhas aleatórias.
  - **Large**: 100.000 células em 26 colunas.
  - **Pathological**: só A1 e A100000 (o caso que tenta um "range scan por bounds").
- O lookup pontual usa 1.000 ids existentes em ordem embaralhada (semente 999) por op.

## Resultados (números reais do BenchmarkDotNet)

### 1. Agregação de coluna inteira — 1 leitura (`WholeColumnScanBenchmarks`)

| Method             | Shape       | Mean         | Ratio | Allocated |
|--------------------|-------------|-------------:|------:|----------:|
| NaiveScan_Alloc    | DenseColumn |   190.8 µs   |  1.00 | 479.568 B |
| NaiveScan_NoAlloc  | DenseColumn |    42.7 µs   |  0.22 |         – |
| LazyIndex_Build    | DenseColumn |   233.8 µs   |  1.23 | 346.136 B |
| LazyIndex_Hit      | DenseColumn |    99.8 µs   |  0.52 |         – |
| Tabular_ColumnEnum | DenseColumn |    11.3 µs   |  0.06 |         – |
| NaiveScan_Alloc    | Sparse      |   171.1 µs   |  1.00 | 319.944 B |
| NaiveScan_NoAlloc  | Sparse      |    47.1 µs   |  0.28 |         – |
| LazyIndex_Build    | Sparse      |   130.4 µs   |  0.76 | 229.392 B |
| LazyIndex_Hit      | Sparse      |     0.485 µs |  0.00 |         – |
| Tabular_ColumnEnum | Sparse      |     0.113 µs |  0.00 |         – |
| NaiveScan_Alloc    | Large       | 1.303,2 µs   |  1.00 | 3.198.128 B |
| NaiveScan_NoAlloc  | Large       |   273.0 µs   |  0.21 |         – |
| LazyIndex_Build    | Large       |   871.9 µs¹  |  0.67 | 1.712.880 B |
| LazyIndex_Hit      | Large       |    29.6 µs   |  0.02 |         – |
| Tabular_ColumnEnum | Large       |     4.4 µs   |  0.00 |         – |

¹ StdDev alto (154 µs) — GC do índice recém-construído; a ordem de grandeza é confiável.

- O parse fiel ao `CellAddress.Parse` (substring + `int.Parse`) custa **4,6×** o filtro no-alloc que
  só lê as letras da coluna. Uma implementação NaiveScan DEVE usar o parse no-alloc: 3,2 MB de
  lixo por scan no shape Large é inaceitável.
- `LazyIndex_Hit` em DenseColumn custa 99,8 µs porque o índice guarda ids e re-consulta o
  dicionário string por id (10.000 probes). Um índice que guardasse `Expression` direto seria mais
  rápido, mas idêntico em complexidade de invalidação — otimização disponível se necessário.

### 2. LazyColumnIndex sob invalidação — 100 leituras, rebuild a cada N (`WholeColumnLazyInvalidationBenchmarks`, shape Large)

| Method                     | InvalidateEvery | Mean      | Ratio | Allocated     |
|----------------------------|-----------------|----------:|------:|--------------:|
| NaiveScan_Repeated         | 1               | 26,81 ms  |  1.00 |             – |
| LazyIndex_WithInvalidation | 1               | 78,56 ms  |  2.93 | 171.288.000 B |
| NaiveScan_Repeated         | 10              | 27,47 ms  |  1.00 |             – |
| LazyIndex_WithInvalidation | 10              | 10,42 ms  |  0.38 |  17.128.800 B |
| NaiveScan_Repeated         | 100             | 41,62 ms² |  1.06 |             – |
| LazyIndex_WithInvalidation | 100             |  6,87 ms² |  0.18 |   1.712.880 B |

² Linhas N=100 com StdDev alto (12,9 ms / 1,8 ms no ShortRun); os ratios N=1 e N=10 são estáveis.

- **Break-even ≈ 4 leituras por época**: build 872 µs + hit 29,6 µs vs scan 273 µs →
  `872 / (273 − 29,6) ≈ 3,6`. Confirmado pela tabela: N=1 é 2,9× PIOR; N=10 já é 2,6× melhor.
- O índice amortiza TAMBÉM entre colunas diferentes (um build serve `A:A`, `B:B`, ...), o que o
  NaiveScan não faz — o break-even real em sheets com várias fórmulas de coluna é ainda menor.
- Invalidar a cada leitura (sheet dominada por voláteis + `Recalculate` frequente) transforma o
  índice em puro overhead + 171 MB de churn de GC. O índice PRECISA ser amarrado à época de
  cache existente, nunca reconstruído por leitura.

### 3. Lookup pontual — O HOT PATH (`WholeColumnPointLookupBenchmarks`, 1.000 lookups/op)

| Method                        | Shape       | Mean      | Ratio | Allocated |
|-------------------------------|-------------|----------:|------:|----------:|
| PointLookup_StringDict        | DenseColumn |  9,81 µs  |  1.00 |         – |
| PointLookup_Tabular_ParseSpan | DenseColumn | 15,34 µs  |  1.56 |         – |
| PointLookup_Tabular_PreParsed | DenseColumn |  6,96 µs  |  0.71 |         – |
| PointLookup_StringDict        | Sparse      |  9,87 µs  |  1.00 |         – |
| PointLookup_Tabular_ParseSpan | Sparse      | 16,95 µs  |  1.72 |         – |
| PointLookup_Tabular_PreParsed | Sparse      |  9,26 µs³ |  0.94 |         – |
| PointLookup_StringDict        | Large       | 11,90 µs³ |  1.00 |         – |
| PointLookup_Tabular_ParseSpan | Large       | 17,29 µs  |  1.46 |         – |
| PointLookup_Tabular_PreParsed | Large       |  7,98 µs  |  0.67 |         – |

³ StdDev alto (0,79 µs / 0,68 µs) — ratio pontual menos confiável, tendência clara mesmo assim.

- **Regressão do hot path com Tabular: +46% a +72%** quando a referência chega como string "A1"
  (o cenário de troca drop-in do storage: parse span→(col,row) + DUAS sondas de dicionário).
- Com coordenadas pré-parseadas o Tabular é até MAIS RÁPIDO (0,67–0,94×) — hash de int é mais
  barato que hash de string. Mas isso exige refatorar a representação de referência do engine
  inteiro (CellExpression, parser, ranges, serialização) para (col,row), não é uma troca de storage.

### 4. Linha inteira (`1:1`) — o caso "simétrico" (`WholeColumnRowEnumBenchmarks`)

| Method             | Shape       | Mean        | Ratio |
|--------------------|-------------|------------:|------:|
| RowScan_StringDict | DenseColumn | 104.582 ns  | 1.000 |
| RowEnum_Tabular    | DenseColumn |      22 ns  | 0.000 |
| RowScan_StringDict | Sparse      |  98.373 ns  | 1.001 |
| RowEnum_Tabular    | Sparse      |     242 ns  | 0.002 |
| RowScan_StringDict | Large       | 704.243 ns  | 1.000 |
| RowEnum_Tabular    | Large       |      81 ns  | 0.000 |

- **Resultado honesto: a hipótese do "problema simétrico ruim" NÃO se confirmou nestes shapes.**
  Sondar 1 `TryGetValue` por coluna populada (6–100 colunas) é ordens de grandeza mais barato que
  varrer todas as chaves string. O custo do Tabular em linha escala com o nº de colunas populadas
  (não com o total de células), então só degradaria em sheets com milhares de colunas — raro.
  O argumento contra o Tabular não é a linha inteira; é o hot path (tabela 3) e o custo de migração.

### 5. Patológico: por que bounds scan não funciona (`WholeColumnBoundsBenchmarks`, A1 + A100000)

| Method                     | Mean            | Ratio      | Allocated   |
|----------------------------|----------------:|-----------:|------------:|
| NaiveScan_PopulatedOnly    |        11,54 ns |       1.00 |           – |
| BoundsRangeScan_StringDict | 2.730.213,65 ns | 236.725,45 | 7.110.448 B |
| BoundsRangeScan_Tabular    |   168.277,31 ns |  14.590,62 |           – |

- Iterar A1..A100000 com `TryGetValue` para achar 2 células é **236 mil vezes** mais lento que
  enumerar as populadas, e aloca 7,1 MB (um id string por linha sondada). Mesmo sem alocação
  (variante tabular), ainda são 14.590×. **A semântica "células populadas" está correta; qualquer
  implementação por bounds (min/max de linha) está desqualificada para dados esparsos.**

### 6. Memória — construção de cada layout (`WholeColumnMemoryBenchmarks`, `Allocated` como proxy de footprint)

| Method                           | Shape       | Mean        | Allocated    | Alloc Ratio |
|----------------------------------|-------------|------------:|-------------:|------------:|
| Build_StringDict                 | DenseColumn |  1.944,1 µs |  1.712,24 KB |        1.00 |
| Build_Tabular                    | DenseColumn |  1.195,0 µs |  1.771,64 KB |        1.03 |
| Build_LazyIndexOnTopOfStringDict | DenseColumn |    128,2 µs |    337,94 KB |        0.20 |
| Build_StringDict                 | Sparse      |    842,9 µs |  1.178,42 KB |        1.00 |
| Build_Tabular                    | Sparse      |    424,1 µs |  1.163,52 KB |        0.99 |
| Build_LazyIndexOnTopOfStringDict | Sparse      |    149,1 µs |    224,02 KB |        0.19 |
| Build_StringDict                 | Large       | 21.345,0 µs⁴| 11.321,48 KB |        1.00 |
| Build_Tabular                    | Large       |  6.692,9 µs | 7.821,35 KB  |        0.69 |
| Build_LazyIndexOnTopOfStringDict | Large       |    729,2 µs |  1.672,73 KB |        0.15 |

⁴ StdDev 2,89 ms (ShortRun + GC Gen2); comparar pelo `Allocated`, que é determinístico.

- Tabular economiza ~31% em Large (sem chaves string), empata nos shapes menores. O LazyIndex
  custa +15–20% de memória EM CIMA do dicionário atual (ele é aditivo, não substituto).

## Análise e recomendação

**(a) O NaiveScan é suficiente?** Sim, como primeira implementação — DESDE QUE com o parse
no-alloc (só as letras da coluna; a variante fiel ao `CellAddress.Parse` é 4,6× mais lenta e aloca
3,2 MB/scan em 100k células). 273 µs por agregação de coluna em 100k células, **1× por época**
graças à memoização, zero estado novo, zero risco. Para o tamanho de sheet que o MySheet atende
hoje, isso não aparece em profile.

**(b) LazyColumnIndex paga o build em quantas leituras?** ~4 leituras de coluna por época (build
872 µs ≈ 3,2 scans; cada hit economiza ~243 µs) — e o build serve todas as colunas, então em
sheets com fórmulas `A:A`, `B:B`, `C:C` na mesma época o break-even cai para ~2 épocas de uso.
Porém com invalidação a cada leitura ele é 2,9× PIOR e gera 171 MB de churn por 100 leituras: só
faz sentido amarrado à época do cache de memoização existente (invalidar junto com
`InvalidateCache()`), nunca por leitura.

**(c) Regressão do lookup pontual no Tabular?** **+46% a +72%** no cenário de troca drop-in (id
chega como string e precisa de parse + duas sondas). Inaceitável: esse é o custo cobrado de TODA
avaliação de célula para beneficiar só as fórmulas de coluna inteira. A variante pré-parseada até
GANHA do atual (0,67–0,94×), mas exige migrar a representação de referências do engine para
(col,row) — um projeto próprio, não um spike.

**(d) Recomendação final:**

1. **Implementar `A:A`/`1:1` agora com NaiveScan no-alloc** sobre `Cells` (semântica "células
   populadas"), resultado memoizado por época como qualquer célula. Simples, sem estado, sem
   quebra de schema.
2. **Evoluir para LazyColumnIndex por época SE profiling mostrar ≥4 leituras de coluna inteira por
   época** (ou várias colunas distintas). É aditivo (+15–20% de memória), cabe no mecanismo de
   invalidação existente e dá ganho de 10–600× no hit.
3. **Rejeitar TabularStorage como troca de armazenamento.** Custos benchmarkáveis: regressão de
   46–72% no hot path universal. Custos não-benchmarkáveis: **quebra do schema MemoryPack de
   `Sheet`** (`Cells` é `Dictionary<string, Expression>` serializado — trocar o tipo invalida todo
   arquivo persistido e exige versionamento/migração) e uma migração invasiva de parser, ranges e
   API pública. O único ponto a favor real (lookup pré-parseado 0,67× + 31% menos memória em
   100k células) só se materializa com a refatoração completa de referências — se um dia o perfil
   justificar, tratar como projeto separado, não como efeito colateral do suporte a `A:A`.
   Nota honesta: o argumento clássico "coluna-major penaliza linha inteira" NÃO se sustentou nos
   shapes testados (tabela 4) — não usar esse argumento contra o Tabular no futuro.

**Disclaimers**: `ShortRunJob` (3 iterações, 1 launch); linhas marcadas ¹–⁴ têm variância alta e
merecem re-run com job completo se a decisão ficar apertada — nenhuma decisão acima depende de
delta < 2×. Ids de lookup em ordem aleatória (semente fixa) para nenhum layout se beneficiar de
localidade sequencial de cache.
