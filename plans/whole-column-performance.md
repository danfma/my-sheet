# Performance de coluna inteira em escala: índice estrutural + range caches por época

Cenário real (usuário, 2026-07-03): sheet com ~506.082 células numa coluna + ~399.776 fórmulas de coluna
inteira (`MATCH/VLOOKUP/XLOOKUP/INDEX`, `SUMIF/COUNTIF/AVERAGEIF`, `IF/ROW/SMALL`) **referenciando a coluna
grande**, ciclo carrega-1×-lê-1× → **~57min só de varreduras**. Meta: colapsar para **segundos**, mantendo
memória enxuta (o foco segue: bater ClosedXML/Aspose em performance E memória). Release: saiu como **2.6.2** (patch — commits perf/test/docs, sem feat: não há API nova; semver correto).

## Diagnóstico (por que o NaiveScan implodiu — lição registrada)
O spike (`plans/whole-column-spike.md`) raciocinou "273 µs/scan, 1× por época via memoização" — verdade POR
FÓRMULA; a memoização é por CÉLULA. Com F=400k fórmulas, cada uma varre N=506k chaves: **O(F×N) ≈ 2×10¹¹**
visitas ≈ 57min (8,5ms/fórmula ≈ 2-3 scans). O spike não modelou multiplicidade de fórmulas — o break-even
do índice ("≥4 leituras/época") foi estourado por 5 ordens de magnitude. E como as fórmulas CONSOMEM a
coluna grande, há um segundo O(F×N) de leitura de valores que nenhum índice estrutural resolve — só caching
semântico por época (a técnica do lookup cache do próprio Excel 2016+ e das engines rápidas).

## Decisões de design (traadas nesta análise)
- **Camada 1 — índice estrutural LAZY read-time** (não write-time): por sheet, `coluna → lista de ids
  ordenada por linha` (+ simétrico de linha sob demanda), construído na 1ª enumeração open-range (1 passada
  O(N)), descartado no `InvalidateCache()`. MESMO contrato de segurança do cache de valores (quem edita sem
  invalidar já vê valores velhos hoje) → **zero encapsulamento de `Sheet.Cells`, zero breaking**. O índice
  write-time do apêndice de `plans/whole-column-references.md` segue adiado (gatilho: grafo de deps).
  Sobrevive ao `Recalculate()` (estrutura ≠ valores).
- **Camada 2 — range caches por época** (internos ao Workbook, sem API nova): keyed pelo record do range
  (`OpenRangeReference`/`RangeReference` têm value-equality — chave de dicionário perfeita) + sheet:
  1. **Snapshot materializado**: os `ComputedValue` populados do range, 1 array, construído 1× via Camada 1
     + `GetCellValue` (memoizado). ~12MB por 506k células — transiente.
  2. **Derivados lazy sob demanda** (cada um só se a função pedir):
     - *hash exato* `valor → primeira posição` (semântica de igualdade do `ValueCoercion.AreEqual`:
       case-insensitive p/ texto, tipos distintos nunca iguais) → MATCH tipo 0/XLOOKUP exato/XMATCH/
       VLOOKUP-FALSE em O(1).
     - *snapshot ordenado* (cópia com posições originais, ordem do `ValueCoercion.Compare`) → MATCH tipo
       1/-1 e VLOOKUP-TRUE por busca binária O(log N); `SMALL/LARGE` por indexação direta; `MEDIAN/
       PERCENTILE/QUARTILE` reutilizam.
     - *mapas de critério de igualdade* `valor → (soma, contagem)` → SUMIF/COUNTIF/AVERAGEIF com critério
       `=` em O(1). (Stretch opcional: prefix-sums sobre o ordenado p/ critérios `>`/`>=`/`<`/`<=` em
       O(log N).) Critérios wildcard/regexy → fallback linear SOBRE O SNAPSHOT (sem re-scan de chaves).
  3. **Memo de agregado idêntico** `(tipo-função, range) → resultado` p/ SUM/COUNT/MAX/MIN/COUNTA de range
     puro repetido em muitas células.
- **Invalidação dos caches de valor (2)**: descartados em `InvalidateCache()` E `Recalculate()` (podem
  conter valores volatile-tainted — regra segura e simples). Camada 1 só no `InvalidateCache()`.
- **Semântica INALTERADA é lei**: propagação de erros, cross-type compare, wildcards, #REF! estrutural de
  sheet-fantasma, política per-função de erro — tudo idêntico. A prova é a suíte existente (795+23) verde
  + testes de equivalência (mesmas fórmulas com cache frio × quente → resultados idênticos).
- **Anti-ClosedXML**: nada disso é API nova nem modelo novo — é cache interno, bounded, descartável. O
  modelo esparso e o hot path de célula única ficam intocados.
- **Thread-safety**: caches em `ConcurrentDictionary` + criação lazy via `Interlocked.CompareExchange`
  (lição do MemoryPack: field initializers são bypassados no deserialize — nunca `= new()` em campo de
  cache do Workbook). Campos `[MemoryPackIgnore]`.

## Estimativa (a validar na Fase 0/3)
Build índice: 1×O(N) ≈ 50-100ms. Snapshot: 506k `GetCellValue` ≈ 0,5-1s (1ª vez; já era pago hoje).
Ordenação: ~100-200ms. 400k lookups O(1)/O(log N) ≈ 0,1-2s. **Total esperado: < 10s (era ~57min).**
Memória de pico: snapshot+ordenado+hash ≈ 30-60MB transiente para 506k — documentar medição real.

## For Future Agents
Marque `- [x]`; ao fechar fase: Status `Complete` + Phase Summary + Verification. TDD. **Verificação SEMPRE
`--no-incremental`**. Fixture `workbook-pre-namespaces.msgpack.bin`/`MemoryPackCompatibilityTests`
intocáveis. Suítes hoje: core **795**, Excel **23**; 0 warnings. Commits inglês, semantic
(`perf(refs): ...`), SEM atribuição a IA. NÃO push (gates do usuário).

---

## Phase 0: Repro sintético + baseline (o benchmark que representa o cenário do usuário)
Status: Complete

- [x] Benchmark `Spike/WholeColumnScale/` no projeto de benchmark: gerador de workbook sintético com uma
      coluna de dados de ~500k células (números ordenados p/ o caso aproximado + strings p/ exato) e blocos
      de fórmulas de coluna inteira parametrizáveis: `MATCH(x_i, A:A)` (tipo 1 e 0), `VLOOKUP(x_i, A:B, 2,
      FALSE)`, `SUMIF(A:A, "=k_i")`, `COUNTIF(A:A, k_i)`, `SMALL(A:A, k_i)`, `SUM(A:A)` repetido.
- [x] Medição em DUAS escalas: reduzida (50k células × 10k fórmulas — roda em segundos, vira teste de
      regressão de perf) e plena (500k × 100k+ — execução manual, documentada). Registrar baseline ATUAL
      (deve exibir o O(F×N)).

### Verification Plan
- Baseline plena reproduz a ordem de grandeza do relato (minutos, extrapolável a ~57min p/ 400k fórmulas).

### Phase Summary
Repro em `benchmarks/Danfma.MySheet.Benchmark/Spike/WholeColumnScale/`: gerador sintético determinístico
(`WholeColumnScaleData` — coluna A ordenada 1..N, coluna B = valor, coluna C = texto `k{r}`, mais uma
tabela ESTREITA K/L/M de 16 células na MESMA sheet), classe BenchmarkDotNet (`WholeColumnScaleBenchmarks`,
ShortRunJob+MemoryDiagnoser, guarda de regressão) e um harness console 1-shot
(`--whole-column-scale [--full]`) para wall-clock, já que a baseline plena é inviável em modo bench.
Cada linha = `InvalidateCache()` + 1 passada completa (o ciclo carrega-1×-lê-1× do relato).

**Baseline REDUZIDA (50k células de dados ≈ 150k células na sheet × 10k fórmulas), pré-índice
(commit `bb50a7c`), Apple Silicon 10 cores, wall-clock ms (duas execuções independentes¹):**

| Fórmula      | BigColumn (ms)    | NarrowColumn (ms) |
|--------------|------------------:|------------------:|
| Match1       | 106.852 / 83.322  | 15.588 / 15.118   |
| Match0       |  96.897 / 83.914  | 15.864 / 15.362   |
| VLookupExact |  23.149 / 22.613  | 11.422 / 11.416   |
| SumIfEqual   |  82.259 / 97.750  | 15.476 / 15.426   |
| CountIfEqual |  83.185 / 100.666 | 15.564 / 15.339   |
| Small        |  81.640 / 95.008  | 15.480 / 15.431   |
| SumRepeated  |  81.558 / 82.114  | 16.500 / 15.728   |

¹ A segunda execução rodou dentro do run `--full`; variância inter-run ~±15%.

O O(F×N) está exposto dos DOIS lados: até a coluna ESTREITA (16 células) custa ~15s por bloco de 10k
fórmulas, porque cada fórmula varre TODAS as ~150k chaves da sheet para achar as 16 da coluna K.

**Baseline PLENA (amostrada): 500k células de dados (~1,5M células na sheet) × 1.000 fórmulas medidas por
tipo, extrapolação linear ×100 → 100k fórmulas** (linear é válido: custo por fórmula constante para N fixo):

| Fórmula      | Big amostra (ms) | Big est. 100k (s) | Narrow amostra (ms) | Narrow est. 100k (s) |
|--------------|-----------------:|------------------:|--------------------:|---------------------:|
| Match1       | 151.404          | 15.140 (~4,2h)    | 15.985              | 1.599                |
| Match0       | 149.185          | 14.918            | 16.150              | 1.615                |
| VLookupExact |  13.588          |  1.359            | 12.572              | 1.257                |
| SumIfEqual   | 156.810          | 15.681            | 16.224              | 1.622                |
| CountIfEqual | 149.700          | 14.970            | 15.547              | 1.555                |
| Small        | 151.385          | 15.139            | 16.170              | 1.617                |
| SumRepeated  | 147.789          | 14.779            | 16.329              | 1.633                |

**~4,1–4,4h por bloco de 100k fórmulas** sobre a coluna grande (VLOOKUP ~23min pelo early-exit do scan da
tabela) — e até a coluna ESTREITA de 16 células custa **~21–27min**, só de varredura de 1,5M chaves por
fórmula. Ordem de grandeza do relato CONFIRMADA: em produto F×N, o relato (400k × 506k = 2,0×10¹¹ visitas
≈ 57min) e o repro (reduzido: 10k × 150k = 1,5×10⁹ ≈ 82–107s; pleno idem em proporção) implicam custo por
visita consistente dentro de ~2-4× (o repro é mais pesado por célula: a sheet sintética carrega 3 colunas
de dados e cada fórmula materializa a lista de 500k valores).

---

## Phase 1: Camada 1 — índice estrutural lazy por sheet
Status: Complete

- [x] Cache estrutural no `Workbook` (`[MemoryPackIgnore]`, lazy Interlocked): por sheet, `Dictionary<int,
      List<string>>` coluna→ids ordenados por linha (+ linha→ids sob demanda, construção independente).
      Invalidação: `InvalidateCache()` descarta; `Recalculate()` NÃO.
- [x] `OpenRangeReference.PopulatedIds`/`TryGetPopulatedBounds`/`Expand`/`ExpandComputedValues` consomem o
      índice (colunas do range → concat das listas) em vez de varrer `Cells.Keys`. Ordem de enumeração
      preservada (col, depois row) — os testes existentes garantem.
- [x] `Subtotal`/demais scans de open-range migram pro índice.

### Verification Plan
- Suíte completa verde (795+23) `--no-incremental` 0 warnings; fixture verde.
- Benchmark reduzido: caso "colunas pequenas em sheet grande" colapsa (o scan de chaves some).

### Phase Summary
`SheetStructuralIndex` (novo, interno): por sheet, `coluna → List<string>` de ids ordenados por linha e o
simétrico `linha → ids` ordenados por coluna, cada mapa construído LAZY e INDEPENDENTE (contadores internos
de build, visíveis aos testes, provam: 1 build por época, reuso entre leituras, whole-row não constrói o
mapa de colunas e vice-versa). Ancorado no `Workbook` via `ConcurrentDictionary<string (sheet, case-insens),
SheetStructuralIndex>` `[MemoryPackIgnore]`, criado lazy com `Interlocked.CompareExchange` (lição MemoryPack:
nunca `= new()` em campo de cache). `InvalidateCache()` descarta (`_structuralIndex?.Clear()`);
`Recalculate()` NÃO toca (estrutura ≠ valores). Consumidores: `PopulatedIds` enumera só as colunas cobertas
(range fechado de colunas → loop direto; lado aberto → chaves do mapa filtradas+ordenadas), com limites de
linha resolvidos por BUSCA BINÁRIA na lista ordenada (nenhum parse por id no caminho quente); whole-row puro
usa o mapa de linhas. `TryGetPopulatedBounds` (a resolução de tabela do VLOOKUP/INDEX/OFFSET, chamada A CADA
avaliação) lê a bounding-box do primeiro/último id de cada lista coberta — O(colunas·log n) por chamada, sem
enumerar ids. `Expand`/`ExpandComputedValues`/`Subtotal` herdam via `PopulatedIds`. Guards de sheet-fantasma
(yield break/false, sem throw) intactos; ordem (coluna, depois linha) agora DETERMINÍSTICA (independe da
ordem de inserção — antes seguia a ordem do dicionário).

**Pós-índice, benchmark reduzido (50k×10k, harness, máquina ociosa, commit `557fcfe`):**

| Fórmula      | BigColumn (ms) | vs baseline | NarrowColumn (ms) | vs baseline |
|--------------|---------------:|------------:|------------------:|------------:|
| Match1       | 106.962        | ~1×         | **59,5**          | **~260×**   |
| Match0       | 95.455         | ~1×         | **68,2**          | **~230×**   |
| VLookupExact | **9.895**      | **~2,3×**   | **43,8**          | **~260×**   |
| SumIfEqual   | 91.739         | ~1×         | **32,8**          | **~470×**   |
| CountIfEqual | 83.163         | ~1×         | **40,0**          | **~390×**   |
| Small        | 77.056         | ~1×         | **28,8**          | **~540×**   |
| SumRepeated  | 66.283         | ~1,2×       | **22,4**          | **~740×**   |

O caso "colunas pequenas em sheet grande" COLAPSOU (15,5s → 22–68ms por bloco: o scan de chaves sumiu; o
que resta é o consumo real das 16 células × 10k fórmulas). O caso "consome a coluna grande" segue lento —
ESPERADO: o custo restante é a LEITURA dos 50k valores por fórmula (o segundo O(F×N), de valores), que é
exatamente o alvo da Camada 2/Fase 2. Bônus: VLOOKUP exato sobre a coluna grande caiu 22,6s → 9,9s porque a
resolução de bounds por fórmula deixou de varrer chaves. **Testes novos:** `StructuralIndexTests` (6) —
construção única/reuso, descarte por `InvalidateCache`, sobrevivência a `Recalculate` (mesma instância, sem
rebuild), ordem coluna-depois-linha independente de inserção, e independência whole-row × whole-column.
**Build:** `--no-incremental` 0 warnings. **Suíte:** core **801** (795 + 6), Excel **23**,
`MemoryPackCompatibilityTests`/fixture intocados e verdes.

---

## Phase 2: Camada 2 — range caches por época (snapshot + derivados)
Status: Complete

- [x] `RangeValueCache` interno no `Workbook` (`ConcurrentDictionary<Reference, RangeSnapshot>`,
      `[MemoryPackIgnore]`, lazy via `Interlocked.CompareExchange`): snapshot materializado + derivados lazy
      (`Lazy<T>`: hash exato; ordenado c/ posições + prefix/suffix-max; mapa de igualdade numérica p/ critérios;
      view de números ordenados + primeiro erro; memo de agregado). Descartado em `InvalidateCache()` E
      `Recalculate()`.
- [x] Ligar nos consumidores (sem mudar semântica; fallback linear sobre o snapshot p/ casos não-indexáveis):
      - `Match` (exato→hash; aproximado tipo 1/-1→prefix/suffix-max; blank-equiv/erro→linear).
      - `XMatch`/`XLookup` (exato forward→hash; resto→`LookupMatching` sobre o array cacheado).
      - `VLookup`/`HLookup` (FALSE→hash da coluna/linha-chave; TRUE→prefix-max); `Lookup` (array cacheado).
      - `SumIf`/`CountIf`/`AverageIf` single-range (igualdade numérica→mapa; resto→linear no snapshot).
      - `Small`/`Large`/`Median`/`Percentile`/`Quartile`/`PercentRank`/`TrimMean` (números ordenados compart.).
      - `Sum`/`Count`/`CountA`/`Max`/`Min`/`Average` de range puro → memo de agregado.
- [x] Testes de equivalência: harness diferencial (`RangeValueCacheEquivalenceTests`) — 7 cenários × 43 fórmulas
      = 301 casos, cada um avaliado BYPASS × FRIO × QUENTE e comparado resultado a resultado; + casos dirigidos
      (aproximado c/ duplicatas → última posição; erro no meio propaga; threshold engaja/pula).
- [x] Stretch (prefix-sums p/ `>`/`<`): NÃO feito — critérios de comparação usam fallback linear sobre o
      snapshot (já elimina o re-scan de valores; suficiente p/ o alvo). Adiado.

### Verification Plan
- Suíte completa verde `--no-incremental` 0 warnings; fixture verde; equivalência frio×quente 100%.
- Benchmark reduzido: MATCH/SUMIF/SMALL sobre coluna grande colapsam de O(F×N) para O(N + F·log N).

### Phase Summary
`RangeSnapshot` (novo, `Danfma.MySheet/RangeValueCache.cs`, interno): materializa 1× o array de `ComputedValue`
de um range populado (na MESMA ordem/valores que os consumidores já veem, via Camada 1 + `GetCellValue`), com
derivados LAZY compartilhados por toda a época: (1) **hash exato** por tipo (`Dictionary<double,int>` números,
`Dictionary<string,int>` OrdinalIgnoreCase texto, posições bool) → 1ª posição; (2) **índice ordenado** por
`ValueCoercion.Compare` (não-blank/não-erro) + `PrefixMaxPosition`/`SuffixMaxPosition` que reproduzem MATCH tipo
1/-1 EXATAMENTE p/ qualquer ordem de entrada (a ÚLTIMA posição entre os empatados = máx posição do prefixo/
sufixo qualificado); (3) **mapa de igualdade numérica** `número → (soma, contagem)`; (4) **view de números
ordenados** + primeiro erro (semântica `NumericAggregation.Fold`); (5) **memo de agregado** `(tipo) →
ComputedValue`. Ancorado no `Workbook` via `ConcurrentDictionary<Reference, RangeSnapshot>` `[MemoryPackIgnore]`,
criado lazy com `Interlocked.CompareExchange`; descartado em `InvalidateCache()` E `Recalculate()` (valores
podem ser volatile-tainted). Limiar de **256 células populadas** (estimativa barata via área do retângulo ou
soma das listas do índice estrutural, SEM materializar): abaixo dele → `null` → caminho linear original. VLOOKUP/
HLOOKUP modelam a coluna/linha-chave como um `RangeReference` derivado do bounding box da tabela (posição do
snapshot = linha/coluna 1-based). Toda semântica preservada bit a bit; onde não indexável (lookup blank-equiv,
critério não-igualdade, wildcard, XLOOKUP aproximado "closest") → fallback linear SOBRE o snapshot cacheado
(elimina o re-scan de valores mesmo assim).

**Reduzido (50k×10k), harness, mesma máquina (Apple Silicon 10 cores), wall-clock ms — BigColumn:**

| Fórmula      | Antes (Fase 1) | Depois (Fase 2) | Ganho     |
|--------------|---------------:|----------------:|----------:|
| Match1       |      80.511,2  |         141,1   | **~570×** |
| Match0       |      73.571,4  |          95,3   | **~770×** |
| VLookupExact |       8.096,7  |         154,7   | **~52×**  |
| SumIfEqual   |      70.490,1  |          68,2   | **~1030×**|
| CountIfEqual |      71.437,7  |          82,3   | **~870×** |
| Small        |      59.455,5  |          74,9   | **~790×** |
| SumRepeated  |      55.167,1  |          89,0   | **~620×** |

(Baseline "Antes" = re-medição desta máquina pré-Fase-2, commit `44c0662`; a Fase 1 já havia colapsado a coluna
ESTREITA, intocada aqui — 60–130ms.) Todos os blocos BigColumn caíram para a **casa das centenas de ms**: build
do snapshot 1× (O(N)) + 10k lookups O(1)/O(log n). VLOOKUP fica mais alto (155ms) porque ainda paga 10k
`CellComputedValueAt` na coluna de retorno + o snapshot da coluna-chave.

**Equivalência:** `RangeValueCacheEquivalenceTests` — **301 casos** (7 cenários: ascendente, com duplicatas,
NÃO-ordenado, tipos mistos, COM erros, zeros/blanks, texto) × 43 fórmulas — BYPASS × FRIO × QUENTE **100%
idênticos**; + 3 testes dirigidos (duplicatas→última posição; erro propaga em SMALL/SUM/MEDIAN; threshold
engaja acima de 256 / pula abaixo).

**Memória (reduzido, 50k):** snapshot de valores ≈ 1,2MB; stack de derivados de UM range de 50k (hash 2,23 +
ordenado 1,59 + mapa numérico 2,60 + números ordenados 0,39) ≈ **6,8MB**; pico absoluto medido com TODOS os
derivados num único range = 15,6MB (inclui o cache de células + índice estrutural pré-Fase-2). Na prática cada
bloco constrói 1–2 derivados → pico real por bloco ~2–4MB. Dentro da projeção do plano (30–60MB p/ 506k).

**Decisões além do plano:** (a) **limiar 256** confirmado no benchmark (coluna estreita de 16 pula, big de 50k
engaja). (b) **Hash de igualdade exata**: só responde p/ lookup NÃO-blank-equivalente (número ≠ 0, texto ≠ "",
TRUE) e não-NaN — a regra intransitiva do blank (`blank = 0 = "" = FALSE`, mas `0 ≠ ""`) e o `NaN` do
`Dictionary<double>` fariam divergir; blank-equiv/blank/erro/NaN → `Unsupported` → linear. (c) **MATCH aproximado
com duplicatas** resolvido por prefix/suffix-max-position sobre o ordenado — reproduz "última posição dos
empatados" p/ QUALQUER ordem de entrada (não só ascendente), sem assumir dados ordenados. (d) **XLOOKUP
aproximado** NÃO acelerado (engine "closest" diverge do MATCH) — linear sobre snapshot; só o exato-forward usa
hash (com guarda `lookupArray.Count ≤ returnArray.Count`). (e) **SUMIF/AVERAGEIF com sum_range separado** ou
critério texto/wildcard/comparação → linear sobre snapshot (mapa só p/ igualdade numérica single-range); o par
(criteria,sum) na chave ficou de fora. (f) `LookupMatching.FindMatch` e `StatisticsMath.Percentile*` migraram de
`List<ComputedValue>`/`List<double>` p/ `IReadOnlyList<...>` p/ receber o array/lista cacheados sem cópia.

**Build:** `--no-incremental` 0 warnings. **Suíte:** core **805** (801 + 4 novos), Excel **23**,
`MemoryPackCompatibilityTests`/fixture intocados e verdes.

---

## Phase 3: Validação em escala plena + regressão comparativa + docs + release
Status: Complete

- [x] Rodar o benchmark PLENO (500k × 100k+): registrar tempo total e pico de memória ANTES × DEPOIS no
      relatório do plano. Meta: **< 60s na escala plena equivalente ao relato (aspiração < 10s)**; memória
      de pico dos caches documentada (~30-60MB esperados).
- [x] **Regressão comparativa (requisito do usuário, 2026-07-03)**: rodar a suíte `SheetBenchmarks`
      (MySheet × ClosedXML: Parse/Compute/Comparison/TextEquality/Conditional, com MemoryDiagnoser) na
      BRANCH e na MAIN, lado a lado. Critérios: (a) nenhuma regressão de tempo/alocação nos caminhos core
      (célula única, cache-heavy — o hot path NÃO foi tocado, deve ser ~0%); (b) **continuamos mais rápidos
      que o ClosedXML em TODAS as categorias** (tempo E memória). Registrar a tabela no plano.
- [x] **Comparativo ClosedXML no cenário de coluna inteira**: PRIMEIRO verificar a capacidade (lição do
      lessons.md: ClosedXML pode não avaliar MATCH/SUMIF/whole-column — testar num spike de 5min). Se
      suportar: medir o cenário reduzido (50k × 10k) nos dois engines e registrar; se não suportar,
      documentar a incapacidade (também é resposta: eles não competem nesse caso de uso). A memória extra
      dos nossos caches deve ser comparada com o footprint do ClosedXML no MESMO workload.
- [x] `docs/performance.md`: seção sobre os range caches por época (o que existe, quando é descartado,
      números medidos honestos). NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [x] `tasks/lessons.md`: a lição do O(F×N) (spike modelou amortização por leitura, não multiplicidade de
      fórmulas — benchmarks de estratégia devem modelar o EIXO DE CARGA COMPLETO: leituras × fórmulas).
- [x] Plano: fases Complete + Final Recap.

### Verification Plan
- Números plenos no plano; suítes verdes; 0 warnings.

### Phase Summary

**Nota de método — por que a escala plena é MEDIDA, não mais extrapolada.** O `--full` do harness (Fase 0)
media 1k fórmulas e multiplicava ×100. Isso era VÁLIDO na baseline (custo verdadeiramente O(F×N): cada
fórmula re-varria N, então o custo por fórmula era constante). Com os caches o custo virou O(N + F·log N):
o build do snapshot é O(N) UMA vez, amortizado por TODAS as fórmulas do bloco. Amostrar 1k e multiplicar
×100 multiplica esse build único ×100 e SUPERESTIMA grosseiramente (o `--full --sampled` chega a "60s" que
são quase todo build). Portanto o número honesto pós-cache tem de ser MEDIDO sobre as 100k fórmulas reais —
o `--full` agora faz exatamente isso (`WholeColumnScaleHarness.RunFullMeasured`, com delta de heap por bloco
+ wall-clock total; `--full --sampled` ainda imprime a extrapolação, só para documentar a armadilha).

**Escala PLENA MEDIDA (500k células de dados ≈ 1,5M na sheet × 100k fórmulas REAIS por bloco), árvore atual
(Camadas 1+2), Apple Silicon 10 cores, wall-clock ms do bloco:**

| Fórmula      | Baseline Big est. 100k (s) | Big AGORA (ms) | Ganho Big | Narrow AGORA (ms) | Big cache Δheap (MB) | Narrow Δheap (MB) |
|--------------|---------------------------:|---------------:|----------:|------------------:|---------------------:|------------------:|
| Match1       | 15.140 (~4,2h)             |    814,7       | ~18.600×  |    344,9          |  0,0¹                |  19,6             |
| Match0       | 14.918                     |    675,2       | ~22.100×  |    310,2          | 92,7                 |  19,6             |
| VLookupExact |  1.359                     |  1.154,7       | ~1.180×   |    344,1          | 122,3                |  19,6             |
| SumIfEqual   | 15.681                     |    775,2       | ~20.200×  |    303,9          | 97,9                 |  19,6             |
| CountIfEqual | 14.970                     |    703,0       | ~21.300×  |    312,4          | 97,9                 |  19,6             |
| Small        | 15.139                     |    636,1       | ~23.800×  |    288,3          | 78,6                 |  19,6             |
| SumRepeated  | 14.779                     |    765,9       | ~19.300×  |    279,1          | 74,8                 |  19,6             |

¹ Match1 é o 1º bloco Big; o Δheap saiu ~0 por ruído de coleta LOH entre blocos (o `SettleHeap` compacta,
mas o high-water do build do 1º bloco engoliu o delta). Os demais blocos Big são consistentes em ~75–122MB.

- **TOTAL medido dos 14 blocos (7 fórmulas × 2 alvos × 100k fórmulas = 1,4M avaliações): 7,71 s.**
- **Pico de heap gerenciado num único bloco: 341,6 MB. Pico de RSS do processo (`/usr/bin/time -l`, run
  inteiro sequencial): 721,5 MB.** Ambos DOMINADOS pelo workbook em si (1,5M células + 100k árvores de
  fórmula), não pelos caches — o Δheap atribuível à AVALIAÇÃO (memo de célula + índice estrutural + snapshot
  do range) é ~75–122MB (Big) / ~19,6MB (Narrow) por bloco, transiente (largado no `InvalidateCache`/
  `Recalculate`). Um pouco acima da projeção de 30–60MB do plano porque o Δheap também inclui o memo das
  100k células de fórmula e o memo de valor das 500k células lidas, não só os derivados do range.

**Meta batida com folga:** o relato era 400k fórmulas × 506k → ~57min. Um bloco Big de 100k fórmulas roda em
**≤1,15s**; extrapolando (agora É linear em F, custo por fórmula O(log N)) 400k fórmulas ≈ **~4,6s**. Mesmo
os 14 blocos juntos = **7,71s** — abaixo da aspiração de <10s, e ~440× abaixo do teto de 60s.

---

**Regressão comparativa (a) — MySheet-branch × MySheet-main, `SheetBenchmarks` ShortRun (N=3),
MemoryDiagnoser, mesma máquina; main = clone limpo de `2cd6727`, pré-caches:**

| Categoria     | Branch (ns) | Main (ns) | Δ tempo | Branch aloc | Main aloc | Δ aloc |
|---------------|------------:|----------:|--------:|------------:|----------:|-------:|
| Comparison    |    71,49    |   69,18   | +3,3%   |    24 B     |   24 B    |  0%    |
| Compute (Sum) |   107,29    |  105,09   | +2,1%   |   144 B     |  144 B    |  0%    |
| Conditional   |    97,44    |  101,42   | −3,9%   |    24 B     |   24 B    |  0%    |
| Parse         |   732,35    |  702,77   | +4,2%   |  2.224 B    | 2.224 B   |  0%    |
| SheetInMemory |   253,70    |  259,84   | −2,4%   |  1.512 B    | 1.512 B   |  0%    |
| TextEquality  |    32,54    |   32,81   | −0,8%   |    24 B     |   24 B    |  0%    |

**Critério (a) — branch ≤ main × 1,05 em tempo E alocação: PASSA.** Alocação **idêntica** em todas as
categorias (0%); tempo dentro de ±4,2% (< 5%). As barras de erro do ShortRun N=3 são largas (ex.: Parse
±137ns), então os deltas de ~2–4% estão dentro do ruído — a alocação, que é exata, prova que o hot path
não foi tocado.

**Regressão comparativa (b) — MySheet-branch × ClosedXML, `SheetBenchmarks`, mesma run:**

| Categoria     | MySheet (ns) | ClosedXML (ns) | MySheet aloc | ClosedXML aloc | Veredito                       |
|---------------|-------------:|---------------:|-------------:|---------------:|--------------------------------|
| Comparison    |    71,49     |    220,91      |    24 B      |    136 B       | MySheet ganha tempo E memória  |
| Compute (Sum) |   107,29     |    206,46      |   144 B      |    136 B       | tempo ganha (1,9×); aloc +8B²  |
| Conditional   |    97,44     |    212,59      |    24 B      |    136 B       | MySheet ganha tempo E memória  |
| SheetInMemory |   253,70     | 37.480,73      | 1.512 B      |  69.531 B      | ganha 148× tempo, 46× memória  |
| TextEquality  |    32,54     |    218,76      |    24 B      |    136 B       | MySheet ganha tempo E memória  |

² Única exceção do gate (b): `MySheetSum` aloca 144B vs 136B do `ClosedXmlSum` (+8 bytes). É PRÉ-EXISTENTE
(idêntico na main — não é regressão desta branch) e reflete workloads diferentes: o benchmark do MySheet
RE-avalia a expressão `=SUM(A1,A2)` a cada op, enquanto o do ClosedXML lê um `double` já computado/cacheado.
Em tempo o MySheet ganha (107 vs 206ns). Parse não tem par no ClosedXML (a lib não expõe parse isolado).

---

**Capability check ClosedXML no cenário de coluna inteira (spike, 16 linhas — falha aqui é de CAPACIDADE,
não de escala):** `XLWorkbook` + `.FormulaA1` + `RecalculateAllFormulas()` + `.Value`:

| Fórmula whole-column       | ClosedXML 0.105.0 |
|----------------------------|-------------------|
| `MATCH(x, A:A, 1)`         | OK                |
| `MATCH(x, A:A, 0)`         | OK                |
| `VLOOKUP(x, A:B, 2, FALSE)`| OK                |
| `SUMIF(A:A, x)`            | OK                |
| `COUNTIF(A:A, x)`          | OK                |
| `SUM(A:A)`                 | OK                |
| `SMALL(A:A, k)`            | **#NAME? (não implementada)** |

Ou seja: ClosedXML COMPETE em whole-column para MATCH/VLOOKUP/SUMIF/COUNTIF/SUM, mas **não implementa
SMALL** (nem LARGE) — MySheet responde onde o ClosedXML nem avalia.

**Comparativo whole-column reduzido (50k células de dados × 10k fórmulas, um engine por processo p/ RSS
limpo via `/usr/bin/time -l`; checksums IDÊNTICOS entre engines → equivalência semântica confirmada):**

| Fórmula        | MySheet eval (ms) | ClosedXML eval (ms) | Ganho tempo | MySheet pico RSS | ClosedXML pico RSS |
|----------------|------------------:|--------------------:|------------:|-----------------:|-------------------:|
| Match1 (aprox) |       123         |      67.294         | **~547×**   |     153 MB       |      156 MB        |
| Match0 (exato) |       105         |       1.821         |   ~17×      |     152 MB       |      155 MB        |
| VLookup        |       154         |       1.463         |   ~9,5×     |     152 MB       |      158 MB        |
| SumIf          |        95         |      16.874         |   ~178×     |     151 MB       |      158 MB        |
| CountIf        |        97         |      16.970         |   ~175×     |     152 MB       |      156 MB        |
| Sum            |       113         |      20.192         |   ~179×     |     153 MB       |      155 MB        |
| Small          |        ~78³       |   n/d (#NAME?)      |     —       |     ~152 MB      |    não avalia      |

³ Small do MySheet: ~78ms no harness reduzido (o `cxcmp` nem mede o par, pois o ClosedXML não avalia).

**MySheet é mais rápido E usa MENOS pico de RSS que o ClosedXML em TODAS as fórmulas suportadas** (2–6MB
abaixo em cada caso, MESMO carregando os caches por época) e ainda responde SMALL, que o ClosedXML não faz.
Achados: o MATCH aproximado do ClosedXML sobre coluna inteira é patológico (67s — ~37× o exato); MATCH
exato/VLOOKUP são rápidos (têm hashing/early-exit); SUMIF/COUNTIF/SUM re-varrem a coluna por fórmula
(~17–20s). O pico de RSS do MySheet neste workload (~152MB) é o número apples-to-apples — o 721MB da escala
plena é de rodar 14 workbooks de 500k sequencialmente num só processo, não um custo por-workload.

**Build:** `--no-incremental` 0 warnings. **Suíte:** core **805**, Excel **23**, fixture/`MemoryPack
CompatibilityTests` intocados e verdes. **Mudança de ferramenta:** `WholeColumnScaleHarness` ganhou o modo
`--full` MEDIDO (100k reais) com instrumentação de memória; `--full --sampled` preserva a extrapolação
antiga só para documentar a armadilha. Os projetos `cxcap`/`cxcmp` (capability + comparativo) ficaram no
scratchpad (fora do repo), como spikes descartáveis.

---

## Phase 4: admissão na 2ª leitura (regressão de produção)
Status: Complete

Relato do usuário (projeto real, 2026-07-03, pós-2.6.2): regressões de **1,5×–6×** em cenários
pequenos/médios; os volumes grandes (o alvo das Fases 1-3) seguiram resolvidos. **Diagnóstico** (confirmado
em `Workbook.TryGetRangeSnapshot`): a Fase 2 admitia o range no cache na **PRIMEIRA** leitura de QUALQUER
range ≥ 256 células. Fórmulas cujo range é lido UMA vez por época pagavam a materialização O(N) do snapshot
(+ um derivado que nunca reusavam) POR CIMA do scan linear que fariam de qualquer jeito — puro desperdício.
As três formas: janelas deslizantes (`SUM(A$1:A500)`, `SUM(A$1:A501)`, … — cada range distinto, lido 1×);
lookup limitado de tiro único (hash exato construído p/ UMA consulta); loop com `InvalidateCache` frequente
+ 1 leitura por época (rebuild constante).

- [x] **Política de admissão na 2ª leitura.** O valor do `_rangeCache` virou `RangeCacheEntry` com estado:
      1ª leitura → grava um marcador leve `Seen` (entry com snapshot nulo) e retorna `null` (o caller segue
      no caminho linear — o mesmo fallback/bypass que TODOS os consumidores já tinham); 2ª leitura → constrói
      o `RangeSnapshot` e o publica; 3ª+ → reusa. Thread-safe: `TryAdd` do marcador + `Interlocked.CompareExchange`
      na construção (corrida benigna na 2ª leitura: dois threads constroem, um publica, o outro é descartado).
- [x] **Limiar de 256 mantido**: abaixo dele nem marca (não polui o dicionário com ranges minúsculos).
      Derivados (hash/ordenado/mapas) continuam lazy pós-build, sem mudança.
- [x] **Cap defensivo de 64k marcadores** (`RangeCacheMarkerCap`, contador `Interlocked` resetado no drop):
      um flood de ranges distintos de uso único para de marcar além do teto e cai no linear — o marcador é
      minúsculo (chave + slot nulo) e some no `InvalidateCache`/`Recalculate`.
- [x] **Repro ANTES do fix** (`--range-cache-admission`, novo harness ao lado do `WholeColumnScale`; toggle
      `RangeCacheDisabled` = baseline pré-cache): mediu a regressão na árvore atual e registrou.
- [x] **Equivalência ajustada à nova política**: o harness diferencial agora cobre os TRÊS caminhos por época
      — leitura 1 (linear/marcada), 2 (build) e 3 (hit) — todos comparados ao BYPASS; o teste de limiar passou
      a afirmar "1ª = null, 2ª = snapshot, 3ª = mesma instância; sub-limiar nunca admite". 100% idênticos.

### Verification Plan
- Os 3 cenários de regressão voltam a ≤ ~1,05× do pré-cache; os blocos Big do harness mantêm os ganhos da
  Fase 2 (a 2ª leitura constrói — com 10k leituras é invisível). Suítes verdes `--no-incremental` 0 warnings.

### Phase Summary

**Regressão (harness `--range-cache-admission`, best-of, Apple Silicon 10 cores; `Enabled/Disabled` =
snapshot-cache LIGADO ÷ pré-cache linear — > 1 = o cache atrapalha):**

| Cenário          | Pré-cache (ms) | 2.6.2 na 1ª leitura | Pós-fix na 2ª leitura |
|------------------|---------------:|--------------------:|----------------------:|
| SlidingWindows   |   4.540–4.796  |     **3,05×**       |    **1,03–1,06×**     |
| BoundedMatchOnce |     256–278    |     **1,75×**       |    **1,03×**¹         |
| InvalidateLoop   |  14.987–18.207 |     **1,06×**       |    **0,99–1,05×**     |

¹ Numa run com carga concorrente o BoundedMatchOnce chegou a 1,16× (tempos absolutos pequenos, ~256ms →
sensíveis a ruído); numa run serial limpa deu 1,03×. O resíduo é bookkeeping do marcador (2 ops de dicionário
por range de uso único), não mais a materialização O(N) — que foi eliminada.

**Blocos Big preservados (harness `--whole-column-scale`, reduzido 50k×10k, BigColumn, ms — run serial limpa):**

| Fórmula      | Fase 2 (1ª leitura) | Fase 4 (2ª leitura) |
|--------------|--------------------:|--------------------:|
| Match1       |        141,1        |       148,9         |
| Match0       |         95,3        |       113,7         |
| VLookupExact |        154,7        |       139,6         |
| SumIfEqual   |         68,2        |        90,6         |
| CountIfEqual |         82,3        |        99,4         |
| Small        |         74,9        |        97,3         |
| SumRepeated  |         89,0        |        90,2         |

Todos na casa das centenas de ms como na Fase 2 (dentro do ruído inter-run desta máquina): a 2ª-leitura só
adiciona UM scan linear extra na 1ª das 10k fórmulas do bloco — invisível. O ganho de ~570×–1030× da Fase 2
sobre o baseline O(F×N) está intacto.

**Decisões além do briefing:** (a) `InternalsVisibleTo` adicionado para `Danfma.MySheet.Benchmark` (como já
existia p/ `.Tests`), para o harness de regressão poder alternar o switch `RangeCacheDisabled` e medir o
baseline pré-cache no MESMO binário (os outros benchmarks seguem só-API-pública). (b) Contador aproximado
`_rangeCacheEntryCount` (`Interlocked`) em vez de `ConcurrentDictionary.Count` (que trava todos os buckets)
para o cap — resetado nos dois pontos de drop. (c) O marcador é um `RangeCacheEntry` por range (não um
sentinela compartilhado): mais simples e a alocação minúscula gen0 não apareceu na medição serial.

**Build:** `--no-incremental` 0 warnings. **Suíte:** core **805**, Excel **23**, fixture/`MemoryPack
CompatibilityTests` intocados e verdes. **Sem API nova**; hot path de célula única intocado.

## Final Recap

Cenário do usuário (sheet de ~506k células numa coluna + ~400k fórmulas de coluna inteira consumindo-a,
ciclo carrega-1×-lê-1× → **~57min só de varreduras**) resolvido em DUAS camadas de cache interno, bounded,
descartável, SEM API nova e SEM tocar o modelo esparso nem o hot path de célula única:

- **Fase 0** — repro sintético determinístico + baseline que EXPÔS o O(F×N) (bloco de 100k fórmulas Big
  ~4,1–4,4h; até a coluna estreita de 16 células ~21–27min, só de varrer 1,5M chaves por fórmula).
- **Fase 1 — índice estrutural lazy por sheet** (`SheetStructuralIndex`): `coluna→ids ordenados` (+ simétrico
  linha→ids), build 1×O(N) por época, sobrevive ao `Recalculate`, largado no `InvalidateCache`. Colapsou o
  caso "colunas pequenas em sheet grande" (15,5s → 22–68ms por bloco: o scan de chaves sumiu).
- **Fase 2 — range caches por época** (`RangeSnapshot`): snapshot materializado 1× + derivados LAZY (hash
  exato, índice ordenado c/ prefix/suffix-max, mapa de igualdade numérica, view de números ordenados, memo
  de agregado), limiar de 256 células, largado no `InvalidateCache` E `Recalculate`. Colapsou o caso "consome
  a coluna grande" (O(F×N) → O(N + F·log N)): blocos Big do reduzido de ~55–80s para a casa das centenas de
  ms. Semântica preservada bit a bit (301 casos de equivalência BYPASS×FRIO×QUENTE 100% idênticos).
- **Fase 3 — validação plena + gate comparativo:** 14 blocos de 100k fórmulas em **7,71s** (bloco Big ≤1,15s;
  ~18.600–23.800× vs baseline nos casos de scan puro); pico de heap 341,6MB / RSS 721,5MB (dominado pelo
  workbook, não pelos caches). Sem regressão vs main (alocação idêntica, tempo ±4,2%); mais rápido E mais
  enxuto que o ClosedXML em TODAS as categorias core E em TODAS as fórmulas whole-column suportadas
  (9,5×–547× em tempo, 2–6MB abaixo em RSS), e resolve SMALL, que o ClosedXML nem avalia.

Meta do plano (< 60s pleno, aspiração < 10s): **batida** (equivalente ao relato ~4,6s; 14 blocos 7,71s).
Suítes verdes (805 + 23), 0 warnings, fixture intocada.

## Deployment Plan
_(mesmo ritual: verificação independente do usuário (rebuild forçado + benchmark) → merge → push → release
**2.7.0** lockstep → `git pull` → refresh `docs/pt-BR/` via Sonnet. PR opcional a critério do usuário.)_
