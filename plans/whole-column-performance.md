# Performance de coluna inteira em escala: índice estrutural + range caches por época

Cenário real (usuário, 2026-07-03): sheet com ~506.082 células numa coluna + ~399.776 fórmulas de coluna
inteira (`MATCH/VLOOKUP/XLOOKUP/INDEX`, `SUMIF/COUNTIF/AVERAGEIF`, `IF/ROW/SMALL`) **referenciando a coluna
grande**, ciclo carrega-1×-lê-1× → **~57min só de varreduras**. Meta: colapsar para **segundos**, mantendo
memória enxuta (o foco segue: bater ClosedXML/Aspose em performance E memória). Release **2.7.0**.

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
Status: Not started

- [ ] Rodar o benchmark PLENO (500k × 100k+): registrar tempo total e pico de memória ANTES × DEPOIS no
      relatório do plano. Meta: **< 60s na escala plena equivalente ao relato (aspiração < 10s)**; memória
      de pico dos caches documentada (~30-60MB esperados).
- [ ] **Regressão comparativa (requisito do usuário, 2026-07-03)**: rodar a suíte `SheetBenchmarks`
      (MySheet × ClosedXML: Parse/Compute/Comparison/TextEquality/Conditional, com MemoryDiagnoser) na
      BRANCH e na MAIN, lado a lado. Critérios: (a) nenhuma regressão de tempo/alocação nos caminhos core
      (célula única, cache-heavy — o hot path NÃO foi tocado, deve ser ~0%); (b) **continuamos mais rápidos
      que o ClosedXML em TODAS as categorias** (tempo E memória). Registrar a tabela no plano.
- [ ] **Comparativo ClosedXML no cenário de coluna inteira**: PRIMEIRO verificar a capacidade (lição do
      lessons.md: ClosedXML pode não avaliar MATCH/SUMIF/whole-column — testar num spike de 5min). Se
      suportar: medir o cenário reduzido (50k × 10k) nos dois engines e registrar; se não suportar,
      documentar a incapacidade (também é resposta: eles não competem nesse caso de uso). A memória extra
      dos nossos caches deve ser comparada com o footprint do ClosedXML no MESMO workload.
- [ ] `docs/performance.md`: seção sobre os range caches por época (o que existe, quando é descartado,
      números medidos honestos); nota em `workbook-and-expressions.md` se necessário. Skill
      `code-documentation-doc-generate`. NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [ ] `tasks/lessons.md`: a lição do O(F×N) (spike modelou amortização por leitura, não multiplicidade de
      fórmulas — benchmarks de estratégia devem modelar o EIXO DE CARGA COMPLETO: leituras × fórmulas).
- [ ] Plano: fases Complete + Final Recap.

### Verification Plan
- Números plenos no plano; suítes verdes; 0 warnings.

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando as fases 0–3 concluírem)_

## Deployment Plan
_(mesmo ritual: verificação independente minha (rebuild forçado + benchmark) → merge → push → release
**2.7.0** lockstep → `git pull` → refresh `docs/pt-BR/` via Sonnet. PR opcional a critério do usuário.)_
