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
Status: Not started

- [ ] `RangeValueCache` interno no `Workbook` (`ConcurrentDictionary<(string Sheet, Reference Range), entry>`,
      `[MemoryPackIgnore]`, lazy): snapshot materializado + derivados lazy (hash exato; ordenado c/ posições;
      mapas de igualdade p/ critérios; memo de agregado). Descartado em `InvalidateCache()` E `Recalculate()`.
- [ ] Ligar nos consumidores (sem mudar semântica; fallback linear sobre o snapshot p/ casos não-indexáveis:
      wildcards, criteria não-igualdade sem stretch, ranges pequenos < limiar ~256 células p/ não pagar
      overhead onde o linear ganha):
      - `Match`/`XMatch`/`XLookup` (exato→hash; aproximado→busca binária; wildcard→linear).
      - `VLookup`/`HLookup`/`Lookup` (FALSE→hash na coluna-chave; TRUE→binária).
      - `SumIf`/`CountIf`/`AverageIf` single-range (igualdade→mapa; resto→linear no snapshot).
      - `Small`/`Large`/`Median`/`Percentile`/`Quartile` (ordenado compartilhado).
      - `Sum`/`Count`/`CountA`/`Max`/`Min` de range puro → memo de agregado.
- [ ] Testes de equivalência: harness que avalia o MESMO conjunto de fórmulas com caches frios (1ª leitura)
      e quentes (2ª) e compara resultado a resultado; + casos de borda (erros no range propagam igual;
      texto case-insensitive; tipos mistos; range vazio; sheet-fantasma → #REF! intacto).
- [ ] Stretch (só se couber limpo): prefix-sums p/ SUMIF/COUNTIF com `>`/`>=`/`<`/`<=`.

### Verification Plan
- Suíte completa verde `--no-incremental` 0 warnings; fixture verde; equivalência frio×quente 100%.
- Benchmark reduzido: MATCH/SUMIF/SMALL sobre coluna grande colapsam de O(F×N) para O(N + F·log N).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Validação em escala plena + docs + release
Status: Not started

- [ ] Rodar o benchmark PLENO (500k × 100k+): registrar tempo total e pico de memória ANTES × DEPOIS no
      relatório do plano. Meta: **< 60s na escala plena equivalente ao relato (aspiração < 10s)**; memória
      de pico dos caches documentada (~30-60MB esperados).
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
