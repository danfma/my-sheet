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
Status: Not started

- [ ] Benchmark `Spike/WholeColumnScale/` no projeto de benchmark: gerador de workbook sintético com uma
      coluna de dados de ~500k células (números ordenados p/ o caso aproximado + strings p/ exato) e blocos
      de fórmulas de coluna inteira parametrizáveis: `MATCH(x_i, A:A)` (tipo 1 e 0), `VLOOKUP(x_i, A:B, 2,
      FALSE)`, `SUMIF(A:A, "=k_i")`, `COUNTIF(A:A, k_i)`, `SMALL(A:A, k_i)`, `SUM(A:A)` repetido.
- [ ] Medição em DUAS escalas: reduzida (50k células × 10k fórmulas — roda em segundos, vira teste de
      regressão de perf) e plena (500k × 100k+ — execução manual, documentada). Registrar baseline ATUAL
      (deve exibir o O(F×N)).

### Verification Plan
- Baseline plena reproduz a ordem de grandeza do relato (minutos, extrapolável a ~57min p/ 400k fórmulas).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 1: Camada 1 — índice estrutural lazy por sheet
Status: Not started

- [ ] Cache estrutural no `Workbook` (`[MemoryPackIgnore]`, lazy Interlocked): por sheet, `Dictionary<int,
      List<string>>` coluna→ids ordenados por linha (+ linha→ids sob demanda, construção independente).
      Invalidação: `InvalidateCache()` descarta; `Recalculate()` NÃO.
- [ ] `OpenRangeReference.PopulatedIds`/`TryGetPopulatedBounds`/`Expand`/`ExpandComputedValues` consomem o
      índice (colunas do range → concat das listas) em vez de varrer `Cells.Keys`. Ordem de enumeração
      preservada (col, depois row) — os testes existentes garantem.
- [ ] `Subtotal`/demais scans de open-range migram pro índice.

### Verification Plan
- Suíte completa verde (795+23) `--no-incremental` 0 warnings; fixture verde.
- Benchmark reduzido: caso "colunas pequenas em sheet grande" colapsa (o scan de chaves some).

### Phase Summary
_(escrever quando a fase concluir)_

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
