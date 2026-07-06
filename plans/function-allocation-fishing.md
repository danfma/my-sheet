# Pesca de alocações por função — A/B MySheet × Aspose com MemoryDiagnoser

Eliminar, função a função, as alocações transientes da avaliação em fórmulas reais (arrays, criteria,
texto) — o eixo que o pipeline fundido expôs: **4,1GB × 2,58GB por execução** (1,59×), com Gen0/Gen1
proporcionais. O hot path escalar já é zero-alloc; o array/criteria/texto NÃO é. Método do dono:
benchmark por função (par MySheet × Aspose), medir → otimizar pontualmente (ArrayPool, stackalloc,
buffers reusáveis) → re-medir, "pescando cada ponto".

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-06.** Iterações de pesca uma a uma, cada uma com
> verificação independente. Alvo inicial apontado pelo dono: `SUMIFS` e `SUMPRODUCT` (listas temporárias).

## Evidência
- Fundido preciso (12 iter, dados reais, GC default): MySheet 4,00s/4,1GB × Aspose 2,57s/2,58GB.
- Suspeitos por família (estrutura da carga real): fórmulas-array (mini-CSE materializa vetores
  `ComputedValue[]` por avaliação — `ArrayEvaluation`), criteria (`CriteriaPairs.Expand`/
  `ConditionalAggregates` — listas), `SUMPRODUCT` (`SumProducts` — listas pareadas), `COUNTIF` sobre
  coluna aberta (snapshots/enumeração), funções de texto (strings novas por op), `XLOOKUP` em colunas.
- Pré-dimensionamento (3.2.1) removeu a DOBRA das listas, não a alocação por avaliação.

## Regras do ciclo
0. **MODO (atualizado pelo dono, 2026-07-06)**: iterações 1-2 rodaram no modo exploratório (proposta →
   liberação). A partir da iteração 3: **autônomo até a fase final** — com gates verdes, o orquestrador
   integra e pusha direto; decisões de design restantes (ranges pequenos) são tomadas por número e
   justificadas no resumo final ao dono. Voltar ao dono apenas se um gate falhar ou um trade-off tocar
   semântica observável.
1. **Nada de otimização sem par de benchmark antes/depois** — o A/B com Aspose ancora o alvo (paridade
   ou melhor por função) e o MemoryDiagnoser é o juiz de alocação.
2. Gates por iteração: alocação da função ↓ SEM regressão de tempo (±5%); suítes completas verdes
   (core **957** baseline / Excel **24**); fixture intocada; build `--no-incremental` 0 warnings.
3. `ArrayPool`/buffers reusáveis: atenção a exception-safety (return no finally), a NUNCA vazar buffer
   alugado para resultado retido, e ao contrato de concorrência (avaliação concorrente existe — pool é
   thread-safe, mas buffer alugado é por-avaliação). `stackalloc` só para tamanhos pequenos limitados
   (criteria de N pares, não ranges de 100k).
4. Ao final do ciclo: re-rodar o pipeline real (fases + fundido) e registrar o efeito agregado; release
   perf (patch) com o ritual completo.

## Phase 1: Suíte FunctionBenchmarks (A/B por função)
Status: Complete
- [x] Nova classe em `benchmarks/Danfma.MySheet.Benchmark`: pares `{Função}_MySheet` × `{Função}_Aspose`
      (baseline) por categoria, `[MemoryDiagnoser]`, formas sintéticas pequenas E grandes inspiradas
      nas famílias reais: SUMIFS multi-criteria, SUMPRODUCT, COUNTIF coluna aberta, fórmula-array
      INDEX/SMALL/IF/ROW, IF-passthrough, OR com muitos disjuntos, XLOOKUP em colunas, texto
      (UPPER/TRIM/LEFT/FIND/concat), MATCH/INT, controle escalar (aritmética).
- [x] Rodar baseline completa e registrar a tabela (tempo + Allocated por função, razões vs Aspose) no
      plano — a ORDEM DA PESCA nasce dela (pior alocador primeiro).
### Verification Plan
- Suíte compila/roda nos dois motores; valores conferem entre motores em cada par (asserção de sanidade);
  tabela baseline registrada.
### Phase Summary
Entregue em `bench/function-fishing` (`8f3ba32`, merged): `FunctionBenchmarks.cs` com 40 benchmarks
(10 famílias × 2 formas × 2 motores), equivalência assertada por par no GlobalSetup. **Correção
metodológica do agente (aprovada)**: medir via `Evaluate` direto da expressão-alvo com inputs quentes
(não `InvalidateCache`+read, que re-materializaria o store inteiro por iteração) — âncora escalar em
**24B** (só o boxing do harness) prova o isolamento. Aspose: `ForceFullCalculation`+chain off.
**Baseline (ShortRun; verificação independente bateu o par SUMIFS)** — tempo: MySheet vence TODAS as
famílias (0,008×–0,69×). Alocação, os alvos da pesca:
| # | Família | MySheet | Aspose | nota |
|---|---|---:|---:|---|
| 1 | SUMIFS @50k | **14,9MB (Gen2!)** | 15,4MB | ~299 B/linha; criteria lists |
| 2 | Array INDEX/SMALL @50k | **14,5MB (Gen2/LOH!)** | 10,9MB | ÚNICA família onde alocamos MAIS (1,33×); vetores mini-CSE 50k×24B > limiar LOH |
| 3 | SUMPRODUCT @50k | 3,6MB | 10,8MB | listas pareadas |
| 4 | Ranges PEQUENOS (XLOOKUP 25,7KB / COUNTIF 12,9KB / MATCH 5,2KB @200) | — | — | inversão: abaixo da admissão do cache, aloca por avaliação |
| 5 | Texto | 272B | 6,4KB | marginal |
| — | Zero-alloc (nada a pescar) | IF-passthrough, OR, escalar, COUNTIF@50k, XLOOKUP/MATCH warm | | |
Sanidade: core 957 / Excel 24 / 0 warnings / produção intocada.

## Phase 2+: Iterações de pesca (uma função/família por vez)
Status: In progress
- [x] **Iteração 1 — SUMIFS/família *IFS (LIBERADA pelo dono e integrada, `bba5cd5`+`ed7514d`)**:
      via híbrida — snapshot admitido = zero-cópia; sem snapshot = cursores posicionais via enumerator
      denso (`CriteriaScan` novo); open-range/union mantém lista. **Achado do 2º alocador escondido:
      `Criteria.Matches` reconstruía o regex de wildcard POR CÉLULA (~10MB @50k×2 critérios)** — regex
      compilado 1× no construtor; SUMIF/COUNTIF/AVERAGEIF de texto ganham de graça. Resultado
      (verificação independente): Sumifs50k **17,2ms→3,3ms (5×)** e **14,2MB→5,36KB (÷2.700), GC
      zerado** (Aspose: 42,5ms/15,0MB — alocamos 0,04% dele). Core **963** (957+6 bordas de criteria),
      Excel 24, fixture intocada, suíte de 40 sem regressão cruzada. SUMPRODUCT NÃO compartilha a
      maquinaria (confirmado) — segue como alvo próprio.
- [x] **Iteração 2 — fórmula-array (LIBERADA pelo dono e integrada, `e6a639a`+`0c22de2`)**: árvore lazy
      de operandos + `ArrayStream` (enumerator struct); SUM-família agrega sem vetor; `INDEX(array,n)`
      anda até o n-ésimo sem buffer (e deixou de avaliar posições não-selecionadas — mais lazy, taint
      espúrio removido); `SMALL/LARGE` com heap-k sobre `ArrayPool` de k slots. Borda de erro-após-k
      provada no código ANTIGO antes do refactor (teste novo). Resultado (verificação independente,
      rebased): Array50k **10,5→1,83ms (~6×)** e **14,2MB→1,03KB (÷14.000), LOH/Gen2 ZERADOS** (Aspose:
      24ms/10,6MB). Micro-divergências documentadas e aceitas: k avaliado antes da varredura (inerte);
      INDEX não avalia não-selecionados. Core **964** (963+1), Excel 24, sem regressão cruzada na suíte.
      Iterações 1+2 somadas: os dois maiores alocadores da carga real de ~29MB → ~6,4KB/avaliação.
- [x] **Iteração 3 — SUMPRODUCT (integrada em modo autônomo, `fc249ac`+`b5ffb19`)**: reusa o
      `PositionalRange` da iteração 1 SEM generalização (um cursor por argumento, N em paralelo);
      bordas de precedência (dims antes de erro de célula; erro antes da regra do zero) provadas no
      código antigo primeiro. Resultado (verificação independente): Sumproduct50k **3,06ms→294µs** e
      **3,5MB→336B (÷10.700), Gen2 zerado** (Aspose: 11ms/10,8MB — 0,003%). `SUMX*`/`PairwiseRanges`
      deliberadamente fora: compartilham lista com as estatísticas bivariadas (CORREL/SLOPE, que
      precisam de dois passes) e não têm par de benchmark (regra 1). Diff: 1 arquivo de produção.
      Core **966** (964+2), Excel 24, suíte de 40 sem regressão cruzada.
- [x] **Iteração 4 — ranges pequenos não-admitidos (integrada em modo autônomo, `10ff799`+`c74de65`)**:
      novo `RangeValueCursor` (streaming do caminho não-admitido) em XLOOKUP/MATCH/COUNTIF; borda de
      XLOOKUP com arrays de tamanhos distintos pinada no código antigo primeiro. Resultado (números
      colhidos pelo orquestrador — o agente estagnou 2× no padrão background e foi dispensado do
      relatório): XLOOKUP@200 **25,7KB→944B**, COUNTIF-open@200 **12,9KB→576B**, closed **5,2KB→376B**,
      MATCH@200 **5,2KB→304B**; tempos 2-4× melhores. Warm @50k SEM regressão (alguns melhoraram:
      COUNTIF 120→96B); trio pescado confirmado (SUMIFS 5.488B / Array 1.056B / SUMPRODUCT 336B).
      Core **967** (966+1), Excel 24, 0 warnings, fixture intocada. Limiar de admissão do cache
      INTOCADO (decisão autônoma: mudança de comportamento global ficaria para o dono — e o streaming
      tornou-a desnecessária). **Atribuição (relatório final do agente)**: 100% do custo era o
      `ArgumentFlattening.ExpandComputedValues` materializando vetor por argumento no fallback
      não-admitido (24B/célula + doubling no open-range + cadeias de iteradores yield). **Residual
      irredutível nesta via (candidatos futuros)**: (a) os ~300-950B restantes são as cadeias de
      iteradores BOXED do open-range — eliminar exigiria enumerator struct para `OpenRangeReference`;
      (b) XLOOKUP aproximado/wildcard/reverso não-admitido segue materializando (precisa de acesso
      indexado no `LookupMatching`; sem par de benchmark e fora da forma da carga real). Juiz de
      semântica: `RangeValueCacheEquivalenceTests` (301 casos diferenciais) verde.
### Verification Plan
- Por iteração: alocação ↓, tempo ±5%, suítes/fixture verdes; MODO EXPLORATÓRIO: proposta → liberação
  do dono → integração.
### Phase Summary
Quatro pescas integradas (1-2 com liberação do dono; 3-4 em modo autônomo com gates verdes). Placar de
alocação por avaliação: SUMIFS 14,2MB→5,4KB; Array 14,2MB→1,03KB (LOH extinta); SUMPRODUCT 3,5MB→336B;
ranges pequenos 5-26KB→304-944B. Todos com tempo 2-6× MELHOR (gate era ±5%). Infra criada e reusada:
`CriteriaScan`/`PositionalRange` (it.1) → reusada it.3; `ArrayStream`+heap-k (it.2); `RangeValueCursor`
(it.4). Suítes finais: core **967** (957+10 bordas novas), Excel 24. Fora deliberadamente: `SUMX*`
(maquinaria bivariada de 2 passes, sem par de benchmark), texto (marginal), limiar de admissão.

## Final: efeito agregado + release
Status: Complete
- [x] Pipeline real re-medido (fases + fundido); release perf patch; refresh pt-BR se docs mudarem.
### Phase Summary
**Efeito agregado (fundido preciso, 12 iter, dados reais, GC default)**: alocação **4,1GB → 1,98GB**
(menos da METADE; agora 0,77× do Aspose — era 1,59×), Gen0 696k→334k/1000 ops; tempo ratio 1,56×→
**1,11×±0,19** (sessão ruidosa; a razão é o dado confiável). **Release v3.4.1 publicado** (patch dos
`perf:`, pré-condição das 4 pescas em chamada separada). Docs não mudaram no ciclo → sem refresh pt-BR.

## Final Recap
Ciclo completo em 2026-07-06, método do dono (A/B por função + pesca pontual) validado ponta a ponta:
suíte de 40 benchmarks com equivalência assertada (Fase 1) → 4 iterações de pesca (SUMIFS ÷2.700;
array ÷14.000 com LOH extinta; SUMPRODUCT ÷10.700; ranges pequenos ÷14-27) → efeito agregado no
pipeline real: alocação do fundido pela metade, abaixo do Aspose. Todas as famílias ficaram 2-6× mais
RÁPIDAS (o gate era só não regredir). Achados de bônus: regex de wildcard reconstruído por célula
(it.1); INDEX avaliava posições não-selecionadas (it.2, taint espúrio removido). Fora por decisão
documentada: SUMX* (sem par de benchmark), limiar de admissão (comportamento global), texto (marginal),
LinkedArraySegment do dono (aprovado-em-princípio para F2/spill). Próximo da fila: multi-targeting
net8.0 (`plans/net8-multitargeting.md`).
