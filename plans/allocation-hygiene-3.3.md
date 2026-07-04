# Higiene de alocação 3.3 — materializações, páginas adaptativas e índice numérico

Fechar os três ralos medidos que sobraram do ciclo 3.2 (princípio do dono: o investimento no-boxing do
`ComputedValue` não pode ser erodido por plumbing). Release-alvo **3.3.0 aditivo**; o re-run K1 do dono
comparará 3.2 × 3.3.

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-04** ("Vamos atacar a próxima fila"). Fases uma
> a uma com verificação independente do orquestrador.

## Evidência (tudo medido no ciclo 3.2 — probes em `benchmarks/`)
- `RangeSnapshot.Build` (Layer-2, 1×/época/range): `List<ComputedValue>` com dobra+recópia + `ToArray`
  = **10,1ms / 21,95MB @ 100k células**; array pré-dimensionado numérico ≈ 1,1ms/2,29MB; block-copy de
  páginas 100% presentes = **0,11ms / 2,29MB (88×)**. Os 2,29MB são o array-resultado, inevitável.
- `ArgumentFlattening.ExpandComputedValues`/SUMPRODUCT/criteria: mesmas `List` intermediárias (bounds
  do range são conhecidos → pré-dimensionar; probe `--range-sequence-probe` é a referência).
- Página de 1024 em sheet pequeno: 10 col × 100 linhas = ~246KB p/ 24KB úteis (10,3×; sweep
  `--dense-store-pagesize`). Página menor não custa tempo no sheet pequeno; K1 denso prefere 1024
  (~1,5× mais rápido que 128). Solução desenhada: **página inicial pequena cobrindo o MESMO intervalo
  de 1024 linhas** (shift/mask intactos) com promoção por realocação do array de slots.
- Open-range: o índice estrutural guarda `List<string>` de ids — a expansão paga 1 parse por id
  (achado do numeric-range-reads; o resto do caminho já é numérico). Armazenar linhas como `int` por
  coluna elimina o último parse do hot path open-range e encolhe o índice (id "C123456" ≈ 38B vs 4B).

## For Future Agents
TDD; build `--no-incremental` ANTES de qualquer `--no-build` (lição); suítes baseline: core **915**,
Excel **24**; fixture `workbook-pre-namespaces.msgpack.bin` + `MemoryPackCompatibilityTests` intocáveis
(NADA aqui muda schema — tudo runtime-only). TUnit/.NET 10 (`dotnet run --project tests/...`).
Semântica inegociável: época/voláteis/taint, seqlock (block-copy DEVE re-checar a versão da página após
a cópia), guarda de esparsidade, equivalência `--k1-endtoend` (agregado 27143285713). Git append-only,
sem push, branch nomeada, commits conventional inglês SEM `!` e SEM atribuição a IA. Verificação e
integração nunca no mesmo bloco Bash (lição do cwd).

## Phase 1: Materializações — Build block-copy + flattening pré-dimensionado
Status: Complete
- [x] `RangeSnapshot.Build`: array pré-dimensionado preenchido via acessador numérico; fast path
      block-copy (`Array.Copy` por página) quando TODAS as páginas cobertas estão 100% presentes
      (bitmap all-set no intervalo), com re-checagem de versão do seqlock pós-cópia (retry por página).
- [x] `ArgumentFlattening`/SUMPRODUCT/criteria: pré-dimensionar pelas bounds quando o argumento é range
      fechado; sem `List` com dobra no caminho quente.
- [x] Gate: probe do Build ≤ ~1,2ms (fallback numérico) e ≤ ~0,2ms (block-copy full) @ 100k; alocação
      do Build ≈ tamanho do resultado (~2,3MB @ 100k); teste dirigido do block-copy com página
      parcialmente presente (usa fallback) e com escrita concorrente (retry do seqlock).
### Verification Plan
- Probe antes/depois; suítes 915+/24 verdes; fixture; k1-endtoend agregado idêntico; harnesses sem regressão.
### Phase Summary
Entregue em `perf/snapshot-materialization` (`80ba67d` Build + `4aa3893` flattening; merged na main).
Achado de layout: snapshot é **column-major** — as páginas verticais por coluna mapeiam 1:1 em fatias
contíguas do destino, `Array.Copy` por página com pré-passe de presença (bitmap all-set do trecho) e
**re-checagem da versão do seqlock pós-cópia** (mudou/ímpar → retry; drop concorrente falha a
re-verificação → fallback célula-a-célula on-demand). `Build` REAL medido pelo orquestrador:
**1,94→0,15ms e 8,29→2,29MB @ 100k** (a alocação virou exatamente o array-resultado; o headline 10,1ms
do plano era pré-numeric-reads). Flattening pré-dimensionado só onde bounds são conhecidas
(`RangeReference` fechado, `PairwiseRanges`); open-range/unions/mistos deliberadamente intocados.
Verificação independente: core **922** (915+7 dirigidos, incluindo concorrência do block-copy e página
parcial), Excel 24, fixture intocada, 0 warnings, k1 agregado idêntico, lifetime/whole-column/mini-cse
sem regressão.

## Phase 2: Promoção adaptativa da 1ª página
Status: Complete
- [x] Página nasce com array de slots pequeno (ex.: 128; knob em `ValueStoreOptions`) cobrindo o MESMO
      intervalo lógico de `RowPageSize` linhas; escrita além do array atual → promoção por realocação
      (dobra até `RowPageSize`), sob o seqlock da página (leitores re-tentam pela versão).
- [x] Gate: sweep `--dense-store-pagesize` re-rodado — sheet pequeno (10×100) cai de ~10× para ≤ ~2×
      de desperdício SEM regressão no K1 denso (≤5% no compute); knob default documentado.
### Verification Plan
- Sweep antes/depois; suítes/fixture verdes; k1-endtoend na banda.
### Phase Summary
Entregue em `perf/adaptive-first-page` (4 commits, merged). Knob `InitialPageSlots` (default 128,
potência de 2 em [16, RowPageSize]). Promoção inteira numa janela ímpar→par do seqlock (Grow com UMA
alocação no tamanho final; leitor valida índices contra o comprimento CAPTURADO e re-checa a versão —
swap de array nunca visível sem retry; stress 0 torn-reads). Convive com o block-copy da F1 (fatia além
do array = não-presente → fallback, teste dirigido). **Refinamento do orquestrador após o flag do
agente** (+12,2MB de churn transiente na dobra do populate denso — vetado pelo princípio do dono):
**página 0 da coluna nasce pequena; páginas seguintes nascem CHEIAS** (coluna que transbordou provou
densidade) — churn K1 26,08→**14,00MB vs 13,92 born-full (+0,08MB ≪ gate 1,5MB)**. Números finais
(verificação independente): sheet pequeno 10×100 **241,9→30,8KB (1,3×)**; médio 50×5k idêntico nos três
regimes (waste 1,0×; ressalva honesta: coluna que termina logo após fronteira de página paga página
cheia — trade aceito); K1 compute +2,9% tempo / +0,1MB, agregado idêntico, e o run do orquestrador
registrou o total K1 batendo o Aspose (923,7 vs 955,2ms). Core **932** (922+10) / Excel 24 / fixture
intocada / 0 warnings / lifetime e Build da F1 nas bandas.

## Phase 3: Índice estrutural numérico (linhas como int) + release
Status: Complete
- [x] `SheetStructuralIndex`: buckets de coluna guardam linhas `int` (id derivado só quando um consumidor
      exigir string); `OpenRangeReference.ExpandComputedValues` vira 100% parse-free; medir memória do
      índice antes/depois e o lifetime harness.
- [x] Docs (`performance.md` — nota curta das três melhorias) + release **3.3.0** (ritual completo;
      dispatch do orquestrador) + refresh pt-BR.
### Verification Plan
- Lifetime harness igual/melhor; suítes/fixture verdes; versionize propõe minor; re-run K1 3.2×3.3 do dono.
### Phase Summary
Entregue em `perf/numeric-structural-index` (4 commits, merged). Buckets `List<string>`→`List<int>`
(`SortBucket` vira `Sort()` puro — o decorate-sort morreu); `OpenRangeReference.ExpandComputedValues`
100% parse-free via `PopulatedCells` numérico + acessador denso; `PopulatedIds`/`Expand` derivam id só
no caminho frio (Subtotal). **Achado da auditoria de ordenação: NÃO havia bug latente** — o decorate já
ordenava pelo eixo secundário como int; dois testes novos travam a ordem numérica (9→10→100; Z→AA).
**Calibração honesta vs plano**: RAM retida do índice **1,94× menor** (14,00→7,23 B/célula; o "~10×" do
plano era a forma SERIALIZÁVEL — 22,1→2,3MB = 9,7×, a alavanca do spike v4 de persistência); open-range
quente só ~2% (o parse era barato no hit; o ganho é estrutural: parse-free + alloc-free + forma
compacta). Verificação independente: core **934** (932+2), Excel 24, fixture intocada, 0 warnings, K1
agregado idêntico (TOTAL 0,83× Aspose no run do agente), lifetime/probes F1-F2 nas bandas.

## Final Recap
Ciclo 3.3 completo em 2026-07-04, mesmo dia do 3.2.0. Três fases: **F1** materializações
(`RangeSnapshot.Build` 1,94→0,15ms / 8,29→2,29MB via block-copy com re-checagem de seqlock; flattening
pré-dimensionado); **F2** página adaptativa (`InitialPageSlots` 128; sheet pequeno 241,9→30,8KB = 1,3×;
refinamento first-page-only cortou o churn de promoção de +12,2MB para +0,08MB após veto do orquestrador
ao flag do agente); **F3** índice estrutural numérico (RAM 1,94×; forma persistível 9,7× menor;
open-range parse-free; sem bug de ordenação). Zero mudança de schema/API além dos knobs aditivos em
`ValueStoreOptions`. Suítes finais: core **934** / Excel **24** / fixture intocada / 0 warnings. K1:
agregado idêntico em toda mudança; total 0,83-0,96× vs Aspose nos runs de verificação. Aceitação
externa: re-run K1 3.2×3.3 do dono. Próximo: spike v4 (persistência do índice + AST numérica).

## Deployment Plan
Executado em 2026-07-04: verificação independente por fase → rebase+ff-merge → sanidade na main com
rebuild `--no-incremental` primeiro → push → `merge-base --is-ancestor` em chamada separada →
`gh workflow run release.yml` (versionize deriva **3.3.0**) → `git pull --tags` → refresh `docs/pt-BR/`
via Sonnet → spike v4 (diretriz do dono).

## Pós-3.3 (diretriz do dono, 2026-07-04): spike v4
Ao concluir as fases e o release 3.3.0, despachar um **spike v4** cobrindo: (1) **persistência do índice
estrutural** — comparar MSWM v3 com seção opcional/descartável (recomendação do orquestrador: preserva
self-healing e não congela a representação no schema core) × membro appendado no schema, com medição do
ganho real (load+primeiro-toque; teto medido hoje: ~0,7ms @ 10k → ~30-36ms @ 500k, 1×/vida) e do custo
de save/tamanho; fazer DEPOIS da Fase 3 (linhas int encolhem a forma persistível ~10×); (2) **revisita
da AST numérica** (`CellReference`/`RangeReference` com `(col,row)` + sheet handle — breaking de wire) —
reavaliar com o codebase pós-3.3: o que resta de custo de string no hot path que só a AST numérica
remove, e se o número justifica o major. Spike = probes/benchmarks + veredito, sem produção; plano v4
próprio nasce do veredito.

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
