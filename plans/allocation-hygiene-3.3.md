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
Status: Not started
- [ ] `RangeSnapshot.Build`: array pré-dimensionado preenchido via acessador numérico; fast path
      block-copy (`Array.Copy` por página) quando TODAS as páginas cobertas estão 100% presentes
      (bitmap all-set no intervalo), com re-checagem de versão do seqlock pós-cópia (retry por página).
- [ ] `ArgumentFlattening`/SUMPRODUCT/criteria: pré-dimensionar pelas bounds quando o argumento é range
      fechado; sem `List` com dobra no caminho quente.
- [ ] Gate: probe do Build ≤ ~1,2ms (fallback numérico) e ≤ ~0,2ms (block-copy full) @ 100k; alocação
      do Build ≈ tamanho do resultado (~2,3MB @ 100k); teste dirigido do block-copy com página
      parcialmente presente (usa fallback) e com escrita concorrente (retry do seqlock).
### Verification Plan
- Probe antes/depois; suítes 915+/24 verdes; fixture; k1-endtoend agregado idêntico; harnesses sem regressão.
### Phase Summary
_(write when phase completes)_

## Phase 2: Promoção adaptativa da 1ª página
Status: Not started
- [ ] Página nasce com array de slots pequeno (ex.: 128; knob em `ValueStoreOptions`) cobrindo o MESMO
      intervalo lógico de `RowPageSize` linhas; escrita além do array atual → promoção por realocação
      (dobra até `RowPageSize`), sob o seqlock da página (leitores re-tentam pela versão).
- [ ] Gate: sweep `--dense-store-pagesize` re-rodado — sheet pequeno (10×100) cai de ~10× para ≤ ~2×
      de desperdício SEM regressão no K1 denso (≤5% no compute); knob default documentado.
### Verification Plan
- Sweep antes/depois; suítes/fixture verdes; k1-endtoend na banda.
### Phase Summary
_(write when phase completes)_

## Phase 3: Índice estrutural numérico (linhas como int) + release
Status: Not started
- [ ] `SheetStructuralIndex`: buckets de coluna guardam linhas `int` (id derivado só quando um consumidor
      exigir string); `OpenRangeReference.ExpandComputedValues` vira 100% parse-free; medir memória do
      índice antes/depois e o lifetime harness.
- [ ] Docs (`performance.md` — nota curta das três melhorias) + release **3.3.0** (ritual completo;
      dispatch do orquestrador) + refresh pt-BR.
### Verification Plan
- Lifetime harness igual/melhor; suítes/fixture verdes; versionize propõe minor; re-run K1 3.2×3.3 do dono.
### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
