# Dense value store — fechar o gap de compute vs Aspose

Substituir a `ConcurrentDictionary<(string,string), ComputedValue>` do cache de valores por um **store
denso paginado por sheet**, endereçado numericamente — o alvo que o profiling do compute K1 revelou.
Meta: compute 687ms → ~250–350ms (paridade com o eval do Aspose), alocação do sweep 162MB → ~10–20MB.

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-03 ("prossiga").** Fases uma a uma com
> verificação independente. Os dois [VETO] serão decididos com os números da Fase 0 (defaults do plano
> valem até lá). Batizado "4.0" pela possibilidade de quebra, mas a Fase 0 pode rebaixá-lo a um 3.x
> NÃO-breaking — ver "Insight central".

## Evidência (medida 2026-07-03, harness `--k1-compute-attrib` + dotnet-trace, verificação independente)
- Compute K1 (400k fórmulas cross-sheet): ~687–780ms. Atribuição: **estrutura do cache domina** —
  `Monitor.Enter_Slowpath` (lock de bucket por insert) ~45% do self-time, `GrowTable` ~20%*, GC induzido
  pelas alocações (162MB/sweep, 1 Node heap por entrada) ~24%. (*GrowTable amplificado pelo loop do
  sampler; o probe de populate único confirma a direção sem o artefato.)
- Hash/equals da chave string-tupla: ~97–115ms (~15%). Resolução de sheet por nome: **3,3ms (0,4% —
  irrelevante, NÃO atacar)**. Lookup de id no `_cells`: ~16ms. AST walk+coerção: ~90ms/400k — quase grátis.
- Escada de estruturas (mesmo tráfego: 1,0M lookup + 600k insert): `ConcurrentDictionary<(string,string)>`
  **574ms** → `ConcurrentDictionary<(int,int)>` 324ms → `Dictionary<(int,int)>` 48ms → **array denso 2,8ms**.
- `HashSet _evaluating` (guarda de ciclo): ~20ms churn; visited-bit por célula ≈ 0.
- Referências: `plans/function-coverage-roadmap.md` (blocos K1) e `benchmarks/.../K1ComputeProfileHarness.cs`.

## Insight central (por que talvez NÃO precise quebrar nada)
O ganho grande NÃO exige numerizar a AST nem o wire format: `CellAddress.TryGetColumnRow` (no-alloc, já
existe) deriva `(col,row)` da string do id NO MOMENTO do acesso ao cache, e o sheet resolve a um handle
int com custo desprezível (3,3ms/600k). Ou seja: o store denso pode viver INTEIRO atrás do
`Workbook.GetCellValue`/`InvalidateCache`, com `CellReference` continuando a guardar strings — API e
MemoryPack intocados. A Fase 0 mede se a derivação on-the-fly come parte do ganho; só se comer é que a
numerização da AST (breaking) entra em pauta. O usuário já autorizou quebrar formato SE necessário — a
ordem correta é provar que não é.

## Design (defaults meus; [VETO] = decisão do usuário)
1. **Store**: por sheet, **duplo nível de páginas nas DUAS dimensões** (revisão do dono, 2026-07-04 —
   o array plano de colunas exigiria grow+copy a cada coluna nova E explodiria com colunas gridless de
   índice alto, ex. `AAAA` ≈ 475k):
   `colGroups[col >> 6][col & 63]` (grupos de 64 colunas) → `column.pages[row >> 10]` →
   `ComputedValue[1024]` (slot = `row & 1023`). Hot path só shift/mask + derefs — ZERO hash (Dictionary
   no diretório seria tolerável mas é 17× pior no nível de célula, medido; e o resize dele re-hasheia
   buckets — pior que copiar ponteiros de grupo). Arrays de grupos/páginas crescem por dobra e são
   publicados via `Interlocked`. Flag de presença por slot (bitmap 128B/página — slot zerado é ambíguo:
   "não computado" × "blank computado"). **Layout medido (2026-07-04)**: `Unsafe.SizeOf<ComputedValue>`
   = 24B (payload real 17B: double 8 + object 8 + kind 1, padding de alinhamento 7); página de 1024 =
   24.600B exatos. SoA (arrays paralelos, 17B/slot) anotado como opção futura, NÃO default (complica o
   seqlock). `InvalidateCache` = descartar/limpar páginas (época de VALORES preservada; índice
   estrutural vitalício do 3.0 intocado). **Tamanhos parametrizados e a explorar (dono, 2026-07-04)**:
   shifts centralizados (`GroupShift`/`PageShift`), sempre potência de 2 (26/alinhamento alfabético
   descartado — custaria div/mod no hot path); a Fase 1 entrega varredura medida de row-page
   128/256/512/1024 em 3 formas (pequena 10×100, média 50×5k, K1) e recomenda o default — desperdício
   de página cheia em sheet pequeno é a preocupação (ex.: 10 col × 100 linhas = 246KB p/ 24KB úteis).
   Alternativa adaptativa a avaliar: página inicial pequena cobrindo o MESMO intervalo de linhas
   (shift/mask intacto) com promoção por realocação quando a coluna crescer.
2. **[VETO] Contrato de concorrência da avaliação** — a decisão que define o teto do ganho:
   (a) *manter avaliação concorrente*: páginas publicadas via `Interlocked`, escrita de `ComputedValue`
   (struct multi-word, torn-write possível) protegida por versão/seqlock por página ou lock striped —
   captura o grosso mas não os 2,8ms ideais; (b) *declarar avaliação single-threaded por época* (locks
   fora do hot path, doc explícito): store vira array puro, ganho máximo. Default meu: (a) striped por
   página — o codebase hoje suporta avaliação concorrente (F1/voláteis) e eu não rebaixo contrato público
   por performance sem o dono mandar.
3. **Visited-bit por célula** no lugar do `HashSet<(string,string)> _evaluating` (guarda de ciclo) e
   avaliação do mesmo para `_volatileTainted` (lista/bitmap por página).
4. **Chaves string somem do hot path** (cache/evaluating/tainted); `Sheets` por nome e `_cells` ficam
   (fora do hot path de leitura repetida; `_cells` lookup é 16ms — reavaliar depois, não agora).
5. **[VETO] Critério de quebra**: se a Fase 0 mostrar que a derivação on-the-fly custa >15% do ganho,
   ativa-se a numerização da AST (`CellReference` com `(int col,int row)` + sheet handle; wire quebra;
   major 4.0 com guia de migração e resave dos MSWM). Senão, isto vira release **3.2.0 aditivo**.

## For Future Agents
TDD; verificação `--no-incremental` 0 warnings; fixture `workbook-pre-namespaces.msgpack.bin` +
`MemoryPackCompatibilityTests` intocáveis (o store é runtime-only `[MemoryPackIgnore]` — se schema mudar,
você está no ramo breaking do item 5 e isso exige o gate do usuário). Suítes baseline: core **893**,
Excel **24**. TUnit/.NET 10 (`dotnet run --project tests/...`, `dotnet test` NÃO funciona). Semântica
INEGOCIÁVEL: memoização por época, voláteis/taint (F1), snapshots Layer-2, equivalência do
`--k1-endtoend` (agregado 27143285713). Git append-only, sem push; commits conventional inglês sem
atribuição a IA. Release dispatch em chamada separada pós merge-base (lição).

## Phase 0: Spike do store + decisão de concorrência
Status: Complete
- [x] Probe (benchmarks) do GetCellValue-shape completo: derivação on-the-fly + store paginado, nas duas
      variantes de concorrência do item 2, contra o baseline — medir compute K1 e alocação.
- [x] Registrar: custo da derivação on-the-fly (decide o item 5); recomendação de concorrência com
      números (decide o item 2 com o usuário).
### Verification Plan
- Tabela probe vs baseline; decisão dos vetos registrada no plano antes da Fase 1.
### Phase Summary
Entregue em `spike/dense-store` (`77e0ea0`, merged na main), harness `--dense-store-spike`, best-of-7,
verificação independente (re-run + suítes 893/24 + 0 warnings). Tráfego GetCellValue-shape do K1
(1,0M lookups, 600k inserts), equivalência de soma entre TODAS as variantes:

| Variante | ms | MB churn |
|---|---:|---:|
| baseline `ConcurrentDictionary<(string,string)>` | 391,5–392,6 | 133,8 |
| denso (a) seqlock por página + derivação | 25,4 | 13,9 |
| denso (b) single-threaded + derivação | 20,6 | 14,0 |

Derivação on-the-fly (DIRETRIZ DO DONO 2026-07-03: span + endereço bit-packed `(col<<32)|row` + página
por shift/mask, zero alocação/split — `CellAddress.TryGetColumnRow` já satisfaz; variante span/packed
escrita e conferida, ambas 0,00MB): **14,3ms sobre 1,0M acessos**.

**[VETO i] RESOLVIDO PELO NÚMERO: derivação = 3,8–3,9% do ganho (≪15%) → o store sai ADITIVO
(release 3.2.0), numerização da AST NÃO disparada, formato MemoryPack intocado.**
**[VETO ii] recomendação registrada, decisão final do usuário na retomada**: manter avaliação
concorrente via (a) seqlock por página custa ~5–6ms absolutos sobre o sweep inteiro (vs teto (b) 7,7ms
pré-derivado) — recomendo (a); (b) só se o dono quiser declarar avaliação single-threaded por época.
Visited-bit: 2,9ms vs 15,2ms do HashSet (5×, zero alloc) — aprovado para a Fase 1.

**RISCO REAL para a Fase 1 — esparsidade patológica**: com páginas de 1024 linhas, 10k células
espalhadas 1-por-página inflam o store a ~240MB (vs ~4MB do dict); clustered/colunas densas ficam em
0,2–3,2MB (melhor que o dict). Fase 1 PRECISA de guarda: fallback híbrido para dict em sheets de
densidade-por-página ultra-baixa (o índice estrutural do 3.0 já conhece bounds/ocupação para detectar)
ou caso patológico documentado — decidir na Fase 1. Torn-write de `ComputedValue` (24B) é real: a
variante (a) EXIGE o seqlock. `InvalidateCache`/snapshots/voláteis: equivalência de época é gate da
Fase 1, não provada pelo spike (que cobre a fatia estrutura ~370ms; AST+coerção ~90ms ficam).

## Phase 1: Implementação no Workbook
Status: Complete
- [x] Store paginado (`SheetValueStore`) substituindo `_cache`; cycle-guard numérico substituindo o
      `HashSet<(string,string)>`; tainted no store; `InvalidateCache`/`Recalculate` preservando época e voláteis.
- [x] Testes: equivalência total das suítes SEM mudança de comportamento; dirigidos de concorrência
      (seqlock stress, mesma-célula-sem-falso-ciclo), Recalculate/taint, esparsidade com teto; fixture verde.
### Verification Plan
- Core 893+/Excel 24 verdes; fixture verde; 0 warnings; `--k1-endtoend` com agregado idêntico.
### Phase Summary
Entregue em `feat/dense-value-store` (`92a26b4` store+wiring+testes, `f77bb70` sweep de page-size).
`SheetValueStore` (runtime-only, `[MemoryPackIgnore]` — schema INTOCADO, fixture + MemoryPackCompat verdes):
por sheet um slab com diretório de colunas em dois níveis `groups[col>>6][col&63]` → `column.pages[row>>10]`
→ página `ComputedValue[1024]` + bitmap de presença (128B) + seqlock (versão par/ímpar + CAS de escritor).
Endereço derivado on-the-fly do id A1 (`CellAddress.TryGetColumnRow`, no-alloc); handle de sheet por nome
(OrdinalIgnoreCase). Diretórios crescem por dobra, publicam via `Volatile`.

**Desvios justificados do design:** (1) **cycle-guard fica thread-local** (não bit-por-slot compartilhado):
`_evaluating` é `[ThreadStatic]` e o contrato documentado é "re-avaliação concorrente benigna da MESMA célula
em threads diferentes NÃO é falso ciclo" — um bit compartilhado quebraria isso. Capturei o ganho da chave
(string,string → tripla `(handle,col,row)`) mantendo a thread-locality; a página NÃO precisa de bitmap de
visited. Teste `SameCell_EvaluatedConcurrentlyAcrossThreads_IsNotAFalseCycle` protege o contrato. (2) **tainted
vira `ConcurrentDictionary<(int,int,int),byte>`** (não bitmap por página): o conjunto tainted é esparso (só
células voláteis), então um dicionário é O(tainted) para marcar E para o drop do Recalculate, enquanto um
bitmap forçaria varredura de todas as páginas tocadas. (3) **fallback overflow** para ids não-A1 (só chamada
direta do host produz um; fórmulas normalizam pra A1) — dicionário `(string,string)` dormante que preserva o
comportamento antigo EXATO. (4) **guarda de esparsidade por contadores de runtime** (páginas alocadas vs
células presentes por slab), não pelo índice estrutural do 3.0: o store precisa funcionar antes do índice
existir e sem acoplar ao Sheet. Limiar `WarmupPages=64`/`MinCellsPerPage=4` — abaixo disso o slab desvia
células novas-de-página para um dicionário por-slab. Cenário patológico (10k espalhadas até 1M) fica em ≤64
páginas (~1.6MB) + dict, provado por `SparseScatter_CapsPagesInsteadOfBallooning` (assertiva de teto).

**Números (`--k1-endtoend`, best-of-3, agregado IDÊNTICO 27143285713):**

| Fase | baseline | dense store |
|---|---:|---:|
| compute | 746,7 ms / 160,3 MB | **127,1 ms / 42,7 MB** |
| extract (a) | 104,7 ms / 28,8 MB | **21,1 ms / 28,8 MB** |
| TOTAL vs Aspose | 1,67× tempo | **0,65× tempo, 0,50× alloc** |

Compute já bate o gate da Fase 2 (≤350ms) com folga (127ms); a alocação do compute (42,7MB) é 3,75× melhor que
o baseline mas ainda acima do alvo ≤20MB da Fase 2 — o resíduo é AST-eval/coerção (que o spike não cobria), não
o store. **Page-size sweep (`--dense-store-pagesize`, diretriz do dono):** nenhum tamanho fixo é ótimo nos dois
extremos — planilha pequena (10×100) desperdiça 10,3× a 1024 vs 1,3× a 128 (sem custo de tempo), planilha densa
desperdiça ~1,0× em qualquer tamanho mas varre ~1,5× mais rápido a 1024 (K1 7,3 vs 11,6 ms). O alvo K1 favorece
1024 → default `PageShift=10` mantido e validado. Fase 2: promoção adaptativa da 1ª página (array de slots
pequeno crescendo até cheio enquanto a página cobre o MESMO intervalo de 1024 linhas — shift/mask intactos).

## Phase 2: Validação + docs + release
Status: Complete
- [x] `--k1-endtoend`: compute ≤ ~350ms e alocação de sweep ≤ ~20MB (gates); total vs Aspose registrado.
- [x] Harnesses do 3.0 (lifetime/write-cost) e SheetBenchmarks sem regressão.
- [x] Docs (`performance.md`); release **v3.2.0 publicado 2026-07-04** + refresh pt-BR despachado.
### Verification Plan
- Gates batidos; suítes/fixture verdes; versionize propõe a versão correta ao ramo.
### Phase Summary
Validação entregue em `feat/dense-store-validation` (`f1bafae` probe `--k1-compute-alloc`, `c7513c6`
docs; merged na main). **AMBOS os gates batidos**: compute 127-183ms ≤ 350ms; alocação — o headline de
42,7MB do sweep decomposto byte-exato (probe reproduz 100%): **28,8MB (67%) são artefato do harness**
(montagem `"C"+r` das strings de id no loop de medição), **13,9MB (33%) são o backing intrínseco das
páginas do store** (≤ 20MB ✓) e o **transiente de leitura/avaliação do engine é 0,0MB** — a premissa
"resíduo é AST/coerção" foi REFUTADA (o hot path escalar não aloca nada; verificação independente do
orquestrador confirmou byte-exato). Nenhuma correção cirúrgica necessária. Regressões: lifetime flat
0,049-0,156ms/época; write-cost Fill puro inalterado e live-index +20% (dentro da banda +23-27% aceita
no 3.0); mini-cse O(range) (array vence escalar @ 10k); aspose-compare consistente; SheetBenchmarks
hot path de leitura idêntico (24B/144B — resíduo é o AsObject do próprio benchmark). `performance.md`
descreve o dense paged store (endereçamento derivado, diretórios em 2 níveis, seqlock, guarda de
esparsidade, época preservada). Nota de processo: o agente pegou e corrigiu (reset local pré-branch,
permitido) um rodapé de atribuição acidental — commits finais limpos, conferido por grep.

## Investigação ReadOnlySequence em ranges (pergunta do dono, 2026-07-04 — VEREDITO MEDIDO)
Probe `--range-sequence-probe` (merged, `3285813`), verificação independente. `SUM(A1:A100000)` denso
já computado: caminho atual (`ExpandComputedValues` com `ToId()` StringBuilder + re-parse + `HandleFor`
POR CÉLULA) 5,6-18,3ms/13,7MB → **acessador numérico por célula (hoist do handle + `TryGetDense(h,c,r)`)
1,14-1,16ms / 0,00MB (−79% tempo, −100% alocação, risco semântico ZERO** — on-demand/cycle-guard/taint/
seqlock intactos) → page-span/`ReadOnlySequence` 0,14ms (+1ms marginal sobre B, zero alocação a mais,
e ativa os 4 obstáculos: slots não-computados, escape do seqlock, fallback de esparsidade, retângulo
multi-segmento por coluna). **Recomendação: (1) acessador numérico nas expansões de range**
(`RangeReference`/`OpenRangeReference`/`CellComputedValueAt`/`ArrayEvaluation.ExpandRange` — alimenta
todos os consumidores de uma vez); **(2) `RangeSnapshot.Build` com block-copy de páginas 100% presentes**
(88-93× no build, 10,1→0,11ms/−19,7MB, uma vez por época por range); **(3) `ReadOnlySequence` público:
NÃO** — o container nunca foi o problema, o endereçamento por string era. `IEnumerable` fica.

**PRINCÍPIO DO DONO (2026-07-04)**: o investimento no-boxing do `ComputedValue` não pode ser erodido
pelo plumbing de enumeração/materialização. Três ralos, por tamanho: (grande) materializações
intermediárias `List<ComputedValue>` com dobra+recópia — `RangeSnapshot.Build` 21,95MB/100k vs 2,29MB
do array pré-dimensionado; `ArgumentFlattening`/SUMPRODUCT idem — atacar no follow-up imediato pós-B;
(médio) dispatch virtual do enumerator por célula ~1-2,7ms/100k — adendo despachado (enumerator struct
ou loop por bounds nos consumidores quentes); (nulo) `IEnumerable<ComputedValue>` genérico NÃO boxa o
elemento (Current devolve a struct por valor).

## Final Recap
Publicado como **v3.2.0 aditivo** em 2026-07-04 — o desfecho não-breaking que a Fase 0 provou possível.
O cache de valores virou o `SheetValueStore`: páginas densas com diretórios em dois níveis nas duas
dimensões (revisão do dono), seqlock por página, guarda de esparsidade com fallback, endereçamento
derivado on-the-fly (span/bits, diretriz do dono) — compute K1 **747→~127ms** (2× mais rápido que o
calc do Aspose; total 0,65-1,04× com METADE da alocação), transiente de avaliação **zero**, backing
13,9MB. Extras do ciclo: geometria configurável via `WorkbookOptions` (config runtime-only, bytes de
serialização idênticos provados); leituras de range numéricas + enumerator struct (pergunta do
`ReadOnlySequence` do dono → veredito medido: container não era o problema; expansão 9,0→1,29ms e
13,7MB→0 por 100k, piso on-demand-safe). Suítes finais: core **915** / Excel **24** / fixture intocada /
0 warnings. Fila 3.3 medida e ranqueada: materializações `List<ComputedValue>` (Build 88× via
block-copy/prealloc + flattening), promoção adaptativa de página, refactor do índice estrutural para
(col,row) numérico no open-range.

## Deployment Plan
Executado em 2026-07-04: verificação independente por entrega (spike F0 → store F1 → validação F2 →
options → numeric-range-reads) → rebase+ff-merge de cada branch → sanidade na main SEMPRE com rebuild
`--no-incremental` antes das suítes (lição dos binários velhos) → push → `merge-base --is-ancestor` dos
commits de feature em chamada separada → `gh workflow run release.yml` (versionize derivou **3.2.0**) →
`git pull --tags` → refresh `docs/pt-BR/` via Sonnet. Futuras releases: mesmo ritual.
