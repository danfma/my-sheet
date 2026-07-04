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
Status: In progress
- [ ] Store paginado substituindo `_cache`; visited-bit substituindo `_evaluating`; tainted por página;
      `InvalidateCache`/`Recalculate` preservando semântica de época e voláteis.
- [ ] Testes: equivalência total das suítes SEM mudança de comportamento; testes dirigidos de
      concorrência conforme o contrato escolhido; fixture verde.
### Verification Plan
- Core 893+/Excel 24 verdes; fixture verde; 0 warnings; `--k1-endtoend` com agregado idêntico.
### Phase Summary
_(write when phase completes)_

## Phase 2: Validação + docs + release
Status: Not started
- [ ] `--k1-endtoend`: compute ≤ ~350ms e alocação de sweep ≤ ~20MB (gates); total vs Aspose registrado.
- [ ] Harnesses do 3.0 (lifetime/write-cost) e SheetBenchmarks sem regressão.
- [ ] Docs (`performance.md`) + release (3.2.0 aditivo OU 4.0 breaking conforme item 5) + refresh pt-BR.
### Verification Plan
- Gates batidos; suítes/fixture verdes; versionize propõe a versão correta ao ramo.
### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
