# Shared Formulas por Nó-Delta — Produtização (G3)

Levar a produção o protótipo aprovado no spike G3 (veredito PROMOTE, 2026-07-10): escravas de shared
formula como `SharedFormulaSlave(master, Δrow, Δcol)` compartilhando 1 árvore ancorada por grupo, em vez
de ~360k árvores expandidas. Resultados do spike no shape shared-formulas: load 546→295ms (−46%),
allocated 251,6→102,5MB (−59%), Gen1+Gen2 −50%; compute +2,3% (<5%); paridade e round-trip verdes.

**Ponto de partida**: branch `worktree-agent-a6eb41a4368d55969` (checkpoint `64fbf3c`, base `4a10c6d`)
— worktree em `.claude/worktrees/agent-a6eb41a4368d55969`. Design: nós `AnchoredCellReference` /
`AnchoredRangeReference` / `SharedFormulaSlave` (union tags 319-321, append-only), modo ancorado no
Parser (preserva `$` numericamente), delta via `EvaluationContext` (WithCell zera; memoização torna o
shift 1×/época), fallback honesto para expansão token-delta quando o master usa formas não suportadas.

## For Future Agents
As work proceeds: mark checkboxes `- [x]`; phase done → status `Complete` + **Phase Summary** + rodar o
**Verification Plan** antes de seguir. Regras do repo: conventional commits sem referência ao assistente;
husky (pre-commit: csharpier+build; pre-push: +suítes); wire MemoryPack é contrato — tags append-only,
membros serializados intocados; `.csharpierignore` já exclui `.claude/`. Release ao final: minor (feat).

## Phase 1: Integrar o spike e revalidar gates na main
Status: Complete

- [x] Criar branch `feat/shared-formula-delta` a partir de main; trazer o diff do spike (merge ou
  cherry-pick de `64fbf3c` do branch do worktree; resolver drift vs main — commits novos: M4, docs)
- [x] Re-rodar TODOS os gates do spike no branch: suítes completas, SharedFormulaDeltaTests,
  MemoryPackCompatibilityTests, ExcelLoadBenchmarks (tabela registrada), --k1-endtoend (<5%)
- [x] Remover o worktree do agente após integração confirmada

### Verification Plan
- Suítes completas verdes; tabela BDN ≈ números do spike (±10% de ruído de máquina)

### Phase Summary
Merge do spike em `feat/shared-formula-delta` (bdf1385; staged por agente que travou em NuGet flaky,
commitado pelo orquestrador; zero conflitos). Interação M4×spike confirmada por código E teste
(CellStoreFormatter.Canonicalize só toca literais — SharedFormulaSlave passa intocado, CellStore.cs:371-403).
Suítes: 1.085 core + 54 Excel verdes; csharpier limpo. Gates re-medidos com worktree removido e máquina quieta:
- Memória (determinístico): Allocated 102,47MB e Gen 17000/6000/3000 — IDÊNTICOS ao spike (−59% / −50% vs baseline) ✓
- Mean 354,7ms vs 295 do spike-day: explicado por máquina ~10% mais lenta hoje (controle: 896 vs 788-838;
  razões normalizadas 0,396 vs 0,374 batem) — métrica de tempo tratada como indicativa
- Compute: A/B mesmo-dia main 185,9ms vs feat 193,2ms = +3,9% (<5%) ✓ (controle Aspose derivou +3% junto)
Worktree e branch do spike removidos após merge confirmado.

## Phase 2: Auditoria de pattern-match nas famílias de funções (o risco central)
Status: Complete

- [x] Varredura EXAUSTIVA de `is CellReference` / `is RangeReference` / switches por tipo de referência
  em Danfma.MySheet/Expressions/** (grep + revisão manual; o spike só adaptou NumericAggregation e
  ReferenceGuard). Suspeitos nomeados no spike: ROW/COLUMN/ADDRESS/CELL, Lookup/*, Subtotal, Npv,
  InformationFunctions; verificar também Offset/Indirect/Index, ArrayEvaluation, CriteriaScan,
  ArgumentFlattening, RangeValueCursor/PositionalRange (aceitam AnchoredRangeReference?)
- [x] Decisão de design na execução: adaptar caso a caso OU introduzir normalização interna única
  (ex. `Expression.TryResolveReference` já existe? estender para os ancorados devolverem a forma
  concreta com delta aplicado) — preferir o ponto único se reduzir a superfície de erro
- [x] Testes por família com fixture de shared formula exercitando cada função suspeita COMO ESCRAVA
  (ex. escrava com ROW(), OFFSET(A1,1,0), SUBTOTAL sobre range ancorado, INDIRECT — definir a matriz)
- [x] Estender SyntheticK1Builder: grupo(s) de shared formula com RANGES (SUM(A1:A3) relativo) e
  funções da matriz — o fixture atual só tem refs de célula

### Verification Plan
- Suítes + novos testes por família verdes; grep de auditoria documentado no summary (lista de sites
  visitados e decisão por site)

### Phase Summary
**Descoberta estrutural que orientou tudo**: a célula MASTER de um grupo shared-formula é sempre
re-parseada como árvore ORDINÁRIA com `CellReference`/`RangeReference` de verdade — só as células
DEPOIS dela (delta ≥ 0) viram `SharedFormulaSlave(master.AnchoredTree, ...)`
(`WorksheetStreamLoader.ReadCell`/`ExpandSlave`). Ou seja, o bug só é alcançável pela ESCRAVA, nunca
pelo master; os primeiros testes escritos erraram exatamente nisso (valor divergente na linha do master,
não da escrava) e passaram "por acidente" no código com bug — corrigido depois de reverter o fix e
confirmar que os testes deveriam falhar (ver abaixo).

**Auditoria — site → classificação → ação** (arquivo:função):
| Site | Classificação | Ação |
|---|---|---|
| `Lookup/Row.cs` `Row.Evaluate` (ROW(ref)) | (a) quebrado | Adicionado `[AnchoredCellReference]`/`[AnchoredRangeReference]` ao switch |
| `Lookup/LookupFunctions.cs` `Column.Evaluate` (COLUMN(ref)) | (a) quebrado | idem, espelha Row |
| `Lookup/LookupFunctions.cs` `FormulaText.Evaluate` (FORMULATEXT(ref)) | (a) quebrado | Arms para os dois tipos ancorados |
| `Information/InformationFunctions.cs` `IsFormula.Evaluate` (ISFORMULA) | (a) quebrado | idem FormulaText |
| `Information/SheetNumber.cs` (SHEET(ref)) | (a) quebrado | Arms — SheetName é literal, sem delta a aplicar |
| `Mathematics/Subtotal.cs` `GatherSkippingSubtotals` | (a) quebrado | `case AnchoredCellReference or AnchoredRangeReference`: resolve via `TryResolveReference` e re-despacha recursivamente (reusa os cases existentes, inclusive a regra "ignora SUBTOTAL aninhado") |
| `Financial/Npv.cs` | (a) quebrado (2 bugs) | range ancorado caía em `default`→#VALOR!; cell ancorada usava regra DIRECT em vez de REFERENCED (texto referenciado devia ser ignorado, não gerar erro) — cases dedicados |
| `ArgumentFlattening.cs` (`FlattenComputedValues`/`ExpandComputedValues`) | (a) quebrado (só range; cell já funcionava via `default`) | `ExpandComputedValues` resolve o range no topo (mesmo código do RangeReference depois); `FlattenComputedValues` ganhou um `case AnchoredRangeReference` mirror |
| `CriteriaScan.cs` `PositionalRange.Open` (SUMIF/SUMIFS/AVERAGEIF/…) | (a) quebrado | Resolve `AnchoredRangeReference` para `RangeReference` concreto no topo do método |
| `RangeValueCursor.cs` `Open` (COUNTIF/MATCH/XLOOKUP) | (a) quebrado | idem, re-despacha via `Open` recursivo |
| `ArrayEvaluation.cs` `Probe`/`TryBuildOperand` (idioma mini-CSE SMALL(IF(range=…,ROW(range)))) | (a) quebrado | `case AnchoredRangeReference` e `case Row{Arguments:[AnchoredRangeReference]}` espelhando os cases de `RangeReference` |
| `Lookup/Offset.cs` (OFFSET) | (b) genérico funciona | `NamedReferences.TryResolveReference` já despacha via virtual `TryResolveReference`, que os nós ancorados sobrescrevem corretamente |
| `Lookup/Index.cs` (INDEX) | (b) genérico funciona | idem; guard `is not Reference` já exclui corretamente os ancorados do branch de array |
| `Lookup/Indirect.cs` (INDIRECT) | (b) nunca toca nós ancorados | Sempre reparseia `ref_text` do zero — nunca vê o master; ROW()/COLUMN() sem argumento (usados no idioma INDIRECT("A"&ROW())) já corretos por construção |
| `Lookup/VLookup.cs`/`HLookup`/`Match`/`XLookup` (tabela/array via `NamedReferences.TryResolveReference`) | (b) genérico funciona | idem Offset/Index |
| `RangeValueCache.cs` `TryGetRangeSnapshot`/`RangeAggregate.Memoize` | (b) degrada com segurança | `range is not (RangeReference or OpenRangeReference)` → `null`, cai no caminho lento (agora corrigido); não é um bug, só não acelera o ancorado |
| `ComputedValue.cs` `EnumerateValues` | (c) inalcançável | Um `AnchoredCellReference`/`AnchoredRangeReference` NUNCA é embrulhado num `ComputedValue.Reference` (nem `Evaluate` produz isso) — só vê CellReference/RangeReference/OpenRangeReference/UnionReference concretos que uma função como OFFSET/CHOOSE já resolveu |
| `UnionReference.cs` `ExpandComputedValues`, `DynamicRange.cs` `TryBox` | (c) inalcançável | `AnchoredFormulaSupport.IsFullyAnchored` rejeita QUALQUER `UnionReference`/`OpenRangeReference`/`DynamicRange` no master inteiro — o grupo cai no fallback legado inteiro; essas formas nunca coexistem com nós ancorados |
| `DirtyGraph/DependencyExtractor.cs` | (c) fora de escopo (Fase 3) | Já documentado como `AlwaysDirty` conservador para os 3 nós — pendência explícita da Fase 3, não da 2 |

**Decisão de design**: por-site, não um funil de normalização único. Os dois sites QUENTES por célula
que o spike já tinha adaptado (`NumericAggregation`, `ReferenceGuard`) permanecem com o case mirror
manual (sem indireção extra). Nos demais sites — que rodam uma vez por chamada de função, não uma vez
por célula dentro do argumento — resolvo o nó ancorado para sua forma concreta com
`AnchoredRangeReference.ToRangeReference(context)` / `AnchoredCellReference.Effective(context)`
(ambos já `internal`, do próprio spike) e então: (i) reuso o case já existente por recursão
(Subtotal, RangeValueCursor — mesmo padrão que UnionReference já usava para suas áreas) ou por
reatribuição da variável local antes do switch (ArgumentFlattening, CriteriaScan — o range resolvido
cai naturalmente no `case RangeReference` já existente, zero código duplicado); ou (ii) adiciono um
arm dedicado quando o corpo é uma expressão curta (Row/Column/FormulaText/IsFormula/SheetNumber/Npv).
Um funil único (ex. `Expression.NormalizeAnchored` chamado incondicionalmente em todo switch) foi
descartado: teria que ser chamado em TODOS os sites de qualquer forma (mesmo esforço de edição), e
adicionaria uma checagem de tipo + possível chamada virtual em todo `CellReference`/`RangeReference`
comum também — o padrão escolhido só paga esse custo quando o nó É de fato ancorado.

**Resultados da matriz de testes** (`tests/Danfma.MySheet.Excel.Tests/SharedFormulaSlaveFunctionTests.cs`,
17 testes): verificado CADA teste revertendo os arquivos de fix para o estado pré-Fase-2 (checkout do
commit-pai) e confirmando que falha — não só que passa depois. 10 dos 17 falharam exatamente nos sites
listados como (a) acima: `Row`, `Column`, `Subtotal`, `Npv`, `SumIf`, `CountIf`, `Match`, `IsFormula`,
`FormulaText` (argumento ancorado), `Sheet`; os outros 7 (`Offset`, `Indirect`, `Index`, `VLookup`,
`Address`/ROW+COLUMN sem argumento, `FormulaText`-da-célula-escrava, e o idioma array
`SMALL(IF(range=…,ROW(range)))`) já passavam antes — travados como testes de regressão.

**O que estava QUEBRADO no spike e foi corrigido** (resumo executivo): `ROW(ref)`/`COLUMN(ref)`,
`SUBTOTAL` sobre range, `NPV` (célula com texto referenciado E range), `SUMIF`/`SUMIFS`/`AVERAGEIF`
sobre range relativo, `COUNTIF`/`MATCH`/`XLOOKUP` sobre range relativo, `ISFORMULA`/`FORMULATEXT` com
argumento ancorado, `SHEET(ref)`, e o idioma mini-CSE `SMALL(IF(range=…,ROW(range)))` — todos
retornavam `#VALOR!` (ou, no caso do NPV com texto, um erro que deveria ter sido ignorado) quando a
fórmula da escrava passava um nó ancorado diretamente para essas funções.

**Fixture sintética**: `tools/SyntheticK1Builder` ganhou um 7º grupo (coluna H, `SUM(B2:D2)`, arrastado
60k linhas) — a primeira RANGE relativa do fixture. Descoberta lateral: um range CROSS-SHEET
(`Data!A1:A3`) no modo anchored do Parser lança `ParseException` (gap documentado em
`ExpressionParser.ParseAnchoredMasterBody`), fazendo o grupo cair inteiro no fallback legado — um
dry-run do `ExcelLoadBenchmarks` (`--job Dry --inProcess`) mostrou 1 exceção capturada nos
diagnósticos de GC com essa formulação; trocando para `SUM(B2:D2)` (mesma sheet) a exceção desapareceu
e o `Allocated` de `LoadSyntheticSharedFormulas` caiu de 155MB para 114MB (o grupo agora de fato
compartilha UMA árvore `AnchoredRangeReference` entre as 60k escravas). Números não são gate (dry-run,
não o BDN completo), só confirmam que o novo grupo carrega e evalua sem erro.

**Números finais das suítes**: 1085 core + 71 Excel (54 pré-existentes + 17 novos) verdes; csharpier
limpo; build 0 warnings.

**SHAs dos commits**:
- `e26c5b9` — `fix(recalc): teach the remaining reference-pattern sites about anchored nodes`
- `626f602` — `test(recalc): slave function-family matrix for anchored-node pattern-match sites`
- `71b9998` — `chore(fixtures): add a relative-range shared-formula group to SyntheticK1Builder`

## Phase 3: Dirty graph preciso + range transiente
Status: Not started

- [ ] `DependencyExtractor`: extrair dependências EFETIVAS dos nós ancorados (delta aplicado) em vez de
  `AlwaysDirty`; teste do RecalculationEngine com workbook carregado de xlsx com shared formulas
  (hoje não existe nenhum — gap de cobertura)
- [ ] `AnchoredRangeReference.Evaluate/expansão`: eliminar o `RangeReference` transiente por avaliação
  (resolver bounds numéricos direto; medir com o fixture range-pesado da Fase 2)

### Verification Plan
- Testes do engine verdes incl. o novo; benchmark do fixture range-pesado sem alocação transiente por
  avaliação (probe de alocação)

### Phase Summary
_(write when phase completes)_

## Phase 4: Fechamento e release
Status: Not started

- [ ] Docs (en + pt-BR): excel-interop.md (como shared formulas são representadas; nota de
  forward-compat: arquivo .myxl novo contendo os nós 319-321 NÃO abre em versões antigas da lib —
  destacar na release), performance.md (números)
- [ ] Nota honesta em docs: ganho é RAM/GC, não disco (wire serializa árvore por escrava)
- [ ] Benchmarks finais (ExcelLoadBenchmarks + k1-endtoend + whole-column) na tabela do summary
- [ ] Merge para main + release minor via workflow

### Verification Plan
- Suítes completas; BDN final registrado; release verde no NuGet

### Phase Summary
_(write when phase completes)_

## Decisões pendentes registradas
- **Arena paginada**: EM ESPERA por decisão do usuário (avaliando outro cenário antes). Nota de design
  no plano audit-fixes (blocos sub-LOH). Os dados do heap profile (ASTs = 61-62% do heap) e o resultado
  desta produtização definem o residual que a arena disputaria.
- **Serialização de tipos do usuário**: pergunta aberta ao usuário (têm tipos próprios persistidos hoje
  ou é necessidade futura?) — pesa contra a arena e a favor de union extensível; aguardando resposta.

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
