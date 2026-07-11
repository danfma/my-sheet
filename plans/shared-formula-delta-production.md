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
Status: Complete

- [x] `DependencyExtractor`: extrair dependências EFETIVAS dos nós ancorados (delta aplicado) em vez de
  `AlwaysDirty`; teste do RecalculationEngine com workbook carregado de xlsx com shared formulas
  (hoje não existe nenhum — gap de cobertura)
- [x] `AnchoredRangeReference.Evaluate/expansão`: eliminar o `RangeReference` transiente por avaliação
  (resolver bounds numéricos direto; medir com o fixture range-pesado da Fase 2)

### Verification Plan
- Testes do engine verdes incl. o novo; benchmark do fixture range-pesado sem alocação transiente por
  avaliação (probe de alocação)

### Phase Summary
**`DependencyExtractor` (Tarefa 1).** `Visit` passou a threadar `(deltaRow, deltaColumn)` por toda a
recursão (default 0/0 na raiz). Três casos novos substituem o antigo bloco `AlwaysDirty`:
`SharedFormulaSlave` entra na sua `Master` tree carregando O PRÓPRIO `(DeltaRow, DeltaColumn)` (mesma
semântica de `SharedFormulaSlave.Evaluate`'s `WithDelta`); `AnchoredCellReference`/`AnchoredRangeReference`
aplicam esse delta aos componentes RELATIVOS via dois novos helpers int-only extraídos das
`Effective(EvaluationContext)`/`ToRangeReference(EvaluationContext)` já existentes — `Effective(int,int)` em
`AnchoredCellReference.cs`, `EffectiveEndpoints(int,int)` em `AnchoredRangeReference.cs` — zero duplicação de
aritmética, comportamento dos sites existentes inalterado (refactor puro). `NameReference` dentro de uma
master tree recebe o delta ambiente sem resetá-lo (o nome é posição-independente por construção — inerte na
prática — espelhando `NamedReferences.EvaluateDefinition`, que também não reseta o contexto).

**Prova de paridade** (`tests/Danfma.MySheet.Tests/DirtyGraph/SharedFormulaDependencyParityTests.cs`, 19
casos): duas camadas. (1) por-fórmula — compara, para o MESMO texto de master + delta, o `DependencyScan` do
caminho ANCORADO (`SharedFormulaSlave` sobre `ExpressionParser.ParseAnchoredMasterBody`, o caminho de
produção) contra o LEGADO (`ExpressionParser.ParseSharedFormulaBody`, o parse por-slave token-delta
pré-spike, preservado intacto para o fallback honesto) — 8 formas (âncoras `$` mistas, range com um corner
absoluto, cross-sheet, IF, função custom sobre argumento ancorado, nome definido, aninhamento range+célula),
múltiplos deltas cada. (2) no grafo — dois workbooks completos (um com `SharedFormulaSlave` real, outro com a
mesma célula usando a expansão legada) comparados via `ReverseDependencyGraph.Build`: `Diagnostics()`
idêntico e `GetAllDependents` idêntico para 4 células de amostra (dependência escalar, de range, e não-lidas).
Verificado que os 19 testes REALMENTE detectam a regressão: revertendo só `DependencyExtractor.cs` ao estado
pré-fix (`git stash` do arquivo), os 19 falham — 3 no `AlwaysDirty` incorreto, 1 na `Diagnostics()`
(`AlwaysDirtyFormulas=6` vs `0`, `DistinctSourceCells=1` vs `5`) — confirmando que não passam "por acidente".

**Teste de engine com xlsx real (Tarefa 2)** —
`tests/Danfma.MySheet.Excel.Tests/SharedFormulaRecalculationEngineTests.cs`, 3 testes, gap de cobertura
fechado (nenhum teste do repo carregava shared formulas de .xlsx através do `RecalculationEngine` antes
disto). Fixture bruta com DOIS grupos (escalar `A{row}*2` e range `SUM(D{row}:F{row})`), 4-6 escravas cada.
O sinal REGRESSIVO forte é `RecalculationResult.DirtyCellCount`: como `DirtyEngine.CalculateDirty` une o
conjunto `AlwaysDirty` a TODO `Recalculate` incondicionalmente, o código pré-fix marcava todas as 6
escravas dirty em QUALQUER edição (valor final ainda correto, mas cone impreciso) — `DirtyCellCount` seria 7
(6 sempre-dirty + a célula editada) em vez de 2 (célula editada + seu único dependente real). Confirmado por
reversão: os 2 testes de edição de input falham no código velho com "Expected to be 2 but found 7". O
terceiro teste (edição da FÓRMULA do master) prova que o rebuild estrutural continua funcionando e que
escravas carregadas do .xlsx mantêm sua árvore ancorada CONGELADA (independente de uma edição posterior,
via API comum, na célula que originalmente foi o master do grupo — comportamento correto, documentado no
teste).

**Range transiente (Tarefa 3).** Probe primeiro: `GC.GetTotalAllocatedBytes` de época fria
(`workbook.InvalidateCache()` + reavaliar) para as 60k células do grupo H (`SUM(B2:D2)`) do
`samples/k1-synthetic.xlsx` — o transiente APARECIA no hot path (`NumericAggregation.Fold`/`FoldA`
chamavam `anchoredRange.ToRangeReference(context).ExpandComputedValues(context)`, materializando um
`RangeReference` + duas strings via `CellAddress.ToId()` POR AVALIAÇÃO). Baseline (5 execuções, mediana):
**81.239.000 bytes** (1354,0 bytes/célula). Fix: `AnchoredRangeReference` ganhou `GetBounds(context)`
(idioma `RangeReference.GetBounds`, direto de `EffectiveEndpoints`) e `ExpandComputedValues(context)`
(gêmeo de `RangeReference.ExpandComputedValues`, monta o `RangeValueSequence` sem o `RangeReference`
intermediário); `NumericAggregation.Fold`/`FoldA` chamam `anchoredRange.ExpandComputedValues(context)`
diretamente. Pós-fix (5 execuções, mediana): **61.719.280 bytes** (1028,65 bytes/célula) — **−24,0%**
(−325,3 bytes/célula), consistente com eliminar 1 `RangeReference` + 2 strings formatadas por avaliação.
Sites "uma vez por chamada" (`ArrayEvaluation`, `CriteriaScan`, `RangeValueCursor`, `TryResolveReference`)
continuam usando `ToRangeReference` — não são hot per-célula (mesmo tier que a Fase 2 já havia classificado
como "genérico funciona"), mantidos como estão por escopo.

**Sanidade de load** (`ExcelLoadBenchmarks --filter '*ExcelLoadBenchmarks*' -i`, in-process): 403,4ms /
114,07MB / Gen 18000-7000-3000 para `LoadSyntheticSharedFormulas` — bate com o número já registrado na Fase 2
(114MB, pós-grupo-H) dentro do ruído; nenhuma explosão. `LoadConvertedFromMyxl`: 1.088ms / 462,88MB (controle,
não afetado por esta fase).

**Suítes**: 1104 core (1085 + 19 novos) + 74 Excel (71 + 3 novos) verdes; csharpier limpo; build 0 warnings.

**SHAs dos commits**: ver Final Recap / histórico do branch.

## Phase 4: Fechamento e release
Status: Complete

- [x] Docs (en + pt-BR): excel-interop.md (como shared formulas são representadas; nota de
  forward-compat: arquivo .myxl novo contendo os nós 319-321 NÃO abre em versões antigas da lib —
  destacar na release), performance.md (números)
- [x] Nota honesta em docs: ganho é RAM/GC, não disco (wire serializa árvore por escrava)
- [x] Benchmarks finais (ExcelLoadBenchmarks + k1-endtoend + whole-column) na tabela do summary
- [x] Merge para main + release minor via workflow

### Verification Plan
- Suítes completas; BDN final registrado; release verde no NuGet

### Phase Summary
_(write when phase completes)_

### Phase Summary
Docs (0296507): excel-interop (subseção da representação delta + forward-compat destacada),
serialization (fronteira one-way dos tags 319-321 + nota RAM/GC-não-disco), performance (modelo de
delta 1×/época + tabela), README bullet — en + pt-BR, âncoras cross-checadas. Benchmarks finais em
máquina limpa (fixture pós-grupo-H, 7 grupos/420k fórmulas): LoadSyntheticSharedFormulas 381ms /
114,05MB / Gen 18000-7000-3000 (consistente com a referência da Fase 3); k1-endtoend compute 173,4ms
(≈ referência pré-G3, gate <5% com folga). M5 (v3.13.0) merged no branch antes dos gates finais —
estado medido = estado shippado.

## Final Recap

Produtização completa do spike G3 em 4 fases, cada uma com provas contrafactuais:
- **F1**: merge + revalidação (memória idêntica ao spike: −59% alloc, −50% Gen1+2 no shape
  shared-formulas; compute +3,9% A/B mesmo-dia).
- **F2**: auditoria exaustiva de pattern-match — 10 bugs reais corrigidos (ROW/COLUMN/FORMULATEXT/
  ISFORMULA/SHEET/SUBTOTAL/NPV/SUMIF/COUNTIF/MATCH em escravas), 17 testes de matriz (10 falham no
  código revertido), fixture ganhou grupo de range relativo.
- **F3**: DependencyExtractor delta-aware (paridade de grafo provada em 19 casos; DirtyCellCount 2 vs
  7 do always-dirty), 1º teste de engine com xlsx real, transiente de range eliminado (−24%/célula).
- **F4**: docs en+pt, gates finais limpos, merge.

Saldo de suítes: 1.085→1.114 core, 54→74 Excel. Wire: tags 319-321 append-only; arquivos antigos
intactos (testado contra binário congelado); forward-compat documentada como fronteira one-way.

## Deployment Plan
Merge em main + release minor via workflow (executados no fechamento). Consumidores: nenhum código
muda; arquivos Excel com shared formulas carregam automaticamente na representação delta; hosts que
serializam .myxl novos devem estar cientes da fronteira one-way (docs/serialization). Medição
recomendada no ambiente real do usuário: load do K1 verdadeiro + Gen counts BDN — o gap vs Aspose
deve ter fechado na proporção do peso de shared formulas no arquivo.
## Decisões pendentes registradas