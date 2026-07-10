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
Status: Not started

- [ ] Criar branch `feat/shared-formula-delta` a partir de main; trazer o diff do spike (merge ou
  cherry-pick de `64fbf3c` do branch do worktree; resolver drift vs main — commits novos: M4, docs)
- [ ] Re-rodar TODOS os gates do spike no branch: suítes completas, SharedFormulaDeltaTests,
  MemoryPackCompatibilityTests, ExcelLoadBenchmarks (tabela registrada), --k1-endtoend (<5%)
- [ ] Remover o worktree do agente após integração confirmada

### Verification Plan
- Suítes completas verdes; tabela BDN ≈ números do spike (±10% de ruído de máquina)

### Phase Summary
_(write when phase completes)_

## Phase 2: Auditoria de pattern-match nas famílias de funções (o risco central)
Status: Not started

- [ ] Varredura EXAUSTIVA de `is CellReference` / `is RangeReference` / switches por tipo de referência
  em Danfma.MySheet/Expressions/** (grep + revisão manual; o spike só adaptou NumericAggregation e
  ReferenceGuard). Suspeitos nomeados no spike: ROW/COLUMN/ADDRESS/CELL, Lookup/*, Subtotal, Npv,
  InformationFunctions; verificar também Offset/Indirect/Index, ArrayEvaluation, CriteriaScan,
  ArgumentFlattening, RangeValueCursor/PositionalRange (aceitam AnchoredRangeReference?)
- [ ] Decisão de design na execução: adaptar caso a caso OU introduzir normalização interna única
  (ex. `Expression.TryResolveReference` já existe? estender para os ancorados devolverem a forma
  concreta com delta aplicado) — preferir o ponto único se reduzir a superfície de erro
- [ ] Testes por família com fixture de shared formula exercitando cada função suspeita COMO ESCRAVA
  (ex. escrava com ROW(), OFFSET(A1,1,0), SUBTOTAL sobre range ancorado, INDIRECT — definir a matriz)
- [ ] Estender SyntheticK1Builder: grupo(s) de shared formula com RANGES (SUM(A1:A3) relativo) e
  funções da matriz — o fixture atual só tem refs de célula

### Verification Plan
- Suítes + novos testes por família verdes; grep de auditoria documentado no summary (lista de sites
  visitados e decisão por site)

### Phase Summary
_(write when phase completes)_

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
