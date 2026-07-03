# Mini-CSE — avaliação elemento-a-elemento de argumentos-array (escopo K1)

Fechar os **305 diffs restantes do K1** (MySheet 2.9.0 × Aspose 26.6 × Excel, concordância 99,946%)
fazendo `SUM(IF(range=…,1,0))`, `SMALL(IF(range=…,ROW(range)),k)` e `INDEX(ROW(range),n)` avaliarem
elemento-a-elemento como no Excel (semântica CSE/implícita), **sem** spill, sem array público e sem nó
AST novo. É o subconjunto do F2 que o K1 exige (evidência em `plans/function-coverage-roadmap.md`,
"Evidência K1 no 2.9.0"); o F2 completo (`FILTER`/`SORT`/spill/`ComputedValueKind.Array` público)
permanece fase futura própria.

> **STATUS: PLANEJADO, NÃO AUTORIZADO PARA EXECUÇÃO.** O usuário pediu plano-somente (2026-07-03).
> Execução só com "prossiga/execute" explícito. Sequenciamento sugerido: após o 3.0 fechar (as frentes
> não conflitam em arquivos — isto vive na avaliação, o 3.0 no armazenamento — mas supervisão de uma
> frente de cada vez é o fluxo estabelecido).

## Fatos do codebase que ancoram o design (mapa auditado 2026-07-03)
- `Expression.Evaluate(EvaluationContext) → ComputedValue` é o contrato único; **não há kind de array**
  em `ComputedValueKind` (`ComputedValue.cs:7`) e o cache por célula (`Workbook.GetCellValue`,
  `Workbook.cs:482`) assume valor ESCALAR por célula — um array-valor vazando para célula quebraria o
  contrato. Arrays devem viver SÓ dentro da avaliação de uma fórmula.
- Hoje: `B2:B5="Show"` → `#VALUE!` (`RangeReference.Evaluate` → `#VALUE!`, `RangeReference.cs:10`;
  `BinaryOperation.Evaluate` propaga, `BinaryOperation.cs:29-35`); `IF` espera condição escalar
  (`If.cs:8-21`); `ROW(range)` → escalar `TopRow` (`Row.cs:16`); `INDEX` exige `RangeReference` concreto
  no 1º arg (`Index.cs:11-17` → `#REF!`); agregadoras avaliam arg-não-range UMA vez, escalarmente
  (`NumericAggregation.Fold` default case, `NumericAggregation.cs:84-100`).
- Maquinaria reusável: `ArgumentFlattening.ExpandComputedValues` (1 arg → `List<ComputedValue>`,
  `ArgumentFlattening.cs:67`), loop posicional do SUMPRODUCT (`SumProducts.cs:26-57`),
  `RangeReference.ExpandComputedValues` (valores memoizados por célula, `RangeReference.cs:50`).
- **Sem impacto MemoryPack**: nenhum nó AST novo (as fórmulas parseiam para nós existentes). Registro:
  tags 312–315 estão OCUPADAS (voláteis F1); a última é 316; qualquer nó futuro usa ≥317 — NÃO vale o
  que o handoff antigo dizia.
- Memo Layer-2 (`RangeAggregate.Memoize`) só ativa com `arguments is [Reference]` — `SUM(IF(...))` não
  entra nele; sem colisão de cache.
- Volatilidade: a reavaliação por célula passa por `GetCellValue` (via `ExpandComputedValues`), então o
  taint volátil sobe pelo choke point normal (`Workbook.cs:538-545`); sub-expressões voláteis dentro do
  array (ex.: `RAND()`) broadcast NÃO passam por célula — verificar taint na Fase C.

## Design (defaults meus; veto do usuário nos pontos marcados)
1. **Avaliador interno `ArrayEvaluation`** (novo arquivo em `Danfma.MySheet/Expressions/`): dado um nó e
   um contexto, produz um vetor `(dims, ComputedValue[])` reavaliando a sub-árvore elemento-a-elemento:
   `RangeReference` → valores por célula (via `ExpandComputedValues`); escalar/literal → broadcast;
   `BinaryOperation` com lado array → comparação/aritmética por elemento (zip + broadcast);
   `If` com condição array → zip condição×ramos (ramos escalares broadcast; `IF(cond,x)` sem else →
   `FALSE` lógico nos elementos falsos, como no Excel — é disso que o idiom `SMALL(IF(...))` depende:
   agregadoras ignoram lógicos/texto em arrays); `Row(range)` → vetor de números de linha. Qualquer nó
   fora desse conjunto → não-array (comportamento atual intacto). Dimensões: R×C row-major (o K1 só usa
   colunas, mas 2-D sai de graça no zip); mismatch de dims → `#VALUE!` por elemento como no Excel.
2. **Detecção no consumidor, não no produtor**: os nós existentes NÃO mudam seu `Evaluate` (célula com
   `=IF(B2:B5="Show",1,0)` seca continua `#VALUE!` — zero regressão, sem implicit intersection nova).
   Quem opta pelo caminho array são os consumidores: (a) `NumericAggregation.Fold` default case — se a
   sub-árvore é array-elegível (`ArrayEvaluation.TryEvaluate`), agrega o vetor (cobre SUM/COUNT/AVERAGE/
   MIN/MAX e SMALL/LARGE via `StatisticsMath.Collect`); (b) `OrderSelection.SortedArrayAndScalar`
   (`OrderStatistics.cs:410-425`) — 1º arg array-elegível → vetor ordenável; (c) `Index.Evaluate` — 1º
   arg array-elegível → materializa vetor e indexa (`n` fora do range → `#REF!` como Excel).
   [VETO?] a lista de consumidores é o dial de escopo: começar por SUM/SMALL/LARGE/INDEX (o que o K1
   exige) e a família COUNT/AVERAGE/MIN/MAX que sai do mesmo `Fold`; NÃO tocar demais funções no 1º passo.
3. **Semântica de agregação sobre array** = a de range: lógicos/texto ignorados por SUM/SMALL etc.
   (FALSE do IF-sem-else é EXCLUÍDO do sorted set do SMALL — validar contra Excel nos testes); erros por
   elemento propagam (primeiro erro vence), como no Excel.
4. **Guarda de custo**: array-elegibilidade requer range FECHADO (A2:A194); coluna inteira/`OpenRange`
   fica FORA do mini-CSE (→ comportamento atual) até o 3.0 provar o custo de varrer coluna usada.
   [VETO?] alternativa: permitir com clamp às células usadas — decidir na execução com benchmark.
5. **Nada de célula-array**: se o resultado final da fórmula da célula for um vetor (célula seca com IF
   array), mantém `#VALUE!` atual. Arrays existem apenas como ARGUMENTO dentro dos consumidores do item 2.

## For Future Agents
TDD estrito (RED com os repros antes de GREEN). Verificação `--no-incremental` 0 warnings. Fixture
`workbook-pre-namespaces.msgpack.bin` + `MemoryPackCompatibilityTests` intocáveis (este plano não muda
schema — se você se pegar precisando de tag MemoryPack, saiu do escopo). Suítes baseline: core 841+
(844 na branch da F1 do 3.0), Excel 24. TUnit/.NET 10: `dotnet run --project tests/... -c Release`
(`dotnet test` não funciona). Git append-only, sem push; commits conventional inglês sem atribuição a IA.
Golden values SÓ de oráculo (Excel real/Aspose — pedir ao usuário na dúvida; lição dos golden values).

## Phase A: Núcleo `ArrayEvaluation`
Status: Not started
- [ ] `ArrayEvaluation.TryEvaluate(Expression, EvaluationContext, out ArrayResult)` interno com o
      conjunto do item 1 (range, broadcast, BinaryOperation, If com/sem else, Row) + dims R×C.
- [ ] Testes unitários diretos do avaliador: `B2:B5="Show"` → `[F,T,F,T]`-like; `IF(cond,1,0)` → vetor;
      `IF(cond,ROW(range))` → números+FALSE; mismatch dims; erro por elemento; range aberto → recusa.
### Verification Plan
- Testes novos verdes; suítes completas verdes (nada de produção consome o avaliador ainda); build 0 warnings.
### Phase Summary
_(write when phase completes)_

## Phase B: Consumidores (SUM-família, SMALL/LARGE, INDEX)
Status: Not started
- [ ] `NumericAggregation.Fold` default case tenta `ArrayEvaluation` antes do caminho escalar;
      `OrderSelection` idem para o 1º arg; `Index.Evaluate` aceita 1º arg array-elegível.
- [ ] Testes RED→GREEN com os repros exatos do K1 (valores-oráculo do doc do usuário):
      `SUM(IF(B2:B5="Show",1,0))=2`; `SMALL(IF(B2:B5="Show",ROW(B2:B5)),1)=3`; `...,2)=5`;
      `INDEX(ROW(B2:B5),1)=2`; + BH25-like completo (IF aninhado com `>` e IF-sem-else, cross-sheet).
- [ ] Regressão: célula seca com IF-array continua `#VALUE!`; `SUM(range)` continua no memo Layer-2.
### Verification Plan
- Repros = valores do oráculo; suítes completas verdes; fixture verde; build 0 warnings.
### Phase Summary
_(write when phase completes)_

## Phase C: Validação K1 + docs + release
Status: Not started
- [ ] Taint volátil: `SUM(IF(range=RAND()>0.5,...))`-like marca a época (teste com contador).
- [ ] Custo: micro-benchmark do BH25-like @ 194 e @ 10k linhas — sem regressão nas suítes de perf.
- [ ] Docs: `function-reference` (notas de array-context nas funções cobertas) + parágrafo em
      `workbook-and-expressions.md`; NÃO tocar `docs/pt-BR/`.
- [ ] Release minor (`feat(eval): ...`) — ritual completo (merge-base em chamada separada; refresh pt-BR).
- [ ] Pós-release: pedir ao usuário o re-run K1 — critério de sucesso: os 305 diffs → ~0.
### Verification Plan
- Suítes/fixture verdes; benchmark sem regressão; versionize propõe minor; re-run K1 do usuário confirma.
### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
