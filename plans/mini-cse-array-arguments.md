# Mini-CSE — avaliação elemento-a-elemento de argumentos-array (escopo K1)

Fechar os **305 diffs restantes do K1** (MySheet 2.9.0 × Aspose 26.6 × Excel, concordância 99,946%)
fazendo `SUM(IF(range=…,1,0))`, `SMALL(IF(range=…,ROW(range)),k)` e `INDEX(ROW(range),n)` avaliarem
elemento-a-elemento como no Excel (semântica CSE/implícita), **sem** spill, sem array público e sem nó
AST novo. É o subconjunto do F2 que o K1 exige (evidência em `plans/function-coverage-roadmap.md`,
"Evidência K1 no 2.9.0"); o F2 completo (`FILTER`/`SORT`/spill/`ComputedValueKind.Array` público)
permanece fase futura própria.

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-03** ("podemos atacar o mini-cse"), após o
> fechamento do 3.0 (v3.0.0 publicado). Fases uma a uma com verificação independente entre elas.
> Baseline pós-3.0: core **862** / Excel **24** verdes.

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
Status: Complete
- [x] `ArrayEvaluation.TryEvaluate(Expression, EvaluationContext, out ArrayResult)` interno com o
      conjunto do item 1 (range, broadcast, BinaryOperation, If com/sem else, Row) + dims R×C.
- [x] Testes unitários diretos do avaliador: `B2:B5="Show"` → `[F,T,F,T]`-like; `IF(cond,1,0)` → vetor;
      `IF(cond,ROW(range))` → números+FALSE; mismatch dims; erro por elemento; range aberto → recusa.
### Verification Plan
- Testes novos verdes; suítes completas verdes (nada de produção consome o avaliador ainda); build 0 warnings.
### Phase Summary
Entregue em **`feat/mini-cse-core`** (fork de `dded8b6`), commit `c99b776`. `ArrayEvaluation.TryEvaluate`
interno + `ArrayEvaluationResult` struct (row-major, dims R×C); operandos com broadcast de escalar;
recursão sobre range fechado (via `CellComputedValueAt` — mesmo choke point `GetCellValue`, memoização/
taint preservados), `BinaryOperation` (semântica compartilhada via `Apply` extraído), `If` 2/3-args com
condição array (sem else → FALSE), `Row(range)`; recusa: nó não-elegível, range aberto. 11 testes
RED→GREEN (RED: 9 falhas + 2 recusas esperadas). Verificação independente: core **873** / Excel 24 /
fixture intocada / 0 warnings. **Decisões do orquestrador sobre as ambiguidades da entrega**:
(1) row-major CONFIRMADO — `INDEX` da Fase B indexa row-major se receber 2-D (K1 é 1-D, indiferente);
(2) mismatch de dims = vetor inteiro `#VALUE!` MANTIDO (para SUM/SMALL "primeiro erro vence" coincide
com o Excel; a nuance broadcast+`#N/A` do Excel real só entraria se um consumidor futuro precisar);
(3) `IF` com condição ESCALAR e ramo array fica FORA (fora do escopo K1; F2 decide);
(4) taint volátil de broadcast = item da Fase C, como planejado.

## Phase B: Consumidores (SUM-família, SMALL/LARGE, INDEX)
Status: Complete
- [x] `NumericAggregation.Fold` default case tenta `ArrayEvaluation` antes do caminho escalar;
      `OrderSelection` idem para o 1º arg; `Index.Evaluate` aceita 1º arg array-elegível.
- [x] Testes RED→GREEN com os repros exatos do K1 (valores-oráculo do doc do usuário):
      `SUM(IF(B2:B5="Show",1,0))=2`; `SMALL(IF(B2:B5="Show",ROW(B2:B5)),1)=3`; `...,2)=5`;
      `INDEX(ROW(B2:B5),1)=2`; + BH25-like completo (IF aninhado com `>` e IF-sem-else, cross-sheet).
- [x] Regressão: célula seca com IF-array continua `#VALUE!`; `SUM(range)` continua no memo Layer-2.
### Verification Plan
- Repros = valores do oráculo; suítes completas verdes; fixture verde; build 0 warnings.
### Phase Summary
Entregue em **`feat/mini-cse-consumers`** (fork de `c99b776`), commit `dfa8cdb`. Portão comum:
`ArrayEvaluation.IsArrayEligible` (walk sintático puro, sem avaliar — evita dupla-avaliação e risco com
voláteis); hot path escalar intacto (short-circuit). Consumidores: `Fold` default case (cobre SUM/COUNT/
AVERAGE/MIN/MAX **e transitivamente SMALL/LARGE/percentis** via `StatisticsMath.Collect` → decisão
ACEITA: não duplicar no `OrderSelection`, seria caminho morto); `Index` com forma-array row-major +
**caso especial documentado `INDEX(ROW(coluna aberta), n)` = identidade `top+n-1` sem materializar**
(o idiom BH25 real; hoje dava `#REF!`; núcleo continua recusando open-range). 12 testes RED→GREEN
(RED: 9 falhas + 3 verdes provando o comportamento atual), incluindo BH25-like cross-sheet completo.
Verificação independente: core **885** (873+12) / Excel 24 / fixture intocada / 0 warnings.

## Phase C: Validação K1 + docs + release
Status: Complete
- [x] Taint volátil: `SUM(IF(range=RAND()>0.5,...))`-like marca a época (teste com contador).
- [x] Custo: micro-benchmark do BH25-like @ 194 e @ 10k linhas — sem regressão nas suítes de perf.
- [x] Docs: `function-reference` (notas de array-context nas funções cobertas) + parágrafo em
      `workbook-and-expressions.md`; NÃO tocar `docs/pt-BR/`.
- [x] Release minor (`feat(eval): ...`) — ritual completo (merge-base em chamada separada; refresh pt-BR).
- [ ] Pós-release: re-run K1 do usuário no 3.1.0 — critério de sucesso: os diffs de CSE → ~0 (dos 301
      diffs no 3.0.0, sobra só o OR/AND com literal de texto, thread separada pendente do critério-Aspose).
### Verification Plan
- Suítes/fixture verdes; benchmark sem regressão; versionize propõe minor; re-run K1 do usuário confirma.
### Phase Summary
Validação entregue em **`feat/mini-cse-validation`** (fork de `dfa8cdb`; commits `fb7b74a` taint,
`1a0cdcf` harness de custo, `f9ab4b9` docs) — ZERO mudança de produção na fase. **Taint volátil PASSA**:
o flag thread-local `_volatileTouched` sobe pelos dois caminhos do avaliador (células de range via
`GetCellValue`; broadcast escalar via `Evaluate` no frame da célula), provado por 4 testes não-vácuos
(refresh pós-`Recalculate` das voláteis + controle não-volátil permanece cacheado = taint preciso).
**Custo O(range)**: `SUM(IF)` array 1,59× o COUNTIF escalar @194 linhas e 0,76× @10k (sub-linear, sem
blow-up); BH25 64× de 194→10k (fator O(n log n) do SMALL, não da maquinaria). Harness lifetime e
SheetBenchmarks nas bandas registradas (produção byte-idêntica à Fase B). Docs com exemplos cobertos por
testes verdes. Verificação independente: core **889** (885+4) / Excel 24 / fixture intocada / 0 warnings.

## Final Recap
Mini-CSE entregue e publicado como **v3.1.0** em 2026-07-03, no mesmo dia do 3.0.0. Três fases: **A**
núcleo `ArrayEvaluation` interno (row-major, broadcast, zip de BinaryOperation/IF, ROW-vetor, recusa de
open-range; 11 testes); **B** consumidores via portão sintático `IsArrayEligible` (família SUM pelo
`Fold`, SMALL/LARGE transitivos, INDEX com forma-array + identidade `INDEX(ROW($A:$A),n)` sem
materializar; 12 repros K1 verdes incluindo o BH25 cross-sheet completo); **C** validação (taint volátil
provado nos dois caminhos; custo O(range) sem blow-up; hot paths escalares byte-idênticos) + docs.
Nenhum nó AST novo → schema MemoryPack intocado (fixture verde do início ao fim). Suítes finais: core
**889** / Excel **24** / 0 warnings. Fecha a causa dominante dos 301 diffs do K1 (Excel: 2/3/5/2 nos
repros); o resíduo esperado é só o OR/AND literal (decisão de oráculo separada).

## Deployment Plan
Executado em 2026-07-03: verificação independente por fase → rebase de `feat/mini-cse-validation`
(cadeia A→B→C) sobre a main (1 conflito trivial em `Program.cs`, flags de harness coexistem) → ff-merge
→ sanidade na main (889/24, 0 warnings) → push → `merge-base --is-ancestor` em chamada separada →
`gh workflow run release.yml` (versionize derivou **3.1.0** do `feat(eval):`) → `git pull --tags` →
refresh `docs/pt-BR/` via Sonnet. Futuras releases: mesmo ritual, dispatch sempre isolado.
