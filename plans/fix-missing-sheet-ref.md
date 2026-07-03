# Bug: referência a sheet inexistente lança KeyNotFoundException em vez de #REF!

Avaliar uma fórmula que referencia um sheet inexistente (`=Ghost!A1`, `SUM(Ghost!A:A)`) lança
`KeyNotFoundException` de dentro de `Workbook.GetCellValue`/enumeração de referências, abortando um batch
inteiro. Deve resolver a **#REF!** (fiel ao Excel e ao contrato documentado: "bad reference → #REF!"),
por célula, sem lançar. Bug report: `~/Downloads/MYSHEET-ISSUE-missing-sheet-ref.md`. Fix → patch **2.6.1**.

## Decisões travadas (usuário, 2026-07-03)
- **#REF! SEMPRE, fiel ao Excel** — sheet inexistente é falha ESTRUTURAL da referência, propagada por TODA
  função consumidora, INCLUSIVE COUNT/COUNTA/COUNTBLANK/COUNTIF (que normalmente ignoram erros de VALOR de
  célula). NÃO tratar sheet-fantasma como "range vazio" (esconderia o erro → SUM=0). A distinção é o cerne:
  - sheet-fantasma → #REF! propaga por tudo (estrutural);
  - célula com `#DIV/0!` DENTRO de um range de sheet existente → COUNT/COUNTIF seguem IGNORANDO (política de
    valor por função, inalterada).
- **`Workbook.TryGetSheet(string name, out Sheet sheet)` público** (aditivo) para o host consultar sem
  try/catch. O indexer `workbook[key]` CONTINUA lançando `KeyNotFoundException` (acesso direto do host, não
  é avaliação — como um `Dictionary`; não-breaking).

## Escopo do bug — os 9 sites que indexam `Sheets[...]` (mapeados 2026-07-03)
O report só cita `GetCellValue`, mas a enumeração de coluna inteira (`PopulatedIds`) indexa o sheet ANTES de
qualquer `GetCellValue` — por isso `COUNTIF(Ghost!D:D,1)` estoura mesmo se só `GetCellValue` for corrigido.
Todos precisam do guard:
1. `Workbook.cs:224` — `GetCellValue` (via `CellReference.Evaluate`). Sheet ausente → `#REF!`.
2. `RangeReference.cs:22` — `Expand` (Expression view) indexa `Sheets[SheetName]`.
3. `OpenRangeReference.cs:73` — `PopulatedIds` (a fonte da enumeração de coluna/linha).
4. `OpenRangeReference.cs:87` — `Expand`.
5. `OpenRangeReference.cs:125` — `TryGetPopulatedBounds` (usado por `ToBoundedRange`, ROWS/COLUMNS, lookups).
6-8. `Subtotal.cs:54,82,96` — scan de nested-subtotal (range, open, cell).
9. `Workbook.cs:188` — `this[key]` público → **mantém lançando** (decisão); NÃO alterar.
`RangeReference.ExpandComputedValues` (:56,:85) já passa por `GetCellValue`, então o guard de (1) o cobre —
MAS confirmar que ele não indexa o sheet antes (não indexa: chama `GetCellValue` por id direto).

## O ponto arquitetural (por que não é só um try/catch)
Modelar sheet-fantasma como "um #REF! no meio do stream de células" NÃO funciona para COUNT-family: elas
ignoram erros do stream (correto para `#DIV/0!` numa célula). Sheet-fantasma é erro do ARGUMENTO, não de uma
célula — tem de ser detectado ANTES da enumeração e curto-circuitar a função para #REF!. A prova de que o
fix está certo é a **matriz de contrato** (abaixo): as mesmas funções devem dar #REF! sobre sheet-fantasma
E manter a política de erro-de-valor sobre sheet existente.

## For Future Agents
Marque `- [x]`; ao fechar fase: Status `Complete` + Phase Summary + Verification. TDD: escreva a matriz de
contrato RED primeiro. **Build de verificação SEMPRE `--no-incremental`.** Fixture
`workbook-pre-namespaces.msgpack.bin`/`MemoryPackCompatibilityTests` intocáveis. Suítes hoje: core **737**,
Excel **22**; 0 warnings. Commits inglês, semantic (`fix(refs): ...`), SEM atribuição a IA. NÃO push.

---

## Phase 1: Resolução de sheet ausente → #REF! (a correção)
Status: Complete

- [x] `Workbook.TryGetSheet(string name, out Sheet sheet)` público (usa o `Sheets` dictionary; case-insensitive
      como o resto). `GetCellValue`: se `!TryGetSheet(sheetName, ...)` → `ComputedValue.Error(Error.Ref)`
      (antes de indexar). Cobre `CellReference` e os paths por-célula (`RangeReference.ExpandComputedValues`).
- [x] Guardar os sites de ENUMERAÇÃO que indexam o sheet direto para ler `.Keys`/`.Cells`
      (`RangeReference.Expand`; `OpenRangeReference.PopulatedIds`/`Expand`/`TryGetPopulatedBounds`;
      `Subtotal` — short-circuit em `Evaluate`): sheet ausente → não lança (yield break / false); a curto-
      circuitação estrutural fica nos choke points.
- [x] **Propagação estrutural**: helper `ReferenceGuard.MissingSheet(arg|args, context) : Error?`
      (reconhece CellReference/RangeReference/OpenRangeReference/UnionReference — todas as áreas — e
      NameReference que resolve para uma dessas). Checado nos choke points: `NumericAggregation.Fold`/`FoldA`
      (cobre SUM/AVERAGE/MIN/MAX/PRODUCT/STDEV/VAR/… que propagam via o canal de erro), e explícito nas
      funções que IGNORAM o canal de erro ou consomem streams: COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS,
      SUMIF, SUMIFS, AVERAGEIF(+CriteriaPairs para AVERAGEIFS/MAXIFS/MINIFS), ROWS, COLUMNS, SUBTOTAL. VLOOKUP/
      HLOOKUP dão #REF! de graça sobre um OPEN range fantasma (o guard de `TryGetPopulatedBounds` faz
      `ToBoundedRange`→null→open range→não é `RangeReference`→#REF!). OFFSET dá #REF! via o próprio
      `TryResolveReference`. A "long tail" (SUMPRODUCT, MATCH/XLOOKUP, financeiras, texto, e VLOOKUP/HLOOKUP
      sobre um range BOUNDED fantasma) ficou coberta na **Phase 3** — paridade total com o Excel.
- [x] **Matriz de contrato (o teste que prova a semântica) — RED primeiro:**
      | Fórmula | Esperado |
      |---|---|
      | `=Ghost!A1` | #REF! (kind Error) |
      | `=UPPER(Ghost!E9)` | #REF! |
      | `=Ghost!A1 + 1` | #REF! (propaga por operador) |
      | `=SUM(Ghost!A:A)` | #REF! |
      | `=COUNT(Ghost!A:A)` | #REF! |
      | `=COUNTA(Ghost!A:A)` | #REF! |
      | `=COUNTIF(Ghost!D:D,1)` | #REF! |
      | `=SUMIF(Ghost!A:A,">5")` | #REF! |
      | `=AVERAGE(Ghost!B2:B9)` | #REF! |
      | `=VLOOKUP(1,Ghost!A:B,2)` | #REF! |
      | `=ROWS(Ghost!A:A)` | #REF! |
      | `=SUM(Main!A:A, Ghost!A:A)` (união com 1 fantasma) | #REF! |
      | **Controles (sheet EXISTE — política de valor inalterada):** | |
      | `=COUNTIF(D:D,1)` (mesma sheet, sem match) | 0 |
      | `=COUNT(A:A)` com `A1=1/0` (célula com erro) | conta os números, IGNORA o erro |
      | `=SUM(A:A)` com `A1=1/0` | #DIV/0! (SUM propaga erro de valor — comportamento atual) |
      | `=Main!A2` (existente, vazia) | Blank |
- [x] Nenhuma célula lança: `MemoryPackCompatibilityTests` verde; suíte inteira verde.

### Verification Plan
- `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` → 0 warnings.
- Suíte core verde incl. a matriz nova; Excel (22) intacta; fixture verde.

### Phase Summary
Matriz de contrato em `tests/…/Expressions/MissingSheetReferenceTests.cs` (20 casos): escrita RED primeiro
(15 falhas, KeyNotFoundException), depois GREEN. Mecanismo:
- `Workbook.TryGetSheet` (público, aditivo) + guard em `GetCellValue` (sheet ausente → `#REF!` cacheado,
  antes de indexar). O indexer `this[key]` continua lançando (decisão).
- `ReferenceGuard.MissingSheet` (novo, `Expressions/ReferenceGuard.cs`) — detecção estrutural pré-enumeração.
- Enumeradores no-throw: `RangeReference.Expand`, `OpenRangeReference.PopulatedIds`/`Expand`/
  `TryGetPopulatedBounds` (`TryGetValue`→yield break/false). Subtotal curto-circuita em `Evaluate`.
- Choke `Fold`/`FoldA` + guards explícitos nas funções error-ignoring/stream (ver acima).
Prova da distinção: COUNT/COUNTIF sobre sheet-fantasma → `#REF!`; COUNT sobre célula `#DIV/0!` em sheet
existente → ignora (conta 2); SUM sobre a mesma → propaga `#DIV/0!`. Build `--no-incremental`: 0 warnings.
Suítes: core **757** (737 + 20), Excel **22** — todas verdes.

---

## Phase 2: Regressão de batch + interop + docs
Status: Complete

**Documentação: usar a skill `code-documentation-doc-generate`.**

- [x] Teste de regressão do cenário REAL do report: `Batch_WithDanglingRefs_NeverThrows` (inclui
      `=UPPER(BOX11MNO_HIDE!E9)`, `=SUM(Ghost!A:A)`, `=COUNTIF(Ghost!D:D,1)`) — loop `GetCellValue` por
      célula, asserindo nenhuma exceção e cada célula pendurada = #REF!, células boas normais.
- [x] Interop: `MissingSheetInteropTests.Load_FormulaReferencingMissingSheet_EvaluatesToRef` — fixture
      ClosedXML com `Ghost!A1`/`SUM(Ghost!A:A)`/`UPPER(Ghost!E9)` → `ExcelFile.Load` + avaliação = #REF!,
      sem lançar; célula bem-formada no mesmo workbook avalia normal.
- [x] `docs/workbook-and-expressions.md`: seção de erros semânticos confirma sheet inexistente → #REF!
      (não lança) e a distinção estrutural vs valor-de-célula; tabela de membros + bullets documentam
      `TryGetSheet` e o contraste com o indexer que lança. `README.md`: bullet de References atualizado.
      `docs/pt-BR/` intocado (refresh no deploy).
- [x] Plano: fases Complete + Phase Summary + Final Recap.

### Verification Plan
- Build `--no-incremental` 0 warnings; ambas as suítes verdes; o teste de batch prova zero-throw.

### Phase Summary
Regressão de batch e interop verdes. Docs via skill `code-documentation-doc-generate`: patch cirúrgico em
`docs/workbook-and-expressions.md` (erros semânticos + `TryGetSheet` na tabela/bullets) e `README.md`
(References). Build `--no-incremental`: 0 warnings. Core **757**, Excel **23** (22 + interop) — verdes.

## Phase 3: Excel-parity da cauda (long tail)
Status: Complete

Estende o #REF! estrutural a TODA função consumidora de referência que ainda degradava para `0`/`#N/A`/`#VALUE!`/
`#NUM!`/`""` sobre um sheet-fantasma (o safety-net dos enumeradores não bastava: um OPEN range fantasma vira
uma sequência VAZIA — silenciosamente engolido — e um BOUNDED range fantasma vira células `#REF!` que os
lookups descartavam, degradando para `#N/A`). Mecanismo uniforme, baixo risco: `if
(ReferenceGuard.MissingSheet(Arguments, context) is { } m) return ComputedValue.Error(m);` no TOPO de cada
`Evaluate`, ANTES da política per-célula — mais dois choke points compartilhados.

- [x] Auditoria exaustiva: `grep` por `ExpandComputedValues`/`FlattenComputedValues`/`PairwiseRanges`/`Fold`,
      cruzando com quem já tinha `ReferenceGuard`. Achado-chave: a família order/dispersion (MEDIAN, MODE,
      STDEV*, VAR*, PERCENTILE*, QUARTILE*, RANK*, LARGE, SMALL, SKEW*, KURT, GEOMEAN, HARMEAN, AVEDEV, DEVSQ,
      TRIMMEAN, PERCENTRANK*, + aliases legados) JÁ estava coberta de graça, pois `StatisticsMath.Collect`/
      `CollectA` roteiam por `NumericAggregation.Fold`/`FoldA` (já guardados na Phase 1). Idem PRODUCT/SUMSQ/
      MULTINOMIAL/GCD/LCM.
- [x] Choke point 1 — `PairwiseRanges.Expand`: guard em x e y ⇒ cobre SUMX2MY2/SUMX2PY2/SUMXMY2 e TODA a
      bivariada (CORREL, PEARSON, COVARIANCE.P/S, RSQ, SLOPE, INTERCEPT, STEYX, FORECAST.LINEAR) + os aliases
      legados COVAR e FORECAST (chamam os `Compute` internos, que passam pelo mesmo helper).
- [x] Early-return `ReferenceGuard.MissingSheet(Arguments, …)` explícito nas demais:
      SUMPRODUCT; PROB; MATCH, XLOOKUP, XMATCH, LOOKUP, VLOOKUP, HLOOKUP (os dois últimos fechando o gap do
      range BOUNDED fantasma); NPV, IRR, XNPV, XIRR, MIRR, FVSCHEDULE; NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY,
      WORKDAY.INTL; SERIESSUM; TEXTJOIN, CONCAT, CONCATENATE; AND, OR, XOR; ROW, COLUMN.
- [x] Matriz de contrato estendida em `MissingSheetReferenceTests` (RED primeiro: 29 falhas → GREEN):
      `SUMPRODUCT(Ghost!A:A,Main!A:A)`, `SUMPRODUCT(Main!A1:A2,Ghost!B1:B2)`, `SUMX2MY2`, `CORREL`, `COVAR`,
      `FORECAST`, `PROB`, `MATCH(1,Ghost!A:A)`, `XLOOKUP`, `XMATCH`, `LOOKUP`, `INDEX`, `VLOOKUP(1,Ghost!A1:B5,2)`,
      `HLOOKUP(1,Ghost!A1:B5,2)`, `TEXTJOIN`, `CONCAT`, `CONCATENATE`, `NPV(0.1,Ghost!A:A)`, `NPV(Ghost!A1,…)`,
      `IRR`, `MIRR`, `XNPV`, `XIRR`, `FVSCHEDULE`, `NETWORKDAYS(…,Ghost!C:C)`, `WORKDAY(…,Ghost!C:C)`, `SERIESSUM`,
      `AND`/`OR`/`XOR(Ghost!A:A)`, `ROW`/`COLUMN(Ghost!A1)`, `MEDIAN`, `PERCENTILE` — todos #REF!.
- [x] Controles de NÃO-regressão adicionados e verdes: `MATCH(1,A:A)` (sheet existente, sem match) → #N/A;
      `XLOOKUP(1,A:A,B:B)` (existente, vazio) → #N/A; `SUMPRODUCT(A1:A2,B1:B2)` → número; `CORREL` normal →
      número. A distinção "sheet-fantasma = #REF! estrutural" vs "range vazio em sheet existente = #N/A/valor"
      segue intacta.

### NÃO tocado (com justificativa)
- **Família error-inspecting** (ISERROR/ISERR/ISNA/ISREF/ISFORMULA/N/TYPE/ERROR.TYPE, IFERROR/IFNA): DEVEM
  capturar o #REF! de uma referência fantasma, não propagá-lo — `ISERROR(Ghost!A1)` = TRUE, fiel ao Excel.
  Uma célula-fantasma já avalia para #REF! (Phase 1) e elas o inspecionam; um guard as quebraria.
- **CHOOSE**: lazy por contrato (só o argumento escolhido é avaliado); um guard em todos os args quebraria a
  preguiça. Uma referência fantasma escolhida é resolvida por quem a consome (agora guardado).
- **INDEX / OFFSET**: já retornam #REF! por si — INDEX sobre bounded fantasma devolve a célula #REF! e sobre
  open fantasma cai no `ToBoundedRange`→null→#REF!; OFFSET resolve a base por `TryResolveReference`.
- **Order/dispersion stats + PRODUCT/SUMSQ/GCD/LCM/MULTINOMIAL**: já cobertos via `Fold`/`FoldA` (Phase 1).

### Verification
Build `Danfma.MySheet.slnx -c Release --no-incremental` → **0 warnings**. Core **795** (todas verdes, incl. a
matriz estendida e os controles), Excel **23**, fixture `MemoryPackCompatibilityTests` verde e intocada.

### Phase Summary
Paridade total com o Excel para sheet-fantasma: QUALQUER função que consome uma referência a um sheet
inexistente retorna #REF!. Funções tocadas (guard explícito, salvo onde indicado): SUMPRODUCT; PROB; MATCH,
XLOOKUP, XMATCH, LOOKUP, VLOOKUP, HLOOKUP; NPV, IRR, XNPV, XIRR, MIRR, FVSCHEDULE; NETWORKDAYS,
NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL; SERIESSUM; TEXTJOIN, CONCAT, CONCATENATE; AND, OR, XOR; ROW, COLUMN; e
— via o choke point `PairwiseRanges.Expand` — SUMX2MY2, SUMX2PY2, SUMXMY2, CORREL, PEARSON, COVARIANCE.P,
COVARIANCE.S, RSQ, SLOPE, INTERCEPT, STEYX, FORECAST.LINEAR, COVAR, FORECAST. Nenhuma cirurgia por-célula:
early-return de 2 choke points + um guard de 4 linhas no topo de cada `Evaluate`. Suítes: core **795**,
Excel **23**, 0 warnings.

## Final Recap
Bug fechado: referência a sheet inexistente resolve para **#REF!** por célula, sem lançar
`KeyNotFoundException` — um batch de workbook inteiro não aborta mais. Mecanismo: `Workbook.TryGetSheet`
(público, aditivo) + guard em `GetCellValue`; helper `ReferenceGuard.MissingSheet` (detecção estrutural
pré-enumeração) checado nos choke points `NumericAggregation.Fold`/`FoldA` e explícito nas funções que
ignoram o canal de erro (COUNT/COUNTA/COUNTBLANK/COUNTIF/COUNTIFS, SUMIF/SUMIFS, AVERAGEIF(S)/MAXIFS/MINIFS,
ROWS/COLUMNS, SUBTOTAL); enumeradores (`RangeReference.Expand`, `OpenRangeReference.PopulatedIds`/`Expand`/
`TryGetPopulatedBounds`) tornados no-throw; VLOOKUP/HLOOKUP e OFFSET dão #REF! pelos seus próprios paths de
resolução. A distinção estrutural-vs-valor está provada na matriz: COUNT/COUNTIF sobre sheet-fantasma =
#REF!; COUNT sobre célula #DIV/0! em sheet existente ainda ignora; SUM propaga. Indexer `this[key]` continua
lançando (decisão). Fixture `workbook-pre-namespaces.msgpack.bin`/`MemoryPackCompatibilityTests` intocáveis
e verdes. Suítes: core **757** (737 + 20), Excel **23** (22 + 1). Build `--no-incremental`: 0 warnings.
Commits (branch `fix/missing-sheet-ref`, sem push — release é gate do usuário): ver Deployment Plan.

## Deployment Plan
Fluxo por **PR** (decisão do usuário, 2026-07-03 — não fazer ff direto na main desta vez):
1. Verificação independente minha com rebuild forçado (`--no-incremental`): a matriz de contrato inteira, os
   dois lados da distinção estrutural-vs-valor, fixture binária, o teste de batch sem throw.
2. `git push origin fix/missing-sheet-ref` (a branch, não a main).
3. `gh pr create` — título/descrição referenciando o bug report (sheet inexistente → #REF!), a matriz de
   contrato e os 9 sites; corpo com o rodapé de PR do harness. Deixar para o usuário revisar e mergear.
4. Após o merge do PR pelo usuário: `gh workflow run release.yml` → **2.6.1** (patch) lockstep → `git pull`
   → refresh `docs/pt-BR/` via sub-agente Sonnet.
