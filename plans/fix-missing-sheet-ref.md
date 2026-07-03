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
      HLOOKUP dão #REF! de graça (o guard de `TryGetPopulatedBounds` faz `ToBoundedRange`→null→open range→não
      é `RangeReference`→#REF!). OFFSET dá #REF! via o próprio `TryResolveReference`. Long tail (SUMPRODUCT,
      MATCH/XLOOKUP/INDEX, financeiras, texto) NÃO lança (safety net dos enumeradores) — degrada, fora da
      matriz.
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
Status: Not started

**Documentação: usar a skill `code-documentation-doc-generate`.**

- [ ] Teste de regressão do cenário REAL do report: workbook com fórmulas cross-sheet penduradas; avaliar
      TODAS as células num loop (`GetCellValue` por célula) — asserir **nenhuma exceção** e cada célula
      pendurada = #REF!. (Prova que o batch de ~495k não aborta mais.)
- [ ] Interop: `.xlsx` com fórmula que referencia sheet ausente → `Load` + avaliação = #REF! (não lança).
      Teste com fixture montada.
- [ ] `docs/workbook-and-expressions.md`: a seção de erros semânticos confirma "referência a sheet
      inexistente → #REF! (não lança)"; documentar `TryGetSheet`. `README.md` se citar semântica de erro.
      NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [ ] Plano: fases Complete + Phase Summary + Final Recap.

### Verification Plan
- Build `--no-incremental` 0 warnings; ambas as suítes verdes; o teste de batch prova zero-throw.

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando as fases 1–2 concluírem)_

## Deployment Plan
Fluxo por **PR** (decisão do usuário, 2026-07-03 — não fazer ff direto na main desta vez):
1. Verificação independente minha com rebuild forçado (`--no-incremental`): a matriz de contrato inteira, os
   dois lados da distinção estrutural-vs-valor, fixture binária, o teste de batch sem throw.
2. `git push origin fix/missing-sheet-ref` (a branch, não a main).
3. `gh pr create` — título/descrição referenciando o bug report (sheet inexistente → #REF!), a matriz de
   contrato e os 9 sites; corpo com o rodapé de PR do harness. Deixar para o usuário revisar e mergear.
4. Após o merge do PR pelo usuário: `gh workflow run release.yml` → **2.6.1** (patch) lockstep → `git pull`
   → refresh `docs/pt-BR/` via sub-agente Sonnet.
