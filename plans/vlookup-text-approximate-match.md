# Approximate-match com chaves de TEXTO (VLOOKUP / MATCH / XLOOKUP)

Corrigir o approximate/closest-match das funções de lookup para comparar chaves de
**texto** (e tipos mistos) com a mesma ordenação Excel-style já usada pelos
operadores `< > <= >=`, eliminando o `#N/A` espúrio quando a chave é uma string.

## Contexto / Causa raiz (confirmada — [Certeza])

Os caminhos de approximate-match comparavam a chave **apenas quando ela era `double`**,
ignorando silenciosamente chaves de texto. O engine já tinha o helper correto —
`ValueCoercion.Compare(left, right)` (`Danfma.MySheet/Expressions/ValueCoercion.cs:153`),
ordenação Excel-style cross-type (número < texto < booleano; texto case-insensitive),
o mesmo que `BinaryOperation` usa para `<= >= < >`
(`Danfma.MySheet/Expressions/BinaryOperation.cs:51-58`). O fix foi rotear os três
caminhos de approximate-match por esse helper.

Repro do bug reportado: `VLOOKUP("Bradbury Creek", Setup!A2:D42, 2, TRUE)` retornava
`#N/A` mesmo com a chave presente na 1ª linha de uma tabela ordenada A→Z; o `#N/A`
propagava por `SUM(B36:B40)` → `total_monthly`.

Os três pontos com o mesmo defeito:
- `Danfma.MySheet/Expressions/VLookup.cs` — `is double key && key <= target`; ainda descartava o retorno de `TryToNumber(lookup, ...)` (texto → `target = 0` silencioso).
- `Danfma.MySheet/Expressions/Match.cs` — coagia `lookup` a número (retornava erro se fosse texto) **e** pulava células com `array[i] is not double value`.
- `Danfma.MySheet/Expressions/XLookup.cs` (`Closest`) — retornava `-1` se `lookup` não era número, **e** pulava células `is not double value`.

## Decisão de escopo

Confirmado com o usuário: corrigidas **as três** funções no mesmo trabalho (raiz
completa), não só a VLOOKUP reportada.

## For Future Agents
Work concluída. Artefato mantido como registro resumível.

## Phase 1: Testes que reproduzem o bug (TDD red)
Status: Complete

- [x] Helper `CalcMixed((string Id, object Value)[])` em `LookupFunctionTests.cs` (seta `StringValue`/`NumberValue` conforme o tipo; `StringValue` qualificado como `Danfma.MySheet.Expressions.StringValue` por causa do conflito com `TUnit.Assertions.StringValue`).
- [x] `VLookup_ApproximateTextKey`: chave de texto na 1ª linha (o caso reportado), chave intermediária, chave "entre nomes" ("Cz" → "Cedar Falls") e chave abaixo da menor (→ `#N/A`).
- [x] `Match_ApproximateTextKey`: `MATCH("Cz",...,1)` → 2; `MATCH("Bradbury Creek",...,1)` → 1.
- [x] `XLookup_ClosestTextKey`: `XLOOKUP("Cz",...,-1)` → next-smaller (10.0); `XLOOKUP("Cz",...,1)` → next-larger (20.0).
- [x] Confirmado: os 3 testes falham com `#N/A`/vazio antes do fix; 7 pré-existentes passam.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/LookupFunctionTests/*"` — esperado: novos testes **falham** antes do fix.

### Phase Summary
RED estabelecido. Descoberta operacional importante: o projeto usa **Microsoft.Testing.Platform (TUnit)**, não VSTest — `dotnet test --filter` NÃO funciona (erro MTP). Rodar testes via `dotnet run --project <test.csproj> -- --treenode-filter "/*/*/Classe/*"` (CI usa `dotnet run ... --no-build`). Também: `StringValue` é ambíguo com TUnit → qualificar. Resultado RED: 3 falham, 7 passam (total 10).

## Phase 2: Fix via `ValueCoercion.Compare` (TDD green)
Status: Complete

- [x] `VLookup.cs` — bloco `approximate` itera linhas e mantém a última onde `Compare(key, lookup) <= 0`; removida a coerção a número; pula `null`/`ErrorValue`; propaga `lookup` se for `ErrorValue`.
- [x] `Match.cs` — removida coerção/erro do `lookup` a número; `matchType>0` → última posição com `Compare(value,lookup) <= 0`; `matchType<0` → `>= 0`; pula `null`/`ErrorValue`; propaga `lookup` `ErrorValue`.
- [x] `XLookup.cs` (`Closest`) — removido early-return de `lookup` numérico; melhor rastreado via `Compare` (`below` = maior `value` ≤ lookup; `!below` = menor `value` ≥ lookup); pula `null`/`ErrorValue`.
- [x] Nenhum caminho approximate coage texto a número.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/LookupFunctionTests/*"` — esperado: **todos** passam.

### Phase Summary
GREEN: 10/10 testes de Lookup passam. Para tipos puramente numéricos, `Compare` é idêntico à comparação numérica anterior (mesma semântica de empate: mantém a primeira ocorrência), então não há mudança de comportamento no caminho numérico. Adicionada propagação de `ErrorValue` no `lookup` (Excel-correto) para evitar que um erro seja classificado como texto.

## Phase 3: Verificação completa / sem regressão
Status: Complete

- [x] Suíte inteira: **177/177 passam**.
- [x] Build limpo: `Build succeeded`, 0 warnings, 0 errors.

### Verification Plan
- `dotnet build Danfma.MySheet/Danfma.MySheet.csproj` — `Build succeeded`, 0 erros. ✅
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj` — 177/177 verdes. ✅

### Phase Summary
Sem regressão. O fix toca 3 arquivos de produção + 1 de teste, todos sob a mesma causa raiz, reusando `ValueCoercion.Compare` (sem reimplementar ordenação).

## Final Recap

O approximate-match de `VLOOKUP`/`MATCH`/`XLOOKUP` só comparava chaves `double`,
retornando `#N/A` (ou `#VALUE!` no MATCH) para chaves de texto. A correção rotea os
três caminhos por `ValueCoercion.Compare` — a mesma ordenação cross-type Excel-style
que os operadores `<= >= < >` já usavam — fazendo texto ordenar lexicograficamente
(case-insensitive). Cobertura por TDD (red → green), suíte 177/177 verde, build limpo.

Arquivos alterados:
- `Danfma.MySheet/Expressions/VLookup.cs`
- `Danfma.MySheet/Expressions/Match.cs`
- `Danfma.MySheet/Expressions/XLookup.cs`
- `tests/Danfma.MySheet.Tests/Parsing/LookupFunctionTests.cs` (helper de texto + 3 testes)

## Deployment Plan

Biblioteca NuGet — sem migração de dados nem passos de runtime.

1. Revisar o diff: `git diff`.
2. Commit (sem co-author, conforme regra do projeto):
   `git commit -am "fix(lookup): approximate-match compara chaves de texto (VLOOKUP/MATCH/XLOOKUP)"`.
3. Push para `main` → CI roda `dotnet run --project tests/... --no-build` (deve ficar verde).
4. Release/versão: bump conforme convenção do repo (correção de bug → patch, ex.: 0.1.0 → 0.1.1) e seguir o fluxo de release existente (tag → publish NuGet via OIDC).
