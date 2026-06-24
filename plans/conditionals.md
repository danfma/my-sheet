# Condicionais e funções lógicas (IF / AND / OR / NOT / IFERROR)

Adicionar funções condicionais e lógicas ao motor de fórmulas, com avaliação curto-circuito nos ramos
de `IF`/`IFERROR`, e habilitar igualdade de texto nos comparadores `=`/`<>` (para destravar
`IF(A1="sim", …)`).

## Context

O parser/avaliador (ver `plans/expression-parser.md`) já suporta literais, referências, intervalos,
aritmética, comparadores numéricos e SUM/AVERAGE/MIN/MAX/COUNT. Falta o ramo condicional — `IF` e
companhia — que é central para qualquer planilha. O bloqueio prático: condições de `IF` quase sempre
comparam texto (`IF(A1="sim", …)`), mas hoje `=`/`<>` só fazem coerção numérica (texto → `#VALUE!`).
Por isso esta iteração inclui uma fatia mínima da paridade de comparadores (igualdade de texto), que
o backlog de Excel-parity já previa.

**Fora de escopo (adiado pelo usuário):** memorização/cache de células já computadas — será tratada
numa iteração própria ("mais complicada"). Também adiado: ordenação de texto (`< > <= >=` em strings),
ordenação cross-type do Excel, e expansão de intervalo dentro de AND/OR.

### Decisões fechadas com o usuário
- Funções desta iteração: **IF, AND, OR, NOT, IFERROR**.
- **Igualdade de texto** em `=`/`<>` (case-insensitive, como Excel); só igualdade, sem ordenação de texto.

### Decisões de design (assumidas — Excel-like)
- **Curto-circuito**: `IF(cond, a, b)` só computa o ramo escolhido; `IFERROR(v, alt)` só computa `alt`
  se `v` for erro. Ex.: `IF(A1=0, 0, 1/A1)` não dá `#DIV/0!` quando `A1=0`.
- **Veracidade da condição** (`ValueCoercion.TryToBool`): `null`(blank)→false; `bool`→ele mesmo;
  `double`→(≠0); `string`→`#VALUE!`; `ErrorValue`→propaga.
- **Igualdade** (`ValueCoercion.AreEqual`): ambos número→`==`; ambos bool→`==`; ambos string→
  `OrdinalIgnoreCase`; tipos diferentes não-blank→false (`1="1"`→false, como Excel); blank equivale a
  `0`/`""`/`false` do outro lado (`A1=0`/`A1=""` com A1 vazio → true).
- **Aridade**: validada no PARSE; aridade errada → `ParseException` (Excel rejeita na entrada).
  `IF`=2–3, `NOT`=1, `IFERROR`=2, `AND`/`OR`=≥1; SUM/AVERAGE/MIN/MAX/COUNT permanecem variádicas (0–∞).
- **Invariantes herdados**: `Compute` continua `object?`; tags `MemoryPackUnion` append-only (novas 14+);
  cada novo nó é `[MemoryPackable] sealed partial record : Function`.

## For Future Agents
Marque `- [x]` ao concluir; ao fechar uma fase mude o Status para `Complete`, escreva o **Phase Summary**
e rode o **Verification Plan**. Use TDD (teste falha → implementa → passa).
Comando de teste: `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` (73 testes verdes hoje).

---

## Phase 1: Igualdade de texto nos comparadores
Status: Complete

Fundação para condições de texto. Refatora o caminho de comparação de `BinaryOperation` sem mexer na
aritmética nem na ordenação numérica.

- [x] Adicionar `ValueCoercion.AreEqual(object? left, object? right) -> bool` com as regras de igualdade
      acima (número/bool/string + equivalência de blank a `0`/`""`/`false`; tipos mistos → false).
- [x] Refatorar `BinaryOperation.Compute`: avaliar Left/Right; se algum for `ErrorValue`, propagar PRIMEIRO;
      `Equal` → `AreEqual`, `NotEqual` → `!AreEqual`; demais operadores (aritmética + ordenação `< > <= >=`)
      seguem pelo caminho numérico atual via `ValueCoercion.TryToNumber` (texto em ordenação ainda → `#VALUE!`).
- [x] Bug correlato corrigido: `CellReference.Compute` usava `.Cells[Id]` (indexer cru do `Dictionary`),
      que lançava `KeyNotFoundException` ao referenciar célula vazia. Passou a usar o indexer do `Sheet`
      (devolve `BlankValue`). Removido o método morto `CellReference.Resolve` (sobra do `Sum` antigo).

### Verification Plan
- Novos/ajustados testes em `tests/MySheet.Tests/Parsing/ExpressionParserTests.cs` (ou um
  `ComparatorTests`):
  - Regressão numérica: `=2=2`→true, `=2<>3`→true, `=1<2`→true (ordenação intacta).
  - Texto: `="a"="A"`→true (case-insensitive), `="a"="b"`→false, `="abc"<>"abd"`→true.
  - Tipos mistos: `=1="1"`→false.
  - Blank: com `A1` vazio, `=A1=0`→true e `=A1=""`→true.
  - Ordenação de texto continua erro: `="a"<"b"`→`#VALUE!`.
  - Propagação: `=(1/0)=1`→`#DIV/0!`.
- `dotnet build MySheet/MySheet.csproj` → 0 Warning(s); suíte completa verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 2: Funções IF / AND / OR / NOT / IFERROR
Status: Complete

- [x] `ValueCoercion.TryToBool(object?, out bool) -> ErrorValue?` (blank→false; bool→ele mesmo;
      número→≠0; texto/outros→`#VALUE!`; erro→propaga).
- [x] Records em `MySheet/Expressions/`: `If`, `And`, `Or`, `Not`, `IfError` (tags **14–18**),
      com curto-circuito em `If`/`IfError`.
- [x] Aridade no `Parser.cs`: `private readonly record struct FunctionSpec(int MinArgs, int MaxArgs, Func<…> Create)`;
      registro com IF(2,3)/AND(1,∞)/OR(1,∞)/NOT(1,1)/IFERROR(2,2) e SUM/AVERAGE/MIN/MAX/COUNT(0,∞).
- [x] `ParseFunctionCall`: desconhecida → `ErrorValue.Name`; conhecida com contagem fora de [Min,Max] →
      `ParseException` (posição do nome).
- [x] Factories não foram necessárias (testes usam `ExpressionParser.Parse`).

### Verification Plan
- Novos testes (`ConditionalTests` em `tests/MySheet.Tests/Parsing/`):
  - IF: `=IF(1>0,10,20)`→10; `=IF(0>1,10,20)`→20; `=IF(1>0,99)`→99; `=IF(0>1,99)`→false (bool);
    curto-circuito `=IF(A1=0,0,1/A1)` com A1=0 → 0 (sem `#DIV/0!`); texto `=IF(A1="sim",1,0)` com A1="sim"→1.
  - Veracidade: `=IF(2,1,0)`→1; `=IF(0,1,0)`→0; `=IF("x",1,0)`→`#VALUE!`; `=IF(1/0,1,0)`→`#DIV/0!`.
  - AND/OR/NOT: `=AND(1>0,2>1)`→true; `=AND(1>0,2<1)`→false; `=OR(1<0,2>1)`→true; `=NOT(1>0)`→false.
  - IFERROR: `=IFERROR(1/0,-1)`→-1; `=IFERROR(5,-1)`→5; curto-circuito (alt só computa em erro).
  - Aridade (`ParseException`): `=IF(1)`, `=IF(1,2,3,4)`, `=NOT(1,2)`, `=IFERROR(1)`.
  - Regressão: `=SUM()`→0 e `=FOO(1)`→`#NAME?` continuam (registro com aridade não quebra variádicas).
- `dotnet build MySheet/MySheet.csproj` → 0 Warning(s); suíte completa verde.
- (Opcional, performance) As novas funções não usam `List`/agregação e fazem curto-circuito → alocação
  só do bool/resultado boxed. Re-rodar `dotnet run -c Release --project benchmarks/MySheet.Benchmark`
  se quiser registrar baseline de um `IF` aninhado.

### Phase Summary
Concluída em TDD. **102 testes, 102 verdes**; `MySheet.csproj` com 0 warnings. Novos arquivos:
`MySheet/Expressions/{If,And,Or,Not,IfError}.cs`; `TryToBool`/`AreEqual` em `ValueCoercion.cs`;
`FunctionSpec` + validação de aridade em `Parser.cs` (tags 14–18). Teste novo: `ConditionalTests`.
Curto-circuito confirmado por teste (`IF(A1=0,0,1/A1)`→0; `IFERROR(5,1/0)`→5). AND/OR/NOT/IFERROR
funcionando; aridade errada lança `ParseException`.

---

## Final Recap
`IF`, `AND`, `OR`, `NOT`, `IFERROR` adicionados, com curto-circuito (Excel-like) nos ramos de IF/IFERROR,
e igualdade de texto (`=`/`<>`, case-insensitive) habilitada para condições. Aridade validada no parse
(erro → `ParseException`). Bug pré-existente corrigido no caminho: referência a célula vazia
(`CellReference.Compute`) lançava `KeyNotFoundException`; agora é blank. Método morto `Resolve` removido.
**Suíte final: 102/102 verde, 0 warnings** nos projetos. Memorização foi avaliada (ver `plans/memoization.md`)
e adiada por decisão do usuário; será a próxima iteração.

Limitações conhecidas (adiadas para a iteração de Excel-parity):
- Ordenação de texto/cross-type em `< > <= >=` (só igualdade de texto entrou).
- AND/OR avaliam cada arg via `TryToBool` (blank→false) e NÃO expandem intervalos nem ignoram blanks
  em referências como o Excel; arg `RangeReference` em AND/OR vira `#VALUE!`.

## Deployment Plan
Biblioteca — sem infraestrutura. Para integrar:
1. `dotnet build` (projetos) → 0 Warning(s)/Error(s).
2. `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → 102/102 verde.
3. Commit em branch + PR (ou na `main`, conforme o padrão do repo). Mensagem sugerida:
   `feat(expressions): add IF/AND/OR/NOT/IFERROR and text equality`.

## Notas de paridade adiadas (próxima iteração)
- Ordenação de texto e cross-type em `< > <= >=` (ver backlog em `plans/expression-parser.md`).
- AND/OR ignorando blanks/texto em referências e aceitando intervalos (`AND(A1:A3)`); nesta iteração
  AND/OR avaliam cada arg via `TryToBool` (blank→false) e um arg `RangeReference` vira `#VALUE!`.
- **Memorização/cache de células computadas** — adiada explicitamente pelo usuário.
