# Biblioteca de funções + extensibilidade

Adicionar ~33 funções do Excel e um mecanismo de extensão para funções customizadas (de domínio),
faseado por complexidade. Built-ins permanecem `records` no union (sem quebrar serialização); funções
customizadas usam um nó genérico `FunctionCall` resolvido por um registro extensível.

## Context
O usuário precisa de uma lista grande de funções para processar planilhas (domínio financeiro/portfólio)
em background. Decisão: priorizar funções sobre memorização (esta fica em `plans/memoization.md`).

Decisões fechadas:
- **Built-ins = records** `: Function` (mantém o formato MemoryPack atual). **Custom = `FunctionCall(nome, args)`**
  genérico, resolvido em runtime por um registro extensível no `Workbook`.
- **Funções de domínio** (INCOME, PORTFOLIO, PROPERTY, GAIN, CONTRIBUTIONS, ORGANIZATION, DETAIL, HIDDEN,
  A_HIDE, EXPR_SUM) NÃO são implementadas aqui: o host as registra via o mecanismo de extensão (Fase 1).
- **Faseado**: arquitetura + matemática/texto/info/contagem condicional primeiro; lookup/referência, LET e
  TEXT-formatação depois.

Decisões de design (assumidas):
- Tokenizer aceita `_` e `.` em nomes (`A_HIDE`, `EXPR_SUM`, `XLFN.XLOOKUP`); prefixo `XLFN.`/`_xlfn.`
  (artefato do Excel) é normalizado para o nome real.
- Nome desconhecido no parse → `FunctionCall` (não `#NAME?`); resolve em runtime, `#NAME?` se não registrado.
  Retrocompatível: `=FOO(1)` continua `#NAME?` no `Compute`.
- Invariantes herdados: `Compute(Workbook) : object?`; tags `MemoryPackUnion` append-only (novas a partir de 19);
  novos built-ins seguem o padrão record-fino + helper estático compartilhado (como `NumericAggregation`).
- Cultura invariante em VALUE/TEXT; locale fica fora de escopo.

## For Future Agents
Marque `- [x]`; ao fechar fase, Status `Complete` + Phase Summary + rode a Verification.
TDD por função. Teste: `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` (102 verdes hoje).
**Paralelização**: após a Fase 1, as Fases 2/3/4 são independentes entre si (categorias distintas, mesmo
mecanismo) → podem rodar em subagentes paralelos. Dentro de uma fase, cada função é uma unidade isolada.
Cada novo built-in: `[MemoryPackable] sealed partial record X(Expression[] Arguments) : Function` + tag
append-only + entrada no `FunctionSpec` do `Parser` (com aridade).

---

## Phase 1: Arquitetura de extensão (fundação)
Status: Complete

Habilita tudo o resto. Entrega o mecanismo de extensão ("como fazer funções de usuário").

- [x] Tokenizer: identificadores aceitam início em `_` e meio com `_`/`.` (`A_HIDE`, `XLFN.XLOOKUP`);
      classificação de cell ref intacta (`^[A-Za-z]+[0-9]+$`).
- [x] `NormalizeFunctionName` no parser: strip de `_xlfn.`/`xlfn.` antes de resolver.
- [x] `FunctionCall(string Name, Expression[] Arguments) : Function` (`[MemoryPackable]`, tag **19**);
      `Compute` resolve `workbook.TryGetFunction` → invoca, senão `ErrorValue.Name`.
- [x] `delegate CustomFunction(Expression[], Workbook)` + `Workbook.RegisterFunction`/`TryGetFunction`
      (`[MemoryPackIgnore]` Dictionary, case-insensitive).
- [x] `ParseFunctionCall`: built-in → record + aridade; senão → `FunctionCall` (retrocompatível: `#NAME?`
      ainda surge no `Compute`).

### Phase Summary
**107/107 verde**, 0 warnings. Retrocompatível (102 testes antigos intactos; `=FOO(1)` segue `#NAME?`).
Novos: `FunctionCall.cs`, `CustomFunction` + registro no `Workbook.cs`, tag 19, tokenizer `_`/`.`,
`NormalizeFunctionName` no `Parser.cs`. Teste: `FunctionExtensionTests` (custom resolve/invoke, não
registrada → #NAME?, XLFN. normalizado, `A_HIDE`, round-trip de `FunctionCall`).
**Decisão de paralelização**: agentes paralelos produzem arquivos NOVOS disjuntos (record + teste por
categoria); a fiação central (tags em `Expression.cs`, entradas no `FunctionSpec` do `Parser.cs`) e a
verificação ficam com a thread principal — evita conflito de arquivo compartilhado e corrupção de `obj/`
por builds concorrentes. Tags reservadas: Fase 2 = 20–26, Fase 3 = 27–36, Fase 4 = 37–42.

### Verification Plan
- `dotnet build MySheet/MySheet.csproj` → 0 Warning(s); suíte atual continua verde (incl. `UnknownFunction_*`
  que ainda dá `#NAME?` no Compute).
- Novos testes: registrar `CUSTOM(x)` via `Workbook.RegisterFunction` e computar; não registrada → `#NAME?`;
  `XLFN.XLOOKUP`/`_xlfn.X` normalizado; identificador `A_HIDE`/`EXPR_SUM` tokenizado como 1 nome de função;
  round-trip `WorkbookTests` intacto (registro é `[MemoryPackIgnore]`).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 2: Matemática + info (fácil, paralelizável)
Status: Complete

- [x] Records + `FunctionSpec` (tags 20–26): `INT`(1), `ROUND`(2), `ROUNDUP`(2), `ABS`(1), `ISNUMBER`(1),
      `ISBLANK`(1), `IFNA`(2). Adicionado `ErrorValue.NotAvailable` (#N/A) para o IFNA.
- [x] `ISBLANK` = arg computa a `null`; `ISNUMBER` = arg computa a `double` (erro → false, não propaga);
      `ROUND` half-away-from-zero, `ROUNDUP` away-from-zero (ambos via fator `10^dígitos`).

### Phase Summary
**114/114 verde**, 0 warnings. Records seguem o padrão fino + `ValueCoercion.TryToNumber`.
Template para os agentes: `Int.cs`/`Round.cs` (math), `IsNumber.cs`/`IsBlank.cs` (info), `IfNa.cs`.
Adicionado `ValueCoercion.TryToText` (compartilhado) para as funções de texto da Fase 3.

### Verification Plan
- Testes: `=INT(2.9)`→2, `=INT(-2.1)`→-3, `=ROUND(2.345,2)`→2.35, `=ROUNDUP(2.001,2)`→2.01, `=ABS(-3)`→3,
  `=ISNUMBER(1)`→true, `=ISNUMBER("x")`→false, `=ISBLANK(A1)` (A1 vazio)→true, `=IFNA(NA-ish, 0)`.
- 0 Warning(s); suíte verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Texto (fácil, paralelizável)
Status: Complete

- [x] `ValueCoercion.TryToText` (número→invariant, bool→TRUE/FALSE, blank→"", erro propaga).
- [x] Records + `FunctionSpec` (tags 27–36): `UPPER`, `LOWER`, `TRIM` (colapsa espaços internos),
      `LEN`, `LEFT`(1–2), `MID`(3), `VALUE`, `CONCAT`(1–∞), `CONCATENATE`(1–∞), `TEXTJOIN`(3–∞).
- [x] Helper `ArgumentFlattening.Flatten` (expande ranges célula a célula) para CONCAT/TEXTJOIN.

### Phase Summary
**123/123 verde**, 0 warnings. Padrão record-fino + `ValueCoercion.TryToText`/`TryToNumber`.
`ArgumentFlattening` reaproveitável pela Fase 4 (COUNTA etc.).

### Verification Plan
- Testes: `=UPPER("ab")`→"AB", `=LOWER("AB")`→"ab", `=TRIM(" a b ")`→"a b", `=LEN("abc")`→3,
  `=LEFT("abcd",2)`→"ab", `=MID("abcd",2,2)`→"bc", `=VALUE("12")`→12, `=CONCAT("a","b")`→"ab",
  `=CONCATENATE("a",1)`→"a1", `=TEXTJOIN("-",TRUE,"a","","b")`→"a-b".
- 0 Warning(s); suíte verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 4: Contagem e soma condicional
Status: Complete

- [x] `Criteria` (motor de critérios): `>`, `<`, `>=`, `<=`, `<>`, `=` + número ou texto; texto
      case-insensitive com curingas `*`/`?` (via regex). `Criteria.Parse(object?)` + `Matches(object?)`.
- [x] Records + `FunctionSpec` (tags 37–42): `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`,
      `SUMIF`(range,crit,[sum_range]), `SUMIFS`(sum_range,…). `ArgumentFlattening.Expand` alinha ranges paralelos.

### Phase Summary
**130/130 verde**, 0 warnings. Ranges paralelos de tamanhos diferentes → `#VALUE!`. Limitações: critério
de ordenação só para números; COUNTIFS/SUMIFS com contagem ímpar de args ignora o último (não valida par).

### Verification Plan
- Testes (com grid montado): `COUNTA` ignora vazios; `COUNTBLANK` conta vazios; `COUNTIF(A1:A3,">1")`;
  `SUMIF(A1:A3,">1")`; `SUMIFS`/`COUNTIFS` com 2 critérios; curinga `COUNTIF(A1:A3,"a*")`.
- 0 Warning(s); suíte verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 5: Referência e lookup (difícil)
Status: Complete (exceto SHEET, adiado)

- [x] `RangeReference` 2D: `RowCount`, `ColumnCount`, `TopRow`, `CellAt` (1-based).
- [x] **Contexto de avaliação introduzido** (`EvaluationContext`: Workbook + sheet/célula atual);
      `Compute(EvaluationContext)` com overload `Compute(Workbook)` de compatibilidade; `CellReference`
      seta a célula atual. (Refactor mecânico via subagente; commit `7c9ab7b`.)
- [x] Records + `FunctionSpec` (tags 43–49): `ROWS`, `ROW` (inclui 0-arg via `context.CellId`),
      `MATCH` (exato/aproximado), `INDEX` (quirk de 1 linha), `VLOOKUP`(3–4, exato/aproximado),
      `XLOOKUP`(3–6, exato + if_not_found), `OFFSET`(3–5, escalar).
- [x] `SHEET`(0–1) (tag 53): `Sheet.Index` (ordem de inserção, atribuída em `Sheets.Add`); SHEET([ref/nome])
      → número 1-based; SHEET() → planilha atual (via `context.SheetName`). Lookups (INDEX/VLOOKUP/OFFSET)
      passaram a acessar células via `RangeReference.CellValueAt`/`GetCellValue` (cache + segurança de ciclo).
- [x] **Enhancements fechados**: OFFSET multi-célula retorna `RangeReference` (expandido por SUM/etc. via
      "valor que é range → expande" em `NumericAggregation`/`ArgumentFlattening`); XLOOKUP com match_mode
      (-1/0/1/2, curinga reusa `Criteria.WildcardMatch`) e search_mode (1/-1). **Argumento omitido** no
      parser (`XLOOKUP(a,b,c,,2)` → blank; vírgula final deixou de lançar) — essencial para dados Excel reais.

### Phase Summary
**142/142 verde**, 0 warnings. O `EvaluationContext` (escolha A do usuário) pavimenta LET (Fase 6) e a
memorização. Limitações registradas: SHEET adiado; OFFSET só escalar; XLOOKUP só modo exato; MATCH/VLOOKUP
aproximado assume ordenação.

### Verification Plan
- Testes: `ROW(A5)`→5, `ROWS(A1:A3)`→3, `MATCH(2,A1:A3,0)`, `INDEX(A1:A3,2)`, `VLOOKUP`, `OFFSET(A1,1,0)`,
  `XLOOKUP` exato. Casos de não-encontrado → `#N/A`.
- 0 Warning(s); suíte verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 6: Avançadas — LET e TEXT (formatação)
Status: Complete

- [x] `NameReference` (tag 50): identificador solto vira referência de nome resolvida no `EvaluationContext`
      (unbound → #NAME?, retrocompatível com o antigo `ErrorValue.Name`).
- [x] `EvaluationContext` ganhou binding de nomes (`WithName`/`TryGetName`, case-insensitive; nomes NÃO
      vazam para células referenciadas via `WithCell`).
- [x] `Let` (tag 51): `LET(nome1, val1, …, calc)`; valores posteriores enxergam nomes anteriores.
- [x] `Text` (tag 52): delega ao `double.ToString(format, invariant)`, que cobre os códigos comuns
      (`0`, `#`, `.`, milhar `,`, `%`). Códigos de data/cor/texto fora de escopo → `#VALUE!`.

### Phase Summary
**147/147 verde**, 0 warnings. LET usa o `EvaluationContext` (escolha A) para o escopo. TEXT limitado a
formatos numéricos compatíveis com .NET.

### Verification Plan
- Testes: `LET(x,2,x*x)`→4, `LET(x,2,y,3,x+y)`→5; `TEXT(1234.5,"#,##0.00")`→"1,234.50", `TEXT(0.5,"0%")`→"50%".
- 0 Warning(s); suíte verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 7: Funções de domínio via extensão
Status: Complete (mecanismo) — a lógica fica com o host

As funções de domínio (INCOME, PORTFOLIO, PROPERTY, GAIN, CONTRIBUTIONS, ORGANIZATION, DETAIL, HIDDEN,
A_HIDE, EXPR_SUM) já são totalmente suportadas pelo mecanismo da Fase 1: parseiam como `FunctionCall`,
serializam com a planilha (nome + args), e resolvem via `Workbook.RegisterFunction(name, impl)`. O usuário
decidiu manter a LÓGICA delas fora desta implementação (o host fornece). Provado por `FunctionExtensionTests`
(registro/invocação, não registrada → #NAME?, round-trip de serialização).

### Phase Summary
Sem código adicional necessário — o mecanismo da Fase 1 cobre o requisito ("registrar + serializar com a
planilha"). Limitação conhecida e comunicada: o *comportamento* (delegate) não serializa; o host re-registra
na carga (comportamento executável não é serializável). Funções definidas como fórmula (estilo LAMBDA) seriam
outra feature, não pedida.

---

## Final Recap
Entregue um mecanismo de extensão de funções + ~40 funções do Excel, em 6 fases, via TDD.
**Suíte final: 147/147 verde, 0 warnings.**

- **Extensão (Fase 1)**: nó `FunctionCall` + `Workbook.RegisterFunction`/`TryGetFunction`; tokenizer aceita
  `_`/`.`; prefixo `XLFN.` normalizado. As 10 funções de domínio do usuário funcionam por aqui (host registra).
- **Funções built-in**: math/info (INT, ROUND, ROUNDUP, ABS, ISNUMBER, ISBLANK, IFNA), texto (UPPER, LOWER,
  TRIM, LEN, LEFT, MID, VALUE, CONCAT, CONCATENATE, TEXTJOIN, TEXT), contagem/soma condicional (COUNTA,
  COUNTBLANK, COUNTIF, COUNTIFS, SUMIF, SUMIFS + motor de critérios), referência/lookup (ROW, ROWS, MATCH,
  INDEX, VLOOKUP, XLOOKUP, OFFSET) e LET — somadas a IF/AND/OR/NOT/IFERROR/SUM/AVERAGE/MIN/MAX/COUNT.
- **Refactor habilitador**: `EvaluationContext` (Workbook + célula/sheet atual + escopo de nomes LET),
  substituindo `Compute(Workbook)` com overload de compatibilidade — pavimenta também a memorização.

Cobertura da lista original: tudo, exceto **SHEET** (adiado — precisa de ordenação de planilhas) e a
**lógica** das funções de domínio (responsabilidade do host).

Limitações conhecidas (registradas): SHEET; OFFSET só escalar (multi-célula adiado); XLOOKUP só exato;
MATCH/VLOOKUP aproximado assume ordenação; TEXT só formatos numéricos compatíveis com .NET; comparadores
sem ordenação de texto/cross-type (backlog em `plans/expression-parser.md`); memorização em `plans/memoization.md`.

## Deployment Plan
Biblioteca — sem infraestrutura. Integração:
1. `dotnet build` (projetos) → 0 Warning(s)/Error(s).
2. `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → 147/147 verde.
3. Host registra funções de domínio via `Workbook.RegisterFunction(name, (args, workbook) => ...)` após
   carregar/desserializar a planilha.
4. Commits já na `main` (histórico linear).

## Fora de escopo (registrado)
- Memorização (`plans/memoization.md`) e o problema de profundidade de recursão (cadeias longas).
- Lógica das funções de domínio (responsabilidade do host via extensão).
- Formatação completa do Excel em TEXT; locale; ordenação cross-type de comparadores.
