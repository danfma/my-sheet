# ExpressionParser — Parser de Fórmulas estilo Pratt para o MySheet

Evoluir `ExpressionParser.Parse(string, Sheet)` (hoje um stub que lança `NotImplementedException`)
para um parser manual estilo Pratt que transforma fórmulas de planilha numa árvore de `Expression`,
ajustando o modelo de `Expression` para suportar aritmética, intervalos, comparadores e múltiplas funções.

## Context

O projeto MySheet é um motor de planilha orientado a desempenho (serialização via MemoryPack,
benchmark contra ClosedXML). O modelo de `Expression` já existe e é serializável, mas só pode ser
construído programaticamente via factories (`Number`, `Cell`, `Sum`). Não há como transformar texto
de fórmula (`"=SUM(A1,A2)"`) na árvore. O teste `ParseSum_ShouldDereferenceCellsToCompute`
(`tests/MySheet.Tests/Expressions/ExpressionTests.cs:24`) já existe e falha em runtime
(`NotImplementedException`).

Decisões fechadas com o usuário:
- **Parser manual (Pratt)** — recursive-descent + binding powers. Sem bibliotecas (sem Parlot/Pidgin).
- **Escopo completo agora**: operadores aritméticos (`+ - * / ^`, unário, parênteses, precedência),
  intervalos (`A1:B2`), comparadores (`= <> < > <= >=`), múltiplas funções
  (SUM, AVERAGE, MIN, MAX, COUNT). O modelo de `Expression` pode ser ajustado.
- **Erros híbridos**: erro de **sintaxe** lança `ParseException` (com posição); erro **semântico** em
  runtime (função desconhecida `#NAME?`, célula inválida, etc.) vira nó `ErrorValue`. Isso exige
  promover `ErrorValue` a `Expression` (corrige o warning morto CS0184 em `Sum.cs:43`).

### Invariantes de design (não violar)
- **`Compute` continua `object? Compute(Workbook)`** retornando escalares CLR (`double`/`string`/`bool`/`null`).
  A ÚNICA subclasse de `Expression` que `Compute` retorna é `ErrorValue` (nó de erro, propagado).
  Não introduzir um `Evaluate : Expression` — quebraria os call-sites `Compute(...) as double?`
  (`ExpressionTests.cs:18,33`, `SheetBenchmarks.cs:36`).
- **Tags de `[MemoryPackUnion]` são append-only.** As tags 0–6 em `Expression.cs:6-13` NUNCA são
  renumeradas, reordenadas ou reutilizadas (quebraria o round-trip de `WorkbookTests.cs` e qualquer
  `.bin` persistido). Novos tipos recebem tags 7+. Idem para a ordem de membros dos novos enums.
- **Todo novo nó concreto**: `[MemoryPackable] public sealed partial record`. Os abstratos
  intermediários (`ValueExpression`, `Reference`, `Function`) permanecem SEM `[MemoryPackable]` —
  o dispatch acontece na raiz `Expression`. Campos/filhos sempre tipados como `Expression`.

## For Future Agents
Conforme o trabalho avança: marque os checkboxes `- [x]` ao concluir itens; quando uma fase terminar,
mude o Status para `Complete` e escreva o **Phase Summary** (o que foi feito, decisões-chave, o que for
preciso para continuar sem contexto); rode o **Verification Plan** da fase e registre o resultado antes
de seguir. Quando todas as fases estiverem prontas, preencha **Final Recap** e **Deployment Plan**.

Comando de teste (net10.0 + TUnit/Microsoft.Testing.Platform — `dotnet test` NÃO funciona):
`dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj`.
Use TDD: escreva o teste, veja-o falhar, implemente, veja-o passar.

---

## Phase 1: Ajustar o modelo de Expression
Status: Complete

Fundação do avaliador. A mudança crítica de correção está aqui (refactor do `Sum`), antes de qualquer
parsing — sem ela o parser produziria árvores que o avaliador não consegue computar.

- [x] Promover `ErrorValue` a `sealed partial record ErrorValue(string ErrorCode) : ValueExpression`,
      `[MemoryPackable]`, `Compute` retorna `this`. Manter `NotValue` (#VALUE!) e adicionar
      `Name` (#NAME?), `Reference` (#REF!), `DivByZero` (#DIV/0!). Registrar union tag **7** em `Expression.cs`.
- [x] Adicionar `BinaryOperation(BinaryOperator Operator, Expression Left, Expression Right)` + enum
      `BinaryOperator { Add, Subtract, Multiply, Divide, Power, Equal, NotEqual, LessThan, GreaterThan,
      LessThanOrEqual, GreaterThanOrEqual }`. Union tag **8**.
- [x] Adicionar `UnaryOperation(UnaryOperator Operator, Expression Operand)` + enum
      `UnaryOperator { Negate, Plus }`. Union tag **9**.
- [x] Criar helper interno `ValueCoercion`: `TryToNumber(object? raw, out double n)`
      tratando `double`/`string` parseável (invariant)/`bool`(1/0)/`null`(0)/`ErrorValue`(propaga);
      coerção falha → `ErrorValue.NotValue`.
- [x] Implementar `BinaryOperation.Compute` (coerção via `ValueCoercion`, propaga erro do operando
      esquerdo primeiro, divisão por zero → #DIV/0!, comparações retornam `bool`) e `UnaryOperation.Compute`.
- [x] Dar campos a `RangeReference(string StartId, string EndId, string SheetName)`. `Compute` escalar
      retorna `ErrorValue.NotValue`. `IEnumerable<Expression> Expand(Workbook)` enumera o retângulo
      (normaliza min/max para intervalos invertidos `A3:A1`).
- [x] Criar `CellAddress` (readonly struct): `Parse("AA12")` → (coluna, linha) e `ToId()`. Normaliza maiúsculas.
- [x] **Refatorar `Sum` (correção crítica)**: agregação via helper `NumericAggregation.Gather` que chama
      `arg.Compute(workbook)` (não mais `Resolve` + switch fechado). Conserta os dois bugs (throw em
      arg `BinaryOperation`/`Sum`; referência encadeada → #VALUE!). Semântica Excel preservada:
      texto literal direto não numérico → #VALUE!; texto/logical/blank referenciado → ignorado.
      `RangeReference` → `Expand` tratando cada célula como referenciada.
- [x] Adicionar `Average`/`Min`/`Max`/`Count` (tags **10/11/12/13**) sobre `NumericAggregation`:
      `AVERAGE()` vazio → #DIV/0!; `MIN`/`MAX` vazio → 0; `COUNT` só numéricos e nunca propaga erro.
      Factories estáticas adicionadas em `Expression` (`Add`/`Subtract`/`Divide`/`Power`/`GreaterThan`/
      `Negate`/`Plus`/`Average`/`Min`/`Max`/`Count`/`Range`).
- [x] Comentário em `Expression.cs` fixando a regra "union tags são append-only".

### Verification Plan
- `dotnet build MySheet/MySheet.csproj` → `0 Error(s)` e **0 Warning(s)** (o CS0184 de `Sum.cs:43` some).
- Novos testes unitários em `tests/MySheet.Tests/Expressions/` (construindo árvores via factories, sem parser):
  - `Add(Number(1), Number(2)).Compute() == 3.0`; `Power(Number(2), Number(3)) == 8.0`;
    `Divide(Number(1), Number(0))` → `ErrorValue("#DIV/0!")`.
  - Comparação: `GreaterThan(Number(2), Number(1)).Compute()` como `bool?` == true.
  - Propagação: `Add(Number(1), <nó que computa ErrorValue>)` → ErrorValue.
  - `Negate(Number(2)) == -2.0`.
  - Range: `Sum(Range("A1","A3", sheet))` com A1=1,A2=2,A3=3 → 6; range invertido `A3:A1` → 6.
  - Encadeamento (regressão do bug): A3=`Sum(A1,A2)`, `Sum(Cell("A3"),Number(1))` → 4.
  - `Average`/`Min`/`Max`/`Count` com casos de vazio e blanks.
- Rodar a suíte: `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` →
  `Workbook_IsSerializable` continua passando (round-trip de union intacto); novos testes verdes.

### Phase Summary
Concluída em TDD (red→green por tipo). Resultado: `dotnet build MySheet/MySheet.csproj` →
**0 Warning(s), 0 Error(s)** (CS0184 morto eliminado pela reescrita do `Sum`). Suíte: **30 testes,
29 verdes**; o único vermelho é `ParseSum_ShouldDereferenceCellsToCompute` (alvo da Fase 3, ainda stub).

Arquivos novos em `MySheet/Expressions/`: `BinaryOperation.cs` (+enum `BinaryOperator`),
`UnaryOperation.cs` (+enum `UnaryOperator`), `ValueCoercion.cs`, `NumericAggregation.cs`,
`CellAddress.cs`, `Average.cs`, `Min.cs`, `Max.cs`, `Count.cs`. Modificados: `ErrorValue.cs`
(agora `: ValueExpression`, MemoryPackable, +constantes), `RangeReference.cs` (campos + `Expand`),
`Sum.cs` (reescrito sobre `NumericAggregation`), `Expression.cs` (tags 7–13 + factories).
Novos testes: `BinaryOperationTests`, `UnaryOperationTests`, `RangeReferenceTests`, `AggregateFunctionTests`.

Decisões-chave para continuar sem contexto:
- **`Compute` continua `object?`**; o único `Expression` que `Compute` retorna é `ErrorValue`
  (escalares CLR para o resto). Toda coerção numérica passa por `ValueCoercion`/`NumericAggregation`.
- **Semântica direto vs referenciado**: argumentos `CellReference`/`RangeReference` são "referenciados"
  (só `double` conta; texto/logical/blank ignorados); o resto é "direto" (texto-numérico e logical contam,
  texto não numérico → #VALUE!). Implementado em `NumericAggregation`.
- **`ErrorValue` ambíguo nos testes** com `TUnit.Assertions.StringValue` → qualifique
  `MySheet.Expressions.StringValue` ao construir strings em testes.
- Factories só foram adicionadas conforme os testes pediram (TDD). A Fase 3 (parser) adicionará as que
  faltarem (ex.: `Multiply`, `LessThan`, `Equal`...) quando os testes do parser as exercitarem, ou o
  parser pode construir `new BinaryOperation(...)` diretamente.

---

## Phase 2: Tokenizer (lexer)
Status: Complete

- [x] `enum TokenType` + `readonly record struct Token(Type, Text, Position)` em `MySheet/Parsing/`.
- [x] `Tokenizer` (internal) sobre o corpo da fórmula; `Tokenize(string) -> List<Token>` terminando
      em `EndOfInput`. Whitespace descartado.
- [x] **Number**: começa só em dígito/`.`; notação científica gananciosa (`1E2`, `1.5E-3`), com back-off
      se o expoente não tiver dígitos.
- [x] **String**: `"..."` com `""` escapado (texto desescapado guardado em `Token.Text`).
- [x] **Identifier**: `[A-Za-z][A-Za-z0-9]*` cru (classificação no parser).
- [x] **Operadores multi-char**: `<>`, `<=`, `>=` antes de `<`, `>`, `=`.
- [x] Caractere inválido / string não terminada → `ParseException` (pública) com posição.

### Verification Plan
- Testes unitários do `Tokenizer` (`TokenizerTests`) cobrindo `SUM(A1,A2)`, `A1:B2`, `-3 + 4 * 2`,
  `1 <= 2`, `1<>2`, `1>=2`, `1E2` (1 token Number, Text "1E2"), `1.5E-3`, `E2`→Identifier,
  `"a""b"`→String("a\"b"), `SUM (A1)` (espaço descartado), `1 # 2`→`ParseException`.
- `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → tudo verde.

### Phase Summary
Concluída em TDD. Suíte: **39 testes, 38 verdes** (só `ParseSum` vermelho, alvo da Fase 3).
Novos arquivos em `MySheet/Parsing/`: `Token.cs` (`TokenType` + `Token`), `Tokenizer.cs`,
`ParseException.cs` (pública). `MySheet.csproj` ganhou `<InternalsVisibleTo Include="MySheet.Tests" />`
para testar os tipos internos. Novo teste: `TokenizerTests`.
Decisão: `Token.Text` guarda o texto JÁ desescapado para strings e o lexema cru para números/identifiers.
O parser da Fase 3 consome `List<Token>` por índice (com `EndOfInput` ao final).

---

## Phase 3: Parser Pratt + ExpressionParser.Parse
Status: Complete

- [x] Parser Pratt (`Parser.cs`, internal) consumindo `List<Token>` por índice. Binding powers:
      comparação 10 < aditivo 20 < multiplicativo 30 < potência 40 (right-assoc) < prefixo unário 45
      < intervalo `:` 50. Unário liga mais forte que `^` (`-2^2 == 4`). `+ - * /` left-assoc.
- [x] **Prefixo (`ParsePrefix`)**: número, string, `TRUE`/`FALSE` (antes do teste de cell ref),
      `(` expr `)`, unário `- +`, e identificador.
- [x] **Referência**: `IsCellReference` (`^[A-Za-z]+[0-9]+$` manual) → `CellReference` com `Id` em
      MAIÚSCULAS (corrige o case-sensitive de `Sheet.Cells`). `SheetName` = `sheet.Name`.
- [x] **Intervalo `:`**: operador infixo de maior precedência, restrito a operandos `CellReference`
      (senão `ParseException`).
- [x] **Chamada de função**: registro `OrdinalIgnoreCase` (SUM/AVERAGE/MIN/MAX/COUNT). Desconhecida
      → `ErrorValue.Name` (#NAME?), sem lançar.
- [x] **`ExpressionParser.Parse`**: `=` → strip + parse; vazio após `=` → `ParseException`. Sem `=`
      → modo literal (número/`TRUE`/`FALSE`/texto; string vazia → `BlankValue`).
- [x] Sobra de tokens (ex.: `=1 2`) → `ParseException`.

### Verification Plan
- O teste alvo `ParseSum_ShouldDereferenceCellsToCompute` (`=SUM(A1,A2)` → 3) passa.
- Matriz `ExpressionParserTests` (precedência/assoc, potência/unário quirk, divisão por zero,
  comparadores, funções+intervalos+case-insensitive, #NAME?/#VALUE!/#DIV/0!, erros de sintaxe,
  modo literal) — ver itens acima.
- `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → toda a suíte verde.

### Phase Summary
Concluída em TDD. **73 testes, 73 verdes** (incluindo `ParseSum` alvo); `dotnet build` de main e testes
com **0 Warning(s)**. Novos arquivos em `MySheet/Parsing/`: `Parser.cs` (Pratt) e `ExpressionParser.cs`
(implementado). Novo teste: `ExpressionParserTests` (34 casos cobrindo a matriz).
Decisões: comparadores são left-assoc com coerção numérica (sem ordenação cross-type do Excel);
posições de `ParseException` são relativas ao corpo após o strip do `=` (off-by-one aceito).

---

## Phase 4: Benchmark de parsing + polimento
Status: Complete

- [x] Adicionado `[Benchmark, BenchmarkCategory("Parse")] MySheetParse()` em `SheetBenchmarks.cs`,
      parseando `=SUM(A1:A3) + AVERAGE(B1,B2)*2 - 3^2` (`[MemoryDiagnoser]` já ativo).
- [x] Revisão de elegância/duplicação: coerção centralizada em `ValueCoercion` (aritmética) e
      `NumericAggregation` (agregação); sem branches mortos; loops inline sem alocação mantidos por
      decisão de desempenho (base serializável / reducer via delegate descartados).

### Verification Plan
- `dotnet build benchmarks/MySheet.Benchmark/MySheet.Benchmark.csproj -c Release` → `0 Error(s)` ✔.
- (Opcional, lento — NÃO executado) `dotnet run -c Release --project benchmarks/MySheet.Benchmark`
  e inspecionar a coluna `Allocated` do benchmark `Parse`.
- Suíte completa verde: `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → 73/73 ✔.

### Phase Summary
Concluída. Benchmark de parsing compila em Release (0 warnings/erros); execução completa deixada como
passo opcional (lento). Revisão de duplicação feita; código mantido sem alocações no caminho quente.

---

## Final Recap
`ExpressionParser.Parse(string, Sheet)` saiu de um stub para um parser Pratt manual completo, com o
modelo de `Expression` estendido. Entregue em 4 fases via TDD; **suíte final 73/73 verde**, todos os
projetos (`MySheet`, testes, benchmark) com **0 Warning(s)**.

Capacidades suportadas: literais (número/texto/boolean/blank), referências de célula `A1`
(case-insensitive), intervalos `A1:B2`, operadores aritméticos `+ - * / ^` com precedência e quirk do
Excel (`-2^2==4`, `^` right-assoc), menos/mais unário, comparadores `= <> < > <= >=` (booleanos),
funções SUM/AVERAGE/MIN/MAX/COUNT (registro extensível, case-insensitive), e modo literal para entrada
sem `=`. Erros híbridos: sintaxe → `ParseException` (pública, com posição); semântica → `ErrorValue`
(#NAME?/#VALUE!/#DIV/0!).

Bugs pré-existentes corrigidos no caminho: (1) `Sum` lançava/errava com argumentos não-escalares e
referências encadeadas (agora agrega via `Compute`); (2) `ErrorValue` não era `Expression` (warning
CS0184 morto); (3) `Sheet.Cells` case-sensitive faria `=sum(a1,a2)` retornar 0 (parser normaliza IDs).

Arquivos novos: `MySheet/Expressions/{BinaryOperation,UnaryOperation,ValueCoercion,NumericAggregation,
CellAddress,Average,Min,Max,Count}.cs`; `MySheet/Parsing/{Token,Tokenizer,ParseException,Parser}.cs`.
Modificados: `MySheet/Expressions/{Expression,ErrorValue,RangeReference,Sum}.cs`, `MySheet/MySheet.csproj`
(`InternalsVisibleTo`), `benchmarks/.../SheetBenchmarks.cs`. Testes novos: `BinaryOperationTests`,
`UnaryOperationTests`, `RangeReferenceTests`, `AggregateFunctionTests`, `Parsing/TokenizerTests`,
`Parsing/ExpressionParserTests`.

Limitações conhecidas (fora de escopo desta iteração): sem ordenação cross-type do Excel em
comparadores; `:` só entre duas células (sem ranges 3D nem união/interseção); posição do
`ParseException` é relativa ao corpo após o `=` (off-by-one); precedência exata do COUNT com logicais
diretos não foi exaustivamente espelhada do Excel.

## Deployment Plan
É uma biblioteca — não há infraestrutura a implantar. Para integrar:
1. `dotnet build` (solução/projetos) → 0 Warning(s), 0 Error(s).
2. `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → 73/73 verde.
3. Commit em branch (não `main` diretamente) e abrir PR. Mensagem sugerida:
   `feat(parsing): add Pratt formula parser and extend expression model`.
4. (Opcional) Rodar o benchmark em Release para registrar baseline de alocação do parser antes do merge.
