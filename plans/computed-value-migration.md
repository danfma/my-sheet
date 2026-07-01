# Migração `Compute → ComputedValue` (+ `Error` smart-enum)

Trocar o retorno de `Expression.Compute` de `object?` (que boxa cada resultado numérico) por um `readonly
struct ComputedValue`, eliminando a alocação/pressão de GC medida no experimento. Breaking change publicável
(bump de major). Design fechado em conversa; execução faseada com TDD, sempre-verde entre fases.

## Contexto e design (fechado)
- **Motivação medida** (`plans/cellvalue-boxing-experiment.md`): cache `object?` boxa 24 B/célula de vida
  longa → dispara Gen1; cache `ComputedValue` → 0 B, zero coletas. Throughput 2,3–4,3× no caminho aritmético,
  4–12% no cache-heavy. Argumento forte = memória/GC num extrator de background.
- **Tipo público `ComputedValue`** (struct de dois campos `double _num` + `object? _ref` + tag, sem `Unsafe`).
  Superfície: `Kind`; `TryGetNumber/Boolean/Text/Error/Reference`; `AsDouble/AsBoolean/AsString` (nullable);
  `ToDouble/ToBoolean/ToText` (**assert estrito, lança**); `EnumerateValues(ctx)` (valores de uma referência,
  via cache); `AsObject()` (escape hatch permanente); fábricas + implícitas **só de entrada**
  (double/bool/string → ComputedValue; string null → Blank; nunca de saída). Coerção fica **interna** ao engine.
- **`Error` smart-enum struct** (alloc-free, `int _code`): well-known nomeados (Null/DivZero/Value/Ref/Name/
  Num/NA), `Display` (`"#VALUE!"`), `ToString()` = Display, `IEquatable`/`==`. `Register()` custom fica para
  depois (YAGNI). O nó de AST `ErrorValue` (serializável) passa a embrulhar um `Error` — um único conceito de
  código de erro.
- **Ranges/referências**: `Kind = Reference` carrega a `Expression` de referência (`RangeReference`);
  `EnumerateValues` devolve `IEnumerable<ComputedValue>` (não expressões). `RangeReference.Compute` continua
  `#VALUE!`; células em massa vêm do sheet/`Expand`, não do `Compute` (decisão do usuário). Compute pode
  retornar a própria referência para casos especiais (OFFSET) — daí o `Kind = Reference`.
- **Escopo do código** (medido): base `abstract object? Compute(EvaluationContext)` (+ overload `Compute(Workbook)`),
  **64 overrides**, **~105 call sites** de `.Compute(`, **~100** usos de `ValueCoercion`, cache
  `ConcurrentDictionary<(string,string), object?>` no `Workbook`. `ComputedValue`/`Error` NÃO entram no
  `MemoryPackUnion` (são runtime; a serialização é do AST/`Expression`).
- **Framework de teste**: TUnit (`[Test] public async Task`, `await Assert.That(...)`). Comando:
  `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release`. Baseline atual: 177 verdes.

## For Future Agents
Marque `- [x]`; ao fechar fase, Status `Complete` + Phase Summary + rode a Verification e registre o resultado.
Branch: `feature/computed-value`. **Sempre-verde**: cada fase mantém a suíte passando. Estratégia de flip
incremental sem big-bang: introduzir `Evaluate(ctx) : ComputedValue` como novo contrato; a base implementa
`Compute` em termos de `Evaluate().AsObject()` e `Evaluate` default em termos de `From(Compute())` — migrar nó
a nó de um lado para o outro, mantendo os dois caminhos consistentes, e só no fim virar a API pública.
NÃO fazer push / abrir PR / publicar no NuGet sem o usuário confirmar.

---

## Phase 1: Tipos core (aditivo, não-breaking)
Status: Complete

Cria `ComputedValue` + `Error` no core, ao lado do `object?` atual, sem tocar `Compute`/`ValueCoercion`/cache.
Nada quebra; a suíte segue verde e ganhamos testes de contrato.

- [x] `Danfma.MySheet/Expressions/Error.cs`: `readonly struct Error : IEquatable<Error>` com `int _code`,
      estáticos well-known (Null/DivZero/Value/Ref/Name/Num/NA), `Display`, `ToString()=Display`, `==`/`!=`,
      `GetHashCode`. Pontes `internal`: `FromCode(int)`, `Code`, `FromDisplay(string)`, `ToErrorValue()`
      (mapeia para os singletons `ErrorValue` existentes; cria `#NULL!` inline).
- [x] `Danfma.MySheet/Expressions/ComputedValue.cs`: `enum ComputedValueKind : byte` + `readonly struct
      ComputedValue` (dois campos + tag). Fábricas (`Number/Boolean/Text/Error/Reference` + `Blank`),
      `TryGet*` (estrito: Number≠Boolean), `As*` (nullable), `To*` (assert lança), `EnumerateValues(ctx)` +
      overload `(Workbook)`, `AsObject()`, `From(object?)` (ponte inversa, ErrorValue→Error), implícitas
      só-de-entrada (double/bool/string/Error → ComputedValue; string null → Blank).
- [x] `tests/Danfma.MySheet.Tests/Expressions/ComputedValueTests.cs` + `ErrorTests.cs`: Kind, TryGet estrito,
      As nullable, To lança (try/catch), Blank, implícitas (string null→Blank), `AsObject`/`From` round-trip
      (incl. ErrorValue↔Error), `Error` Display/ToString/==, `EnumerateValues` sobre range montado.

### Verification Plan
- `dotnet build Danfma.MySheet/Danfma.MySheet.csproj -c Release` → 0 Warning(s)/Error(s).
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → todos verdes.
  Nada do engine (`Compute`/`ValueCoercion`/`Workbook`/`ErrorValue`) alterado (`git diff --stat` só novos).

### Phase Summary
**Concluída.** Build core 0 Warning(s)/0 Error(s); suíte **235/235 verde** (177 antigos intactos + novos de
contrato). `git status` mostra só arquivos novos; `git diff` do engine (`Expression.cs`/`ValueCoercion.cs`/
`Workbook.cs`/`ErrorValue.cs`) **vazio** — Fase 1 é 100% aditiva, zero risco de regressão. Commit `feat`.

Decisões/aprendizados registrados:
- `TryGet*`/`As*`/`To*` são **estritos por tipo exato**: `TryGetNumber` num Boolean → `false` (não cruza).
  A coerção Excel (bool→1) NÃO está nesta superfície — fica interna ao engine (Fase 2, em `ValueCoercion`).
- `Error` (struct de identidade) ↔ `ErrorValue` (nó de AST): `AsObject()` mapeia Error→ErrorValue reusando
  os singletons; `From(ErrorValue)` mapeia por `ErrorCode`. Round-trip provado em `ErrorTests`.
- Adicionado overload `EnumerateValues(Workbook)` (espelha `Compute(Workbook)`) para ergonomia/testes.
- `InternalsVisibleTo("Danfma.MySheet.Tests")` já existe → testes exercitam as pontes `internal`.
- Nenhum `MemoryPackUnion` novo: `ComputedValue`/`Error` são runtime, não entram na serialização do AST.

---

## Phase 2: Contrato `Evaluate` + coerção nativa + nós-valor
Status: Not started

- [ ] Base `Expression`: adicionar `public virtual ComputedValue Evaluate(EvaluationContext ctx) =>
      ComputedValue.From(Compute(ctx));` (default bridged). Mantém `Compute` como está → tudo verde.
- [ ] `ValueCoercion` ganha overloads nativos sobre `ComputedValue` (`ToNumber(in ComputedValue)` etc.),
      coexistindo com os de `object?`. Coerção continua **interna**.
- [ ] Migrar nós-folha (`NumberValue`, `StringValue`, `BooleanValue`, `BlankValue`, `ErrorValue`,
      `CellReference`, `RangeReference`, `NameReference`, `UnionReference`) para **override `Evaluate`
      nativo**; o `Compute` de cada um passa a `=> Evaluate(ctx).AsObject()`.
- [ ] `ErrorValue` passa a embrulhar/expor um `Error` (reconciliação do código de erro).

### Verification Plan
- Suíte verde. Um teste novo prova consistência: para um conjunto de nós, `node.Evaluate(ctx).AsObject()`
  equivale ao antigo `node.Compute(ctx)`.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Migrar nós compostos (math/texto/info/lógica/contagem)
Status: Not started

- [ ] Migrar para `Evaluate` nativo: operadores (`BinaryOperation`/`UnaryOperation`), agregação
      (`Sum/Average/Min/Max/Count` + `NumericAggregation`), lógica (`If/And/Or/Not/IfError/IfNa`), math
      (`Int/Round/RoundUp/Abs`), info (`IsNumber/IsBlank`), texto (`Upper/Lower/Trim/Len/Left/Mid/Value/
      Concat/Concatenate/TextJoin/Text`), contagem/soma condicional (`CountA/CountBlank/CountIf(s)/SumIf(s)`
      + `Criteria`/`ArgumentFlattening`). Cada nó: `Evaluate` nativo, `Compute` delega, testes verdes.

### Verification Plan
- Suíte verde a cada nó migrado; `git diff` toca só os nós da fase + helpers compartilhados.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 4: Migrar lookup/referência + financeiras + LET
Status: Not started

- [ ] Migrar `Row/Rows/Match/Index/VLookup/XLookup/Offset/SheetNumber`, `Let`/`NameReference`, financeiras
      (`Pmt/Pv/Fv/Nper/Ipmt/Ppmt/Npv/Rate/Irr` + `TimeValueOfMoney`), `FunctionCall` (custom). Fechar os 64.
- [ ] `Offset` multi-célula e lookups passam a produzir/consumir `Kind = Reference`/`EnumerateValues`.

### Verification Plan
- Suíte verde; nenhum override de `Compute` nativo restante (todos delegam a `Evaluate`).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 5: Virar a API pública + cache `ComputedValue`
Status: Not started

- [ ] `Workbook._cache`: `ConcurrentDictionary<(string,string), ComputedValue>` (valor deixa de ser boxed —
      origem do ganho de Gen1). `GetCellValue` passa a `ComputedValue`.
- [ ] Tornar `Evaluate` o contrato primário; **`Compute` público passa a retornar `ComputedValue`** (virar a
      assinatura — breaking). Remover os overrides `object?` e os adaptadores internos; `AsObject()` permanece
      como escape hatch público.
- [ ] Converter os ~105 call sites (testes/benchmark inclusos) para a API `ComputedValue` (`TryGet*`/`To*`).

### Verification Plan
- Suíte verde já usando a API nova. `dotnet run` do benchmark `SheetBenchmarks` mostra queda de `Allocated`
  em fórmulas de produção vs. baseline `main`.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 6: Publicação (docs + versão + benchmark real)
Status: Not started

- [ ] Atualizar `README.md` (API nova, exemplo com `ToDouble()`/`TryGet*`), guia de migração (breaking).
- [ ] Bump de **major** (`chore(release)`), changelog do breaking change.
- [ ] Rodar `SheetBenchmarks` antes/depois e registrar o ganho real de alocação.
- [ ] (Somente com aval do usuário) push / PR / publish no NuGet.

### Verification Plan
- `dotnet build` 0 warnings; suíte verde; README compila os exemplos mentalmente; benchmark real registrado.

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando todas as fases concluírem)_

## Deployment Plan
_(escrever quando concluir: bump de major, publish NuGet via CI/OIDC, guia de migração — só com aval)_
