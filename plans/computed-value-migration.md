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
Status: Complete

- [x] Base `Expression`: `public virtual ComputedValue Evaluate(EvaluationContext ctx) =>
      ComputedValue.From(Compute(ctx));` (default bridged) + overload `Evaluate(Workbook)`. `Compute`
      intacto → tudo verde.
- [x] `ValueCoercion` ganhou coerção nativa sobre `ComputedValue` como **extension methods**
      (`CoerceToNumber`/`CoerceToBool`/`CoerceToText` `(this in ComputedValue, out …) : Error?`) → chamada
      **fluente** `value.CoerceToNumber(out var n)`, mantendo a coerção `internal` (fora da API pública do
      struct). **Não** são `Try*` (não retornam `bool`): devolvem o `Error` a propagar (ou `null`), via
      `if (value.CoerceTo…(out var v) is { } error) return error;`. Os `TryTo*(object?)` legados continuam
      estáticos (extension sobre `object?` seria ruim) e serão deletados nas Fases 3–5.
- [x] Migrar nós-valor literais (`NumberValue`, `StringValue`, `BooleanValue`, `BlankValue`) para override
      `Evaluate` nativo; `Compute` de cada um passa a `=> Evaluate(ctx).AsObject()`.
- [x] `ErrorValue`: override `Evaluate => ComputedValue.Error(AsError())` + `internal Error AsError()`
      (reconciliação). `Compute` mantido como `=> this` (zero-risco; todos os `ErrorValue` são singletons
      well-known → ponte lossless, confirmado por grep de `new ErrorValue(`).
- **Escopo ajustado (registrado):** as **referências** (`CellReference`/`RangeReference`/`NameReference`/
  `UnionReference`) NÃO foram migradas aqui — elas leem do cache `object?` e não ganham nada nativo até o
  cache virar `ComputedValue` (Fase 5). Ficam no bridge da base (correto) até lá.

### Verification Plan
- Suíte verde. Teste `EvaluateBridgeTests` prova: `node.Evaluate(wb).AsObject()` == `node.Compute(wb)` para
  os nós-valor; nó não-migrado (`Sum`) roteia pelo bridge da base; coerção nativa (número/bool/texto/blank/
  erro) espelha a de `object?`.

### Phase Summary
**Concluída.** Build core 0 Warning(s)/0 Error(s); suíte **240/240 verde** (235 + 5 novos). Sempre-verde
mantido: `Compute` de todos os call sites (~105) continua funcionando — os nós migrados roteiam por
`Evaluate().AsObject()`, os não-migrados usam o bridge `From(Compute())` da base.

Decisões/aprendizados:
- **Seam de migração**: a base tem os dois caminhos consistentes por construção. Um nó está em UM de dois
  estados: (a) `Compute` nativo + `Evaluate` herdado (bridge), ou (b) `Evaluate` nativo + `Compute` delega.
  Migrar = mover de (a) para (b), um nó por vez, sempre verde.
- **`ErrorValue.Compute` intocado de propósito**: preservar a identidade exata do nó (não normalizar via
  `Error`) é zero-risco; a normalização só aconteceria em erros não-well-known, que não existem no engine.
- **Coerção nativa retorna `Error?`** (não `ErrorValue?`) — é o que os nós compostos vão propagar na Fase 3.
- Referências adiadas para depois do cache (Fase 5) — evita `From()` wrapping sem ganho.

---

## Phase 3: Migrar nós compostos (math/texto/info/lógica/contagem)
Status: In progress

Faseada em lotes coesos (sempre-verde a cada lote). Ordem escolhida: primeiro os que **não** dependem dos
helpers de range (escalares), depois os que exigem migrar `NumericAggregation`/`ArgumentFlattening`/`Criteria`.

- [x] **3a — Lógica** (`If/And/Or/Not/IfError/IfNa`): `Evaluate` nativo consumindo `CoerceToBool` + checagem
      de erro; `Compute` delega. Curto-circuito preservado (`If`/`IfError`/`IfNa` só avaliam o ramo tomado).
      Zero helper novo. **Concluído** (241/241 verde).
- [ ] **3b — Operadores** (`BinaryOperation`/`UnaryOperation`): aritmética via `CoerceToNumber`; comparação
      exige adicionar overloads nativos `AreEqual`/`Compare(in ComputedValue, in ComputedValue)` em `ValueCoercion`.
- [ ] **3c — Math/Info escalares** (`Int/Round/RoundUp/Abs`, `IsNumber/IsBlank`).
- [ ] **3d — Texto escalar** (`Upper/Lower/Trim/Len/Left/Mid/Value/Text`).
- [ ] **3e — Agregação + variádicos + condicional** (`Sum/Average/Min/Max/Count`, `Concat/Concatenate/TextJoin`,
      `CountA/CountBlank/CountIf(s)/SumIf(s)`): exige migrar `NumericAggregation`/`ArgumentFlattening`/`Criteria`
      para produzir/consumir `ComputedValue`.

### Verification Plan
- Suíte verde a cada lote; `git diff` toca só os nós do lote + helpers compartilhados. Os testes antigos de
  cada função (que exercitam `Compute` → agora roteado por `Evaluate`) continuam verdes, provando semântica.

### Phase Summary
_(escrever quando a fase concluir — lote 3a concluído)_

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
