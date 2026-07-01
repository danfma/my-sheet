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
Status: Complete

Faseada em lotes coesos (sempre-verde a cada lote). Ordem: primeiro os que **não** dependem dos helpers de
range (escalares), depois os que envolvem `NumericAggregation`/`ArgumentFlattening`/`Criteria`.

- [x] **3a — Lógica** (`If/And/Or/Not/IfError/IfNa`): `CoerceToBool` + checagem de erro; curto-circuito preservado.
- [x] **3b — Operadores** (`BinaryOperation`/`UnaryOperation`): aritmética via `CoerceToNumber`; comparação via
      novos `AreEqual`/`Compare(in ComputedValue, in ComputedValue)` nativos (cross-type, sem boxing).
- [x] **3c — Math/Info escalares** (`Int/Round/RoundUp/Abs`, `IsNumber/IsBlank`).
- [x] **3d — Texto escalar** (`Upper/Lower/Trim/Len/Left/Mid/Value/Text`).
- [x] **3e — Agregação + variádicos + condicional** (`Sum/Average/Min/Max/Count`, `Concat/Concatenate/TextJoin`,
      `CountA/CountBlank/CountIf(s)/SumIf(s)`): nós migrados **envolvendo** os helpers `object?` existentes
      (`NumericAggregation`/`ArgumentFlattening`/`Criteria`) — resultado deixa de re-boxar; o boxing **interno**
      dos helpers some quando o cache virar `ComputedValue` (Fase 5). Baixo risco, helpers testados intactos.

### Verification Plan
- Suíte verde a cada lote; os testes antigos de cada função (que exercitam `Compute` → agora roteado por
  `Evaluate`) continuam verdes, provando semântica. Novos testes nativos: `Evaluate_LogicNodes_Native`,
  `Evaluate_Operators_Native`.

### Phase Summary
**Concluída (5 lotes, sempre-verde).** Todos os ~30 nós compostos de math/texto/info/lógica/contagem têm
`Evaluate` nativo; `Compute` delega. **242/242 verde, 0 warnings.** `ValueCoercion` ganhou `AreEqual`/`Compare`
nativos. Decisão registrada: os nós que dependem de range (3e) **envolvem** os helpers `object?` em vez de
migrá-los — completa a fase (nós nativos) com risco mínimo; a eliminação do boxing interno dos helpers pega
carona na virada do cache (Fase 5). Restam para a Fase 4: lookup/referência, financeiras, LET, `FunctionCall`
e as referências (`Cell/Range/Name/Union`).

---

## Phase 4: Migrar lookup/referência + financeiras + LET
Status: Complete

- [x] **4a — Financeiras** (`Pmt/Pv/Fv/Nper/Ipmt/Ppmt/Rate` escalares via transform; `Npv/Irr` range). 41
      provas contra o oráculo `ExcelFinancialFunctions` verdes.
- [x] **4b — Lookup/LET** (`Row/Rows/Match/Index/VLookup/XLookup/Offset/SheetNumber`, `Let`, `FunctionCall`).
      `Offset` multi-célula devolve `Kind = Reference` (agregadores expandem via `AsObject()` → range).
- Referências (`Cell/Range/Name/Union`) adiadas para a Fase 5 (viram nativas com a virada do cache).

### Verification Plan
- Suíte verde (242/242); `git diff` toca só os nós do lote.

### Phase Summary
**Concluída.** Todas as financeiras e lookup/LET/FunctionCall têm `Evaluate` nativo. Financeiras escalares
migradas por script (regex regular) + `Npv/Irr` à mão; validadas pelas 41 provas do oráculo. Lookups mantêm
os helpers `object?` (chaves/arrays via cache) e produzem `ComputedValue`.

---

## Phase 5: Cache `ComputedValue` + centralização (feito NÃO-breaking)
Status: Complete

- [x] **5a — Cache** `Workbook._cache` → `ConcurrentDictionary<_, ComputedValue>`; `GetCellComputedValue`
      computa via `Evaluate` e cacheia o struct inline (fim dos boxes de vida longa → Gen1). `GetCellValue`
      público delega via `AsObject()` (não-breaking). Referências (`Cell/Range/Name/Union`) migradas.
- [x] **5b — Centralização** `Evaluate(ctx) : ComputedValue` virou o contrato **abstrato** primário (todos os
      64 nós o implementam); `Compute(ctx) : object?` virou um **delegate único na base** (`=> Evaluate().AsObject()`),
      removendo as 64 delegações duplicadas por nó.
- **Desvio consciente do plano original:** a virada saiu **NÃO-breaking**. Descobriu-se que dá para entregar o
      ganho de GC (cache = `ComputedValue`) + a API tipada (`Evaluate : ComputedValue`) mantendo `Compute : object?`
      como interop — em vez de renomear `Compute` para retornar `ComputedValue` (que quebraria os ~105 call
      sites/README/consumidores). O rename breaking fica como decisão em aberto do usuário (ver Final Recap).

### Verification Plan
- Suíte 242/242 verde; benchmark e testes (que usam `.Compute`) compilam sem mudança → prova não-breaking.

### Phase Summary
**Concluída, não-breaking.** O cache guarda `ComputedValue` (ganho de GC entregue). A API pública ganhou
`Evaluate/Evaluate(Workbook) : ComputedValue` (sem boxing) e manteve `Compute : object?` (interop). Base
simplificada (−64 delegações). Nenhum call site externo quebra.

---

## Phase 6: Docs + versão + benchmark
Status: Complete (exceto publish — aguarda aval)

- [x] `README.md` atualizado: bullet de "allocation-free evaluation" + quick-start com `Evaluate`/`ToDouble`/
      `TryGetError` e o form `Compute : object?` para interop.
- [x] Solução inteira compila (core + tests + benchmark) 0 warnings; benchmark usa `.Compute` e compila →
      confirma que, **até a Fase 6**, a migração era não-breaking (aditiva). A Fase 8 (remoção do `Compute`,
      decidida depois) tornou-a **breaking → major** (ver Deployment Plan).
- [ ] **(Aguarda aval)** push / PR / publish no NuGet — passo externo/irreversível, não executado.

### Verification Plan
- `dotnet build` (todos os projetos) 0 warnings; suíte 242/242 verde; exemplos do README coerentes com a API.

### Phase Summary
**Concluída** (docs + verificação). Publish deixado para aval do usuário.

## Final Recap
Migração `Compute → ComputedValue` **concluída em 6 fases, sempre-verde (242/242, 0 warnings)** na branch
`feature/computed-value`, e — de forma inesperada e melhor — **não-breaking**.

**O que foi entregue:**
- **Tipos core** (`ComputedValue` struct de dois campos + tag; `Error` smart-enum alloc-free) com a superfície
  acordada: `TryGet*`/`As*`/`To*`, `EnumerateValues`, `AsObject`/`From`, `Error` com `ToString()=Display`.
- **Todos os 64 nós** implementam `Evaluate(ctx) : ComputedValue` nativamente; a coerção é `internal` e fluente
  (`value.CoerceToNumber(out …) is { } error`).
- **Cache `ComputedValue`**: `Workbook._cache` guarda o struct inline → elimina os boxes de vida longa que o
  experimento mediu promovendo a **Gen1**. Este é o ganho de GC, agora no core de produção.
- **API pública**: `Evaluate(ctx/Workbook) : ComputedValue` (sem boxing) + `Compute(ctx/Workbook) : object?`
  (interop, inalterado). `GetCellValue : object?` inalterado.

**Decisão tomada (Fase 8) — remover `Compute` de vez:** por ser uma lib pré-1.0 sem consumidores, o usuário
optou por cortar a dívida de interop em vez de mantê-la. `Expression.Compute(object?)` foi **removido**;
`Evaluate(ctx/Workbook) : ComputedValue` é a **única** API de avaliação. Todas as superfícies `object?` irmãs
(`Workbook.GetCellValue`, `RangeReference`/`UnionReference`/`ArgumentFlattening`/`Criteria`/`ValueCoercion`)
também caíram — `GetCellValue` agora é público retornando `ComputedValue`. Quem quer `object?` chama `.AsObject()`
no resultado. Isto é um **breaking change (major bump)**, assumido de propósito enquanto é barato.

**Caminho de range fechado (Fase 7):** os helpers `NumericAggregation`/`ArgumentFlattening`/`Criteria` e as
leituras `ExpandComputedValues`/`CellComputedValueAt` agora consomem `ComputedValue` **direto do cache** (sem
`AsObject`), e o resultado-range do `OFFSET` (`Kind = Reference`) é enumerado via `EnumerateValues` — exatamente
o mecanismo para o qual o `ComputedValue` foi desenhado (carregar referências). `SUM/AVERAGE/MIN/MAX/COUNT`,
`CONCAT/TEXTJOIN`, `COUNTIF(S)/SUMIF(S)`, `MATCH/VLOOKUP/XLOOKUP/INDEX` e `NPV/IRR` ficaram **0 box** na leitura
de range, batendo com o `CvEngine` do experimento. Nenhum consumidor `object?` interno resta (as versões
`ExpandValues`/`CellValueAt`/`Flatten`/`Expand` `object?` sobrevivem só como interop público, delegando via
`AsObject`). Resíduo mínimo: o binding de nomes do `LET` guarda `object?` (escalar, 1× por binding — não-range).

## Phase 7: Caminho de range/agregação end-to-end ComputedValue
Status: Complete

- [x] `RangeReference`/`UnionReference`: `ExpandComputedValues`/`CellComputedValueAt` (leem `GetCellComputedValue`);
      `ExpandValues`/`CellValueAt` `object?` viram views de interop (delegam via `AsObject`).
- [x] `ComputedValue.EnumerateValues` lê `ComputedValue` direto (sem `From`).
- [x] `NumericAggregation`/`ArgumentFlattening`/`Criteria` consomem `ComputedValue` (switch por `Kind`).
- [x] Agregação (`Sum/Average/Min/Max/Count`), variádicos (`Concat/Concatenate/TextJoin/CountA/CountBlank`),
      condicionais (`CountIf(s)/SumIf(s)`), lookups (`Match/Index/VLookup/XLookup`) e `NPV/IRR` migrados.
- **Verificação:** 242/242 verde a cada lote (7a agregação, 7b variádicos/condicional, 7c lookups+NPV/IRR).
  Grep confirma: zero chamadas internas às versões `object?` dos helpers de range.

## Phase 8: Remoção do `Compute` (breaking, major)
Status: Complete

- [x] Migrado o escopo de nomes do `LET` para `ComputedValue` (`EvaluationContext.WithName`/`TryGetName`,
      `NameReference`) — elimina o último `object?` interno.
- [x] Removido `Expression.Compute(ctx)`/`Compute(Workbook)`; `Evaluate` é a única API.
- [x] Removidas as superfícies `object?` irmãs (todas mortas, 0 uso): `Workbook.GetCellValue:object?` (o
      `ComputedValue` virou público como `GetCellValue`), `RangeReference`/`UnionReference` `ExpandValues`/
      `CellValueAt` `object?`, `ArgumentFlattening.Flatten`/`Expand` `object?`, `Criteria.Parse`/`Matches`
      `object?`, `ValueCoercion.TryTo*`/`AreEqual`/`Compare` `object?`.
- [x] Migrados ~102 call sites de testes + 4 do benchmark (`.Compute(x)` → `.Evaluate(x).AsObject()`); README
      atualizado (só `Evaluate` + `.AsObject()` para interop). `ComputedValue.From`/`AsObject` permanecem (a
      ponte usada por `FunctionCall`, cujo delegate `CustomFunction` retorna `object?`).
- **Verificação:** build de todos os projetos 0 warnings; **241/241 verde**; grep confirma zero `Compute`
  público no core.

## Deployment Plan
Biblioteca — a mudança é **breaking** (semver **major** → `1.0.0`; ou `versionize` derivando do commit
`feat!`/`BREAKING CHANGE`). Passos (só com aval do usuário; nada externo executado):
1. Merge de `feature/computed-value` → `main` (PR ou fast-forward).
2. Deixar o CI (`versionize` + release GitHub Actions/OIDC) publicar no NuGet — o commit `feat!` dispara major.
3. Nota de release / guia de migração: `Compute(...)` foi removido → use `Evaluate(...)` (retorna
   `ComputedValue`); extraia com `TryGetNumber`/`ToDouble`/`AsDouble`/`TryGetError`, ou `.AsObject()` para o
   `object?` de interop. `Workbook.GetCellValue` agora retorna `ComputedValue`.
4. Resíduo opcional (não-breaking, futuro): migrar o delegate `CustomFunction` para retornar `ComputedValue`
   (hoje retorna `object?`, embrulhado via `From`), eliminando o último `object?` da superfície pública.
