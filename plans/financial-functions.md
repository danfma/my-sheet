# Funções Financeiras (PMT, PV, FV, NPER, IPMT, PPMT, NPV, RATE, IRR)

Implementar as 9 funções financeiras periódicas do Excel no motor de fórmulas do
MySheet, seguindo o padrão existente (um `record : Function` por arquivo, registro
no `Parser`, serialização MemoryPack, propagação de `ErrorValue`). Compatibilidade
com Excel é o critério de corretude.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done,
set its status to `Complete` and write its **Phase Summary** (what was done, key
decisions, anything needed to continue with zero context); run the phase's
**Verification Plan** and record the result before moving on. When all phases are
done, fill in **Final Recap** and **Deployment Plan**.

### Contexto fixo do projeto (não muda entre fases)

- **Padrão de função**: cada função é um `[MemoryPackable] public sealed partial record Nome(Expression[] Arguments) : Function` com `override object? Compute(EvaluationContext context)`. Um arquivo por função em `Danfma.MySheet/Expressions/`.
- **Coerção numérica**: `ValueCoercion.TryToNumber(value, out double n)` retorna `ErrorValue?` (`null` = ok). Propagar o erro retornando-o. Blank→0, bool→1/0, string numérica→parse invariante.
- **Argumentos opcionais omitidos**: o parser injeta `BlankValue.Instance` para vírgulas vazias/ausentes; `TryToNumber(blank)`→0. Para defaults diferentes de 0 (ex.: `type` default 0 já é 0; `guess` default 0.1), checar `Arguments.Length` antes de avaliar o argumento.
- **Erros existentes** (`ErrorValue.cs`): `#VALUE!`, `#NAME?`, `#REF!`, `#DIV/0!`, `#N/A`. **Falta `#NUM!`** — adicionado na Fase 1.
- **Registro**: dicionário `Functions` em `Danfma.MySheet/Parsing/Parser.cs` (linhas ~33-75): `["NOME"] = new(minArgs, maxArgs, args => new Nome(args))`. Aridade errada lança `ParseException` no parse.
- **Serialização**: TODA função precisa de uma tag `[MemoryPackUnion(N, typeof(Nome))]` em `Expression.cs`. Tags são **append-only**; a maior tag atual é **54** (`UnionReference`). Novas tags começam em **55**.
- **Achatamento de ranges**: `ArgumentFlattening.Expand(...)` expande RangeReference/UnionReference para escalares. Usado por SUM e similares. NPV/IRR dependem da **ordem** (período do desconto) — verificar que a expansão preserva ordem linha/coluna antes de confiar.
- **Build/test**: SDK fixado em `global.json` (10.0.301). Rodar `dotnet test` a partir da raiz do repo.
- **Framework de teste**: **TUnit** (não xUnit). Testes são `public class` com métodos `[Test] public async Task` e asserts `await Assert.That(x).IsEqualTo(y)`. Ver `LookupFunctionTests.cs` para o padrão (helpers `Calc`/`CalcMixed`). Novo arquivo por área: `FinancialFunctionTests.cs`.
- **Oráculo de teste (decidido via spike)**: usar o pacote **`ExcelFinancialFunctions`** (namespace `Excel.FinancialFunctions`, classe estática `Financial`) como **dependência test-only** do projeto de testes (NÃO referenciar na lib publicada). Cobre os 9 e replica o Excel. Os testes computam o esperado via `Financial.Pmt/Pv/Fv/NPer/IPmt/PPmt/Npv/Rate/Irr(...)` e assertam MySheet ≈ EFF com tolerância (`1e-6`). API: `Financial.Pmt(rate,nper,pv,fv,PaymentDue.EndOfPeriod|BeginningOfPeriod)`; `Financial.Npv(rate, double[])`; `Financial.Irr(double[], guess)`; `Financial.Rate(nper,pmt,pv,fv,PaymentDue,guess)`.
  - **ClosedXML descartado como oráculo**: spike mostrou que só computa PMT/FV/IPMT; PV/NPER/PPMT/NPV/RATE/IRR → `#NAME?` (não implementados).

### Equação base TVM (Time Value of Money)

`type ∈ {0 = fim do período, 1 = início}`, `r` = taxa, `n` = nper, `factor = (1+r)^n`.

```
r ≠ 0:  pv·factor + pmt·(1 + r·type)·(factor − 1)/r + fv = 0
r = 0:  pv + pmt·n + fv = 0
```

Rearranjos (no helper `TimeValueOfMoney`):

```
Fv(r,n,pmt,pv,type):
  r≠0: -(pv·factor + pmt·(1 + r·type)·(factor − 1)/r)
  r=0: -(pv + pmt·n)

Pv(r,n,pmt,fv,type):
  r≠0: -(fv + pmt·(1 + r·type)·(factor − 1)/r) / factor
  r=0: -(fv + pmt·n)

Pmt(r,n,pv,fv,type):
  r≠0: -(pv·factor + fv) / ((1 + r·type)·(factor − 1)/r)
  r=0: -(pv + fv) / n

Nper(r,pmt,pv,fv,type):
  r=0: -(pv + fv)/pmt          (pmt = 0 → #NUM!)
  r≠0: k = pmt·(1 + r·type)/r
       nper = ln((k − fv)/(k + pv)) / ln(1 + r)
       argumento do ln ≤ 0  ou  1+r ≤ 0  → #NUM!
```

IPMT/PPMT (referência LibreOffice/Apache-POI = match de-facto do Excel):

```
per < 1 ou per > nper → #NUM!
pmt = TVM.Pmt(rate, nper, pv, fv, type)
se per == 1:
    fvBefore = (type==1) ? 0 : -pv
senão:
    fvBefore = (type==1) ? TVM.Fv(rate, per-2, pmt, pv, 1) - pmt
                         : TVM.Fv(rate, per-1, pmt, pv, 0)
IPMT = fvBefore · rate
PPMT = pmt − IPMT
```

NPV / IRR:

```
NPV(rate, cf[]) = Σ_{i=1..n} cf_i / (1+rate)^i      (1º fluxo descontado 1 período)
                  (1+rate) == 0 → #DIV/0!

IRR: achar r tal que Σ_{i=0..n} cf_i/(1+r)^i = 0     (1º fluxo no período 0)
     f'(r) = Σ_{i=0..n} (−i·cf_i)/(1+r)^{i+1}
     sem mudança de sinal nos fluxos → não converge → #NUM!
```

Solver (replicar Excel — Fase 2): **Newton-Raphson**, `guess` default `0.1`,
**máx. 20 iterações**, tolerância **1e-7**; sem convergência → `#NUM!`.
RATE usa derivada numérica (diferença central) sobre a equação TVM, avaliando o
ramo `r=0` quando `|r|` for desprezível para evitar divisão por zero. IRR usa a
derivada analítica acima.

### Assinaturas e aridades

| Função | Assinatura | (min,max) | Tag MemoryPack |
|---|---|---|---|
| PMT  | `(rate, nper, pv, [fv], [type])`              | (3,5) | 55 |
| PV   | `(rate, nper, pmt, [fv], [type])`             | (3,5) | 56 |
| FV   | `(rate, nper, pmt, [pv], [type])`             | (3,5) | 57 |
| NPER | `(rate, pmt, pv, [fv], [type])`               | (3,5) | 58 |
| IPMT | `(rate, per, nper, pv, [fv], [type])`         | (4,6) | 59 |
| PPMT | `(rate, per, nper, pv, [fv], [type])`         | (4,6) | 60 |
| NPV  | `(rate, value1, [value2, ...])`               | (2,∞) | 61 |
| RATE | `(nper, pmt, pv, [fv], [type], [guess])`      | (3,6) | 62 |
| IRR  | `(values, [guess])`                           | (1,2) | 63 |

### Escopo (YAGNI)
Somente as 9 funções acima. **Fora**: XNPV, XIRR, CUMIPMT, CUMPRINC, day-count,
MIRR. Não introduzir refactor não relacionado.

---

## Phase 1: Fundações + funções closed-form (PMT, PV, FV, NPER, IPMT, PPMT, NPV)
Status: Complete

- [x] Adicionar `public static readonly ErrorValue Number = new("#NUM!");` em `ErrorValue.cs`
- [x] Adicionar factory `public static StringValue String(string value) => new(value);` em `Expression.cs` (o `Number(double)` já existia)
- [x] Criar `Danfma.MySheet/Expressions/TimeValueOfMoney.cs` (`internal static`) com `Pmt/Pv/Fv/Nper/IPmt`
- [x] Implementar records: `Pmt.cs`, `Pv.cs`, `Fv.cs`, `Nper.cs`, `Ipmt.cs`, `Ppmt.cs`, `Npv.cs`
- [x] NPV: percorrer argumentos em ordem espelhando `NumericAggregation` (direto conta texto/lógico, referenciado ignora); `(1+rate)==0` → `#DIV/0!`
- [x] Adicionar tags `[MemoryPackUnion(55, typeof(Pmt))]` … `[MemoryPackUnion(61, typeof(Npv))]` em `Expression.cs`
- [x] Registrar PMT/PV/FV/NPER/IPMT/PPMT/NPV no dicionário `Functions` do `Parser.cs`
- [x] ~~Spike ClosedXML~~ → **feito**: ClosedXML cobre só 3/9. Decisão: oráculo = `ExcelFinancialFunctions` (test-only).
- [x] Adicionar `PackageReference Include="ExcelFinancialFunctions"` ao `tests/Danfma.MySheet.Tests/*.csproj`
- [x] Criar `tests/Danfma.MySheet.Tests/Parsing/FinancialFunctionTests.cs` (29 testes + 1 round-trip MemoryPack)

### Verification Plan
- `dotnet test` a partir da raiz → **todos verdes**, incluindo os ~N novos testes da Fase 1 (registrar N no Phase Summary).
- Golden values **confirmados via `ExcelFinancialFunctions` no spike** (tolerância `1e-6` nos asserts):
  - `PMT(0.05, 10, -1000)` = `129.5045749654566`
  - `FV(0.05, 10, -100, -1000)` = `2886.683880332326`
  - `PV(0.05, 10, -100, -1000)` = `1386.0867464592409`
  - `NPER(0.05, -100, 1000)` = `14.206699082890461`
  - `IPMT(0.05, 1, 10, -1000)` = `50.0`  e  `PPMT(0.05, 1, 10, -1000)` = `79.50457496545661`
  - `NPV(0.1, -10000, 3000, 4200, 6800)` = `1188.4434123352216`

### Phase Summary
**Concluída.** As 7 funções closed-form estão implementadas, registradas, serializáveis e testadas
contra o oráculo `ExcelFinancialFunctions`. Suíte completa: **207/207 verde** (177 anteriores + 30
financeiros).

Arquivos criados: `Expressions/TimeValueOfMoney.cs` (helper `internal static` com `Pmt/Pv/Fv/Nper/IPmt`),
`Expressions/{Pmt,Pv,Fv,Nper,Ipmt,Ppmt,Npv}.cs`, `tests/.../Parsing/FinancialFunctionTests.cs`.
Modificados: `Expressions/ErrorValue.cs` (+`#NUM!`), `Expressions/Expression.cs` (+factory `String`,
+tags MemoryPackUnion 55–61), `Parsing/Parser.cs` (+7 registros), `tests/.../*.csproj`
(+`ExcelFinancialFunctions` 3.2.0 test-only).

Decisões-chave (zero-contexto para a Fase 2):
- **`TimeValueOfMoney` é matemática pura `double`** — não retorna `ErrorValue`. Entradas degeneradas
  (ex.: `nper=0`, log de número ≤ 0) viram `NaN`/`Infinity`, e **cada record mapeia não-finito →
  `ErrorValue.Number` (#NUM!)** via `double.IsFinite(result)`. A Fase 2 deve seguir o mesmo contrato.
- **Padrão de leitura de argumentos é inline** (sequência de `ValueCoercion.TryToNumber(...) is { } err`
  → `return err`), espelhando `Round.cs`/`VLookup.cs`. Mantido por consistência, mesmo verboso. RATE/IRR
  devem seguir igual.
- **`type` é normalizado** com `type != 0 ? 1 : 0` antes de chamar o helper.
- **NPV não usa `ArgumentFlattening`**; replica a classificação direto/referenciado de `NumericAggregation`
  (texto/lógico direto conta; referenciado ignora), com desconto por período acumulando `discount *= (1+rate)`.
- **IPMT/PPMT** seguem a referência LibreOffice (`TimeValueOfMoney.IPmt`); `PPMT = Pmt − IPmt`; guarda
  `per < 1 || per > nper → #NUM!`.
- **Oráculo confirmado**: `ExcelFinancialFunctions` cobre os 9 e replica o Excel; ClosedXML cobre só 3.
- **Comando de teste** (TUnit, não `dotnet test`): `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/FinancialFunctionTests/*"`.
- **Formatação**: CSharpier 1.3.0 (`dotnet csharpier format <arquivos>`). CI **não** checa formato (só
  build+test), mas mantemos o padrão. `LookupFunctionTests.cs` está desformatado desde antes (commit
  3460bb3) — não mexer aqui.

### Verificação executada
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → **207/207 verde**.

## Phase 2: Funções iterativas (RATE, IRR)
Status: Complete

- [x] Implementar o solver iterativo em `TimeValueOfMoney.Solve` (compartilhado por RATE e IRR)
- [x] Implementar `Rate.cs`: resolve `Fv(rate,…) = fv` para `r` (reusa o helper `Fv`); `guess` default 0.1; não-convergência → `#NUM!`
- [x] Implementar `Irr.cs`: range único achatado em ordem (`ArgumentFlattening.Expand`); `guess` default 0.1; guarda mudança de sinal; não-convergência → `#NUM!`
- [x] Adicionar tags `[MemoryPackUnion(62, typeof(Rate))]` e `[MemoryPackUnion(63, typeof(Irr))]` em `Expression.cs`
- [x] Registrar RATE/IRR no `Parser.cs` com aridades (3,6) e (1,2)
- [x] Estender `FinancialFunctionTests.cs`: convergência normal, efeito do `guess`, não-convergência → `#NUM!`, IRR sem mudança de sinal → `#NUM!`, round-trip MemoryPack

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → **todos verdes**.
- Golden values **confirmados via `ExcelFinancialFunctions`** (tolerância `1e-6`):
  - `RATE(10, -129.5046, 1000)` ≈ `0.0500000398` (≈ 5%; inverso aproximado do PMT acima)
  - `RATE(360, -600, 100000)` ≈ `0.0050058250` (financiamento de 30 anos — o caso que quebrou o Newton ingênuo)
  - `IRR({-10000, 3000, 4200, 6800})` = `0.16340560068897625` (≈ 16,34%)
  - cross-check independente: NPV avaliado na taxa IRR retornada deve ser ≈ 0.

### Phase Summary
**Concluída.** RATE e IRR implementadas, registradas, serializáveis e testadas contra o oráculo
`ExcelFinancialFunctions`. Suíte completa: **218/218 verde** (207 da Fase 1 + 11 de RATE/IRR).

Arquivos criados: `Expressions/Rate.cs`, `Expressions/Irr.cs`. Modificados: `Expressions/TimeValueOfMoney.cs`
(+`Solve`), `Expressions/Expression.cs` (+tags 62–63), `Parsing/Parser.cs` (+2 registros),
`tests/.../FinancialFunctionTests.cs` (+11 testes).

**DESVIO IMPORTANTE da decisão original do solver (registrar para o usuário):**
- A escolha original era *"replicar o Excel exatamente: Newton ~20 iterações, tol 1e-7, #NUM! senão"*.
- **Newton ingênuo (e o secante do POI/LibreOffice) NÃO replicam o Excel**: ambos falham em
  `RATE(360,-600,100000)` (financiamento de 30 anos), um caso trivial que o Excel resolve. Com `guess=0.1`,
  `(1.1)^360 ≈ 8e14` domina o resíduo e o método local diverge (Newton → `NaN`) ou converge falsamente
  perto de zero (secante). Ou seja, a opção "Newton ingênuo" era **autocontraditória** com o objetivo
  "bater com o Excel".
- **Solução adotada**: `TimeValueOfMoney.Solve` faz **bracketing + bisseção** — robusto, sempre encontra
  a raiz única, convergência 1e-7 na *taxa* (não no resíduo em unidades de moeda). Isso é o que de fato
  reproduz os resultados do Excel/EFF, inclusive o mortgage. É essencialmente a opção "robusta" que havia
  sido descartada — descartada com base na premissa (falsa) de que Newton ingênuo ≈ Excel.
- **Limitação conhecida e aceita**: para IRR com *múltiplas* mudanças de sinal (múltiplas raízes reais), a
  bisseção retorna *uma* raiz no bracket, não necessariamente a mais próxima do `guess` como o Excel faria.
  Todos os casos de teste (e a esmagadora maioria de uso real) têm raiz única, onde o resultado é idêntico
  ao do Excel. Documentado no XML-doc de `Solve`. Se for preciso casar o "guess-nearest" do Excel em
  cenários multi-raiz no futuro, adicionar uma fase Newton-a-partir-do-guess antes do fallback de bisseção.
- O parâmetro `guess` é honrado como centro do bracket; para raiz única não altera o resultado (consistente
  com o Excel nesses casos).

### Verificação executada
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → **218/218 verde**.
CSharpier aplicado aos arquivos novos/alterados.

## Final Recap
As **9 funções financeiras** (PMT, PV, FV, NPER, IPMT, PPMT, NPV, RATE, IRR) estão implementadas no motor
de fórmulas do MySheet, compatíveis com o Excel e verificadas por comparação direta contra a biblioteca
`ExcelFinancialFunctions` (oráculo independente). Cobertura: **+41 testes** (218/218 na suíte total).

Estrutura final:
- `Expressions/TimeValueOfMoney.cs` — helper `internal static`: `Pmt/Pv/Fv/Nper/IPmt` (matemática pura de
  anuidade) + `Solve` (bracketing+bisseção para RATE/IRR).
- `Expressions/{Pmt,Pv,Fv,Nper,Ipmt,Ppmt,Npv,Rate,Irr}.cs` — um `record : Function` por função.
- `Expressions/ErrorValue.cs` — novo código `#NUM!` (`ErrorValue.Number`).
- `Expressions/Expression.cs` — factory `String(string)` + tags MemoryPackUnion **55–63**.
- `Parsing/Parser.cs` — 9 registros novos com aridades Excel.
- `tests/.../Parsing/FinancialFunctionTests.cs` — 41 testes vs. oráculo EFF (test-only `ExcelFinancialFunctions` 3.2.0).

Contrato de erros: argumentos com `ErrorValue` propagam; entradas degeneradas/sem solução → `#NUM!`;
`NPV` com `rate=-1` → `#DIV/0!`; `IPMT/PPMT` com `per` fora de `[1,nper]` → `#NUM!`; `IRR` sem mudança de
sinal → `#NUM!`.

Tags MemoryPack são append-only; o round-trip save/load foi testado para as 9 funções.

## Deployment Plan
Biblioteca NuGet (`Danfma.MySheet`); o release é dirigido por commits convencionais + `versionize` no
workflow `release.yml`. Para publicar:

1. **Pré-check local**:
   - `dotnet build Danfma.MySheet.slnx -c Release`
   - `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → 218/218 verde.
2. **Commit** com mensagem convencional `feat` (ex.: `feat(financial): adiciona PMT, PV, FV, NPER, IPMT,
   PPMT, NPV, RATE, IRR`) — o `feat` faz o `versionize` subir a minor (0.1.0 → 0.2.0). Sem referência ao
   assistente na mensagem (regra do usuário).
3. **Push para `main`** → o CI (`ci.yml`) roda build+test; o `release.yml` versiona, cria a tag e publica no
   NuGet via Trusted Publishing (OIDC).
4. **Dependência test-only**: `ExcelFinancialFunctions` está só no `.csproj` de testes — **não** entra no
   pacote publicado. Confirmar que o `.nupkg` gerado não a referencia (é transitiva de teste apenas).
5. **Pós-publicação**: validar o pacote no NuGet.org e, opcionalmente, smoke-test `=PMT(...)`/`=IRR(...)`
   num consumidor.
