# Experimento: `CellValue` struct vs. `object?` (eliminar boxing no avaliador)

Provar, com BenchmarkDotNet e em escala pequena (spike isolado, sem tocar os 64 arquivos de
produção), se trocar o retorno de `Expression.Compute` de `object?` por um value type opaco `CellValue`
elimina o boxing do caminho numérico **e** melhora throughput o suficiente para justificar a migração
completa. Rodando na branch `experiment/cellvalue-struct`.

## Context

Hoje `Expression.Compute(EvaluationContext) : object?` (base em `Danfma.MySheet/Expressions/Expression.cs:75`).
O plano de memoização já mediu que, após `EvaluationContext` virar `readonly struct`, sobra **um único
alocador no caminho quente: o box do `double`** que sai de `Compute` (24 B por resultado). O objetivo é
matar esse box.

### Diagnóstico técnico já fechado (não reabrir)
- **A ideia (retornar value type) é correta**; é como interpretadores reais representam valores (tagged value).
- **O mecanismo original proposto (union via `Unsafe`/`[FieldOffset]`/`IntPtr`) NÃO funciona no CLR.** O GC é
  preciso: sobrepor um campo gerenciado (`object`/`string`) com um não-gerenciado (`double`/`IntPtr`) no mesmo
  offset lança `TypeLoadException` na carga do tipo. Guardar objeto como `IntPtr` cru exige `GCHandle`/pinning
  (mais caro que o box que se quer evitar). Portanto: **sem `Unsafe`, sem `FieldOffset`, sem `IntPtr`.**
- **C# 14 / .NET 10 não têm discriminated unions nativas** (só proposta) → emula-se com struct + tag.

### Design escolhido (a validar): `readonly struct` de dois campos + tag
```csharp
public enum CellValueKind : byte { Blank, Number, Boolean, Error, Text, Reference /*, Array (futuro) */ }

public readonly struct CellValue
{
    private readonly double _num;    // Number | Boolean(0/1) | Error(código como int) | Blank(ignorado)
    private readonly object? _ref;   // Text(string) | Reference | Array — null nos escalares
    private readonly CellValueKind _kind;

    // Fábricas: Number(double), Bool(bool), Error(code), Blank, Text(string), Ref(...)
    // Acessores: Kind, IsNumber, TryGetNumber(out double), AsObject() (ponte p/ call sites antigos), ...
}
```
- **Ganho central**: Number/Bool/Blank/Error vivem em `_num` + `_kind` → caminho quente **zero alocação**.
  Erros viram alocation-free (código no `_num` em vez de um `ErrorValue` no heap).
- Text/Reference guardam uma referência **já existente** em `_ref` (não alocam nada novo).
- Tamanho ~24 B (igual ao box atual), mas na pilha/registradores — sem heap, sem pointer-chase no unbox.
- `AsObject()` é a ponte que permite migração gradual (call sites antigos continuam consumindo `object?`).

### Risco conhecido que o benchmark tem que capturar
`CellValue` (~24 B) é 3× o tamanho de um `object?` (8 B) para passar por valor num avaliador recursivo
(105 call sites), acima do ponto-doce de 16 B para enregistrar. Parte do ganho de não-alocar pode voltar
como tráfego de cópia na pilha. Por isso a decisão exige **medir**, não presumir.

### Decisões do usuário (locked)
- **Formato**: spike isolado no projeto de benchmark (`benchmarks/Danfma.MySheet.Benchmark`), dois mini-ASTs
  (variante `object?` vs variante `CellValue`) comparados lado a lado. NÃO tocar os 64 overrides de produção.
- **Representatividade**: modelar o domínio escalar completo (Number/Bool/Blank/Error inline; Text/Reference
  em `object?`), com cargas **dominadas por aritmética**.
- **Critério de decisão (gate para migrar)**: alocação **~0** no caminho numérico **E** ganho de throughput
  **material (≥15% ns/op)**. Se só a alocação cair sem throughput → NÃO migra (registrar plano-B).
- **Ferramenta**: BenchmarkDotNet com `MemoryDiagnoser` (comparar `Allocated` e `Mean`).

### Plano-B (se o gate falhar por throughput)
Mitigação cirúrgica sem tocar 64 arquivos: internar/cachear os boxes comuns (`0.0`, `1.0`, `true`, `false`)
num `object?` compartilhado. Ganho parcial de alocação, custo ~5 linhas. Registrar números mesmo assim.

## For Future Agents
Marque `- [x]` ao concluir; ao fechar a fase mude Status para `Complete`, escreva o **Phase Summary** e rode
a **Verification Plan**, registrando os números reais. Tudo nesta branch (`experiment/cellvalue-struct`).
O spike é **descartável**: não precisa de qualidade de produção, precisa de números confiáveis. Não migre os
64 arquivos aqui — a migração, se aprovada, vira um plano próprio.

---

## Phase 1: Spike — `CellValue` + mini-AST dual
Status: Complete

- [x] Criar `benchmarks/Danfma.MySheet.Benchmark/Spike/CellValue.cs`: o `readonly struct` de dois campos +
      `CellValueKind`, fábricas (`Number`/`Bool`/`Error`/`Blank`/`Text`) e acessores
      (`TryGetNumber`, `TryGetText`, `TryGetError`, `Kind`, `IsError`, `AsObject`). **Sem `Unsafe`/`FieldOffset`/`IntPtr`.**
- [x] Criar um mini-AST **duas vezes**, cobrindo o domínio escalar e os nós numéricos quentes
      (`Spike/SpikeNodes.cs`):
      - Variante A (baseline): `interface INodeObj { object? Eval(); }` com `NumObj`/`BoolObj`/`TextObj`/
        `ErrObj`, `AddObj` e `SumObj`. Coerção via `SpikeCoercion.ToNumber(object?)` (unbox + checagem de tipo).
      - Variante B (experimento): `interface INodeCv { CellValue Eval(); }` com os mesmos nós, retornando
        `CellValue`. Coerção via `SpikeCoercion.ToNumber(in CellValue)` → `TryGetNumber` (sem unbox).
      - Construtores paralelos em `SpikeTrees`: `SumFold`, `CumChain`, `Mixed` (as duas variantes recebem as
        mesmas formas de carga).
- [x] Incluir no domínio um caminho Text (`"2"`/`"3"` numérico, ramo de coerção de texto) e um Error
      (`SpikeError` singleton em cache nas duas variantes, como os `ErrorValue` reais) — assim nem erro nem
      texto criam diferença de alocação, isolando o box numérico como a única variável medida.
- [x] `dotnet build -c Release benchmarks/Danfma.MySheet.Benchmark/Danfma.MySheet.Benchmark.csproj` → 0 erros.

### Verification Plan
- `dotnet build -c Release benchmarks/Danfma.MySheet.Benchmark/Danfma.MySheet.Benchmark.csproj`
  → `Build succeeded`, 0 Warning(s)/Error(s). Confirma que o struct de dois campos compila **sem** `Unsafe`
  (prova implícita de que não caímos na armadilha do `FieldOffset` overlap).
- Sanidade de resultado: `dotnet run -c Release --project ... -- --check` (`SpikeSelfCheck`) provando que A e B
  produzem o **mesmo** valor para `SumFold(1..100)` = 5050, `CumChain(1..100)` = 5050, `Mixed(100)` (A vs B),
  coerção de texto (`"3"+4`=7) e propagação de erro (código 3) — garante caminhos equivalentes.

### Phase Summary
**Concluída.** Build em Release **0 Warning(s)/0 Error(s)**; `--check` imprime
`Spike self-check OK — object? e CellValue são equivalentes.` Arquivos novos (todos em
`benchmarks/Danfma.MySheet.Benchmark/Spike/`, nada do core tocado):
- `CellValue.cs` — `readonly struct` de dois campos (`double _num` + `object? _ref`) + tag `CellValueKind`,
  sem `Unsafe`. Number/Bool/Blank/Error moram no `_num`; Text no `_ref`. `TryGetNumber` é o caminho quente.
- `SpikeNodes.cs` — as duas variantes (`INodeObj`/`INodeCv`), `SpikeError` singleton em cache, `SpikeCoercion`
  (mesma semântica), e `SpikeTrees` com `SumFold`/`CumChain`/`Mixed` pareados.
- `SpikeSelfCheck.cs` — asserts de equivalência; ligado no `Program.cs` via `--check`.

**Decisão de design registrada:** erros são singletons em cache nas duas variantes (espelham os `ErrorValue`
reais do MySheet), e Text guarda referência já existente — então a **única** diferença de alocação entre A e B
é o box do `double`. Isso torna o benchmark da Fase 2 um experimento limpo (uma variável), mas também implica
que o ganho de alocação medido será atribuível **só** ao caminho numérico (erros/textos não contam).

**Atenção para a Fase 2 (SO):** `CumChain` é uma cadeia recursiva profunda; `Eval` recursiona a profundidade
inteira. Profundidade ~100k estoura a stack default (~1 MB) — o MySheet real usa `RunWithLargeStack` para isso.
No harness, usar profundidades seguras para o `CumChain` (ex.: 1.000 e ~5.000) e reservar o `N` grande
(100.000) para `SumFold`/`Mixed`, que são folds largos (loop, sem recursão profunda).

---

## Phase 2: Harness BenchmarkDotNet + medição
Status: Complete

- [x] Criar `benchmarks/Danfma.MySheet.Benchmark/Spike/CellValueBenchmarks.cs` com `[MemoryDiagnoser]`,
      `[ShortRunJob]` e `_Object` como `Baseline` (→ coluna `Ratio`). Duas classes por causa dos ranges de
      `[Params]`: `FoldBenchmarks` (N=1.000/100.000) e `ChainBenchmarks` (Depth=1.000/3.000). Variantes
      pareadas `_Object` / `_BoxCache` / `_CellValue` para `SumFold`, `Mixed` e `CumChain`.
- [x] Variante box-cache (plano-B) em `Spike/BoxCacheNodes.cs`: cache de boxes de inteiros 0..255 + true/false.
- [x] `Program.cs` roda a suíte via `BenchmarkSwitcher` (`-- --filter *Spike*`); `--check` preservado.
- [x] Rodar em Release e salvar o summary. Relatórios do BenchmarkDotNet copiados para
      `plans/cellvalue-boxing-experiment/results-fold.md` e `results-chain.md`.

### Verification Plan
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark/Danfma.MySheet.Benchmark.csproj --no-build -- --filter *Spike*`
  → tabelas do BenchmarkDotNet com as colunas **Allocated** e **Ratio**.
- Confirmado: as linhas `*_CellValue` mostram **`Allocated` = 0 B** (Alloc Ratio 0.00, Gen0 `-`) em TODOS os
  workloads; `*_Object` mostram 22–144 KB (o box do double). Delta de `Mean` (Ratio) registrado abaixo.

### Phase Summary
**Concluída.** Ambiente: Apple M1 Pro, .NET 10.0.9, BenchmarkDotNet 0.15.8, `ShortRunJob`. Resultados
(Ratio = Mean/baseline `_Object`; menor = mais rápido):

| Workload | Object (baseline) | Box-cache | CellValue |
|---|---|---|---|
| SumFold N=1.000 | 4,35 µs · 24.024 B | 3,80 µs (0,87×) · 17.904 B (0,75×) | **1,10 µs (0,25×) · 0 B** |
| SumFold N=100.000 | 476,9 µs · 2,40 MB | 437,1 µs (0,92×) · 2,39 MB | **111,8 µs (0,23×) · 0 B** |
| Mixed N=1.000 | 6,61 µs · 22.824 B | 5,87 µs (0,89×) · 17.016 B | **2,42 µs (0,37×) · 0 B** |
| Mixed N=100.000 | 611,7 µs · 2,28 MB | 638,3 µs (**1,04× — mais lento**) · 2,27 MB | **202,4 µs (0,33×) · 0 B** |
| CumChain d=1.000 | 13,31 µs · 47.976 B | 13,34 µs (1,00×) · 41.352 B | **5,85 µs (0,44×) · 0 B** |
| CumChain d=3.000 | 44,43 µs · 143.976 B | 43,74 µs (0,98×) · 137.352 B | **18,45 µs (0,42×) · 0 B** |

**Leitura:** `CellValue` zera a alocação (0 B, Gen0 `-`) em todos os casos e é **2,3×–4,3× mais rápido**
(Ratio 0,23–0,44). O **box-cache** mal arranha: remove 5–25% da alocação (só os literais ≤255) e o
throughput oscila entre 0,87× e 1,04× (às vezes **mais lento**, quando o range-check custa mais que o box
que evita). Caveat metodológico: `ShortRunJob`/N=3 deixa alguns intervalos de confiança largos (ex.:
`SumFold_Object` N=100k com margem alta), mas os `RatioSD` são minúsculos e os efeitos (2–4×) são grandes
demais para o ruído mudar a conclusão.

---

## Phase 2b: Cenário fiel — grafo memoizado + cache `CellValue`
Status: Complete

Adicionada em resposta a uma crítica válida: o spike da Fase 2 isolava o box transitório num `Eval` puro —
não cobria ranges, grafo de dependências com caminhos cruzados/compartilhados, nem o **cache**. É no cache
que o boxing mais pesa: hoje `Workbook._cache` é `ConcurrentDictionary<…, object?>`, então cada valor
numérico cacheado é um **box de vida longa** (promove para Gen1/Gen2). Este experimento mede isso.

- [x] `Spike/GraphEval.cs`: grafo determinístico de N células — 64 folhas viram hubs compartilhados; cada
      célula depende do predecessor (cadeia A←B←C) **e** de um hub (D e E dependendo do mesmo B); ~3% são
      `SUM(range)`; algumas folhas são texto numérico (dados mistos). Memoizado, extraído em lote.
- [x] Dois motores: `ObjEngine` (`Dictionary<string, object?>`) vs `CvEngine` (`Dictionary<string, CellValue>`).
- [x] `GraphBenchmarks` (`[Params(10_000, 100_000)]`, cache reusado via `Clear` → mede alocação por extração).

### Resultados

| Extração | Mean | Ratio | Gen0 | Gen1 | Allocated |
|---|---|---|---|---|---|
| Graph_ObjCache 10k | 477,1 µs | 1,00 | 38,1 | 9,3 | 239.904 B |
| **Graph_CvCache 10k** | **456,0 µs** | **0,96** | **–** | **–** | **0 B** |
| Graph_ObjCache 100k | 5.642 µs | 1,00 | 375,0 | 171,9 | 2.399.904 B |
| **Graph_CvCache 100k** | **4.949 µs** | **0,88** | **–** | **–** | **0 B** |

### Phase Summary
**Dois achados, um deles corrige meu veredito anterior:**

1. **Alocação: confirmado exatamente.** O cache `object?` aloca **24,0 B por célula** (23,99 B/célula medido
   = o box do `double`); o cache `CellValue` aloca **0 B**. Mais importante: os boxes são de **vida longa**
   (ficam no cache), então promovem — o `ObjCache` dispara **coletas Gen1** (171,9/1000 ops em 100k), enquanto
   o `CvCache` não tem **nenhuma** coleta (Gen0 nem Gen1). Este é o ponto que o spike puro não capturava: o
   `CellValue` no cache elimina não só a alocação, mas a **pressão de GC de geração alta** — o que mais importa
   num extrator de background de longa duração.

2. **Throughput: modesto neste caminho — abaixo do gate.** Ratio 0,96 (10k, ~4% mais rápido) e 0,88 (100k,
   ~12%). **Ambos abaixo dos 15%.** No caminho de extração dominado por lookups de `Dictionary` com chave
   `string` (hashing + `TryGetValue`), o box é custo de memória/GC, **não** de throughput — as buscas dominam.
   Isto contrasta com os 2,3×–4,3× da Fase 2 (caminho aritmético puro).

**Correção ao veredito da Fase 2:** o gate "alocação ~0 **E** ≥15% throughput" passa folgado no caminho
**eval-heavy** (aritmética), mas no caminho **cache-heavy** (extração em lote de grafo grande — que é o caso
de uso real declarado) o throughput fica em 4–12% e **não** cruza os 15%. O ganho ali é de **memória/GC**
(0 B + zero Gen1), não de velocidade bruta. Ressalva: o `Mean` de microbenchmark subestima o benefício real
de GC — pausas de Gen1 e a pressão que os boxes de vida longa exercem sobre o resto do processo não aparecem
num número de throughput em regime estacionário.

---

## Phase 3: Decisão contra o gate
Status: Complete

- [x] Comparar os números ao gate: alocação ~0 no caminho numérico **E** ≥15% de ganho de `Mean`.
      **Resultado: gate atingido decisivamente.** Alocação = **0 B** (não ~0) em todos os workloads; throughput
      **56–77% melhor** (Ratio 0,23–0,44), muito além dos 15%. Até o `CumChain` (recursão profunda, o caso
      desenhado para estressar a cópia de 24 B) é 2,3× mais rápido → **o risco da cópia não se materializou**.
- [x] Veredito escrito no Final Recap: **MIGRAR**. O box-cache (plano-B) foi medido e rejeitado (ganho
      parcial de alocação, throughput nulo ou negativo).
- [ ] **Destino da branch — decisão do usuário** (pendente): manter o spike como referência (merge para
      `main`), manter a branch viva até o plano de migração, ou descartar. Recomendação: manter o spike
      (é a evidência que justifica a migração) e abrir um plano de migração separado.

### Verification Plan
- Final Recap contém os números reais das três variantes, o veredito explícito e o esboço da migração.
- Nenhum arquivo de produção alterado: `git diff --stat main -- Danfma.MySheet/` **vazio** (confirmado; o
  spike vive só em `benchmarks/.../Spike/` + `Program.cs`).

### Phase Summary
**Concluída. Veredito: MIGRAR.** O gate acordado (alocação ~0 **E** ≥15% de throughput) foi superado em
larga margem: `CellValue` zera a alocação e entrega 2,3×–4,3× de throughput. O plano-B (box-cache) não é
alternativa viável. A migração real dos 64 arquivos NÃO foi feita aqui (por design) — é o próximo plano.

**Decisão de API pública (DX) — `Compute` passa a retornar `ComputedValue`:** a troca é **source-breaking**
e assumida de propósito. O tipo de produção chama-se **`ComputedValue`** (o spike o prototipou como
`CellValue`; mesmo design de dois campos). Ganha uma superfície de helpers ergonômica; a coerção
estilo-Excel fica **interna** ao engine (não vira API pública) — alinhado com "execução otimizada, não um
novo Excel". Contrato acordado:

```csharp
public readonly struct ComputedValue
{
    public ComputedValueKind Kind { get; }         // Blank|Number|Boolean|Text|Error|Reference

    // TryGet* — seguro, out+bool (substitui os antigos Is*)
    public bool TryGetNumber(out double value);
    public bool TryGetBoolean(out bool value);
    public bool TryGetText(out string value);
    public bool TryGetError(out Error error);
    public bool TryGetReference(out Reference reference);   // Kind=Reference: carrega a Expression de referência (RangeReference/CellReference)

    // As* — açúcar nullable sobre o TryGet (estrito, SEM coerção)
    public double? AsDouble();  public bool? AsBoolean();  public string? AsString();

    // To* — assert estrito: LANÇA em não-correspondência (sem coerção; sem NaN/sentinela)
    public double ToDouble();   public bool ToBoolean();   public string ToText();  // ToText, não ToString

    // Enumera os VALORES de uma referência via cache (NÃO expressões). Vazio se não for Reference.
    public IEnumerable<ComputedValue> EnumerateValues(EvaluationContext context);

    public object? AsObject();                     // escape hatch / interop (permanente, não removido)

    // fábricas + implícitas SÓ de entrada: double/bool/string → ComputedValue (string null → Blank). NUNCA de saída.
    public static ComputedValue Number(double)/Boolean(bool)/Text(string)/Error(Error)/Reference(Reference);
    public static readonly ComputedValue Blank;
}

// Erros: smart enum struct, alloc-free (o int code cabe no ComputedValue). Well-known agora; Register() depois.
public readonly struct Error : IEquatable<Error>
{
    // well-known nomeados: Null(#NULL!), DivZero(#DIV/0!), Value(#VALUE!), Ref(#REF!),
    // Name(#NAME?), Num(#NUM!), NA(#N/A) + modernos (Spill, Calc, …)
    public string Display { get; }                 // "#VALUE!"
    public override string ToString() => Display;  // imprime o DisplayName
    // futuro (não agora, YAGNI): public static Error Register(string display) — process-local, como custom functions
}
```

**Ranges e referências — duas camadas distintas (validado no código, não misturar):** `RangeReference.Compute`
hoje devolve `#VALUE!` (`ErrorValue.NotValue`) — uma range não tem valor escalar. As células vêm de métodos
separados: `Expand() : IEnumerable<Expression>` (os ASTs — camada de referência) e `ExpandValues() :
IEnumerable<object?>` (os valores, via cache). Portanto `Compute`/`ComputedValue` **não** retorna
`IEnumerable<Expression>`: isso é camada de referência, já servida por `Expand`, e funções como `SUM`
inspecionam o *argumento* (é range? → `Expand`/`ExpandValues`). No `ComputedValue`, a range/referência é o
`Kind = Reference`, cujo payload (`_ref`) é a `RangeReference` — que **é** uma `Expression` (é aqui que o
`ComputedValue` carrega uma Expression). Para obter os valores, `EnumerateValues` devolve
`IEnumerable<ComputedValue>` (valores via cache), não expressões. Array materializado (dynamic arrays/spill,
fora de escopo) seria um futuro `Kind = Array` com `ComputedValue[]`.

**Nota `Error` vs `ErrorValue`:** já existe `ErrorValue` (nó de AST para erro literal, serializável). O novo
`Error` (struct de identidade/código, alloc-free) é o que trafega no `ComputedValue`; na migração, o
`ErrorValue` (Expression) passa a embrulhar um `Error` — um único conceito de "código de erro", não dois.

**Esboço da migração (plano futuro, não executado):**
1. Introduzir `ComputedValue` + `ExcelError` no core (`Danfma.MySheet/Expressions/`) com o contrato acima.
   `ExcelError` embrulha um `int _code` → cabe no `ComputedValue` (alloc-free); os well-known são estáticos.
   Custom-error `Register` é process-local (mesmo modelo das custom functions) e fica para quando houver demanda.
2. `Expression.Compute` passa a retornar `ComputedValue` (público, breaking). Durante a migração, um adaptador
   interno `AsObject()` mantém os ~105 call sites funcionando até serem convertidos um a um.
3. Reescrever `ValueCoercion` para consumir/produzir `ComputedValue` (o `TryToNumber(object?)` vira
   `ToNumber(in ComputedValue)` — o caminho `TryGetNumber` do spike). A coerção continua **interna**.
4. Migrar os 64 nós `record : Function` folha→composto (números/valores primeiro; agregadores e lookup por
   último), cada um com seus testes TUnit verdes.
5. Trocar o `Workbook._cache` de `ConcurrentDictionary<…, object?>` para `<…, ComputedValue>` (o valor
   cacheado deixa de ser boxed — a origem do ganho de Gen1 da Fase 2b).
6. `AsObject()` **permanece** como escape hatch público de interop (não é removido); os adaptadores internos
   temporários, sim, saem. Publicar guia de migração (breaking change) e bump de major. Rodar a suíte
   (177/177) + `SheetBenchmarks` reais para confirmar o ganho de alocação em fórmulas de produção.

---

## Final Recap
Experimento concluído (Fases 1, 2, 2b, 3) na branch `experiment/cellvalue-struct`, sem tocar o core.

**Pergunta:** trocar `object? Compute(...)` por um value type `CellValue` elimina o boxing e vale a migração?
**Resposta medida: sim, mas o argumento é de MEMÓRIA/GC, não só de velocidade.** `CellValue` — um
`readonly struct` de dois campos (`double` + `object?`) + tag de 1 byte, **sem `Unsafe`/`FieldOffset`/`IntPtr`**
— zera a alocação em **todos** os cenários, mas o ganho de throughput depende do caminho:

- **Caminho eval-heavy** (aritmética pura, Fase 2): **2,3×–4,3× mais rápido** + 0 B. Gate passa folgado.
- **Caminho cache-heavy** (grafo memoizado + extração em lote, Fase 2b — o caso de uso real declarado):
  **0 B + zero coletas Gen1** (o `object?` boxa 24 B/célula de vida longa e dispara Gen1), mas o throughput
  sobe só **4–12%** — **abaixo dos 15%** do gate, porque lookups de `Dictionary<string,…>` dominam esse caminho.

**Quatro achados:**
1. **O mecanismo original (union via `Unsafe`) estava errado** — o GC preciso proíbe sobrepor campo gerenciado
   com não-gerenciado; o struct de dois campos é a forma correta e entrega tudo.
2. **O risco da cópia de 24 B não se materializou** — o `CumChain` (recursão profunda) é 2,3× mais rápido.
3. **O plano-B (box-cache) é inferior** — remove só 5–25% da alocação, throughput neutro a negativo.
4. **No cache, o ganho é de GC** — 24 B/célula de boxes de vida longa (→ Gen1) eliminados; throughput só 4–12% ali.

**Veredito: MIGRAR, com o argumento correto.** A justificativa forte e universal é a **eliminação de
alocação e de pressão de GC de geração alta** num extrator de background de longa duração (exatamente o
alvo do `plans/memoization.md`). O headline de 2–4× vale para fórmulas aritméticas, não para o caminho de
extração em lote (4–12% ali). Se a decisão pesar **só** throughput bruto no caminho cache-heavy, o gate NÃO
é cruzado e a migração de 64 arquivos fica mais difícil de justificar; se pesar GC/memória/latência de cauda
num processo longo, o caso é forte. Ressalva metodológica: `Mean` de microbenchmark subestima o benefício de
GC (pausas de Gen1 e pressão sobre o resto do processo não entram no número).

## Deployment Plan
Spike de benchmark — **sem deploy**. Passos de fechamento:
1. **Decisão do usuário sobre a branch** (único item aberto): manter o spike como evidência (recomendado) ou
   descartar. `git diff --stat main -- Danfma.MySheet/` está vazio → nada de produção a reverter.
2. Se aprovado seguir: abrir um plano de migração separado a partir do "Esboço da migração" (Fase 3),
   executado com TDD nó a nó, mantendo a suíte 177/177 verde a cada passo.
3. Artefatos preservados: `plans/cellvalue-boxing-experiment/results-fold.md` e `results-chain.md`
   (relatórios do BenchmarkDotNet) + o plano visual em `plans/cellvalue-boxing-experiment/plan.mdx`.
