# Spike: Dirty flags + Grafo de Dependências Reversa

Spike para medir se um grafo de dependências reversa + marcação dirty permite recomputar
**somente** o que uma edição pontual afeta (em vez do `InvalidateCache()` global que hoje descarta
todo o cache quente), quanto isso acelera edições pontuais numa planilha de 500k+ células, e a que
custo de memória/build do grafo. Inclui a extração input→output (quais células "no topo" um conjunto
de edições afeta) e uma análise escrita da separação AST/dados (stateless).

**Isto é um SPIKE de medição, não uma feature de produção.** A entrega é um relatório com números e
uma recomendação go/no-go, não um merge para `main`. Todo código novo é aditivo e vive atrás de uma
API nova (`CalculateDirty`), sem substituir `InvalidateCache`/`Recalculate`.

## For Future Agents
Conforme o trabalho avança: marque os checkboxes `- [x]` ao completar; ao terminar uma fase, mude o
Status para `Complete` e escreva o **Phase Summary** (o que foi feito, decisões-chave, tudo que é
preciso para continuar com zero contexto); rode o **Verification Plan** da fase e registre o resultado
antes de seguir. Ao terminar todas as fases, preencha **Final Recap** e **Deployment Plan**.

## Decisões travadas (contexto de projeto)
- **Grafo reverso NÃO materializa aresta por célula para ranges.** `=SUM(A:A)` × muitos agregados
  explodiria a memória (o codebase já brigou contra 35MB/24MB de strings duplicadas). Ranges são
  guardados como `(fórmula, sheet, retângulo/coluna-aberta)` e consultados por **contenção de ponto** —
  mesma filosofia do `SheetStructuralIndex`.
- **Referências dinâmicas → sempre-dirty.** `OFFSET`/`INDIRECT`/`INDEX`-que-retorna-ref/`DynamicRange`/
  `RAND`/`NOW`/`TODAY`/`RANDBETWEEN` não são extraíveis estaticamente da AST; reusam o mecanismo de taint
  volátil e recomputam sempre. (Hoje só `INDIRECT`/voláteis marcam; `OFFSET` passa a marcar também.)
- **Recomputação**: implementar e comparar **duas** estratégias — (4a) *evict-and-pull* (reusa o motor
  pull-based memoizado) e (4b) *scheduler bottom-up topológico* explícito.
- **Contrato de concorrência**: edição → `CalculateDirty` → leitura é single-thread na fase de mutação,
  igual ao contrato de manutenção do índice estrutural. Sem grafo mutando concorrente com avaliação.
- **Portão de corretude (inegociável)**: `CalculateDirty` = bit-idêntico a `InvalidateCache()+ComputeAll()`
  sobre lotes de edições aleatórias.
- **Fixture**: o dono fornece um `.xlsx` de template REAL (~500k células) para a Fase 5. Gerador sintético
  parametrizável é o fallback e serve às Fases 0–4.
- **Stateless (AST/dados)**: entregue como **análise escrita** (Fase 6); o código da separação é um spike
  seguinte.

## Ponto de partida no código (mapa para o executor)
- Choke points de mutação já existem e são os pontos de ancoragem documentados: `Sheet.SetCell(id, expr)`
  e `Sheet.Remove(id)` em `Danfma.MySheet/Workbook.cs` — os comentários já os chamam de *"attach point for
  the future reverse dependency graph"*.
- Endereçamento numérico e store denso: `SheetValueStore` (`SheetValueStore.cs`); o conjunto de células
  volátil-tainted (`_tainted`) e `DropTainted()` são o modelo direto para o conjunto dirty.
- Extração de dependências: as referências estáticas vêm de `CellReference`, `RangeReference`,
  `OpenRangeReference`, `UnionReference`; as dinâmicas de `Offset`/`Indirect`/`Index`/`DynamicRange`/
  `NameReference`. A árvore é percorrível como em `FormulaWriter.Call` / `Workbook.HasUnqualifiedReference`.
- Índice espacial de referência: `SheetStructuralIndex` (`SheetStructuralIndex.cs`) é o molde (buckets por
  coluna/linha, mantidos por escrita, escopo de vida) — mas ele indexa células *populadas*, não *ranges de
  fórmulas*; o índice de contenção da Fase 2 é novo.
- Testes: TUnit, rodados via `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj
  -c Release -- --treenode-filter "..."`. Benchmarks: `dotnet run -c Release --project
  benchmarks/Danfma.MySheet.Benchmark -- --<flag>` (ver `Program.cs`).

---

## Phase 0: Fixture K1 anonimizado (.myxl) + baseline + harness
Status: Complete

O dono forneceu o K1.xlsx real como JSON em `samples/` (dois arquivos CONFIDENCIAIS, não versionar):
`synthetic-sheet-as-json.json` (97MB, 26 sheets, 566.086 células, 549.266 fórmulas; a sheet "Standard
Footnotes 2025" tem 506k células — o caso coluna-inteira) e `synthetic-input.json` (mapa de 2577 chaves→
valor). Fatos apurados: 17 funções distintas, **todas já implementadas** (zero dependências novas); as 2577
chaves de interpolação `{Key}` = exatamente as 2577 chaves do input; 100% das fórmulas começam com `=`;
todo `customExpression` é interpolação; interpolações NUNCA aparecem em células `formula`.

- [x] `.gitignore`: ignorar os dois JSON crus de `samples/` (dados confidenciais) + `k1-anonymized.json` + `*.myxl`.
- [x] Ferramenta `tools/K1FixtureBuilder` (console net10, referencia `Danfma.MySheet`) que transforma o JSON
      cru → JSON anonimizado + `.myxl`:
      - **Anonimização determinística por-valor** (lorem ipsum semeado por hash do original → mesma entrada dá
        mesma saída, preservando igualdade): todo conteúdo de string (scalars E literais `"..."` de fórmulas)
        vira lorem ipsum, então `COUNTIF(...,"Show")` continua casando com a célula "Show" (ambos → mesmo
        lorem); valores numéricos do input e scalars numéricos → número fake de magnitude similar.
      - **Nomes de sheet anonimizados** → `Sheet_{n}` (identificador seguro: não parece ref de célula por
        causa do `_`, e é simples → sem aspas). Os qualificadores cross-sheet nas fórmulas que casam com um
        nome renomeado são reescritos (só 6, os `BOX*_HIDE` → `Sheet_14!`…); os qualificadores por codeName
        (`Sheet8` etc.) não existem como sheet e passam verbatim (já eram `#REF!`, comportamento preservado).
      - **Estrutura preservada** (é o grafo de dependências, o objeto do spike): referências de célula,
        nomes de função, e literais numéricos que fazem parte de refs — intocados.
      - **Sheet "Input"** (3 colunas): para cada uma das 2577 chaves distintas, linha `r`: `A{r}` = nome
        genérico (`IN0001`…), `B{r}` = valor de input anonimizado (ou blank), `C{r}` =
        `=XLOOKUP(A{r},$A$1:$A$2577,$B$1:$B$2577)`.
      - **Reescrita de interpolações**: toda célula scalar/customExpression com `{Key}` vira fórmula
        concatenando os literais anonimizados (`"..."`) com `Input!C{r}` via `&` (ex.: `-{FDK1_USWTHTAX}` →
        `="-"&Input!C{r}`; `{ENT_EIN}` → `=Input!C{r}`). Cria a aresta de dependência para o Input.
      - **Chaves genéricas**: `{Key}` original nunca aparece na saída — nem em `A`/`C` do Input nem nas refs.
      - **Fórmulas estranhas**: em `ParseException`, substituir a célula por um scalar aleatório e contar.
- [x] Rodar a ferramenta: `samples/k1-anonymized.json` (116.7 MB, referência) e `samples/k1.myxl`
      (**4.9 MB**, Brotli — o lorem repetitivo comprime muito bem) gerados; `--verify` recarrega (1.5s) e
      afirma 27 sheets, Input=7731 células, XLOOKUP resolve, S2 (ex-`{ENT_EIN}`)→Text, nomes de sheet
      anonimizados. Fórmulas: 0 falhas de parse; 2/2000 células da sheet grande dão `#REF!` (cross-sheet
      por codeName inexistente — comportamento fiel).
- [x] JSON crus permanecem em disco (gitignored) como fonte de regeneração — NÃO deletados (não foram criados
      por nós; deletar sem forma de regenerar é pior). Decisão de apagar fica com o dono.
- [x] Harness de spike no benchmark (`Spike/DirtyGraph/DirtyGraphHarness.cs`, flag `--dirty-graph [N]`,
      despachado em `Program.cs`) que carrega `k1.myxl`, mede o `ComputeAll` frio e o **baseline** de N
      edições pontuais (editar → `InvalidateCache()` → `ComputeAll()`, com restauro entre iterações).

### Verification Plan
- `dotnet build tools/K1FixtureBuilder`: 0 erros. ✓
- Rodar a ferramenta; conferir no stdout: 26 sheets + 1 Input, ~566k células, 0 funções não implementadas,
  interpolações reescritas ≈ 3271, falhas de parse registradas. ✓ (0 falhas)
- `dotnet run … --verify` recarrega `samples/k1.myxl` e afirma: 27 sheets, `Sheets["Input"].Count` = 7731,
  interpolação conhecida resolve, sheets anonimizadas. ✓
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --dirty-graph 6` imprime o
  baseline sem exceção. ✓

### Phase Summary
**Feito.** Fixture K1 real (566.086 células, 26 sheets; a maior — ex-"Standard Footnotes 2025", agora
`Sheet_3` — tem 506k células) convertido de JSON confidencial para `samples/k1.myxl` **anonimizado** (4.9MB,
Brotli): texto → lorem ipsum determinístico com igualdade preservada, números → fake de magnitude similar,
nomes de sheet → `Sheet_{n}` (com os 6 qualificadores cross-sheet reais reescritos), 2577 chaves `{Key}` →
sheet `Input` (A=nome genérico `IN####`, B=valor, C=`XLOOKUP`) + interpolações reescritas como fórmulas
`…&Input!C{r}`. Ferramenta: `tools/K1FixtureBuilder` (modo default gera; `--verify` valida). Os 2 JSON crus
ficam gitignored em disco como fonte de regeneração.

**Descobertas-chave que moldam o resto do spike:**
- **Zero funções faltando** — o K1 usa 17 funções distintas, todas implementadas. Sem dependências a codar.
- **0 falhas de parse** em 549k fórmulas.
- **Sheet_3 concentra 89% das células (506k)** — é onde o fan-in de coluna-inteira, se existir, vai decidir a
  memória do grafo reverso (Fase 2). Confirmar o perfil dela lá.

**BASELINE (o número a bater):** carregar o fixture = 2.3s; `ComputeAll` frio = 1933ms / 313MB. Uma edição
pontual no modelo atual (tudo-ou-nada: `InvalidateCache` + `ComputeAll`) custa **~1196ms e ~308MB de
alocação, por edição**, independente de quantas células a edição realmente afeta. É exatamente essa
insensibilidade ao tamanho do impacto que o CalculateDirty ataca — headroom enorme se uma edição pontual
tocar só dezenas de células.

**Decisão pendente do dono:** apagar (ou não) os 2 JSON crus de `samples/`.

## Phase 1: Extração de dependências (forward) + classificação de nós dinâmicos
Status: Complete

- [x] `DependencyExtractor.Extract(Expression, Workbook?)` em `Danfma.MySheet/DirtyGraph/DependencyExtractor.cs`
      percorre a AST e retorna um `DependencyScan` (`CellDep`s, `RangeDep`s, flag `AlwaysDirty`): `CellReference`
      → dep de célula; `RangeReference`/`OpenRangeReference` → dep de range (retângulo/aberto); `UnionReference`,
      operadores e args de qualquer função built-in → recursão.
- [x] Nós **dinâmicos/voláteis** → `AlwaysDirty`: `OFFSET`, `INDIRECT`, `DynamicRange`, os voláteis via
      `Expression.IsVolatile`, `FunctionCall` custom (comportamento host desconhecido), e `NameReference`
      irresolúvel. `INDEX`/`CHOOSE`/`VLOOKUP` NÃO são dirty — a dep é o range inteiro que varrem (enumerável).
- [x] Decisão documentada no código: extrai as deps estáticas E marca `AlwaysDirty` (super-aproximação é
      SEGURA para dirty — falso-dirty é inofensivo, dirty-perdido é stale). `NameReference` resolve via
      `Workbook.DefinedNames` (com guarda de ciclo) quando um workbook é passado.

### Verification Plan
- `DependencyExtractorTests` (14 casos): arithmetic→{A1,B2}; SUM(A1:C3)→range; SUM(A:A)→coluna aberta;
  união; OFFSET/INDIRECT/NOW/DynamicRange/custom→always-dirty; INDEX→range (não-dirty); cross-sheet;
  defined-name resolve. `dotnet run … --treenode-filter "/*/*/DependencyExtractorTests/*"` → **14/14 ✓**.
- Suíte completa: **1016/1016 ✓** (código só aditivo, nada existente afetado).

### Phase Summary
**Feito.** `DependencyExtractor` extrai o grafo forward de uma fórmula com super-aproximação segura. 14 testes
dirigidos + suíte completa (1016) verdes.

**Bug pré-existente descoberto E CORRIGIDO (publicado na `main`):** `FormulaWriter.Call` não tinha case para
`Indirect` (MemoryPackUnion tag 318) → lançava `NotSupportedException`; qualquer workbook com `INDIRECT`
quebrava em `FORMULATEXT`, `expression.ToFormula(...)` e `SaveAsExcel(FormulaMode.Formulas)` — possivelmente o
cenário de produção reportado. Corrigido num worktree/subagente isolado, diff minimalista (1 arm + 2 casos de teste
de round-trip), suíte 1004/1004, e **mergeado na main** (commit `935aed0`, `fix(formula-writer): render
INDIRECT …`); o branch do spike foi rebasado por cima. O extrator ainda trata `Indirect` explicitamente (por
clareza + o guard de `VisitArguments`).

## Phase 2: Grafo reverso + índice de contenção de ranges
Status: Complete

- [x] Estrutura de arestas reversas célula→dependentes (`ReverseDependencyGraph`, endereçamento numérico
      `CellDep`). `Build(workbook)` varre cada fórmula 1× via `DependencyExtractor` e inverte as arestas.
- [x] **Índice de contenção de ranges** bucketizado por `(sheet, coluna)`: uma aresta-de-range de eixo-coluna
      fechado indexa em cada bucket de coluna do span (largura pequena, nunca altura); coluna aberta / span
      largo (>256) vai para um fallback por-sheet. Consulta por ponto testa só a linha no bucket da coluna.
- [x] Sonda de footprint (`--dep-graph`) na K1 (566k): footprint + fan-in direto + distribuição.
- [ ] Manutenção incremental em `Sheet.SetCell`/`Remove` — MOVIDA para a Fase 3 (é onde as edições acontecem).

### Verification Plan
- `ReverseDependencyGraphTests` (7 casos) valida `GetAllDependents` contra **oráculo por força bruta** (fecho
  transitivo re-derivado): cadeia, retângulo, coluna-inteira (linha vazia inclusa), cross-sheet, diamante,
  volátil, + comparação num workbook randomizado (40 edições). `--treenode-filter
  "/*/*/ReverseDependencyGraphTests/*"` → **7/7 ✓** (mantidos após o bucketizing).

### Phase Summary
**Grafo + memória VALIDADOS; a enumeração transitiva revelou o achado-chave que reformula a Fase 4.**

Números na K1 (566k células): build do grafo **1.3s**; 1.1M arestas de célula; **816k deps de range** (uma por
fórmula-range — a mitigação funciona: expandir por célula daria bilhões), das quais **408k são coluna-inteira**;
footprint **~86MB estimado / 145MB real**. → **Memória é viável**, não explode.

**Fan-in por edição pontual — bimodal (o achado):**
- **Mediana = 3 dependentes diretos.** A EDIÇÃO TÍPICA afeta quase nada → o CalculateDirty tem potencial de
  ganho de ~2-3 ordens de grandeza sobre o baseline de 1.2s/edição. A tese se sustenta para o caso comum.
- **Cauda pesada real:** o MAIOR bucket de coluna tem **490.705** entradas. Ou seja, existe uma coluna "quente"
  cujas ~490k fórmulas-range a referenciam; editar uma célula lá afeta MESMO ~metade do workbook. Nesses edits
  o dirty NÃO ganha do full-recompute (o impacto É grande) — e a BFS transitiva ingênua (re-varre o bucket por
  célula da fronteira) fica **quadrática**. Custo da consulta DIRETA (limitada): mediana 0.2ms, ok.

**Re-plano da Fase 4 (decisão):** a estratégia **evict-and-pull** fica claramente preferida — ela NUNCA enumera
o fecho transitivo (evita o quadrático): só evicta as células dirty e deixa o motor pull-based memoizado
recomputar o cone, cada célula 1×. O grafo reverso serve para (a) saber o que evictar e (b) enumerar outputs —
e essa enumeração, quando precisa, deve ser um **sweep por fixpoint** sobre os buckets (O(entradas × profundidade),
limitado), não a BFS por-célula. A alternativa "scheduler topológico" perde força (mesma armadilha de enumeração).

**FECHADO:** `GetAllDependents` reescrito como **sweep fixpoint** (BFS nas arestas de célula + varredura de
range 1×/passe com índice de linhas-dirty por coluna via `SortedSet.GetViewBetween`) — não-quadrático, oráculo
7/7 mantido, 110-127ms/edição na K1.

**Impacto transitivo re-medido nos DOIS designs de fixture (o achado que decide o valor do spike):**
| edição | fixture XLOOKUP-coluna (acoplado) | fixture ref-direta (desacoplado) |
| --- | --- | --- |
| **input (Input!B)** — o caso real | mediana **542k (96%)** | **mediana 5 células**, p90 496k, média 141k (bimodal) |
| célula interna qualquer | mediana 231k | mediana 208k |

O XLOOKUP sobre `$B$1:$B$N` era o vilão: acoplava todo input a tudo. Com `Input!C{r}=B{r}` direto, a **edição
típica de um input afeta ~5 células** → dirty seria ~5 en 566k, **~5 ordens de grandeza** mais rápido que o full
de 1.2s. **A premissa se sustenta fortemente para o caso input→output.** MAS a distribuição é **bimodal**:
~10-15% dos inputs alimentam a coluna quente REAL do K1 (490k fórmulas-range) e cascateiam para ~metade do
workbook — para esses o dirty não ganha (o impacto É grande). Editar uma célula INTERNA aleatória também cai
na cauda (mediana 208k), mas esse não é o caso de uso.

**Conclusão da Fase 2:** grafo + memória viáveis (~86-145MB, build 0.8s), fixpoint eficiente e correto, e o
motor dirty **vale a pena** — a maioria das edições de input vai de 1.2s → microssegundos. A Fase 4 fica
**evict-and-pull** (uma estratégia). A manutenção incremental do grafo vai para a Fase 3.

## Phase 3: Marcação dirty + invalidação seletiva
Status: Complete

- [x] **Decisão de armazenamento:** NÃO há storage de dirty-flag separado. "Dirty" = "evictado do cache".
      `SheetValueStore.EvictDense(handle,col,row)` dropa o bit de PRESENÇA da célula (+ marca tainted) —
      reusa o bitmap de presença que já existe. `Workbook.EvictDense(sheet,col,row)` é o ponto de entrada.
- [x] `DirtyEngine.CalculateDirty(edited)` propaga transitivamente via o grafo reverso (editadas ∪ cone
      transitivo ∪ sempre-dirty) e evicta esse conjunto.
- [x] API (interna, aditiva): `CalculateDirty` (= MarkDirty + evict, retorna o cone dirty), `PullAll` (força
      o recompute), `GetAffectedOutputs` (os *sinks* do sub-DAG — "as células no topo").
- [x] `InvalidateCache`/`Recalculate` intocados; `EvictDense` é aditivo.

### Verification Plan
- `DirtyMarkingTests` (5): cone dirty == editadas+dependentes; **evict-and-pull propaga a edição SEM
  `InvalidateCache`** (A3 vai de 3→12 após editar A1); eviction é seletiva (célula não-afetada fica HIT);
  outputs == sinks; sempre-dirty sempre incluído. **5/5 ✓**.

### Phase Summary
**Feito.** O evict-and-pull nasceu: `EvictDense` (dropa 1 célula do cache) + `DirtyEngine` (computa o cone e
evicta). Nenhum storage novo de dirty — "evictado" É "dirty". `GetAffectedOutputs` dá os sinks (o conjunto
input→output). Corretude do propagate provada (A1 editado → A3 fresco sem InvalidateCache).

## Phase 4: Recompute evict-and-pull + PORTÃO DE CORRETUDE
Status: Complete

- [x] **4a — evict-and-pull**: `DirtyEngine.CalculateDirty` evicta o cone dirty; o pull-based memoizado (via
      `RunWithLargeStack`/`PullAll`) recomputa só o cone, cada célula 1×. Enumeração **output-sensitive**
      (`GetAllDependentsBounded`, BFS por-célula via buckets) com **cap de 50k**: cone pequeno → evict-and-pull;
      cone grande (coluna quente) → `null` ⇒ fallback a `InvalidateCache`+ComputeAll.
- [x] **4b — scheduler topológico: DESCARTADO** (decisão da Fase 2 — a estratégia evict-and-pull evita a
      enumeração/quadratura; o scheduler cairia na mesma armadilha). Uma estratégia, não duas.
- [x] **PORTÃO DE CORRETUDE (fuzz):** `DirtyRecomputeEquivalenceTests` — 60 lotes de edições aleatórias num
      workbook com células+ranges+coluna-inteira+cadeias: todas as células após `CalculateDirty` são
      **bit-idênticas** às após `InvalidateCache()+ComputeAll()`. **PASSOU.**

### Verification Plan
- `DirtyRecomputeEquivalenceTests` — **1/1 ✓** (60 iterações internas). Suíte completa **1031/1031 ✓**.
- Harness `--dirty-recompute` na K1 mede o speedup end-to-end (abaixo).

### Phase Summary
**Feito, e o payoff está medido.** Speedup end-to-end na K1 editando INPUTS reais (Input!B):
- Build do motor (grafo): **1.1s, uma vez** (amortizado sobre milhares de edições).
- Baseline (InvalidateCache+ComputeAll): **1264 ms/edição**.
- **31/40 edições → cone pequeno (mediana 4 células recomputadas) → mediana 1.05 ms ⇒ ~1200× mais rápido.**
- **9/40 edições → cone grande (cauda da coluna quente) → fallback a full-recompute** (~1264 ms; sem ganho,
  mas correto — ali o impacto É ~metade do workbook).
- **Corretude bit-idêntica provada** pelo portão fuzzed.

O híbrido (bounded + fallback) casa exatamente com a distribuição bimodal: a maioria dos inputs ganha ~1200×,
a minoria "hub" cai para o full de forma correta.

## Phase 5: Medição em escala + go/no-go
Status: Not started

- [ ] Rodar o harness no **template real de 500k do dono** (fallback: sintético de 500k). Medir, para 4a e 4b:
      (1) speedup de recompute de edição pontual (dirty vs `InvalidateCache`+releitura); (2) tempo de build do
      grafo; (3) memória do grafo + índice de contenção + dirty set; (4) custo da consulta de contenção por edição.
- [ ] Demo **input→output**: marcar um conjunto de células de input como dirty, extrair `GetAffectedOutputs()`,
      e confirmar que popular "outra planilha/PDF" a partir daí evita varrer o sheet inteiro — medir a diferença
      vs. a extração atual (varre todas as células).
- [ ] Escrever o **relatório**: tabela de números, o ponto de equilíbrio (a partir de que fan-out/tamanho de
      edição o dirty ganha), o custo de memória do fan-in de coluna-inteira, e a recomendação go/no-go (e qual
      das duas estratégias de recompute).

### Verification Plan
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --dirty-graph --scale <fixture>`
  produz a tabela completa de métricas para 4a e 4b sem exceção. O relatório é anexado ao Phase Summary.

### Phase Summary
_(escrever quando a fase completar)_

## Phase 6: Análise escrita — separação AST/dados (stateless)
Status: Not started

- [ ] Medir o custo de re-parse a 500k: parsear a AST uma vez vs. N vezes (proxy para processar N documentos
      contra 1 template), em tempo e memória.
- [ ] Tabela classificando CADA peça de estado como **Definition** (compartilhável/imutável entre runs) vs
      **Memory** (per-run): AST das células, nomes definidos, custom functions, value store denso, cache de
      range, índice estrutural, `_epochNow`/RNG/taint volátil, dirty flags, arestas estáticas vs dinâmicas do
      grafo reverso.
- [ ] Registrar as fronteiras de refactor que o grafo reverso revelou (arestas estáticas = Definition; dirty/
      taint/arestas dinâmicas = Memory) e os problemas: entrelaçamento `CellStore`↔`Sheet`, contaminação
      cruzada entre documentos (`NOW()`/RNG/cache de range), destino de nomes/custom functions.
- [ ] Recomendação + esboço de um spike seguinte para a separação em código.

### Verification Plan
- Revisão do dono: a seção de análise responde "vale separar AST dos dados?" com números de re-parse e a
  tabela Definition/Memory completa, sem lacunas. (Verificação humana; não há comando autônomo.)

### Phase Summary
_(escrever quando a fase completar)_

## Final Recap
_(escrever quando todas as fases completarem)_

## Deployment Plan
_(escrever quando todas as fases completarem — sendo um spike, "deploy" = decisão go/no-go, promoção do
código validado para uma feature branch com escopo de produção, ou descarte do spike com o relatório
arquivado em `plans/`)_
