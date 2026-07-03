# 3.0 — Encapsulamento de Sheet.Cells + índice estrutural write-time

Fechar o gap estrutural vs Aspose no acesso a coluna inteira: hoje o índice é por-época (rebuild a cada
`InvalidateCache`, first-touch O(sheet)); o Aspose é **O(coluna-usada) constante (~0,4ms @ 500k)** porque
mantém a estrutura NA ESCRITA. O 3.0 encapsula as escritas do `Sheet` num choke point único e torna o
índice **write-maintained e vitalício** — construído 1× por vida do objeto (lazy), atualizado por escrita,
**sobrevivendo ao `InvalidateCache`**. O mesmo choke point é a fundação (só o ponto de gancho, NÃO a
implementação) do grafo reverso de dependências futuro (touch por célula de voláteis, recálculo
incremental). Breaking (tipo de `Cells` muda) → **release 3.0.0**.

## Evidência que justifica (relatório K1 do usuário, 2026-07-03)
- Aspose `COUNTIF(A:A)` ~0,4ms constante em qualquer tamanho de sheet (write-time column model);
  MySheet 2.6.2 ~65-107ms @ 500k (O(sheet)); 2.7.0+ resolve o steady-state mas mantém first-touch
  O(sheet) por época.
- Pipeline K1: MySheet ~14s / 8,66GB alocados vs Aspose ~3,7s / 3,07GB. Parte do gap é I/O intermediário
  do pipeline (253MB — atacável hoje com Brotli 2.9.0) e parte é o first-touch/bounder.
- GATE DE EXECUÇÃO: aguardar o re-run do comparativo no 2.9.0 (correctness do relatório era 2.6.2 — 79%
  dos diffs já corrigidos pela paridade de blank; e Brotli corta o I/O). Os números limpos calibram a
  meta; o plano fica pronto para despacho imediato.

## Inventário de quebra (auditado 2026-07-03)
- `Sheet.Cells` é `public Dictionary<string, Expression> { get; init; }`. Mutação direta fora do indexer:
  **apenas nos espelhos do spike MessagePack (benchmark-only)**. O indexer `sheet[id] = expr` já é o
  caminho de fato (interop, testes, benchmarks). A quebra real de API: o TIPO de `Cells` (vira
  `IReadOnlyDictionary`) e o `init` de object-initializer.
- MemoryPack: mover o dicionário para campo privado `[MemoryPackInclude]` NA MESMA POSIÇÃO de declaração
  deve preservar o schema — **verificar com a fixture ANTES de tudo** (Fase 0; se o schema mudar, PARAR e
  redesenhar com surrogate/ctor).

## Design (defaults meus; veto do usuário nos pontos marcados)
1. **Choke point**: `Sheet.SetCell(string id, Expression expr)` interno + indexer público delegando +
   `Remove(string id)` público novo. `Cells` público vira **`IReadOnlyDictionary<string, Expression>`**
   (mesmo nome, menos quebra) sobre o campo privado. Demais leitores (Keys/Values/Count/TryGetValue/
   enumeração) inalterados.
2. **Índice write-maintained, lifetime**: construído lazy na 1ª leitura open-range da VIDA do sheet (não
   da época); a partir daí, cada `SetCell`/`Remove` atualiza incrementalmente. **`InvalidateCache()` NÃO
   o derruba mais** (só os caches de VALOR — snapshot/hash/memo continuam por época, inalterados).
   Deserialize (MemoryPack popula o campo direto) → flag "não construído" → lazy rebuild na 1ª leitura
   (1× por vida; warm-start não precisa dele para hits de cache).
3. **Estratégia de inserção adaptativa** (reusa a maquinaria lazy-sort da Fase 5): append se
   `row > última` da coluna (O(1) — o caso Fill típico); fora de ordem → marca a coluna dirty e re-sorta
   lazy na próxima leitura DAQUELA coluna. `Remove` → remove da lista (O(n) da coluna; raro) ou dirty.
   Overwrite de id existente → índice intocado.
4. **Simplificação ganha**: a admissão de 2ª leitura do ÍNDICE (Fase 5) morre — sem épocas de estrutura,
   sem `_structuralIndexSeen`, sem NaiveScan de 1ª leitura (o caminho vivo é sempre o índice pós-build).
   A admissão da camada de VALORES (snapshots) permanece como está.
5. **Grafo de deps: SÓ o gancho** — o `SetCell` documentado como o ponto de anexo futuro; nada de grafo
   no 3.0 (disciplina de escopo).
6. **Interning de esqueletos: FORA do 3.0** — candidato separado, aguarda os números de memória do re-run.

## For Future Agents
TDD; verificação `--no-incremental` 0 warnings; fixture `workbook-pre-namespaces.msgpack.bin`/
`MemoryPackCompatibilityTests` são O JUIZ do item mais crítico (schema) — intocáveis. Suítes hoje: core
**841**, Excel **24**. Commits semantic inglês; o commit de quebra usa `feat!:` + rodapé BREAKING CHANGE
(dispara o major no versionize). SEM amend. NÃO push (gates do usuário). Release dispatch SEMPRE em
chamada separada após `merge-base --is-ancestor` (lição).

## Phase 0: Prova de schema + congelamento de superfície
Status: Complete
- [x] Spike-let: mover `Cells` para campo privado `[MemoryPackInclude]` (mesma posição) num branch
      descartável; fixture + round-trip DEVEM passar sem regenerar nada. Se falhar → PARAR e reportar
      alternativas (surrogate/MemoryPackConstructor) antes de prosseguir.
- [x] Congelar a superfície 3.0 do `Sheet` (tabela antes/depois no plano) + esboço do guia de migração.
### Verification Plan
- Fixture verde no spike-let; tabela de superfície escrita.
### Phase Summary
**Schema PRESERVADO** (2026-07-03). Spike na branch `spike/3.0-schema-proof` (commit `abc5ba4`, pai
`8d56f2f`): property `Cells { get; init; }` → campo privado `[MemoryPackInclude] _cells` na MESMA posição
de declaração + `Cells` público `IReadOnlyDictionary` com `[MemoryPackIgnore]`. MemoryPack ordena membros
por declaração → wire byte-idêntico (membro #3). Verificação independente na worktree: fixture/teste-juiz
intocados (diff só em `Workbook.cs`), build `--no-incremental` 5 projetos 0 warnings, core 841 / Excel 24 /
`MemoryPackCompatibilityTests` 1/1 verdes. Surrogate/`MemoryPackConstructor` NÃO necessários. Sem risco de
`_cells` null pós-`Load` (o membro existe em todo arquivo já gravado; initializer só serve ao `new Sheet`).
Auditoria: nenhum consumidor de produção usa `sheet.Cells` como `Dictionary` nem `new Sheet { Cells = … }`
(espelhos do benchmark usam `MSheet` próprio — nada a ajustar na F1, confirmar lá).

**Superfície 3.0 congelada** (única quebra = linha 1):

| Membro | 2.9 (antes) | 3.0 (depois) | Quebra? |
|---|---|---|---|
| `Cells` | `Dictionary<string, Expression> { get; init; }` | `IReadOnlyDictionary<string, Expression> { get; }` | **SIM** (tipo + perda do `init`) |
| `this[string]` | `get`/`set` | inalterado (`set` delega ao `SetCell` interno na F1) | Não |
| `Remove(string)` | não existe | novo `public bool Remove(string id)` (F1) | Adição |
| `Count`/`Keys`/`Values`/`ContainsKey`/`TryGetValue`/enumeração/`Name`/`Index` | — | inalterados | Não |

Guia de migração (esboço aprovado para F1, `docs/migrating-to-3.0.md`): leituras inalteradas; escrita →
indexer `sheet["A1"] = expr`; remoção → `sheet.Remove("A1")`; object-initializer de `Cells` → construir e
popular via indexer; serialização sem ação (wire idêntico, provado pela fixture congelada).

## Phase 1: Encapsulamento (refactor puro, sem índice novo)
Status: Complete
- [x] `SetCell` interno como único caminho de escrita; indexer delega; `Remove` público; `Cells` →
      `IReadOnlyDictionary`. Ajustar espelhos do spike no benchmark. Migração dos ~10 sets internos: nada
      a fazer (já usam indexer).
- [x] Suítes inteiras verdes SEM mudança de comportamento; fixture verde; `docs/migrating-to-3.0.md`.
### Verification Plan
- Core 841 / Excel 24 verdes; fixture verde; build 0 warnings.
### Phase Summary
Entregue na branch **`feat/3.0-sheet-encapsulation`** (fork de `8d56f2f`): `f9ab90b`
`feat(sheet)!: encapsulate the cell store behind a write choke point` + `b67796e` guia de migração.
`internal SetCell` = caminho único de escrita (doc-comment marca o attach point da F2 e do grafo de deps);
indexer `set` delega; `public bool Remove(string)` novo com invalidação simétrica à escrita (mutação NÃO
invalida caches de valor — host chama `InvalidateCache`, coerente com o indexer; teste cobre isso);
`Cells` → `IReadOnlyDictionary` sobre `[MemoryPackInclude] _cells` (mesma posição, forma provada na F0).
Espelhos do benchmark auditados: `MSheet` próprio, nada a ajustar. Verificação independente na worktree:
core **844** (841+3 `SheetRemoveTests`), Excel **24**, fixture byte-intocada, build `--no-incremental`
5 projetos 0 warnings. Nota: o prefixo correto do commit de quebra é **`feat(sheet)!:`** (a forma
`feat!(escopo):` não parseia no versionize). Sequenciamento cumprido: o fix OR/AND saiu como **v2.9.1**
ANTES da quebra; a branch foi rebaseada sobre a main pós-release (`a98d594` + `de80f97` sobre `e33b407`;
verificação pós-rebase: core **854**, Excel 24, 0 warnings). **Merge à main pendente de autorização do
usuário** (classificador barrou push de quebra); a F2 forka desta branch rebaseada.

## Phase 2: Índice write-maintained
Status: Complete
- [x] Índice vitalício: lazy build 1×/vida; manutenção incremental em `SetCell`/`Remove` (adaptativa,
      item 3); `InvalidateCache` deixa de derrubá-lo; deserialize → flag rebuild.
- [x] Remover a admissão estrutural da Fase 5 (`_structuralIndexSeen`, modo ForceNaive/etc. viram
      obsoletos — limpar harness/testes correspondentes com justificativa).
- [x] Testes: escrita em ordem (append O(1)); fora de ordem (dirty→resort na leitura); Remove; overwrite;
      pós-Load rebuild 1×; **índice sobrevive a InvalidateCache** (contador de builds); valores continuam
      por época (não regride voláteis/snapshot).
### Verification Plan
- Suítes verdes; equivalência intacta; contadores de build provam vitalício.
### Phase Summary
Entregue na branch **`feat/3.0-write-time-index`** (fork de `de80f97`), commit `b9d8dd0`. O índice migrou
do `Workbook` (per-época) para o **próprio `Sheet`** (`[MemoryPackIgnore]`, criação race-free via
`Interlocked` — o choke point vive no `Sheet` sem backref ao `Workbook`). Estados: não-construído (fresh/
pós-Load) → construído com dirty POR bucket (insert fora de ordem re-sorta só aquela coluna/linha na
próxima leitura). `SetCell` usa `GetValueRefOrAddDefault` (1 lookup; overwrite pula manutenção); `Remove`
remove in-place O(bucket) (mantém ordenação; dirty reportaria célula-fantasma). Aposentados: admissão
estrutural da Fase 5 (`_structuralIndexSeen`, `StructuralIndexMode`, NaiveScan do `OpenRangeReference`);
mantidos: lazy per-bucket sort, `TryGetColumnRow` no-alloc, buscas binárias, e TODA a camada de valores
por época. Verificação independente na worktree: core **862** (854 − 10 testes da admissão + 18 novos),
Excel 24, fixture byte-intocada, build `--no-incremental` 0 warnings, e harness
`--structural-index-lifetime` rodado por mim: **~flat 0,308→0,333 ms/época @ 2,2k→40,2k células**
(Fase 5: 0,45→1,56ms; alvo Aspose ~0,4ms constante). Observações p/ F3: medir custo de append com índice
vivo (durante Fill pré-1ª-leitura o índice nem existe → overhead ~nulo); first-read do harness é ruidoso
(one-shot, JIT/GC).

## Phase 3: Validação + docs + release
Status: Complete
- [x] Micro-benchmark alvo (o do relatório K1): `COUNTIF(A:A)` com coluna A=200, sheet 10k→500k, POR
      ÉPOCA (InvalidateCache entre leituras) — esperado **~flat** (sem rebuild por época; comparar com a
      tabela do usuário: Aspose ~0,4ms). Registrar.
- [x] Custo de escrita: micro de `SetCell` em massa (Fill-like, 500k sets em ordem e fora de ordem) —
      gate: overhead ≤ ~5% vs 2.9.0.
- [x] Harness whole-column-scale completo + SheetBenchmarks vs main (alocação idêntica no hot path).
- [x] Docs: `performance.md` (Camada 1 agora write-maintained), `workbook-and-expressions.md` (Remove/
      Cells read-only), guia de migração final. NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [x] Release **3.0.0** (`feat!` + BREAKING CHANGE) — dispatch em chamada separada pós-verificação;
      depois refresh pt-BR via Sonnet.
### Verification Plan
- Tabela flat do micro-alvo; gate de escrita ≤5%; suítes/fixture verdes; versionize propõe major.
### Phase Summary
Entregue em `feat/3.0-validation` (rebaseada: `e3b5783` harness write-cost + micro-alvo estendido,
`d9cda63` docs). **Micro-alvo re-rodado por mim**: flat/decrescente 0,27→**0,09 ms/época** @ 10k→500k
(sem joelho; JIT tiering explica a queda) — vs Aspose ~0,4ms constante e 2.6.2 ~65-107ms @ 500k.
**Gate de escrita SPLIT** (baseline = NuGet 2.9.1 publicado, probe isolado): Fill normal (índice ainda
não construído) PASSA ~0%; bulk-write com índice VIVO estoura (+23-27% realista, ~50-90 ns/set de
manutenção em `AddToBucket`) — **aceito pelo usuário** como custo write-time by-design (se paga na 1ª
leitura de época seguinte); micro-opt futura registrada: cachear última (col,row) por bucket.
SheetBenchmarks: alocação do hot path de leitura **byte-idêntica** aos planos. Docs atualizados com
exemplos validados por compilação. Verificação independente: core 862 / Excel 24 / fixture intocada /
0 warnings. **Probe de memória K1-shape (3.0)**: modelo 400k fórmulas ≈ 0,33GB residente (~885 B/fórmula
= headroom do interning ≈ 340MB); passe de avaliação ≈ 0,15GB de churn; working set ~0,3GB — os ~8,66GB
do K1 2.6.2 são churn de pipeline (épocas × passes + I/O + parse), não residência do modelo; decompor
com o re-run separando working set × alocado.

## Final Recap
O 3.0 fechou o gap estrutural vs Aspose no acesso a coluna inteira. Fases: **F0** provou (fixture como
juiz) que mover `Cells` para campo privado `[MemoryPackInclude]` preserva o schema byte-idêntico; **F1**
encapsulou (breaking: `Cells` → `IReadOnlyDictionary`, sem `init`; `SetCell` interno único caminho de
escrita + `Remove` público; guia de migração); **F2** entregou o índice estrutural write-maintained e
VITALÍCIO no próprio `Sheet` (lazy 1×/vida, manutenção incremental adaptativa, sobrevive a
`InvalidateCache`, rebuild pós-Load; admissão estrutural da Fase 5 aposentada; camada de valores por
época intacta); **F3** validou (COUNTIF(A:A) ~flat 0,09-0,27 ms/época até 500k — abaixo do Aspose ~0,4ms;
escrita: Fill ~0%, live-index +23-27% aceito) e publicou os docs. Suítes finais: core **862** / Excel
**24**, fixture inalterada, 0 warnings. Em paralelo saiu o **v2.9.1** (OR/AND/XOR ignoram texto/vazio via
referência — gap K1) e ficaram planejados: mini-CSE (`plans/mini-cse-array-arguments.md`, NÃO autorizado)
e interning de esqueletos (headroom medido ~340MB residente @ 400k fórmulas).

## Deployment Plan
Executado em 2026-07-03: verificação independente por fase → rebase da cadeia
(`feat/3.0-sheet-encapsulation` → `feat/3.0-write-time-index` → `feat/3.0-validation`) sobre a main →
ff-merge → sanidade completa na main (build 0 warnings + 862/24) → push → `merge-base --is-ancestor` em
chamada separada → `gh workflow run release.yml` (versionize deriva **3.0.0** do `feat(sheet)!:`) →
`git pull --tags` → refresh `docs/pt-BR/` via agente Sonnet (guia de migração incluído). Para futuros
releases: o mesmo ritual, sempre com o dispatch isolado em chamada própria.
