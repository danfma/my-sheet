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
`feat!(escopo):` não parseia no versionize). **Integração à main ADIADA de propósito**: o fix escalar
OR/AND (K1) deve entrar na main ANTES da quebra 3.0 para viabilizar release 2.9.x; a F2 forka desta
branch.

## Phase 2: Índice write-maintained
Status: Not started
- [ ] Índice vitalício: lazy build 1×/vida; manutenção incremental em `SetCell`/`Remove` (adaptativa,
      item 3); `InvalidateCache` deixa de derrubá-lo; deserialize → flag rebuild.
- [ ] Remover a admissão estrutural da Fase 5 (`_structuralIndexSeen`, modo ForceNaive/etc. viram
      obsoletos — limpar harness/testes correspondentes com justificativa).
- [ ] Testes: escrita em ordem (append O(1)); fora de ordem (dirty→resort na leitura); Remove; overwrite;
      pós-Load rebuild 1×; **índice sobrevive a InvalidateCache** (contador de builds); valores continuam
      por época (não regride voláteis/snapshot).
### Verification Plan
- Suítes verdes; equivalência intacta; contadores de build provam vitalício.

## Phase 3: Validação + docs + release
Status: Not started
- [ ] Micro-benchmark alvo (o do relatório K1): `COUNTIF(A:A)` com coluna A=200, sheet 10k→500k, POR
      ÉPOCA (InvalidateCache entre leituras) — esperado **~flat** (sem rebuild por época; comparar com a
      tabela do usuário: Aspose ~0,4ms). Registrar.
- [ ] Custo de escrita: micro de `SetCell` em massa (Fill-like, 500k sets em ordem e fora de ordem) —
      gate: overhead ≤ ~5% vs 2.9.0.
- [ ] Harness whole-column-scale completo + SheetBenchmarks vs main (alocação idêntica no hot path).
- [ ] Docs: `performance.md` (Camada 1 agora write-maintained), `workbook-and-expressions.md` (Remove/
      Cells read-only), guia de migração final — skill `code-documentation-doc-generate`. NÃO tocar
      `docs/pt-BR/` (refresh no deploy).
- [ ] Release **3.0.0** (`feat!` + BREAKING CHANGE) — dispatch em chamada separada pós-verificação;
      depois refresh pt-BR via Sonnet.
### Verification Plan
- Tabela flat do micro-alvo; gate de escrita ≤5%; suítes/fixture verdes; versionize propõe major.

## Final Recap
_(ao concluir)_
## Deployment Plan
_(verificação minha → merge → push → pré-condição merge-base em chamada separada → release 3.0.0 →
`git pull` → refresh pt-BR. Guia de migração publicado junto.)_
