# Funções voláteis (F1) — TODAY/NOW/RAND/RANDBETWEEN com época + Recalculate()

Adicionar as 4 funções voláteis puras ao engine, sem quebrar a memoização: um modelo de **época** onde
voláteis (e seus dependentes, por contágio) são cacheados por época e regenerados no avanço da época;
relógio/RNG vêm de fontes injetáveis no `Workbook` (testáveis). Release aditivo **2.5.0** (não-breaking).

## Decisões travadas (com o usuário, 2026-07-02)
- **Cache por época**, não "sempre gerar": obrigatório por correção — `RAND()` regenerado a cada leitura
  daria valores diferentes na mesma célula (incoerente). Voláteis computam 1× por época e ficam cacheados.
- **`Recalculate()`** (novo) = o "touch global": limpa só as células voláteis marcadas + reseta relógio/RNG;
  NÃO recomputa (lazy on-read). `InvalidateCache()` continua limpando tudo (novos inputs) e também reseta a
  época. No loop de servidor (setar inputs → InvalidateCache → ler), voláteis já regeneram por batch.
- **`touch` por célula: BARRADO** — volatilidade é contagiosa; sem grafo reverso de dependências, tocar só
  `A1=NOW()` deixaria `B1=A1+1` cacheado e velho. Feature futura, condicionada ao grafo de deps (mesma infra
  do recálculo incremental). Documentar.
- **Escopo = TODAY, NOW, RAND, RANDBETWEEN.** `OFFSET` **NÃO** vira volátil (no Excel é volátil por
  segurança de auto-recálculo; aqui a invalidação é explícita → marcá-lo contaminaria meia planilha por
  nada — divergência consciente, documentada). `INDIRECT` fora (o difícil é resolver referência a partir de
  texto — feature própria; fase futura).

## Micro-decisões (default meu; usuário pode vetar na revisão)
- **Amostragem lazy do relógio**: `_epochNow` é amostrado na PRIMEIRA leitura de volátil da época (não na
  chamada do `Recalculate()`). Uniforme (sem caso especial da época inicial), não amostra se ninguém lê
  volátil, e `NOW()` reflete a hora em que o valor foi de fato produzido. (Alternativa eager: amostrar no
  `Recalculate()` → `NOW()` = timestamp do recálculo; trocável em 1 linha.)
- **Hora local via TimeProvider**: `NOW()`/`TODAY()` usam `TimeProvider.GetLocalNow()` (Excel usa hora
  local); o fake de teste controla relógio E fuso.
- **RNG reproduzível**: `Workbook.RandomSeed` (int?, default null). Uma instância `Random` persistente
  criada a partir do seed (ou time-based se null) e avançada AO LONGO das épocas (não re-semeada por época)
  — assim épocas dão valores diferentes naturalmente, e com seed fixo a execução inteira é reproduzível.
  Dentro da época, cada célula `RAND` cacheia seu saque (re-leitura estável).
- **Versão 2.5.0** (minor): tudo aditivo — `IsVolatile` virtual default false, novos nós, `Recalculate()`/
  `TimeProvider`/`RandomSeed` novos, tags MemoryPackUnion append (312–315). Compat binária preservada.

## For Future Agents
Marque `- [x]`; ao fechar fase: Status `Complete` + Phase Summary + rode a Verification. TDD rigoroso
(RED antes de GREEN). **Verificação SEMPRE com `dotnet build ... --no-incremental`** (builds incrementais
mascaram warnings de analyzer — lição registrada em tasks/lessons.md). Não tocar na fixture
`tests/Danfma.MySheet.Tests/Fixtures/workbook-pre-namespaces.msgpack.bin` nem em `MemoryPackCompatibilityTests`.
Suítes hoje: core **676**, Excel **20**; build **0 warnings**. Commits: inglês, curto + corpo, semantic
(`feat(volatile): ...`), SEM atribuição a IA. NÃO fazer push (gates externos são do usuário).

Contexto de código (layout pós-Fase R): nós de função vivem em `Danfma.MySheet.Expressions.{Categoria}`;
`ComputedValue`/`Error` na raiz `Danfma.MySheet`. `Workbook` já tem: `_cache`
(`ConcurrentDictionary<(string,string), ComputedValue>`), `_evaluating` (`[ThreadStatic] HashSet` — detector
de ciclos, o PADRÃO a espelhar para a marca de volatilidade), `GetCellValue` (compute-on-read + memoização),
`InvalidateCache()`. `EvaluationContext` carrega `Workbook`. Categorias do Excel: TODAY/NOW → `Expressions.Dates`;
RAND/RANDBETWEEN → `Expressions.Mathematics`.

---

## Phase 1: Infra de época + TODAY/NOW (caminho do relógio)
Status: Complete

- [x] `Expression`: `public virtual bool IsVolatile => false` (introspecção; default false).
- [x] `Workbook`: `public TimeProvider TimeProvider { get; set; } = TimeProvider.System;` e o estado de
      época: `_epochNow` (nullable), `[ThreadStatic] static bool _volatileTouched;`, e um conjunto
      concorrente `_volatileTainted` (`ConcurrentDictionary<(string,string), byte>` como set — escrito
      durante avaliação concorrente).
- [x] Método interno `MarkVolatileTouched()` (seta a flag thread-local) e `EpochNow()` (amostra `_epochNow`
      lazy via `TimeProvider.GetLocalNow()`, com publicação thread-safe — `Interlocked`/lock; retorna serial
      OADate). Os nós voláteis chamam ambos no `Evaluate`.
- [x] `GetCellValue`: em volta do `Evaluate` da célula, salvar/zerar/propagar `_volatileTouched` (padrão do
      `_evaluating`): se a avaliação da célula tocou volátil → adicionar a chave a `_volatileTainted` (mas
      AINDA cachear normalmente — cache por época); propagar a flag ao chamador (contágio transitivo).
- [x] `Recalculate()` (público): remover do `_cache` só as chaves em `_volatileTainted`, limpar
      `_volatileTainted`, resetar `_epochNow = null`. NÃO recomputa. `InvalidateCache()`: além de limpar tudo,
      limpar `_volatileTainted` e resetar `_epochNow`.
- [x] `Today` (`Expressions/Dates/`): 0 args, `IsVolatile => true`, `Evaluate` = parte-data do serial de
      `EpochNow()` (via `Math.Floor`), chamando `MarkVolatileTouched`. `Now`: 0 args, serial completo.
- [x] `Parser.Functions` (`TODAY` 0/0, `NOW` 0/0) + `FormulaWriter.Call` + corpus exaustivo. Tags MemoryPack
      append (verificar maior atual = 311 → 312/313).
- [x] Teste-double `FixedTimeProvider : TimeProvider` (override `GetUtcNow` + `LocalTimeZone`), sem pacote novo.

### Verification Plan
- `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` → 0 warnings.
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → verde, incl.
  os novos: (a) `NOW()` com provider fixo → serial esperado; (b) avança o provider SEM `Recalculate` → mesmo
  valor (coerência de época); (c) avança + `Recalculate()` → novo valor; (d) `A1=NOW()`, `B1=A1` →
  após `Recalculate`, B1 também muda (contágio); (e) duas células `NOW()` na mesma época concordam;
  (f) célula NÃO-volátil (custom function com contador de efeito colateral) NÃO recomputa após `Recalculate`
  — contador estável — provando que Recalculate preserva o cache estável; (g) `MemoryPackCompatibilityTests`
  verde (IsVolatile é comportamento, não serializado).
- Suíte Excel (20) intacta.

### Phase Summary
Concluída em 2026-07-02 (branch `feature/volatile-functions`, TDD RED→GREEN: os 10 testes de
`VolatileClockTests` entraram primeiro — RED por símbolos ausentes: `Workbook.TimeProvider`,
`Workbook.Recalculate`, `Now`/`Today`, `Expression.IsVolatile` — e viraram verdes ao registrar a infra +
os nós). **2 funções novas** (NOW, TODAY; Parser 300 → 302). Tags MemoryPackUnion **312 (Now)/313 (Today)**
(append-only; próximo livre = 314; comentário do cabeçalho atualizado). Núcleo `Danfma.MySheet.Expressions`:
`Expression.IsVolatile` (virtual, default false, `[MemoryPackIgnore]` — comportamento, não estado; overridden
`true` em Now/Today). `Workbook`: campo `[ThreadStatic] static bool _volatileTouched` (a flag de contágio, no
mesmo padrão save/reset/propagate do `_evaluating`), set concorrente `_volatileTainted`
(`ConcurrentDictionary<(string,string),byte>`), `_epochNow` (`double?`), e um `TimeProvider` injetável
(`[MemoryPackIgnore]`). `GetCellValue` agora, em volta do `Evaluate` da célula: salva o flag do chamador,
zera, avalia, e se a célula tocou volátil (direta ou transitivamente) grava a chave em `_volatileTainted`
(mas AINDA cacheia — cache por época) e restaura o flag como `outer || cellTouched` (o `outer` é o que carrega
o contágio entre irmãos). `Recalculate()` remove do `_cache` só as chaves marcadas + limpa o set + reseta
`_epochNow`; `InvalidateCache()` limpa tudo + o set + reseta `_epochNow`. Nós NOW/TODAY em
`Expressions/Dates/VolatileClock.cs`: leem `TimeProvider.GetLocalNow().DateTime` (hora LOCAL, como o Excel),
via `Math.Floor` no TODAY; ambos chamam `MarkVolatileTouched()`. Amostragem LAZY do relógio (default do
plano): `EpochNow()` amostra `_epochNow` na PRIMEIRA leitura de volátil da época.

**Thread-safety (decisão além do plano):** o plano dizia "Interlocked/lock para amostrar `_epochNow` uma
única vez". Resolvi com um único `lock (VolatileLock)` cobrindo a amostragem de `_epochNow` (`_epochNow ??= …`
dentro do lock — leitura/escrita atômica da amostra) e o reset em `Recalculate`/`InvalidateCache`. Evitei o
fast-path lock-free (que teria uma torn read do `double?`), preferindo o lock sempre — não é caminho quente
(1× por célula volátil por época) e é provadamente correto. O `_volatileTainted` e o `VolatileLock` são
criados via `Interlocked.CompareExchange` (lazy race-free), porque **MemoryPack ignora os field initializers
na desserialização** (mesmo motivo do null-handling do `DefinedNames`) — sem isso o lock viria null após um
`Load`. Idem `TimeProvider`: getter lazy `_timeProvider ?? TimeProvider.System` (robusto ao `Load`, sem hook).
Escrita no set via indexer do `ConcurrentDictionary` (thread-safe).

**Compat binária:** `TimeProvider` e (fase 2) `RandomSeed` são `[MemoryPackIgnore]` — o schema
serializado do `Workbook` NÃO mudou (`DefinedNames` continua o último membro serializado). A fixture
`workbook-pre-namespaces.msgpack.bin` abre e reavalia idêntica (guarda verde). Um round-trip novo
(`VolatileFormula_RoundTripsThroughMemoryPack`) prova que os tags 312/313 sobrevivem ao Save/Load.

Suíte core: 676 → **686 verdes** (10 novos em `VolatileClockTests`: NOW hora-local, TODAY=floor, coerência
intra-época sem/​com `Recalculate`, contágio `B1=A1+1`, duas células NOW concordam, célula não-volátil com
contador de efeito colateral estável através do `Recalculate`, `InvalidateCache` re-amostra, introspecção
`IsVolatile`, round-trip MemoryPack); corpus do FormulaWriter +2 (`NOW()`/`TODAY()`); suíte Excel intacta
(20); build da solução `--no-incremental` **0 warnings**.

---

## Phase 2: RAND/RANDBETWEEN (caminho do RNG)
Status: Not started

- [ ] `Workbook`: `public int? RandomSeed { get; set; }` + `Random` persistente (`_random`, criado lazy de
      `RandomSeed ?? time-based`, NÃO re-semeado por época), acessado via método interno `NextRandom()` que
      também chama `MarkVolatileTouched`. `InvalidateCache()`/`Recalculate()` NÃO resetam o `_random` (a
      sequência continua → épocas diferentes; o cache por época garante estabilidade intra-época).
- [ ] `Rand` (`Expressions/Mathematics/`): 0 args, `IsVolatile => true`, `[0,1)` de `NextRandom()`.
- [ ] `RandBetween`: 2 args (bottom, top), inteiros; `bottom > top` → `#NUM!`; inclusivo nas duas pontas;
      trunca args não-inteiros (confirmar regra na doc MS). `IsVolatile => true`.
- [ ] `Parser.Functions` (`RAND` 0/0, `RANDBETWEEN` 2/2) + `FormulaWriter.Call` + corpus. Tags append (314/315).

### Verification Plan
- Build `--no-incremental` 0 warnings; suíte verde, incl.: (a) `RandomSeed` fixo → sequência determinística
  reproduzível; (b) `RAND()` na mesma época lido 2× → MESMO valor (cache por época); após `Recalculate` →
  valor diferente; (c) duas células `RAND` na mesma época → valores DIFERENTES (saques distintos);
  (d) `RANDBETWEEN(1,6)` sempre em [1,6]; `RANDBETWEEN(6,1)` → `#NUM!`; (e) fixture binária verde.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Documentação + finalização
Status: Not started

**Usar a skill `code-documentation-doc-generate`** para a atualização da documentação (extrair fatos do
código → gerar/atualizar → validar exemplos por compilação), como no trabalho inicial de docs.

- [ ] `docs/function-reference.md`: `TODAY`/`NOW` de ⬜ (deferred) para ✅ (Date and time 23/25 → 25/25);
      `RAND`/`RANDBETWEEN` ✅ (Math); contagens do topo 300 → 304; blocos de cobertura.
- [ ] `docs/workbook-and-expressions.md`: seção "Volatile functions" — o modelo de época, `Recalculate()` vs
      `InvalidateCache()`, `TimeProvider`/`RandomSeed` injetáveis, coerência intra-época, e o limite do
      `touch` por célula (precisa de grafo de dependências). Nota sobre `OFFSET` não-volátil (divergência).
- [ ] `README.md`: bullet de features (voláteis + `Recalculate`/`TimeProvider`).
- [ ] NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [ ] `plans/function-coverage-roadmap.md`: F1 → Complete (nas "Fases futuras"); apontar para este plano.

### Verification Plan
- Build `--no-incremental` 0 warnings; ambas as suítes verdes; contagens do reference conferem com o Parser
  (`grep -cE '^\s+\["' Parser.cs` = 304).

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando as fases 1–3 concluírem)_

## Deployment Plan
_(quando concluir, com aval do usuário — mesmo ritual das ondas: verificação independente minha com rebuild
forçado → merge `feature/volatile-functions` → `main` → push → `gh workflow run release.yml` → **2.5.0**
lockstep dos 2 pacotes → `git pull` do bump → refresh `docs/pt-BR/` via sub-agente Sonnet.)_
