# Audit Fixes — Correções e Otimizações da Auditoria de 2026-07-10

Executar os achados da auditoria completa (4 relatórios, sintetizados na sessão de 2026-07-10) em fases
incrementais, **cada fase entregue como release parcial** via `gh workflow run release.yml`. Regra do
usuário: sem breaking changes; qualquer candidato a breaking fica para o final e é avaliado com ele.
Execução delegada a subagentes (modelo apropriado por tarefa); o orquestrador verifica e commita.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to
`Complete`, write its **Phase Summary** and run its **Verification Plan** (recording results) before
moving on. When all phases are done, fill **Final Recap** and **Deployment Plan**.

Regras do repo: conventional commits SEM referência ao assistente; hooks husky (pre-commit: csharpier
check + build; pre-push: + as duas suítes); `dotnet csharpier format .` antes de commitar. Wire format
MemoryPack é contrato de byte-identidade: NUNCA tocar membros serializados, ordem de declaração, ou o
CellStoreFormatter; campos novos em tipos serializados exigem `[MemoryPackIgnore]`.

Releases: versionize bumpa por conventional commits (fix/perf → patch; feat → minor). Cada fase fecha
com push + `gh workflow run release.yml --ref main` + verificação do run.

## Phase 1: Correções de corretude → release patch
Status: Complete

- [x] **F1 (DefineName × dirty graph)**: `DefineName` não bumpa versão estrutural → `RecalculationEngine`
  serve valores stale após redefinição de nome (Workbook.cs:1002; IsStale em RecalculationEngine.cs:178
  só olha sheets). Fix: versão de nomes no Workbook (`[MemoryPackIgnore]`), bump em `DefineName`,
  `IsStale` compara no snapshot. Teste: redefinir nome usado por fórmula → engine recomputa via caminho
  parcial (sem InvalidateCache manual).
- [x] **F2 (merge × @r implícito)**: `ExcelMerge` assume `@r` presente (`int.Parse(...!)` em
  ExcelMerge.cs:253,344) — arquivo que o loader lê OK (posição implícita) quebra o merge. Fix: replicar
  o tracking currentRow/nextColumn do `WorksheetStreamLoader`. Teste: fixture bruto sem `@r`
  (padrão de StreamingLoadEdgeTests) passa por load E merge.
- [x] **F3 (guard de profundidade)**: Parser (Parser.cs:378-413) e FormulaWriter (FormulaWriter.cs:48)
  recursam sem limite → StackOverflow não-capturável em fórmula patológica. Fix: contador de
  profundidade (limite generoso, ex. 256) → `ParseException` no parse; guard equivalente no writer.
  Testes: profundidade N-1 passa, N+1 lança ParseException; sem SO.

### Verification Plan
- Build 0 warnings; 1.051 core + 46 Excel tests verdes + testes novos das 3 correções
- `--excel-memory` sem regressão (números ~iguais aos de 2026-07-10)
- Push + release workflow verde + tag nova no NuGet

### Phase Summary
Três fixes delegados a agentes Sonnet (1 por vez, árvore compartilhada), revisados e commitados pelo
orquestrador:
- **F1** (db9c647): `_namesVersion` runtime-only ([MemoryPackIgnore]) bumpado nos 2 overloads de
  DefineName; IsStale do engine compara. Prova contrafactual: teste falha sem o fix (StructureRebuilt
  = false → valor stale). Mutação direta do dicionário DefinedNames documentada como não-rastreada.
- **F2** (bcd3406): tracking currentRow/nextColumn no merge espelhando o loader; verbatim das
  existentes PROVADO seguro (posição implícita = anterior+1; não há inteiro entre as duas → inserção
  nossa nunca desloca implícita). 2 testes novos com fixture bruto.
- **F3** (da54344): MaxDepth=256 nos DOIS ciclos de recursão do Parser (ParseExpression + 
  ParseQualifiedReference — o segundo descoberto pelo agente: ranges cross-sheet encadeados recorrem
  sem passar pelo hub) → ParseException capturável; FormulaWriter com depth por parâmetro →
  InvalidOperationException. 7 testes novos.

Suítes: 1.059 core + 48 Excel, 0 falhas. Zero toque em wire format.

## Phase 2: Superlineares (RANK + regex cache) → release patch
Status: Complete

- [x] **P1 (RANK O(n²))**: RANK.EQ/RANK.AVG re-escaneiam o ref por célula (OrderStatistics.cs:142-160).
  Fix: usar o sorted-view do `RangeSnapshot` (SortedNumericValues, lower/upper bound binário) para
  derivar `outranking`/`equal` em O(log n) por célula quando snapshot admitido; fallback linear mantido.
  CUIDADO: paridade de tie-break de RANK.AVG e ordem asc/desc — congelar testes de paridade ANTES.
- [x] **P2 (regex recompilado)**: `ExcelRegex.Create` (RegexFunctions.cs:198), wildcard de SEARCH
  (TextManipulation.cs:165) e `Criteria.WildcardMatch` estático (Criteria.cs:148) compilam Regex por
  avaliação. Fix: cache estático bounded `ConcurrentDictionary<(string,RegexOptions),Regex>` (cap ~256
  entradas, eviction simples), mantendo o timeout. Benchmark próprio simples antes/depois (coluna
  sintética de REGEXTEST) para provar o ganho.

### Verification Plan
- Testes de paridade RANK congelados antes passam depois; suítes completas verdes
- Micro-bench demonstra ganho (RANK coluna 10k; REGEX coluna 10k)
- Push + release verde

### Phase Summary
- **P1** (3f1c133): RANK.EQ/AVG via 2 buscas binárias no sorted-view do snapshot (novo
  `RangeSnapshot.NumericRankCounts` + LowerBound/UpperBound no idioma dos vizinhos); fórmulas
  algebraicamente idênticas ao scan linear (EQ = outranking+1; AVG = outranking+(equal+1)/2);
  propagação de erro com paridade (first-error-in-scan-order); ranges pequenos mantêm o linear.
  Paridade congelada ANTES (8 testes: ties, asc/desc, #N/A, não-numéricos, travessia do threshold de
  admissão). Bench contrafactual: coluna 5k → 740ms (revertido) vs ~25ms (fix) ≈ 20-40x.
- **P2** (fec1742): `RegexCache` bounded (256, clear-all eviction documentada) usado por ExcelRegex,
  SEARCH wildcard e Criteria (instância + WildcardMatch estático); timeout 1s preservado. BÔNUS de
  segurança: Criteria/WildcardMatch NÃO tinham timeout (ReDoS real com wildcards encadeados) — agora
  fail-safe como no-match. 50k criações mesmo pattern: 110ms → 13ms (8,2x). 6 testes novos.

Suítes: 1.073 core + 48 Excel, 0 falhas. Zero wire format.

## Phase 3: Quick wins de alocação → release patch
Status: Complete

- [x] **A1**: `RangeReference` corners — parse único no-alloc via `TryGetColumnRow` (helper TryGetBounds);
  consertar RowCount/ColumnCount/TopRow/LeftColumn (RangeReference.cs:86-113) e os multiplicadores
  (ArgumentFlattening.cs:75, CriteriaScan.cs:57, ArrayEvaluation.cs:454). Campos cacheados = `[MemoryPackIgnore]`.
- [x] **A2**: VLOOKUP fallback — hoistar bounds+handle do laço (VLookup.cs:87,119,129); `keyColumn` lazy (:64).
- [x] **A3**: `CellId.Parse` com `AsSpan` (CellId.cs:19); `CellId.Format` sem concat em loop.
- [x] **A4**: Solvers financeiros — hoistar cronograma de cupom/year-fractions invariantes dos lambdas
  (BondMath.cs:338,982,1250) e reduzir walks duplicados em Price (:161-207,281-314).
- [x] **A5**: `EvaluationContext.WithName` — encadeamento sem cópia O(k²) (EvaluationContext.cs:36-45).
- [x] **A6**: locks separados para `EpochNow` vs `NextRandom` (Workbook.cs:333,349).
- [x] **A7**: NETWORKDAYS — serial incremental no loop diário, sem ToOADate por passo (WorkdayFunctions.cs:141,330).

### Verification Plan
- Suítes completas verdes; `--k1-endtoend` e `--excel-memory` sem regressão (compute deve melhorar ou empatar)
- Push + release verde

### Phase Summary
Três lotes sequenciais (agentes Sonnet), todos com micro-medição contrafactual e paridade provada:
- **A1+A2** (f70fd14): `RangeBounds` struct (corners parseados 1× no-alloc, semântica de range invertido
  preservada) + hoist de bounds/handle nos fallbacks de VLOOKUP/HLOOKUP + keyColumn lazy sob o threshold
  de admissão. VLOOKUP fallback: 22.096 → 0,1 B/aval; SUM range: 288 → 120 B/aval.
- **A3+A4+A7** (80dba42): CellId com AsSpan/stackalloc; `CouponSchedule`/`OddFPriceContext`/year-fractions
  hoistados dos lambdas dos solvers SEM reordenar aritmética (checksums idênticos). YIELD×2k: ~126 → ~40ms;
  NETWORKDAYS×2k: ~23 → ~18ms (serial incremental; weekday continua via DateTime — decisão conservadora
  documentada).
- **A5+A6** (commit desta entrada): LET com cadeia imutável de escopos (1 nó/binding, shadowing grátis,
  comparador OrdinalIgnoreCase preservado) — 20 bindings ×10k: 77,6ms/16,8KB → 18,3ms/1,1KB; locks de
  clock e RNG separados (prova: nenhum ponto reseta ambos atomicamente; RNG nunca é re-semeado por época).

Suítes: 1.074 core + 48 Excel, 0 falhas. Zero wire format.

## Phase 4: Streaming uniforme + export dispatch → release patch
Status: Not started

- [ ] **S1**: SUMIF/AVERAGEIF/AND/OR/XOR para cursores streaming (padrão de COUNTIF/CriteriaScan)
  (SumIf.cs:18, ConditionalAggregates.cs:24, LogicalReduction.cs:40,95). Paridade semântica coberta
  pelos 1.051 testes existentes (famílias bem testadas).
- [ ] **S2**: `FormulaWriter.Call` — dispatch por `FrozenDictionary<Type, ...>` em vez do switch de 304
  braços (FormulaWriter.cs:282-598). Sem mudança de output (FormulaWriterTests congelam).
- [ ] **S3**: wildcard estático (`Criteria.WildcardMatch`) compilado 1× antes do scan (LookupMatching.cs:53).

### Verification Plan
- Suítes completas verdes (FormulaWriterTests = oráculo do S2)
- Export em modo Formulas: benchmark antes/depois no fixture sintético
- Push + release verde

### Phase Summary
_(write when phase completes)_

## Phase 5: Robustez do interop → release minor
Status: Not started

- [ ] **R1**: `ExcelLoadOptions` (aditivo: `Load(path/stream, options?)`) com coletor de warnings
  (`IReadOnlyList<LoadWarning>` exposto via options) — defined names inválidos deixam de sumir em silêncio
  (ExcelFile.cs:104).
- [ ] **R2**: merge — investigar escrita direta sem `MemoryStream` intermediário (ExcelMerge.cs:143-154):
  part temporário ou arquivo temp; SE inviável no SDK, documentar a limitação na doc da classe e
  registrar decisão aqui.
- [ ] **R3**: comentário do next-free-tag da union (Expression.cs:15 → 319) + doc do invariante das 3
  estruturas de estrutura (structural index / dirty buckets / dense pages) num comment central.

### Verification Plan
- Suítes verdes; teste novo de warnings; `--excel-memory` merge sem regressão de pico
- Push + release verde

### Phase Summary
_(write when phase completes)_

## Phase 6: FunctionRegistry (AVALIAR BREAKING com o usuário antes de executar)
Status: Not started — GATED em decisão do usuário

- [ ] Extrair catálogo de funções do Parser para `FunctionRegistry` (mantendo FrozenDictionary interno)
- [ ] `FormulaWriter` deriva nomes do registry (mata a 4ª cópia manual)
- [ ] Custom functions: validação de aridade opcional em `RegisterFunction` (aditivo)
- [ ] DECISÃO COM O USUÁRIO: permitir override de built-ins? (mudança de comportamento observável)
- [ ] Extrações do Workbook.cs (serialização, época volátil, admissão de range cache, Sheet) — só as
  `[MemoryPackIgnore]`-safe; membros serializados NÃO se movem

### Verification Plan
- Suítes completas + MemoryPackCompatibilityTests (wire intacto) + FormulaWriterTests
- Push + release (minor se houver feat)

### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete: já embutido — cada fase publica sua release)_
