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
Status: Complete

- [x] **S1**: SUMIF/AVERAGEIF/AND/OR/XOR para cursores streaming (padrão de COUNTIF/CriteriaScan)
  (SumIf.cs:18, ConditionalAggregates.cs:24, LogicalReduction.cs:40,95). Paridade semântica coberta
  pelos 1.051 testes existentes (famílias bem testadas).
- [x] **S2**: `FormulaWriter.Call` — dispatch por `FrozenDictionary<Type, ...>` em vez do switch de 304
  braços (FormulaWriter.cs:282-598). Sem mudança de output (FormulaWriterTests congelam).
- [x] **S3**: wildcard estático (`Criteria.WildcardMatch`) compilado 1× antes do scan (LookupMatching.cs:53).

### Verification Plan
- Suítes completas verdes (FormulaWriterTests = oráculo do S2)
- Export em modo Formulas: benchmark antes/depois no fixture sintético
- Push + release verde

### Phase Summary
- **S1** (bf9a23d): SUMIF/AVERAGEIF → PositionalRange; AND/OR/XOR → RangeValueCursor por ref (sem
  IEnumerable boxado). Paridade congelada: erros de célula não propagam no par critério, blanks
  ignorados, shape mismatch = Min-length (≠ SUMIFS #VALUE! — documentado e testado). List de 50k
  eliminada (~1,2MB/leitura). DESCOBERTA: dupla sondagem do TryGetRangeSnapshot (stateful) anula o
  streaming da 1ª época — corrigida via overload que passa o snapshot já sondado.
- **S2+S3+bônus** (commit desta entrada): FormulaWriter com FrozenDictionary sobre os 304 braços
  uniformes + switch residual de 2 (FunctionCall, Sum) — output byte-idêntico, dispatch ~35% mais
  rápido; wildcard de lookup compilado 1× antes do scan (fail-safe de timeout preservado); dupla
  sondagem corrigida em CountIf, MATCH e XLOOKUP (varredura completa dos call sites — demais já
  corretos). COUNTIF 50k 1ª leitura: 9,57MB → 8,36MB (idêntico ao SUMIF).

Suítes: 1.077 core + 48 Excel, 0 falhas. Zero wire format; zero mudança de texto emitido.

## Phase 5: Robustez do interop → release minor
Status: Complete

- [x] **R1**: `ExcelLoadOptions` (aditivo: `Load(path/stream, options?)`) com coletor de warnings
  (`IReadOnlyList<LoadWarning>` exposto via options) — defined names inválidos deixam de sumir em silêncio
  (ExcelFile.cs:104).
- [x] **R2**: merge — investigar escrita direta sem `MemoryStream` intermediário (ExcelMerge.cs:143-154):
  part temporário ou arquivo temp; SE inviável no SDK, documentar a limitação na doc da classe e
  registrar decisão aqui.
- [x] **R3**: comentário do next-free-tag da union (Expression.cs:15 → 319) + doc do invariante das 3
  estruturas de estrutura (structural index / dirty buckets / dense pages) num comment central.

### Verification Plan
- Suítes verdes; teste novo de warnings; `--excel-memory` merge sem regressão de pico
- Push + release verde

### Phase Summary
- **R1+R3** (182a1f4): `ExcelLoadOptions` aditivo com callback `OnWarning` (ExcelLoadWarning record:
  Kind/Subject/Detail; kinds: InvalidDefinedName, UnparsableDateLiteral com cell id como Subject) —
  defined names inválidos deixaram de sumir em silêncio; sem options, comportamento byte-idêntico.
  5 testes novos; docs en+pt atualizadas. Comentário do next-free-tag corrigido (319, com nota para
  recontar); nota-âncora do invariante triplo de estrutura em SheetStructuralIndex com cross-refs.
- **R2** (commit desta entrada): buffer do merge trocado de MemoryStream para FileStream temp
  (DeleteOnClose, 64KB): allocated 248→176 MB (−29%), peak-live 416→344 MB (−17%), tempo estável;
  round-trip 0 mismatches; sem temporários órfãos. Opção A (part swap) descartada por risco de
  renomear part referenciado.

Suítes: 1.077 core + 53 Excel, 0 falhas.

## Phase 6: FunctionRegistry (AVALIAR BREAKING com o usuário antes de executar)
Status: Complete — escopo executado: tudo menos override (decisão do usuário)

- [x] Extrair catálogo de funções do Parser para `FunctionRegistry` (mantendo FrozenDictionary interno)
- [x] `FormulaWriter` deriva nomes do registry (mata a 4ª cópia manual)
- [x] Custom functions: validação de aridade opcional em `RegisterFunction` (aditivo)
- [x] DECISÃO COM O USUÁRIO: override de built-ins NÃO entra (fica documentado como possível evolução opt-in)
- [x] Extrações do Workbook.cs (serialização, época volátil, admissão de range cache, Sheet) — só as
  `[MemoryPackIgnore]`-safe; membros serializados NÃO se movem

### Verification Plan
- Suítes completas + MemoryPackCompatibilityTests (wire intacto) + FormulaWriterTests
- Push + release (minor se houver feat)

### Phase Summary
- **P3** (b2f8d83): admissão população-aware p/ retângulos via structural index (early-exit no
  threshold; nunca força build do índice; legado preservado sem índice). Retângulo esparso 1M×20
  populadas: 192 MB → 888 bytes.
- **M1-M3** (46b7aed): SUBTOTAL passe único numérico (−29% alloc, −45% tempo); fast-paths 0/1-arg no
  ParseFunctionCall (opção (b) do plano REJEITADA com prova — List já nasce cap 4 e 0-args regride);
  NUMBERVALUE por span; presize do snapshot warm; IRR REJEITADO (exigiria alargar
  TimeValueOfMoney.Solve em 3 sites por alocação cold).
- **P4** (commit desta entrada): Values de retângulo fechado → view zero-copy sobre o dense store
  (equivalência PROVADA: build força presença; épocas caem juntas). 40 ranges de 250k: −228,8 MB
  (bate 24B/célula exato). OpenRangeReference deliberadamente adiado (design próprio; whole-column
  flagship intocado, sem regressão).

Suítes: 1.084 core + 53 Excel, 0 falhas.

### Phase Summary
- **6a** (ea4b4fa): `FunctionRegistry` com 305 entradas unificadas → índices ByName (Parser) e ByType
  (FormulaWriter) derivados de UMA lista; Sum absorvido via accessor custom (só FunctionCall permanece
  especial, por ser o fallback runtime); Parser 976→642 linhas, FormulaWriter 673→328.
  `RegisterFunction(minArgs, maxArgs)` opcional validado na avaliação (#VALUE! em vez de exceção do
  host); defaults preservam comportamento legado (provado por teste).
- **6b** (commit desta entrada): Workbook.cs 1.384→607 linhas — Sheet.cs, CollectionExtensions.cs,
  Workbook.Serialization.cs, Workbook.VolatileEpoch.cs e admissão do range cache movida para
  RangeValueCache.cs (subsistema Layer-2 inteiro num arquivo). Wire safety PROVADA por emissão do
  source generator (só Sheets+DefinedNames serializam); MemoryPackCompatibilityTests (binário congelado)
  verdes após CADA extração. Flake de isolamento do RegexCacheTests corrigido ([NotInParallel]).

Suítes finais: 1.081 core + 53 Excel, 0 falhas.


## Phase 7: Cauda de desempenho + micro pendentes → release patch
Status: Complete

- [x] **P3 (admissão população-aware)**: `EstimatePopulatedCells` para RangeReference fechado é cego
  (retorna área capada) enquanto open ranges consultam o structural index para contagem exata
  (Workbook — região movida p/ RangeValueCache.cs na 6b; era Workbook.cs:244-288). Fix: usar o
  structural index também para retângulos fechados (contagem por interseção coluna×linhas), tornando a
  admissão consistente e evitando snapshot de retângulo esparso. Medir: cenário de retângulo 1000×1 com
  10 células populadas NÃO deve mais materializar 1000 slots.
- [x] **M1 (Parser: 2 alocações por function call)**: `ParseFunctionCall` aloca `List<Expression>` +
  `ToArray()` nos dois ramos (Parser.cs ~:772,794,805 pré-refactor; localizar pós-6a). Avaliar: buffer
  pooled, ou contagem em duas passadas, ou aceitar 1 alocação (o array final é retido pelo nó — só a
  List é lixo). Meta: 1 alocação (o array) por chamada.
- [x] **M2 (SUBTOTAL)**: dupla materialização + `ToId()` string por célula (Subtotal.cs:~40,70-84,170).
  Passe único, id via span/CellRef.TryFormat ou consulta numérica direta.
- [x] **M3 (triviais)**: NUMBERVALUE sem LINQ sobre chars (TextFormatting.cs:153,180); presize de
  `SnapshotComputedValues` (Workbook.Serialization.cs, era Workbook.cs:855); closure do IRR (Irr.cs:52)
  se trivial.
- [x] **P4 (INVESTIGAÇÃO — RangeSnapshot.Values duplica o dense store)**: o snapshot materializa cópia
  física dos valores por época (RangeValueCache.cs:107,163-209) além dos índices derivados. Investigar
  (estilo R2, com critério): é viável construir os índices derivados lendo o dense store sem reter
  `Values`, mantendo os consumidores de iteração via cursor? SE a medição mostrar regressão de tempo
  >10% nos benchmarks whole-column, REJEITAR e documentar. Rodar --whole-column-scale e --k1-endtoend
  antes/depois.

### Verification Plan
- Suítes completas verdes; `--excel-memory`, `--k1-endtoend` e `--whole-column-scale` sem regressão
- P3: teste do retângulo esparso; P4: decisão documentada com números
- Push + release patch

### Phase Summary
_(write when phase completes)_


## Phase 8: Pressão de GC no load (Gen1/Gen2 vs Aspose) → release patch
Status: Not started

Contexto (medição externa do usuário, BDN, leitura de K1 ~500k células): Aspose (99000, 30000, 5000)
vs MySheet (77000, 43000, 8000) em Gen0/1/2 → nossa lib aloca MENOS transiente mas PROMOVE mais
(sobrevivência), ficando ~33% mais lenta. Causa estrutural: modelo = milhões de objetos pequenos
(1 Expression/célula + strings + dict entries) vs modelo colunar do Aspose (poucos arrays grandes);
custo de mark ∝ objetos vivos + promoção de todo objeto retido nascido durante o load.

- [x] **G1 (instrumento — espec. refinada pelo usuário)**: benchmark BDN com MemoryDiagnoser para
  ExcelFile.Load com DOIS casos: (a) xlsx CONVERTIDO do .myxl do repo (GlobalSetup: Workbook.Load do
  k1.myxl|k1-synthetic.myxl → SaveAsExcel com FormulaMode.Formulas → temp xlsx) — fórmulas completas
  por célula, sem shared formulas (nosso export não as emite); (b) k1-synthetic.xlsx (grupos de shared
  formula reais). Os dois shapes isolam parse pleno vs expansão por delta. Gen0/1/2 + tempo + allocated;
  baseline registrada ANTES do G2.
- [x] **G2 (dedup de literais no load)**: no WorksheetStreamLoader/LoadContext, dedup de
  `NumberValue` por valor (dicionário por load; inteiro pequenos e valores repetidos dominam dados),
  `StringValue` 1 wrapper por instância de shared string, `BooleanValue.True/False` singletons (se já
  não existirem). Mesma prova de segurança do FormulaCache (imutável, estado por (sheet,col,row) fora
  do nó; MemoryPack sem reference-tracking → wire por célula idêntico). Medir com G1: meta = queda
  visível de Gen1/Gen2 e de objetos vivos pós-load.
- [x] **G3 SPIKE EXECUTADO — VEREDITO: PROMOTE** (todos os gates com folga): escravas de shared formula como nó-delta
  `(masterTree, deltaRow, deltaCol)` em vez de 360k+ árvores expandidas — como o Excel armazena.
  Colapsa a contagem de objetos do load na maior alavanca disponível. EXIGE: união tag nova
  (append-only OK), resolução de referências delta-aware na AVALIAÇÃO (mudança profunda), FormulaWriter
  (ToFormula da escrava = shift on-demand), paridade com SharedFormulaDeltaTests. Especificar spike com
  critérios ANTES de codar; avaliar com números do G1/G2 na mão.

**Resultados do spike G3 (worktree agent-a6eb41a4368d55969, checkpoint 64fbf3c, base 4a10c6d):**
load shared-formulas: 546→295ms (−46%), 251,6→102,5MB (−59%), Gen1+Gen2 18000→9000 (−50%);
controle sem shared inalterado; compute +2,3% (<5%); paridade e round-trip .myxl verdes.
Design: AnchoredCellReference/AnchoredRangeReference/SharedFormulaSlave (tags 319-321),
modo ancorado no Parser, delta no EvaluationContext (WithCell zera). Pendências de produção:
(1) auditoria de pattern-match `is CellReference` nas 305 funções (só MAX/IF/ROUND adaptadas);
(2) DependencyExtractor conservador (AlwaysDirty); (3) AnchoredRange aloca transiente;
(4) ganho é RAM/GC, não disco (wire ainda duplica master por escrava).

**Heap profile (gcdump, processos isolados, xlsx vs .myxl do mesmo conteúdo):**
- HIPÓTESE DAS ASTs CONFIRMADA: nós Expressions.* + Expression[] = 61,7%/60,8% do heap (A/B),
  ~75% dos objetos vivos — igual nos dois formatos de load.
- MEMORYPACK: suspeito real mas SECUNDÁRIO (+3,9% heap no caminho .myxl): o deserializer DESFAZ o
  dedup de StringValue do G2 (209→60.008 wrappers, +59.799 strings, ~4,7MB) + ~64KB de sobras
  (buffer 65.560B retido, formatters). FIX BARATO IDENTIFICADO: dedup no CellStoreFormatter.Deserialize
  espelhando o GetOrAddString do loader → novo item M4.
  **M4 EXECUTADO** (commit perf(serialization)): dedup read-side no formatter (write intocado, bytes
  de disco idênticos); probe: 209→209 instâncias através do round-trip (antes 209→60.008); teste
  permanente RoundTrip_SharesRepeatedLiteralInstances; 1.085+53 verdes.
- Bônus: caminho xlsx retém ~237KB de reflexão/emit do OpenXml (one-time, irrelevante).

### Verification Plan
- G1 rodado antes/depois de G2 (tabela no summary); suítes completas verdes; wire intacto
  (MemoryPackCompatibilityTests)
- Push + release patch (G2); G3 só após decisão

### Phase Summary (parcial — G3 pendente de decisão)
- **G1** (ec696b7): `ExcelLoadBenchmarks` permanente (MemoryDiagnoser): convertido-do-myxl 1.579ms /
  420MB / Gen 60000-19000-5000; shared-formulas 547ms / 253MB / Gen 43000-14000-4000. Razões de
  promoção (Gen1/Gen0 ≈ 0,32) reproduzem a assinatura externa do usuário.
- **G2** (commit desta entrada): singletons de BooleanValue + dedup de StringValue (60.008→209
  instâncias no K1, provado por identidade) e ErrorValue no load. NumberValue dedup MEDIDO E
  REVERTIDO (52% duplicação < break-even ~54-70% do overhead do dict; matemática no call site).
  Efeito agregado marginal (~0,4% Allocated; Gen counts inalterados) — a população dominante de
  objetos retidos são as ~360k árvores de fórmula → o gap de Gen1/Gen2 do usuário aponta para o G3.
  Armadilha de metodologia documentada: GetTotalMemory enganado por FragmentedBytes; PromotedBytes e
  contagem por identidade são os sinais confiáveis.

**Estado ao pausar (2026-07-10, fim do dia)**: G1+G2 entregues (v3.12.2); decisões do usuário:
G3 = SPIKE AGORA (especificar com critérios: paridade via SharedFormulaDeltaTests + ExcelLoadBenchmarks
como gate; protótipo atrás do loader; se provar ganho → /real-work próprio para produção);
intern do Sheet.Name no load = NÃO (só o caveat multi-tenant no backlog). Nenhum agente em voo;
working tree limpo; próxima ação = disparar o spike G3.

**Nota de design (usuário, 2026-07-10, pós-pausa) — arena de fórmulas (candidata a v5):** se a
representação bytecode/RPN for adiante após os números do spike G3 + heap profile, a arena deve ser
PAGINADA (blocos ~32-64KB, sob o limiar de 85KB do LOH; handle = bloco+offset+len; bump allocator;
fórmula nunca cruza bloco; blocos owned, não ArrayPool), não um array contíguo único — mesmo idioma já
provado no SheetValueStore (páginas de ~24,6KB, geometria em WorkbookOptions). Promoção dos blocos a
Gen2 é desejável (modelo residente estável); o ganho anti-LOH é fragmentação/compactação. MemoryPack
continua quase-blit (lista de byte[] = memcpy por bloco). Decisão entre incremental (G3+refs numéricas)
vs arena SÓ com o heap profile na mão.

## Backlog (triado da auditoria completa — válido, não planejado)

Itens dos 4 relatórios que NÃO subiram ao plano, registrados para não se perder:
- **OpenRangeReference sem cópia Values** (follow-up do P4): exige inversão posição→(col,row) sobre os
  buckets do structural index (abstração de offsets segmentados); flagship whole-column — design próprio.
- **Spike EvictDense × range cache** (achado do P4): EvictDense evicta célula densa SEM limpar
  _rangeCache — snapshot stale se o caminho spike do dirty graph for usado com ranges; pré-existente,
  documentado como "not host API", mas é armadilha se o spike for promovido.
- **Dirty graph é ilha** (decisão de produto): integrar ao SetCell/InvalidateCache OU medir e documentar
  o crossover rebuild O(F) vs InvalidateCache; hoje há dois mundos de invalidação paralelos.
- **AST = modelo serializado**: spans/trivia p/ error-recovery e parser incremental exigem CST paralelo
  (BREAKING-RISK no wire se feito nos nós). Limitação estratégica.
- **Override de built-ins por custom functions** (decisão do usuário: adiado; seria opt-in).
- **Tokens-por-span no Tokenizer** (~1-1,5 dia; gated na medição do usuário no ambiente real — ver
  plans/excel-load-streaming.md Fase 8).
- Caminhos de manutenção de época sem guarda de thread (`SheetValueStore.Clear`,
  `EnumerateNonTainted` — contrato single-thread só em prosa; Save warm concorrente pode ler torn).
- Boilerplate de extração de argumentos financeiros (Bonds/OddBonds/CouponSchedule — manutenção;
  ordem de avaliação de erros é observável, abstração precisa preservá-la).
- `ReferenceGuard.MissingSheet` re-resolve NameReference/DynamicRange que a função resolve de novo.
- LOOKUP forma-array materializa keys+results (baixo; forma-vetor já é zero-copy).
- Seqlock: nota de ABA teórico a 2³¹ writes (comentário, não fix).
- **Intern pool de nomes de sheet é process-lifetime** (Parser string.Intern + MemoryPack
  InternStringFormatter, mesmo pool global): num servidor multi-tenant carregando milhares de arquivos
  com nomes distintos, o pool cresce para sempre. Trade-off deliberado ("tiny bounded set" vale por
  workbook, não por processo); trocar por pool por-workbook quebraria a convergência com o formatter.
  Nuance menor: ExcelFile.Sheets.Add não interna o Sheet.Name → até 2 instâncias por nome pós-load.
- Export para XmlWriter puro (unificação com merge; ganho modesto, decisão adiada na sessão de load).

## Final Recap

Plano executado integralmente em 6 fases / 6 releases parciais, tudo delegado a subagentes Sonnet com
revisão e commit do orquestrador. Zero breaking changes (override de built-ins deliberadamente fora,
registrado como evolução opt-in futura). Zero toque em wire format (provado por source generator +
binário congelado).

| Release | Conteúdo | Destaques medidos |
|---|---|---|
| v3.10.1 | 3 correções de corretude | DefineName×grafo stale; merge×@r implícito; StackOverflow→ParseException |
| v3.10.2 | Superlineares | RANK 5k: 740→~25ms; regex 50k: 110→13ms; +fix ReDoS sem timeout |
| v3.10.3 | Alocação (7 itens) | VLOOKUP 22KB→0,1B/aval; YIELD ~3x; LET 14,7x menos alocação |
| v3.10.4 | Streaming + dispatch | Writer ~35% + output idêntico; dupla sondagem corrigida em 3 funções |
| v3.11.0 | Interop robusto | ExcelLoadOptions/warnings; merge −29% alloc/−17% pico (temp file) |
| (próxima) | Registry + god-file | Catálogo 4×→1×; Workbook.cs 1.384→607 linhas |

Descobertas de agentes além do escopo pedido: 2º ciclo de recursão no Parser (F3), bug de dupla
sondagem do snapshot em CountIf/MATCH/XLOOKUP (S1/S2), ausência de timeout no wildcard (P2), flake de
isolamento do RegexCacheTests (6b).

Pendências documentadas para ciclos futuros: override opt-in de built-ins; tokens-por-span no
Tokenizer (plans/excel-load-streaming.md); AST=wire como limitação estratégica; crossover
rebuild-vs-InvalidateCache do dirty graph (documentar número).

## Deployment Plan

Cada fase já foi publicada no NuGet pela release workflow (v3.10.1→v3.11.0). A release final
(registry + extrações, refactor: não bumpa sozinho — sai como patch junto do próximo fix/perf, OU
disparada manualmente se desejado; o commit refactor: estará no CHANGELOG da próxima release).
Consumidores: nenhum código precisa mudar; hosts que queiram aridade validada passam
minArgs/maxArgs no RegisterFunction; hosts que queiram warnings de load passam ExcelLoadOptions.
