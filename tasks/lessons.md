# Lessons

Padrões aprendidos com correções e descobertas, para não repetir erros.

## MySheet — funções financeiras (2026-06-29)

- **Não confiar em golden values calculados de cabeça.** Eu cravei `IRR{-10000,3000,4200,6800} ≈ 0.201`
  no plano; o valor real é `0.16341`. E "corrigi" um PV correto para um errado. Regra: usar um oráculo
  (`ExcelFinancialFunctions`) e verificar antes de afirmar números.
- **Verificar a capacidade de uma lib antes de depender dela.** ClosedXML só computa 3 das 9 financeiras
  (PV/NPER/PPMT/NPV/RATE/IRR → `#NAME?`). Spike rápido evitou apostar numa estratégia de teste furada.
- **"Replicar o Excel com Newton ingênuo (~20 iter)" é uma armadilha para RATE/IRR.** Newton e o secante
  do POI/LibreOffice falham em `RATE(360,-600,100000)` (financiamento de 30 anos) — caso trivial que o
  Excel resolve. `(1+guess)^360` domina o resíduo e o método local diverge ou converge falso. O Excel usa
  solver robusto; bracketing + bisseção (convergência na *taxa*, não no resíduo em moeda) é o que de fato
  bate com o Excel. Lição: quando o objetivo é "igual ao Excel", validar contra um caso stiff (mortgage)
  antes de fixar o algoritmo.
- **Checar o framework de teste antes de rodar.** O projeto usa **TUnit** (`[Test]`, `await Assert.That().IsEqualTo().Within()`),
  não xUnit. `dotnet test` falha no .NET 10 (VSTest removido); o comando é
  `dotnet run --project tests/...Tests.csproj -c Release -- --treenode-filter "/*/*/Classe/*"`.
- **`.Within(tolerance)` do TUnit exige `double` (não `double?`).** Extrair o número com um helper que vira
  `NaN` quando vier `ErrorValue`, para o assert falhar limpo no RED em vez de lançar exceção no cast.

## MySheet — design de API pública / refactor de representação (2026-07-01)

- **Separar mudança de representação interna da forma da API pública.** No experimento `ComputedValue`
  (trocar `object?` por struct para matar boxing), meu esboço de migração assumiu de cara "manter `object?`
  público / remover a ponte `AsObject`" — um default conservador apresentado como se fosse obviamente certo.
  O usuário queria justamente o oposto: **retornar** o novo tipo publicamente, com helpers ergonômicos. Lição:
  o ganho de GC vem do interno (cache + nós passando o struct); a forma da API pública é decisão de produto
  do usuário — apresentar como opções (A: preservar / B: opt-in / C: trocar), não cravar a conservadora.
- **Validar a semântica do domínio de valor no código antes de opinar.** Sobre "qual o valor de uma range?":
  o código já separa duas camadas — `RangeReference.Compute` → `#VALUE!`, `Expand()` → `IEnumerable<Expression>`
  (referência), `ExpandValues()` → valores via cache. `IEnumerable<Expression>` NÃO sai do `Compute`; é camada
  de referência. Ler `RangeReference.cs` antes de propor evitou confirmar uma modelagem que misturava camadas.
- **Boxing no cache pesa mais que o transitório.** Medido: cache `object?` boxa 24 B/célula de **vida longa**
  → dispara **Gen1**; cache `Dictionary<string, ComputedValue>` → 0 B, zero coletas. O throughput no caminho
  cache-heavy sobe só 4–12% (lookups de Dictionary dominam), mas o ganho de GC é o argumento forte, não a
  velocidade bruta. Não vender um refactor de perf só pelo throughput sem medir o eixo de alocação/GC.

## MySheet — convenções e processo (2026-07-01)

- **Commits: inglês, sujeito curto + corpo descritivo, semantic/conventional commits.** Correção do usuário
  em 2026-07-01 — os commits em português eram a convenção antiga do repo; o padrão agora é
  `tipo(escopo): resumo em inglês` + parágrafo curto de contexto. Nunca incluir atribuição de IA.
- **Contagens em planos: recontar antes de publicar.** Escrevi "~85 exclusões" no roadmap de funções quando
  a lista explícita somava 35 — o usuário aprovou a LISTA, mas o número errado contaminou a meta de
  cobertura (~435 vs ~485 viáveis). Números derivados de listas devem ser contados por script/soma real,
  nunca estimados de memória (mesma família da lição dos golden values).

- **Verificar builds de agentes com `--no-incremental`.** O relatório da Onda 5 alegou 0 warnings, mas o
  build incremental escondia 14 avisos de analyzer nos testes (recompilar não reemite warnings de projetos
  up-to-date). A verificação independente agora força rebuild. Corolário: nunca encadear `git commit` atrás
  de build sem condicionar ao sucesso — um sed meu quebrou o build e o commit passou junto.

## MySheet — financeiras de título / oráculo como fonte de verdade (2026-07-02, Onda 6)

- **Quando o oráculo É a fonte de verdade, porte a lógica DELE — não assuma a fórmula do livro-texto.**
  A fórmula canônica de PRICE (`(1+y/f)^(k-1+DSC/E)` com `E = COUPDAYS`) NÃO reproduz o
  `ExcelFinancialFunctions` (o oráculo). O oráculo computa `dsc = e - a` com `a = DaysBetween(pcd, settlement)`
  e `e = CoupDays`, e — pior — os day-counts de título usam uma variante `ModifyStartDate`/`ModifyBothDates`
  do 30/360-US e um "days in year" actual/actual próprios que DIVERGEM do `DayCount` do YEARFRAC (onda 5)
  em casos de fim-de-mês/fevereiro (ex.: início Feb-end + fim dia-31 → 1 dia de diferença). Eu perdi tempo
  testando variantes de fórmula às cegas; a virada foi BUSCAR O CÓDIGO-FONTE do oráculo (F# no GitHub) e
  portá-lo verbatim. Regra: se a tolerância exigida é 1e-9 contra uma lib específica, leia a lib.
- **Fuzz valor-a-valor contra o oráculo ANTES de portar ao codebase.** Montei um probe console que
  implementava os candidatos em C# e comparava contra o oráculo em dezenas de milhares de casos (5 bases ×
  3 frequências × datas aleatórias) por função. Pegou: (a) loop de agenda de cupom invertido (off-by-one no
  passo); (b) que a agenda ANDA iterativamente para trás (clamp de fevereiro é "sticky", ≠ computar do
  maturity direto); (c) off-by-one no `findDepr` do AMORDEGRC (entra com countedPeriod=1 e incrementa antes
  de computar). Portei só depois de maxErr=0 (closed-form) / ~1e-9 (solver). Resultado: 46 funções entraram
  verdes de primeira no codebase — o único bug restante (índice do XNPV values/dates) foi de PLUMBING do
  record, não da matemática, e foi o teste (não o fuzz) que pegou.
- **Preconditions do Excel ≠ do "seria razoável".** `ACCRINT` do oráculo EXIGE `first_interest >= settlement`
  (o ramo multi-período do código-fonte é inalcançável pela API pública); `ODDFPRICE` tem uma precondição
  NÃO-documentada (o first_coupon precisa alinhar com a agenda vinda do maturity) que a própria lib comenta
  como "not in the docs, but nevertheless is needed". Um teste meu quebrou por usar first_interest < settlement
  "multi-período" — inválido. Ler as `calc*` (preconditions) do oráculo evita golden values impossíveis.

## MySheet — voláteis / MemoryPack e thread-safety (2026-07-02, F1)

- **MemoryPack IGNORA os field/property initializers na desserialização.** Um membro com initializer (`= new()`,
  `= TimeProvider.System`) vem NULL depois de um `Load` — a evidência já estava no codebase (o `RestoreComparers`
  trata `DefinedNames` null "older files carry no DefinedNames"). Consequência para estado NÃO-serializado
  (`[MemoryPackIgnore]`) que precisa existir em runtime (locks, providers, sets concorrentes): NÃO confie no
  initializer. Padrões robustos: (a) getter lazy com default (`_timeProvider ?? TimeProvider.System`);
  (b) criação lazy race-free via `Interlocked.CompareExchange` (locks/dicionários concorrentes). Assim funciona
  tanto no `new Workbook()` quanto no `Load`, sem depender do hook `[MemoryPackOnDeserialized]`.
- **Compat binária de config runtime = `[MemoryPackIgnore]`, não membro no fim do schema.** `TimeProvider`/
  `RandomSeed` são config, não estado: `[MemoryPackIgnore]` mantém o schema serializado inalterado (a fixture
  antiga abre) — diferente do `DefinedNames`, que É estado e foi appendado como último membro. `IsVolatile` é
  comportamento: property virtual get-only com `[MemoryPackIgnore]` (o analyzer do MemoryPack aceita, mesmo
  padrão do `Sheet.Count`). Regra: pergunte "isto é ESTADO do documento ou CONFIG/COMPORTAMENTO?" antes de
  decidir append-ao-schema vs ignore.
- **Amostra-uma-vez sob concorrência = lock simples, não fast-path lock-free de `double?`.** Para amostrar
  `_epochNow` (um `double?`) 1× por época com avaliação concorrente, um `lock` sempre (`_epochNow ??= …` dentro)
  é correto e barato (não é caminho quente: 1× por célula volátil por época). O fast-path `if (_epochNow is {})`
  sem lock teria torn read do `double?` (struct não-atômica). Não otimize prematuramente o que não é hot.

- **Agentes nunca devem usar `git commit --amend` (nem rewrite) em branch compartilhada.** Na fase de perf
  de coluna inteira, o agente A amendou o próprio commit de docs (44c0662→6b85577) DEPOIS que o agente B já
  tinha bifurcado do commit original → históricos divergentes e rebase forçado na integração. Conteúdo era
  benigno (tabela de resultados), mas a forma certa era um commit NOVO ("docs: add full-scale numbers").
  Regra para briefings: append-only também no git — amend só se o commit nunca saiu do próprio agente E
  nenhum outro trabalho partiu dele (na prática: nunca).

## MySheet — benchmarks de estratégia / eixo de carga (2026-07-03, coluna inteira)

- **Um benchmark de ESTRATÉGIA tem de modelar o EIXO DE CARGA COMPLETO: leituras × fórmulas × tamanho.**
  O spike de coluna inteira (`plans/whole-column-spike.md`) mediu "273 µs/scan, 1× por época via memoização"
  e concluiu break-even do índice em "≥4 leituras/época". Verdade — mas POR FÓRMULA: a memoização é por
  CÉLULA, não por época global. Com F=400k fórmulas cada uma varrendo N=506k chaves, o custo real é
  O(F×N) ≈ 2×10¹¹ visitas ≈ 57min — o break-even foi estourado por 5 ORDENS DE GRANDEZA. O spike modelou o
  eixo da AMORTIZAÇÃO POR LEITURA (quantas vezes um MESMO scan é reusado) e esqueceu o eixo da MULTIPLICIDADE
  DE FÓRMULAS (quantos scans distintos existem). Regra: antes de cravar um limiar/estratégia, enumere TODOS
  os eixos que multiplicam o custo (nº de fórmulas × nº de células por fórmula × nº de leituras) e ponha o
  pior deles no gerador do benchmark — senão o número "por unidade" mente sobre o total.
- **Pós-otimização, extrapolação linear de amostra pode INVERTER de válida para enganosa.** O harness `--full`
  amostrava 1k fórmulas e multiplicava ×100. Na baseline (custo O(F×N), linear por fórmula) isso era CORRETO.
  Depois dos caches (custo O(N + F·log N)) o build do snapshot é O(N) UMA vez amortizado por TODO o bloco —
  amostrar 1k e multiplicar ×100 multiplica esse build único ×100 e superestima grosseiramente (deu "60s" que
  eram quase todo build). Regra: quando a mudança que você está medindo altera a COMPLEXIDADE (não só a
  constante), a extrapolação de amostra que valia na baseline deixa de valer — MEÇA a carga real.
- **Verificar a capacidade do concorrente ANTES de afirmar que você "ganha" (corolário da lição financeira).**
  Antes de comparar whole-column com o ClosedXML, um spike de 16 linhas confirmou o que ele avalia: MATCH/
  VLOOKUP/SUMIF/COUNTIF/SUM sim, mas SMALL/LARGE → `#NAME?`. A incapacidade é resposta (eles não competem em
  SMALL), e evita medir o par inexistente e reportar um "ganho" inventado.

- **Antes de disparar release: confirmar que o commit-alvo ESTÁ na main.** No fix de admissão do range
  cache, mergeei a branch da worktree (que não tinha o commit — o agente criou uma branch própria nomeada
  no relatório) e disparei o release: o v2.6.3 saiu VAZIO (docs-only) e o fix teve que sair como v2.6.4.
  Regras: (1) integrar a branch que o agente NOMEIA no relatório, não a branch da worktree; (2) `git log`
  confirmando o hash do fix na main é pré-condição do `gh workflow run release.yml`; (3) nunca engolir
  falha de `git branch -d` com `|| true` — a recusa "not fully merged" era o aviso.

- **Reincidência do release vazio (2.8.1) — a causa raiz era o ENCADEAMENTO, não a falta de checagem.** Eu
  tinha a lição ("confirmar hash na main antes do release") e até imprimi o git log — mas o `gh workflow
  run` estava no MESMO bloco Bash que o rebase que falhou, então disparou de qualquer jeito. Regra
  reforçada: o dispatch de release vive SEMPRE numa chamada separada, DEPOIS de eu ler a verificação da
  pré-condição (merge-base --is-ancestor do hash da feature). Blocos compostos param de valer para
  qualquer sequência que contenha uma ação irreversível.
