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

## MySheet — supervisão multi-worktree (2026-07-03, ciclo 3.0)

- **A forma do breaking commit é `tipo(escopo)!:`, nunca `tipo!(escopo):`.** Meu briefing da Fase 1 do 3.0
  pediu o literal `feat!(sheet):` — forma INVÁLIDA de conventional commit que o versionize não parseia
  (o commit seria ignorado na derivação do major). O agente pegou e corrigiu para `feat(sheet)!:` + rodapé
  `BREAKING CHANGE:`. Regra: o `!` vem DEPOIS do parêntese de escopo; conferir a forma no briefing antes
  de despachar.
- **Com várias worktrees ativas, comando git de integração/commit SEMPRE com `cd` absoluto explícito no
  próprio bloco.** Um commit do plano quase caiu na worktree da Fase 1 porque o cwd do shell persiste
  entre chamadas e eu tinha "sobrado" na worktree do agente (o `git add` de caminho relativo não achou
  nada e o erro salvou). O cwd após verificação de entrega de agente é IMPREVISÍVEL — tratar todo bloco
  git de supervisão como se o cwd estivesse errado. REINCIDÊNCIA (2026-07-04, 2×): (a) probe rodou na
  main em vez da worktree (flag "desconhecida" do BDN foi o sintoma); (b) o `merge --ff-only` da branch
  de docs rodou DENTRO da worktree (na própria branch → "Already up to date" e push vazio — a main não
  recebeu nada). Nuance: um `cd` para a worktree NO INÍCIO do bloco envenena o merge no FIM do mesmo
  bloco. Regra operacional: verificação (na worktree) e integração (na main) NUNCA no mesmo bloco Bash;
  o bloco de integração começa com `cd /Volumes/Work/Develop/MySheet`.
- **O número do release nasce dos TIPOS de commit, não do plano.** O ciclo "3.3" saiu como **3.2.1**
  porque tudo foi commitado como `perf:`/`test:`/`docs:` — e `perf` deriva PATCH no versionize; só
  `feat` deriva minor. Havia um `feat` legítimo mal rotulado (`InitialPageSlots` é API pública nova,
  commitado como `perf(store)`). Regra: o briefing de cada fase declara o TIPO exato do commit principal
  em função da versão-alvo (`feat` p/ minor, `feat!`+BREAKING p/ major, `perf`/`fix` p/ patch), e o
  supervisor confere o tipo ANTES do merge — o versionize não lê intenção.
- **`--no-build` imediatamente após um merge roda binários VELHOS.** Na integração do dense store, a
  "sanidade na main" reportou 893 testes quando a branch mergeada tinha 901 — o `dotnet run --no-build`
  reusou binários pré-merge e testou o código antigo (contagem menor foi o único aviso; podia ter
  passado despercebido com contagem igual). Regra: sanidade pós-merge SEMPRE começa com
  `dotnet build --no-incremental` no próprio bloco, ANTES de qualquer `--no-build`; e a contagem
  esperada de testes é parte da verificação (anotar o esperado antes de rodar).
- **Agente que roda benchmark DEVE rodar em foreground — dois agentes seguidos encerraram o turno
  "aguardando" o próprio run em background, e o watcher morre com a sessão do agente (a entrega fica
  pela metade e o orquestrador paga um resume).** Regra de briefing para qualquer missão com
  benchmark/harness: "rode em foreground e espere; NUNCA termine o turno com trabalho pendente em
  background". O orquestrador, ao retomar um agente nesse estado, instrui re-rodar em foreground.

## MySheet — issue #8 (2026-09-08)

- **Rodar a fórmula representativa da issue de ponta a ponta, não só as linhas da tabela de sintomas.**
  A issue isolava dois defeitos (`+texto`, `$1:$1`) e afirmava que "todo o resto da fórmula funciona". Após
  corrigir os dois, a fórmula real (`LET(hdr, Data!$1:$1, MATCH(x, hdr, 0), …)`) ainda dava `#N/A`: o `LET`
  vinculava ranges a `#VALUE!` — bug independente do `$`, que só apareceu porque escrevi a fórmula inteira
  como teste. Regra: o critério de aceite é a fórmula do usuário funcionar, não os sintomas listados sumirem.
- **Antes de procurar uma issue, confirmar o tracker.** `gh issue view 3163` no repo, nos repos vizinhos,
  Trello e memória não achou nada; a issue real era a #8 no próprio repo (número trocado pelo usuário).
  Uma pergunta objetiva ("onde está?") resolveu em um turno; meia dúzia de buscas às cegas não resolveram.
- **Separar hunks por commit sem `git add -p`:** `git diff -- arquivo` → filtrar hunks por regex →
  `git apply --cached --recount`. Arquivo novo cujo conteúdo pertence a dois commits: gravar a versão
  parcial, `git add`, restaurar a completa. Permite commits conventional independentes (changelog do
  versionize) quando um arquivo acumula mudanças de defeitos diferentes.
- **Toda afirmação em doc/comentário precisa de um teste que a exercite literalmente.** Escrevi em
  `UnaryOperation` e na doc que "`SUM(+A1:A3)` continua vendo a range", mas o único teste era
  `SUM(+OFFSET(...))` — um VALOR de referência, caminho que já funcionava. Um `A1:A3` sintático virava
  `#VALUE!`. A revisão externa (zclaude/GLM) pegou porque comparou a frase da doc com os testes do diff.
  Regra: ao escrever "X funciona" num comentário, o exemplo X vira um `[Arguments]` do teste, sem substituto.
- **Revisões externas valem pelo que trazem de novo, não pelo volume.** Copilot (lite) achou 2 pontos de API
  e overflow; o zclaude com o diff completo achou o bug semântico que os dois primeiros passes não viram.
  Rodar o segundo revisor DEPOIS de corrigir o primeiro, com o diff atualizado, evita relatórios duplicados.


## MySheet — fallback de fórmula não parseável / tabelas (2026-09-08)

- **Fallback que decodifica dado que antes era ignorado precisa ser total ANTES de eu chamar o fix de pronto.**
  Fiz o loader degradar uma fórmula não parseável para o `<v>` em cache — e o `<v>` de célula de fórmula
  *nunca era decodificado* até então. `DecodeLiteral` usava `double.Parse`/`int.Parse` sem guarda, então
  `<f>…</f><v/>` trocou `ParseException` por `FormatException`: a carga continuava morrendo. Reportei como
  "completo e verificado". Regra: quando um fix passa a alimentar um decoder com um shape novo, enumerar os
  shapes daquele canal (vazio, tipo errado, índice fora de faixa) e testar cada um — o caminho de fallback é
  exatamente o que só roda em arquivo estranho, ou seja, nunca é exercitado pela suíte existente.
- **Valor de fixture que coincide com o valor correto é asserção que não falha.** Dois testes meus cachearam
  `42` numa célula cujo `SUM` real também dava `42`: passavam vindo do cache OU do cálculo, pinando nada.
  Regra: em teste de fallback, o valor em cache tem que ser uma MENTIRA deliberada (999) para que a asserção
  distinga a origem.
- **Desfazer efeito no `catch` é pior que não causá-lo.** Eu registrava o grupo de fórmula compartilhada e no
  `catch` fazia `SharedFormulas.Remove(si)` — mas o `Remove` não sabe se a entrada é a que eu escrevi. Com si
  reusado (fora de spec) apagava o grupo legítimo e as escravas dele perdiam a fórmula em silêncio. A solução
  elegante foi reordenar: parsear primeiro, registrar só depois do sucesso. Regra: preferir "não causar o
  efeito" a "causar e reverter" — reverter exige identidade que geralmente não se tem.
- **`catch (ParseException)` não cobre "não parseou".** O caminho anchored faz `int.Parse` da linha, então
  `A99999999999+1` sobe `OverflowException` e escapa. Regra: ao capturar por tipo de exceção para significar
  uma CONDIÇÃO ("não representável"), auditar todo o caminho por outros tipos que expressam a mesma condição,
  e ajustar a doc se algum ficar de fora.
- **Fidelidade ao Excel/Aspose ganha de conveniência de implementação (diretriz do usuário).** Eu havia
  decidido `#VALUE!` uniforme para expressão multi-célula em contexto de célula porque é simples e consistente.
  Não é comportamento do Excel: sem spill, o fiel é interseção implícita (o que o `@` significa hoje).
  `EvaluationContext` já carrega `SheetName`/`CellId`, então era possível — e "possível" é o critério.
  Regra: para cada decisão semântica, declarar QUAL comportamento do Excel está sendo reproduzido; se a
  resposta for "nenhum, é mais simples assim", redesenhar.
- **Subagente de review adversarial com instrução de PROVAR por execução acha o que a revisão por leitura não
  acha.** Os cinco defeitos acima vieram de um revisor que montou projeto de probe fora do repo e executou
  cada alegação. Instrução que fez a diferença: "onde for barato, PROVE com probe; medido vence argumentado".

## MySheet — Fase 1 por subagentes (2026-09-09)

- **`TaskOutput` num agente local despeja o transcript JSONL inteiro no contexto.** Usei uma vez com `block:true`
  e recebi dezenas de KB de transcript. Regra: nunca chamar `TaskOutput` para agente local; a notificação de
  conclusão chega sozinha, e o `.output` de agente é o transcript, não o resultado.
- **Fato de repositório no prompt de despacho vem do repositório, não da memória.** Escrevi o GUID da página
  SUMPRODUCT de cabeça (`…f4b6`) e o repo tinha `…fd2e`; o implementador precisou escolher. Regra: qualquer
  identificador/valor citado num brief é copiado por `grep` do repo na hora do despacho.
- **Agente travado no meio de "experimento de mutação" pode ter deixado a mutação no disco.** Antes de retomar,
  verifiquei a árvore eu mesmo (csharpier, build, suíte) e só então mandei "finalize, não refaça". Regra: estado
  da working tree é medido, não inferido do último log do agente.
- **A review final de branch acha o que a review por tarefa não pode ver.** Cinco reviews por tarefa aprovaram
  ROW e ROWS separadamente; só a review do branch inteiro viu que `ROWS(INDEX(Ghost!…))` = 1 enquanto
  `ROW(INDEX(Ghost!…))` = `#REF!` — a família divergindo no mesmo argumento, o objetivo declarado do item 27.
  Regra: a review final compara COMPORTAMENTOS entre tarefas para o mesmo shape de entrada, não só o diff.
- **Item de plano "verbatim" não é autoridade quando a review oferece a forma elegante.** O plano mandava
  `ColumnNumbersOperand` espelhando `RowNumbersOperand`; colapsar em `PositionNumbersOperand` parametrizado por
  eixo encolheu o arquivo e eliminou o drift por construção. Regra: o plano é o argumento; a regra de elegância
  do projeto e a evidência do revisor podem revogá-lo — registrado como Ruling no ledger com o custo se errado.
- **`zclaude` é o `claude` CLI com ambiente GLM; o wrapper não repassa argumentos.** Chamei `zclaude --print …` e
  falhou; o script (12 linhas, que eu já tinha impresso) termina em `claude` sem `"$@"`. Regra: para invocar o
  GLM não-interativo, exportar os `export` do wrapper e chamar `$HOME/.local/bin/claude --print --model opus
  "<prompt>" < /dev/null` (`opus` → `glm-5.3[1m]`). E: quando imprimo um script para entender a invocação,
  ler até a última linha antes de chamar — a resposta estava lá.
- **Ruling sobre "o que o Excel faz" sem rodar o oráculo disponível.** Decidi que SUBTOTAL/AGGREGATE 1-13 dobram
  um argumento array "como o SUM" a partir de uma afirmação não medida de um verificador de design ("Excel dá 6").
  O Aspose.Cells 26.6.0 — designado no plano como oráculo — estava no cache NuGet e responde `#VALUE!`; o revisor
  Fable rodou e derrubou o ruling. Regra: toda decisão que afirme comportamento do Excel roda o oráculo Aspose
  ANTES de virar ruling; "P0-driven" sem medição é só inferência com capa. Custo real: um commit `fix(eval)` na
  direção errada, cinco pins e três valores nos docs a inverter.

## 2026-09-09 — "Excel" means Aspose (user ruling)

- The user's compatibility target is **Aspose.Cells as measured**, not the Microsoft page and not an imagined Excel. When a page and the oracle disagree, the oracle wins; a page is consulted only where the oracle cannot be made to answer. I had ruled "page wins" on the nested-AGGREGATE skip in Phase 2 — reversed into Phase 11.
- Every recorded divergence is a WORK ITEM under this rule, never a documented limitation, unless the engine structurally cannot match (no spill model, no hidden-row model). Do not offer "keep and document" as an option for a measured divergence; offer the phase that fixes it.
- Controller edits to `plans/`/`tasks/` while an implementer works IN PLACE share the tree: stage and commit immediately (`--no-verify` is legitimate for a docs-only commit blocked by the implementer's unformatted WIP), or the staged files get swept into the implementer's next commit. Better: run implementers in a worktree when the controller expects to edit plans concurrently.
- An oracle can be internally inconsistent, and "match the oracle" then needs a human ruling, not a coin flip. Phase 11 measured Aspose treating a union as a flattened array (ISREF FALSE, ROWS #VALUE!) while accepting `ROWS(UnionName)` = 2 — matching it would delete shipped, tested behaviour. Bring the contradictory measurements to the user as an explicit exception request; never decide it silently in either direction.
- A mutation experiment inside a working tree with uncommitted work is a footgun: `git checkout <file>` to revert the mutation also wipes the task's own edits (a Phase 8 implementer lost ~24 minutes this way and had to re-apply from its scripts). Brief implementers to copy the file to the scratchpad before mutating, or to commit first and mutate after.
- A guard built from a hand-written list rots; a guard that reads its list by reflection over the tests' own `[Arguments]` rows and asserts the uncovered set is EMPTY registers new names automatically and names the offender when one is missing. Ask for the closing assertion, not just the list.
- A mutation experiment must simulate the FULL mistake, not the minimal edit. I mutated only a registry factory name and the stale count assertions fired, so I concluded the guard names its offender; a reviewer mutated the way a contributor actually would (wrong flag + delete the name row + bump the counts, all natural when adding a function) and the suite went green with a range-aware function wrongly lifted. Ask: what would the real mistake look like, including every edit the mistaken developer would make on purpose?
- Fifth instance in this project of a comment asserting a fact the code does not carry: a test comment said four measured values were "pinned so the open question cannot move silently" while the test asserted none of them. Pattern to watch: whenever a comment says "pinned", "verified", "measured" or "guaranteed", check that the artifact it names exists. Reviewers should grep for the cited value, not read the sentence.
- A reviewer sent to confirm "there is no rule" found the rule: Aspose's lone-format-token behaviour is run-length dependent (one letter reads the raw map, two or more read the phantom-aware map). "No derivable rule" is a claim about search effort, not about the oracle — treat it as provisional until someone has tried the obvious axis.
- NEVER tell an agent that an uncommitted change is yours without proving it. I read `git status`, saw a modified test file, remembered editing it, and told the running fix wave I would drop "my" edit — the edit was the wave's work in progress, and had it obeyed literally it would have discarded its own uncommitted work, which no gate catches. Proof before the claim: `git diff <file>` and check whether the commit you are thinking of already contains the change (`git log -1 --stat -- <file>`). In a shared tree, assume every unstaged change belongs to whoever is running, not to you.
- A sentence written to REASSURE invites a claim nobody measured, and hedging one clause is not a fix while the replacement still asserts closure. Phase 9 ran this loop three times on the same paragraph: "exactly four rows differ, everything else agrees" (false), then "the boundary and the rows depending on the phantom day agree" (false in its own terms, a fifth counterexample), then finally "agreement below the boundary is a row-by-row fact, not a rule to extend" plus a pin for the counterexample. Write the negative claim only where a test enforces it; where you cannot enumerate, say the set is not characterised, and never write a closing sentence that promises every number is pinned unless every number is a formula result with an assertion behind it.
- Comparing two oracle measurements taken with DIFFERENT entry modes manufactures phantom contradictions. A Phase 10 probe reported Aspose answering 4 for a nested SUM and 6 for the INDEX sum over the same array; entered consistently as CSE both are 9, and plain entry gives 2. Before recording an oracle self-contradiction, re-measure both halves in the same mode and say which mode in the note — otherwise the "contradiction" is your own methodology.
- A python replacement bounded by "the next list item" or "the next blank line" silently swallows the rest of a bullet. I corrected one sentence in a docs twin that way and dropped four others; a reviewer caught it, and the check that would have caught it at the time is one grep for a distinctive token from the deleted text (`grep -c tableParts` returned 0 after my edit and 1 on main). When editing prose by script: replace an EXACT full block you have printed and read, and afterwards grep for one distinctive word from every part you did not mean to touch.
- Editing a docs twin means diffing the twin against its original afterwards, not trusting that the same edit was applied to both. My edit was right for the English file and lossy for the Portuguese one, and the acceptance note I wrote ("both now separate what is missing from what works") was true of the half I had checked.
- A false mechanism spreads: "MemoryPack bypasses field initializers on deserialize" had propagated to two engine comments, one test comment and three plan files before anyone read the GENERATED formatter, which materializes the object with `new T() { … }` over the parameterless constructor — so the initializer runs and is then overwritten. The consequence everyone relied on was true, which is why it survived. When a comment explains WHY code is shaped a certain way, check the mechanism, not just the consequence, and grep the whole repo for the sentence rather than the one site you were sent to.
- A shared scratchpad directory is not private: a fix wave wrote `scratchpad/probe/` over another agent's tree of the same name and the overwritten files were unrecoverable. Briefs must tell every agent to write under a path unique to its own task, and to check whether the directory exists before writing.
- **A test's NAME is an assertion, and flipping its expected values does not update it.** Task 8 turned three pins from a divergence to parity and left the method called `AProducer_UnderAScalarConditionIf_CollapsesToItsTopLeft_ADivergence`, with a comment ending "these three lines turn red the moment someone fixes the arm — if you are here because they went red, that is the fix landing". Both survived the commit that made the collapse stop happening. When a pin moves from "wrong on purpose" to "right", the name, the comment and any `_ADivergence`/`_KnownLimit` suffix all move with it, and the comment's job changes from "why this is wrong" to "what it was and what moved it".
- **A one-fixture measurement cannot retire a claim someone made on a different fixture.** I recorded that Phase 7's `SUM(FILTER(A:A,A:A>0))` = `#VALUE!` claim was "NOT reproducible — the oracle answers 14 in both modes". It reproduces exactly as written, on the wider fixture the claim was made on: over `A1:A3` = 5, 0, 9 alone it is 14 / 14, but add `A5`=7, `A7`="t", `A8`=7 and it is `#VALUE!` plain / 28 CSE. Before writing "not reproducible", reconstruct the ORIGINAL fixture, not a convenient one — and if you cannot find it, say the fixture is unknown rather than that the claim is false.
- **An oracle number that varies by fixture must never be recorded without its fixture.** Two agents measured `AVERAGE(UNIQUE(...))` and reported 3 and 0; both were right, on 5,0,9 and on 9,5,9,0. A third reader would have logged an oracle self-contradiction that does not exist. The rule was derivable once five fixtures were laid side by side (the numerator drops every element up to and including the last zero, the denominator keeps the full count), and a table of five rows is shorter than the argument the single number would have started. Same discipline as the entry-mode lesson above: the number is meaningless without the conditions.
- **A "next free tag" computed against `main` is wrong whenever another branch merges first.** A re-verification correctly found the next free MemoryPack union tag was 323 on main, and correctly flagged that the phase design's 322 was taken. It was still the wrong answer for the phase in question, because the branch merging ahead of it takes 323 through 326. Any append-only allocation — union tags, error codes, enum members — must be computed against the state at MERGE time, so name the branches queued ahead and count past them.
- **`grep -c <identifier>` overcounts when a policy comment names the identifier.** `grep -c MemoryPackUnion Expression.cs` prints 328 where 327 attributes exist, because the append-only warning comment above them contains the word. Anchor the pattern to the syntax (`^\[MemoryPackUnion`) whenever a count is going to be written into a document as a fact.
