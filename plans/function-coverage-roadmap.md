# Roadmap de cobertura de funções — ondas escalares + datas (1.1.0 → 1.6.0)

Elevar a cobertura de 52 → ~230 funções em 6 ondas SEM mudança estrutural no engine, liberando **um
release minor por onda**. Decisões estruturais (voláteis, arrays, LAMBDA, distribuições) ficam ANALISADAS
aqui com proposta + alternativas, mas são executadas em fases futuras próprias — correto antes de remendo.

## Context / decisões (fechadas com o usuário em 2026-07-01)
- **Escopo agora = ondas 1–6 (escalares + datas)**: nenhuma mudança no `ComputedValue`/cache/parser além de
  novos nós. Arrays (onda F2) e LAMBDA (F3) ficam desenhados, não executados.
- **Voláteis adiados**: TODAY/NOW/RAND/RANDBETWEEN/INDIRECT ficam FORA das ondas (a onda de datas entrega
  23/25 funções). A infra de volatilidade (análise §A1) é a fase futura F1.
- **Exclusões permanentes (35 funções — recontado; a estimativa inicial de ~85 estava errada)**: Cubes (7), Web (3), `RTD/STOCKHISTORY/IMAGE/HYPERLINK`,
  pivô (`GETPIVOTDATA/PIVOTBY/GROUPBY`), ambiente de UI (`CELL/INFO`), `CALL/REGISTER.ID/EUROCONVERT`,
  byte-duplo/locale CJK (`DBCS/ASC/BAHTTEXT/PHONETIC/LENB/FINDB/LEFTB/MIDB/RIGHTB/SEARCHB/REPLACEB`),
  serviços de tradução (`DETECTLANGUAGE/TRANSLATE`), e `ISOMITTED` (só faz sentido com LAMBDA → move para
  F3). Documentar no README como **fora de escopo com justificativa** (engine de servidor, sem UI/serviços
  externos/locale CJK); a cobertura passa a ser medida também contra o catálogo viável (~485).
- **Release por onda**: cada onda fecha → merge na main → push (aval do usuário) → `gh workflow run
  release.yml` (versionize deriva o minor dos commits `feat`). Lockstep dos 2 pacotes continua.
- **Processo por função (TDD, lição do lessons.md)**: golden values de ORÁCULO, nunca de cabeça — Excel
  real/LibreOffice para math/text/datas (documentar a origem no teste), `ExcelFinancialFunctions` (já no
  csproj de testes) para as financeiras. RED → record + entrada no `Parser.Functions` + entrada no
  `FormulaWriter.Call` + ✅ no README → GREEN.
- **Convenções de implementação**: 1 função = 1 record `: Function` (padrão atual; uniformidade e
  serialização estável), mas ARQUIVOS agrupados por família (ex.: `Expressions/Trigonometry.cs` com vários
  records) para não criar 180 arquivos. `MemoryPackUnion` é append-only: próximo tag livre = 64, nunca
  reordenar. Coerção via `CoerceToNumber/...` existentes; erros Excel corretos (`#NUM!` para domínio,
  `#DIV/0!`, `#VALUE!`).

## A. Análise estrutural (proposta bem embasada + alternativas)

### A1. Funções voláteis (TODAY, NOW, RAND, RANDBETWEEN, INDIRECT) — ADIADO → fase F1
**Problema:** o cache por célula congela valores até `InvalidateCache` — `TODAY()` de ontem persistiria; e
não basta não-cachear a célula volátil: um DEPENDENTE cacheado (`B1=A1+1` com `A1=TODAY()`) congelaria do
mesmo jeito. Volatilidade precisa propagar pelos avaliadores transitivos.
**Proposta (quando F1 ativar):** flag thread-local "avaliação tocou volátil" no `Workbook` (mesmo padrão do
detector de ciclos `_evaluating`): `GetCellValue` seta a flag ao avaliar nó volátil (propriedade virtual
`IsVolatile` na `Expression`, default false) e, se a flag subiu durante a avaliação de uma célula, NÃO grava
essa célula no cache e propaga a flag ao chamador — O(1), local, sem grafo de dependências. Tempo vem de um
**`TimeProvider` injetável no `Workbook`** (default `TimeProvider.System`): TODAY/NOW testáveis e o host
controla o relógio (consistência intra-batch: o host congela o TimeProvider durante um batch). RAND idem com
seed injetável.
**Alternativas rejeitadas:** (a) semântica snapshot documentada (voláteis cacheiam até invalidar) — zero
custo, mas diverge do Excel e surpreende (data velha); (b) grafo de dependências com dirty-tracking —
correto e mais poderoso, porém uma reescrita do modelo de recálculo; desproporcional aqui, reavaliar se um
dia houver recálculo incremental.

### A2. Funções que retornam arrays (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TRANSPOSE, MMULT, MINVERSE,
MDETERM, MUNIT, FREQUENCY, TEXTSPLIT, TOCOL/TOROW, WRAP*, TAKE/DROP, CHOOSEROWS/COLS, H/VSTACK, EXPAND,
RANDARRAY, MODE.MULT, ARRAYTOTEXT, PERCENTOF, SEQUENCE) — FORA deste roadmap → fase F2
**Problema:** o modelo é 1 célula = 1 valor; `ComputedValue` não tem forma de array materializado (só
`Reference`, que aponta para células existentes).
**Proposta (F2):** novo `ComputedValueKind.Array` carregando `ComputedValue[,]` no campo `_ref` (mesma
técnica do `Reference`) + `EnumerateValues` cobrindo Array → agregações/lookups consomem de graça
(`SUM(FILTER(...))`, o caso de servidor). Célula cujo RESULTADO é array exibe o canto superior-esquerdo
(semântica pré-spill do Excel), sem spill físico.
**Alternativas rejeitadas:** (a) spill físico em células vizinhas (semântica moderna do Excel) — exige
células fantasma, invalidação de vizinhança e erro `#SPILL!`; só faz sentido DEPOIS do Kind.Array e de um
grafo de dependências; (b) materializar arrays como ranges temporários na sheet — mutaria dados do usuário
(gambiarra clássica, rejeitada).

### A3. LAMBDA + BYROW/BYCOL/MAP/REDUCE/SCAN/MAKEARRAY/ISOMITTED — FORA → fase F3
Exige **funções-como-valor** no `ComputedValue` (closure carregando `Expression` + escopo de nomes) — a
maior mudança do catálogo. Depende de F2 (essas funções produzem/consomem arrays). Sem análise adicional
até F2 concluir.

### A4. Distribuições estatísticas (NORM.*, T.*, CHISQ.*, F.*, GAMMA*, BETA*, BINOM*, POISSON, WEIBULL,
EXPON, LOGNORM, NEGBINOM, HYPGEOM, CONFIDENCE.*, CRITBINOM, FORECAST.ETS*, LINEST/LOGEST/TREND/GROWTH,
Z.TEST/T.TEST/F.TEST/CHISQ.TEST) — FORA → fase F4
**Problema:** exigem funções especiais (gamma incompleta regularizada, beta incompleta, erf, inversas por
busca numérica) com precisão validada. FORECAST.ETS é um modelo de séries temporais inteiro. LINEST exige
álgebra de mínimos quadrados + retorno em array (depende de F2).
**Proposta (F4):** implementar special functions internas (Lanczos para gamma, frações continuadas para as
incompletas, bisseção/Newton salvaguardado para inversas — mesmo padrão robusto que usamos no RATE/IRR),
validadas contra oráculo (Excel/R/scipy tabelados). Sem dependência externa (não há pacote .NET mantido que
cubra tudo; e "correto > remendo" vale para precisão também: testes de tolerância documentada).

### A5. SUBTOTAL/AGGREGATE — parcial na onda 4, com limite de modelo declarado
Os códigos 101-111 ("ignorar linhas ocultas") não têm semântica no MySheet: o modelo não tem linhas
ocultas. NÃO é gambiarra implementar 1-11 e 101-111 com o mesmo comportamento (é o limite do modelo de
dados, documentado); a parte real de SUBTOTAL — ignorar SUBTOTALs aninhados no range — será implementada
de verdade (detectar nós SUBTOTAL nas células do range). AGGREGATE fica para F2 (vários códigos exigem
array/ordenação com opções que dependem de LARGE/SMALL/PERCENTILE — reavaliar depois da onda 4).

### A6. Datas como número serial (onda 5) — decisão de representação
**Proposta:** datas SÃO doubles seriais (fiel ao Excel; zero mudança no ComputedValue). Helper interno
`DateSerial` no core: serial ↔ `DateTime` via OADate (base 30/12/1899, que compensa o bug do Excel de
tratar 1900 como bissexto). **Limite documentado:** seriais 1..59 (jan-fev/1900) divergem do Excel em 1
dia e o serial 60 (29/02/1900, data inexistente) não é representável — irrelevante na prática, registrado
no doc. `DATEDIF` implementa os 6 modos com os quirks documentados do Excel; `YEARFRAC` implementa as 5
bases de day-count (0=US 30/360, 1=actual/actual, 2=actual/360, 3=actual/365, 4=EU 30/360) — essa infra é
REUTILIZADA pelas financeiras de título da onda 6. NETWORKDAYS/WORKDAY aceitam range de feriados
(ExpandComputedValues já resolve).
**Alternativa rejeitada:** `ComputedValueKind.Date` dedicado — quebraria a equivalência data≡número do
Excel (SUM sobre datas, comparações), duplicaria coerções, e o interop xlsx já trafega serial.

### A7. Locale — contrato invariant
`FIND/SEARCH/SUBSTITUTE/PROPER/EXACT` usam ordinal/invariant (SEARCH case-insensitive + wildcards `*?~`
reutilizando `Criteria.WildcardMatch`). `DOLLAR/FIXED` formatam com invariant culture e símbolo `$` —
documentado (a lib não é locale-aware por design; `TEXT` já segue isso).

### A8. Escala do Parser/FormulaWriter/serialização
~180 records novos: tags MemoryPackUnion 64..~245 (ushort, sem risco de esgotar; append-only). O switch
`FormulaWriter.Call` e o mapa `Parser.Functions` crescem em paralelo — o teste exaustivo
`AllBuiltInFunctions_RoundTripStructurally` DEVE ser atualizado a cada onda (uma fórmula mínima por função
nova); ele é o guarda que impede função parseável sem un-parse. Nomes com ponto (`CEILING.MATH`,
`STDEV.S`, `RANK.EQ`) já são suportados pelo tokenizer (aceita `.` em identificadores).

## For Future Agents
Marque `- [x]`; ao fechar onda: Status `Complete` + Phase Summary + Verification + atualizar README
(✅ nas funções + contagem) + release (aval do usuário para push/dispatch). TDD por função com golden
values de oráculo citado no teste. Testes: core
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` (544 após a
Onda 4); Excel `.../Danfma.MySheet.Excel.Tests... -c Release` (16 hoje). Build da solução: 0 warnings
sempre. Pós-Fase R: nós de função novos nascem no namespace da categoria
(`Danfma.MySheet.Expressions.{Categoria}`, pasta espelhada); as tags do union continuam append-only.

---

## Phase 0: Preparação
Status: Complete

- [x] Merge `feature/shared-formulas` → `main` + push (feito em 2026-07-01; a documentação completa
      — README novo + docs/ — também entrou na main no mesmo push, CI verde).
- [x] Cobertura (agora em `docs/function-reference.md`, movida pelo trabalho de docs): seção
      "Out of scope (by design)" com as 35 exclusões permanentes e justificativa por grupo; funções
      marcadas ✖ (removidas das listas ⬜), contagem dupla no cabeçalho (52/~520 catálogo, ~485 viável),
      e nota de que custom functions são o escape hatch para supri-las.
- [x] Onda 1 delegada a sub-agente em worktree isolada (branch `feature/functions-wave-1`).

### Verification Plan
- CI verde na main; README renderiza com a nova seção; contagens conferem com o Parser.

### Phase Summary
Concluída em 2026-07-01. O merge da `feature/shared-formulas` e a documentação completa (README novo +
9 guias em `docs/`, gerados por sub-agente com exemplos compilados/verificados) entraram na main no mesmo
push (CI verde). A tabela de cobertura vive agora em `docs/function-reference.md`; ganhou legenda ✖,
seção "Out of scope (by design)" (35 funções, justificativa por grupo, custom functions como escape
hatch) e contagem dupla. Correção de honestidade: a estimativa "~85 exclusões" do plano estava errada —
recontagem real = 35 (catálogo viável ~485, não ~435); lição registrada em tasks/lessons.md. Convenção
nova de commits (inglês, curto+descrição, semantic) registrada e em vigor a partir daqui.

---

## Phase 1 (Onda 1 → 1.1.0): Math & Trigonometria escalar (~52 funções)
Status: Complete

Zero estrutura nova — tudo `CoerceToNumber` + `Math.*`. Agrupar em arquivos por família.

- [x] Potência/raiz/log: `SQRT` `POWER` `EXP` `LN` `LOG` (1-2 args) `LOG10` `SQRTPI`
- [x] Arredondamento: `ROUNDDOWN` `TRUNC` (1-2 args) `MROUND` `CEILING` `CEILING.MATH` `CEILING.PRECISE`
      `ISO.CEILING` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `EVEN` `ODD` (semânticas de sinal do Excel:
      CEILING com número negativo/significância negativa → casos de erro `#NUM!` conferidos no oráculo)
- [x] Aritmética: `MOD` (sinal do divisor, como Excel: `MOD(-3,2)=1`) `QUOTIENT` `SIGN` `PI` `PRODUCT`
      `SUMSQ` `MULTINOMIAL` `SERIESSUM`
- [x] Combinatória: `FACT` `FACTDOUBLE` `COMBIN` `COMBINA` `GCD` `LCM`
- [x] Trig: `SIN` `COS` `TAN` `COT` `SEC` `CSC` `ASIN` `ACOS` `ATAN` `ATAN2` (ordem de args do Excel:
      x,y!) `ACOT` + hiperbólicas `SINH` `COSH` `TANH` `COTH` `SECH` `CSCH` `ASINH` `ACOSH` `ATANH` `ACOTH`
      + `DEGREES` `RADIANS`
- [x] Bases: `BASE` `DECIMAL` `ROMAN` `ARABIC`
- [x] FormulaWriter.Call + corpus exaustivo + README ✅ (Math foi para 67/82 — contado por script,
      a estimativa "~59" do plano estava baixa)

### Verification Plan
- Suítes verdes; golden values citando oráculo (Excel/LibreOffice) nos casos não-triviais (MOD negativo,
  CEILING negativo, ATAN2, COMBINA, ROMAN); round-trip exaustivo do FormulaWriter cobre as novas.

### Phase Summary
Concluída em 2026-07-01 (branch `feature/functions-wave-1`, TDD RED→GREEN). **60 funções novas** (a
estimativa "~52" do título estava baixa; a lista explícita dos bullets sempre teve 60): 7 potência/log +
12 arredondamento + 8 aritmética + 6 combinatória + 23 trig + 4 bases. Total registrado no Parser: 112
(era 52); Math and Trigonometry 67/82. Arquivos por família em `Expressions/` (`PowerAndLog.cs`,
`Rounding.cs`, `Arithmetic.cs`, `Combinatorics.cs`, `Trigonometry.cs`, `NumberBases.cs`) + helper
`ExcelMath.cs`; tags MemoryPackUnion 64–123 (append-only; próximo livre = 124). Golden values: páginas
oficiais da Microsoft fetchadas em 2026-07-01 e citadas nos testes (MOD/CEILING/FLOOR/CEILING.MATH/
FLOOR.MATH/MROUND/EVEN/ODD/ATAN2/COMBIN(A)/GCD/LCM/FACT(DOUBLE)/MULTINOMIAL/SERIESSUM/BASE/DECIMAL/
ROMAN/ARABIC/COT/ACOT/ACOTH/LOG/POWER/SQRTPI/TRUNC/QUOTIENT/SIGN); `Math.*` do .NET como oráculo do trig
puro. Decisões: (1) `ExcelMath.Snap` (G14) reproduz o arredondamento cosmético do Excel nos quocientes —
sem isso MROUND(1.3,0.2)=1.2 em IEEE-754, contra 1.4 documentado; (2) **ROMAN só na forma clássica**
(0/TRUE/omitida) — formas concisas 1-4/FALSE devolvem `#VALUE!` porque o algoritmo não é verificável
contra oráculo além do único exemplo documentado (limitação registrada no doc e testada); (3) COMBINA não
impõe n≥k (a doc da MS declara a restrição, mas a fórmula COMBIN(n+k-1,k) é válida para k>n e o Excel
real aceita — casos documentados cobertos por teste); (4) limite documentado de 2^27 aplicado a
COT/SEC/CSC; overflow numérico (EXP/SINH/FACT) → `#NUM!`. Novos folds: `NumericListFold` (GCD/LCM/
MULTINOMIAL) reusando `NumericAggregation` (texto/lógicos referenciados ignorados, como SUM). Suíte core:
274 → **350 verdes** (76 novos: 74 casos em 6 arquivos de teste por família + 2 round-trips canônicos);
suíte Excel intacta (16); build 0 warnings. Corpus exaustivo do FormulaWriter cobre as 60.

---

## Phase 2 (Onda 2 → 1.2.0): Logical + Information + Text escalar (~36 funções)
Status: Complete

- [x] Logical: `TRUE` `FALSE` (funções, além dos literais) `XOR` `IFS` `SWITCH` (2 formas: com default)
- [x] Information: `NA` `ISERROR` `ISERR` `ISNA` `ISTEXT` `ISNONTEXT` `ISLOGICAL` `ISEVEN` `ISODD` `ISREF`
      `ISFORMULA` (via sheet do contexto) `N` `T` `TYPE` `ERROR.TYPE` `SHEETS`
- [x] Text: `RIGHT` `FIND` `SEARCH` (case-insensitive + wildcards `? * ~` — busca posicional própria, não
      o full-match do `Criteria.WildcardMatch`) `REPLACE`
      `SUBSTITUTE` (com instance_num) `REPT` `PROPER` `EXACT` `CHAR` `CODE` `UNICHAR` `UNICODE` `CLEAN`
      `FIXED` `DOLLAR` `NUMBERVALUE` `TEXTBEFORE` `TEXTAFTER` `VALUETOTEXT`
- [x] Regex (modernas, escalar): `REGEXTEST` `REGEXEXTRACT` `REGEXREPLACE` (System.Text.RegularExpressions,
      timeout defensivo)
- [x] FormulaWriter.Call + corpus + README ✅

### Verification Plan
- Suítes verdes; SEARCH/FIND diferenciados por caso e wildcard; SUBSTITUTE com instance_num contra oráculo.

### Phase Summary
Concluída em 2026-07-01 (branch `feature/functions-wave-2`, TDD RED→GREEN por família). **43 funções
novas** (contado por script no `Parser.Functions`; o "~36" do título subestimava — a lista explícita
sempre teve 43): 5 lógicas + 16 de informação + 13 de manipulação de texto + 6 de formatação/extração +
3 regex. Total registrado no Parser: 155 (era 112); Logical 12/19, Information 18/22, Text 34/49.
Arquivos por família em `Expressions/` (`LogicalFunctions.cs`, `InformationFunctions.cs`,
`TextManipulation.cs`, `TextFormatting.cs`, `RegexFunctions.cs`); tags MemoryPackUnion 124–166
(append-only; próximo livre = 167). Golden values: páginas oficiais da Microsoft fetchadas em 2026-07-01
e citadas nos testes (a página do FIND está aposentada no site — usado o archive da mesma URL; os GUIDs
corretos das páginas REGEX* diferem dos "canônicos" e estão citados nos testes). Decisões: (1) SEARCH
implementa busca POSICIONAL de wildcard via regex não-ancorada (leftmost match = posição 1-based), não
reuso do `Criteria.WildcardMatch` (que é full-match) — com escape `~` e timeout defensivo de 1s;
(2) IFS/SWITCH lazy como o If (só o ramo que casa avalia; SWITCH compara via `ValueCoercion.AreEqual`);
(3) XOR segue a doc: paridade de TRUEs, texto/blank em ranges ignorados, nenhum valor lógico →
`#VALUE!`; (4) CHAR/CODE usam code point Unicode (Latin-1 em 1-255) em vez da página ANSI do Windows —
contrato locale-invariant §A7, documentado no reference; (5) FIXED/DOLLAR invariant (`.`/`,`/`$`,
negativos do DOLLAR em parênteses, arredondamento away-from-zero em posição de dígito possivelmente
negativa); (6) TEXTBEFORE/TEXTAFTER com engine único (instance_num negativo conta do fim, match_mode,
match_end como delimitador virtual na borda, if_not_found lazy; split documentado `#VALUE!`/`#N/A`);
(7) REGEXEXTRACT só no modo escalar 0 — modos 1/2 retornam array → `#VALUE!` até a F2; ocorrência
inexistente no REGEXREPLACE devolve o texto inalterado (caso não documentado na página, decisão
registrada); (8) T propaga erros (norma do engine; a tabela da página só cobre não-erros). Suíte core:
350 → **408 verdes** (58 novos: 5 arquivos de teste por família + corpus do FormulaWriter cobrindo as
43); suíte Excel intacta (16); build 0 warnings. Docs: `function-reference.md` → 155 (tabelas
Logical/Information/Text reescritas, notas de limitação regex/CHAR/SHEETS) + contagem 155 espelhada no
README e nos guias em inglês (`docs/pt-BR/` intocado — espelho é de outro agente).

---

## Phase 3 (Onda 3 → 1.3.0): Lookup & Reference escalar (~9 funções)
Status: Complete

- [x] `CHOOSE` `HLOOKUP` (espelho do VLookup) `LOOKUP` (formas vetor e array) `COLUMN` `COLUMNS` `XMATCH`
      (modos de match/search como XLOOKUP) `ADDRESS` `AREAS` `FORMULATEXT` (reusa `FormulaWriter` — só
      funciona porque a onda 0 do un-parser existe)
- [x] `INDIRECT` NÃO entra (volátil — F1).
- [x] FormulaWriter.Call + corpus + README ✅

### Verification Plan
- Suítes verdes; HLOOKUP/LOOKUP/XMATCH com os mesmos casos de borda dos testes do VLOOKUP/XLOOKUP
  (approximate text keys — regressão do bug 3460bb3).

### Phase Summary
Concluída em 2026-07-02 (branch `feature/functions-wave-3`, TDD RED→GREEN pela família). **9 funções
novas** (contado por script no `Parser.Functions`: 155 → 164; Lookup and Reference 7/40 → 16/40):
CHOOSE, HLOOKUP, LOOKUP (vetor + array), COLUMN, COLUMNS, XMATCH, ADDRESS, AREAS e FORMULATEXT.
Arquivo `Expressions/LookupFunctions.cs` + helper compartilhado `Expressions/LookupMatching.cs`
(FindMatch/Wildcard/Closest EXTRAÍDOS do XLookup — refactor protegido pelos testes existentes do
XLOOKUP; XLOOKUP, XMATCH e LOOKUP usam o mesmo engine de match); tags MemoryPackUnion 167–175
(append-only; próximo livre = 176). Golden values: páginas oficiais da Microsoft fetchadas em
2026-07-02 e citadas nos testes (CHOOSE/HLOOKUP/LOOKUP vetor/ADDRESS/XMATCH/COLUMN/COLUMNS/AREAS/
FORMULATEXT); HLOOKUP("B",...,TRUE)→5 e XMATCH("Gra",...,1)→2 cobrem chave TEXTO com match
aproximado (regressão 3460bb3). Decisões: (1) CHOOSE lazy como IF (só o argumento escolhido avalia —
testado com função custom que lança) e um range escolhido vira `ComputedValue.Reference` para
consumidores range-aware (`SUM(CHOOSE(…))`), mesma técnica do OFFSET; a forma `A2:CHOOSE(…)` da
página não parseia (o `:` exige células) — fora de escopo; (2) HLOOKUP segue o contrato de erro da
doc (row_index < 1 → `#VALUE!`, além da tabela → `#REF!`) — nota: o VLookup existente devolve
`#REF!` nos dois casos, divergência pré-existente não tocada; (3) LOOKUP array form implementa a
regra documentada (mais largo que alto → busca 1ª LINHA/retorna última LINHA; senão 1ª COLUNA/última
COLUNA) — a página aposentou os exemplos da forma array, casos derivados mecanicamente da regra
citada; (4) ADDRESS com a1=FALSE implementa APENAS a forma absoluta documentada `R2C3`; as formas
R1C1 relativas (`R2C[3]`) → `#VALUE!` (limitação declarada no doc); quoting do sheet_text reusa
`FormulaWriter.IsSimpleSheetName` (tornado internal); (5) AREAS é checagem sintática do nó
(UnionReference conta áreas, recursivo; referência → 1; não-referência → `#VALUE!`); o operador de
interseção por espaço (`AREAS(B2:D4 B2)`) não existe no parser — fora de escopo; (6) FORMULATEXT
un-parseia com o contexto de sheet da célula REFERENCIADA (referências locais sem qualificação);
binários XMATCH search_mode 2/-2 degradam para linear (mesmo comportamento do XLOOKUP). Suíte core:
408 → **440 verdes** (32 novos casos em `LookupReferenceFunctionTests` + corpus do FormulaWriter
cobrindo as 9); suíte Excel intacta (16); build 0 warnings. Docs: `function-reference.md` → 164
(tabela Lookup reescrita com 16 assinaturas, coverage 16/40) + contagem 164 espelhada no README e
nos guias em inglês (`docs/pt-BR/` intocado — espelho é de outro agente).

---

## Phase R: Reorganização dos nós da AST em namespaces semânticos (breaking → 2.0.0)
Status: Complete
<!-- Pedido do usuário em 2026-07-02: "separar os nós da nossa AST em namespaces semânticos, por
categoria — o namespace está ficando muito inflado". EXECUTAR logo após a Onda 3 integrar e ANTES da
Onda 4 (menor churn acumulado; adoção do 1.x ainda mínima). As ondas 4-6 viram 2.1.0/2.2.0/2.3.0. -->

Proposta de forma (fazer certo UMA vez — decisões a validar com o usuário no gate de execução):
- `Danfma.MySheet` (raiz): + `ComputedValue`, `ComputedValueKind`, `Error` — são os tipos de RESULTADO da
  API pública (par de `Workbook`/`Sheet`), não "expressões". **APROVADO pelo usuário em 2026-07-02.**
- `Danfma.MySheet.Expressions` (núcleo enxuto): `Expression`, `ValueExpression`, `Function`, `Reference`,
  `FunctionCall`, `EvaluationContext`, nós-valor (`NumberValue`/`StringValue`/`BooleanValue`/`BlankValue`/
  `ErrorValue`), referências (`CellReference`/`RangeReference`/`UnionReference`/`NameReference`),
  operações (`BinaryOperation`/`UnaryOperation`) e helpers internal.
- `Danfma.MySheet.Expressions.{Logical|Mathematics|Statistical|Text|Information|Lookup|Financial|Dates}`:
  os records de função DIRETO sob Expressions (decisão do usuário em 2026-07-02: o segmento `.Functions.`
  é redundante — as categorias só conterão funções). `Mathematics` (não `Math`) porque um segmento `Math`
  faria o identificador `Math` resolver para o namespace dentro daqueles arquivos, quebrando `Math.*`
  (mesma razão de `Dates`, não `DateTime`). Categorias espelham as seções publicadas do
  function-reference (Excel): SUM/SUMIF(S)/PRODUCT → Mathematics; AVERAGE/COUNT*/MAX/MIN → Statistical.
  Pastas espelham namespaces. `FunctionCall` fica no núcleo (é o nó de extensibilidade, não categoria).

- [x] Mover os records de função para os namespaces por categoria (pastas correspondentes); atualizar
      usings de `Parser`, `FormulaWriter`, testes (preferir `GlobalUsings` nos csproj de teste).
- [x] Decisão do usuário: mover `ComputedValue`/`Error` para a raiz `Danfma.MySheet` — SIM (2026-07-02).
- [x] **Teste de compat binária MemoryPack**: ANTES do refactor, serializar um workbook representativo
      (fórmulas de todas as categorias) e commitar o blob como fixture; DEPOIS, teste que o deserializa e
      reavalia — prova que o union por tags sobrevive à mudança de namespace. (Tags não mudam: append-only
      continua valendo.)
- [x] Docs: guia de migração 1.x → 2.0 (tabela de namespaces), atualização dos guias EN (o refresh
      pt-BR fica com o agente do espelho, como nas ondas — `docs/pt-BR/` não foi tocado).
- [x] Atualizar os briefings das ondas seguintes: nós novos nascem no namespace da categoria.
- [ ] Commit `refactor!:` (BREAKING CHANGE) → release **2.0.0**.

### Verification Plan
- Suítes completas verdes (core + Excel) + teste de compat binária com a fixture pré-refactor.
- Build 0 warnings; `versionize --dry-run` propõe major.

### Phase Summary
Executada em 2026-07-02 (branch `refactor/ast-namespaces`, worktree isolada; 3 commits: fixture →
`refactor!:` → docs). **Moves**: `ComputedValue`/`ComputedValueKind`/`Error` para a raiz
`Danfma.MySheet`; **164 records de função** para 7 namespaces por categoria do function-reference —
Logical 12, Mathematics 67, Statistical 8, Text 34, Information 18, Lookup 16, Financial 9 — com
pastas espelhando namespaces; núcleo `Danfma.MySheet.Expressions` mantém os nós estruturais e TODOS os
helpers internal. Julgamento de categoria: **T saiu de `InformationFunctions.cs` para
`Expressions/Text/T.cs`** — o function-reference (espelho do Excel) lista T em Text, e a seção do doc
é o árbitro. **Nenhuma tag `MemoryPackUnion` mudou**: a fixture `workbook-pre-namespaces.msgpack.bin`
(gerada no layout 1.x, commit próprio ANTES do refactor) deserializa e reavalia 14 células idênticas
nos dois layouts — suíte core **442/442 antes e depois** (a onda 3 fechou em 440, mas o fix do VLOOKUP
`0fb6af5` somou 1 → 441; +1 do guard de compat = 442), Excel **16/16**, build 0 warnings. Colisões
resolvidas: `Text`/`Lookup` são namespace E record homônimo → qualificação `Text.Text`/`Lookup.Lookup`
dentro de `Expressions` (nas attrs do union), e `Lookup.Index` vs `System.Index` qualificado no
Parser/FormulaWriter; `GlobalUsings.cs` por projeto de teste (no de Excel, só `Mathematics` — importar
tudo colidiria com `Row`/`Column`/`Text` do OpenXML). Docs EN: `docs/migrating-to-2.0.md` (tabela
tipo→namespace, before/after de usings, garantia de compat binária) + links no README/docs/README +
snippets ajustados (computed-value, workbook-and-expressions, serialization); `docs/pt-BR/` intocado —
refresh delegado ao agente do espelho. Release 2.0.0 pendente de merge + aval do usuário (checkbox
aberto); ondas 4–6 renumeradas para 2.1.0/2.2.0/2.3.0.

---

## Phase 4 (Onda 4 → 2.1.0): Condicionais, SUMPRODUCT, estatística descritiva + aliases (~55 funções)
Status: Complete

Namespaces (pós-Fase R): records novos nascem em `Expressions.Statistical` (condicionais AVERAGEIF*/
MAXIFS/MINIFS, variantes A, ordem/posição, dispersão, bivariadas, escalares e aliases) e
`Expressions.Mathematics` (SUMPRODUCT/SUMX*, SUBTOTAL — categorias Math do Excel).

- [x] Condicionais (Criteria existente): `AVERAGEIF` `AVERAGEIFS` `MAXIFS` `MINIFS`
- [x] Variantes A: `AVERAGEA` `MAXA` `MINA`
- [x] Pairwise multi-range (helper novo `PairwiseRanges` em NumericAggregation): `SUMPRODUCT` `SUMX2MY2`
      `SUMX2PY2` `SUMXMY2` (ranges de shapes diferentes → `#VALUE!` no SUMPRODUCT; `#N/A` nos SUMX*,
      como documentado nas páginas oficiais)
- [x] `SUBTOTAL` (códigos 1-11 = 101-111; ignora SUBTOTAL aninhado de verdade; sem hidden rows — limite de
      modelo documentado, ver §A5)
- [x] Ordem/posição (helper de sorting sobre `List<double>`): `MEDIAN` `MODE.SNGL` `LARGE` `SMALL`
      `RANK.EQ` `RANK.AVG` `PERCENTILE.INC` `PERCENTILE.EXC` `PERCENTRANK.INC` `PERCENTRANK.EXC`
      `QUARTILE.INC` `QUARTILE.EXC` `TRIMMEAN`
- [x] Dispersão/momentos: `STDEV.S` `STDEV.P` `STDEVA` `STDEVPA` `VAR.S` `VAR.P` `VARA` `VARPA` `AVEDEV`
      `DEVSQ` `GEOMEAN` `HARMEAN` `SKEW` `SKEW.P` `KURT` `STANDARDIZE`
- [x] Bivariadas (pairwise): `CORREL` `PEARSON` `COVARIANCE.P` `COVARIANCE.S` `RSQ` `SLOPE` `INTERCEPT`
      `STEYX` `FORECAST.LINEAR`
- [x] Escalares simples: `FISHER` `FISHERINV` `PHI` `PERMUT` `PERMUTATIONA` `PROB`
      (⚠️ **`GAUSS` movido para F4**, erro do plano corrigido em 2026-07-02: `GAUSS(z) = Φ(z) − 0,5`
      exige a CDF normal — erf —, que pertence à fase de distribuições; `PHI` é só a densidade e ficou)
- [x] Aliases Compatibility dos implementados (**records DISTINTOS**, não a mesma factory — o un-parse
      preserva o nome legado que o usuário escreveu; lógica compartilhada via `Compute` internal):
      `MODE` `STDEV` `STDEVP` `VAR` `VARP` `RANK` `PERCENTILE` `PERCENTRANK` `QUARTILE` `COVAR`
      `FORECAST` (⚠️ `CONFIDENCE`/`CRITBINOM` e os *DIST/*INV ficam para F4 — dependem de distribuições)
- [x] FormulaWriter.Call + corpus + README ✅

### Verification Plan
- Suítes verdes; PERCENTILE.INC/EXC e QUARTIS contra oráculo (interpolações diferem entre ferramentas —
  fixar Excel); CORREL/SLOPE contra oráculo com dataset fixo; SUMPRODUCT shape-mismatch → `#VALUE!`.

### Phase Summary
Concluída em 2026-07-02 (branch `feature/functions-wave-4`, TDD RED→GREEN por família: os testes das
6 famílias entraram primeiro — 78 casos falhando com as funções fora do Parser — e o registro
Parser+FormulaWriter virou tudo verde de primeira). **67 funções novas** (contado por script no
`Parser.Functions`: 164 → **231**): 4 condicionais + 3 variantes A + 5 Math (SUMPRODUCT/SUMX*/
SUBTOTAL) + 13 ordem/posição + 16 dispersão/momentos + 9 bivariadas + 6 escalares + 11 aliases
Compatibility. `GAUSS` FICOU FORA (movido para F4 — exige erf/CDF normal; anotado no checkbox).
Arquivos por família: `Expressions/Statistical/{ConditionalAggregates,AVariants,OrderStatistics,
Dispersion,Bivariate,StatisticalScalars}.cs`, `Expressions/Mathematics/{SumProducts,Subtotal}.cs` e
o namespace NOVO `Expressions/Compatibility/CompatibilityAliases.cs` (records distintos que delegam
aos `Compute` internal das formas modernas — STDEV(...) nunca vira STDEV.S(...) no un-parse/
FORMULATEXT, testado). Helpers internal novos no núcleo: `PairwiseRanges` (pares posicionais com
política ignorar-par vs zero — bivariadas/SUMX* vs SUMPRODUCT), `StatisticsMath` (coleta em ordem
de varredura p/ MODE, interpolações PERCENTILE.INC em k(n−1) e .EXC em k(n+1), variâncias
amostral/populacional reusadas por STDEV*/VAR*/SUBTOTAL) e `NumericAggregation.FoldA` (regra A:
texto referenciado → 0, lógico → 1/0). SUBTOTAL implementa o nested-subtotal REAL (varre célula a
célula e pula as cujo Expression raiz é nó Subtotal, inclusive via referência de OFFSET/CHOOSE);
101-111 = 1-11 sem hidden rows (limite de modelo documentado no reference). Tags MemoryPackUnion
176–242 (append-only; próximo livre = 243; fixture de compat binária intocada e verde). Golden
values: 40+ páginas oficiais da Microsoft fetchadas em 2026-07-02 por 4 sub-agentes de pesquisa e
citadas por função nos testes (PERCENTILE/QUARTILE/PERCENTRANK com truncagem, RANK com empates,
SKEW/SKEW.P/KURT, CORREL/COVARIANCE/RSQ/SLOPE/INTERCEPT/STEYX/FORECAST.LINEAR com os datasets das
páginas, TRIMMEAN, SUBTOTAL 303/75.75, SUMX* −55/521/79, AVERAGEIF(S)/MAXIFS/MINIFS). Documentado
nos testes: 2 exemplos oficiais internamente inconsistentes foram evitados (AVERAGEIFS ex.1 imprime
75 vs 80.5 na própria descrição; AVERAGEA fórmula 2 contradiz a regra de célula vazia) e o exemplo
do SUMPRODUCT (dataset só em imagem) foi coberto pela EQUIVALÊNCIA documentada com a forma longa.
Correção de doc pré-existente: `FLOOR` estava ⬜ na categoria Compatibility do function-reference
apesar de implementado desde a onda 1 → ✅ (Compatibility 1 → 13/41). Suíte core: 442 → **544
verdes** (102 novos em 6 arquivos de teste + corpus do FormulaWriter cobrindo as 67); suíte Excel
intacta (16); build 0 warnings. Docs: `function-reference.md` → 231 (Statistical 60/111 com tabela
de 59 assinaturas, Math 72/82, seção nova "Compatibility — legacy aliases", nota GAUSS→F4) +
contagem 231 espelhada no README (`docs/pt-BR/` intocado — espelho é de outro agente).

---

## Phase N: Named ranges (prioridade do usuário, 2026-07-02 → 2.2.0)
Status: Complete
<!-- Avaliação: médio-pequeno. Parser/un-parser já produzem/imprimem NameReference (mecanismo do LET);
agregações consomem nome→range de graça via ComputedValue.Reference/EnumerateValues (mecanismo do OFFSET).
Valor alto: xlsx real com definedNames hoje avalia #NAME? silencioso. -->

- [x] `Workbook.DefinedNames` (case-insensitive, nome→`Expression`) + `DefineName(name, Expression)` e
      conveniência `DefineName(name, formulaText)` (refs DEVEM ser sheet-qualified — ArgumentException se
      não). **Membro MemoryPack APPENDED AO FIM** — a fixture `workbook-pre-namespaces.msgpack.bin`
      continuou verde (arquivos 1.x/2.0 abrem com DefinedNames vazio); comparer restaurado + null tratado
      em `[MemoryPackOnDeserialized]`.
- [x] `NameReference.Evaluate`: escopo LET primeiro (shadowing), depois DefinedNames — avaliação com
      guarda de ciclo nome→nome (thread-local `NamedReferences._resolving`, padrão do detector de células;
      ciclo → `#REF!`).
- [x] Helper interno de resolução (`NamedReferences.TryResolveReference`, unwrap `NameReference`→`Reference`
      com a mesma guarda) para as funções que exigem nó de referência SINTÁTICO: `VLOOKUP`/`HLOOKUP`
      (table), `INDEX`, `OFFSET`, `ROWS`, `COLUMNS`, `AREAS`, `ISREF` — cada uma testada com named range.
- [x] Interop: `ExcelFile.Load` lê `definedNames` workbook-scope (pula `LocalSheetId` e builtin `_xlnm.*`
      — limite documentado); `SaveAsExcel` escreve DefinedNames (refersTo totalmente qualificado —
      `ToFormula` com contexto vazio qualifica tudo).
- [x] Testes: SUM/VLOOKUP por nome, shadowing do LET, ciclo, case-insensitive, round-trip MemoryPack novo
      + fixture antiga, round-trip xlsx com oráculo ClosedXML.
- [x] Docs EN (workbook-and-expressions + excel-interop + README) — pt-BR no refresh do release.

### Verification Plan
- Suítes completas verdes (incl. fixture binária antiga); build 0 warnings; round-trip xlsx de nomes.

### Phase Summary
Concluída em 2026-07-02 (branch `feature/named-ranges`, TDD RED→GREEN). **Nenhum nó de AST novo** — o
parser/un-parser já produziam/imprimiam `NameReference` (mecanismo do LET), confirmado por teste de
round-trip; **nenhuma tag MemoryPackUnion nova** (próximo livre segue 243). Núcleo: `Workbook.DefinedNames`
(`Dictionary<string, Expression>`, case-insensitive) é o **último membro serializado** do Workbook
(append-only de schema) — a fixture `workbook-pre-namespaces.msgpack.bin` abriu verde ANTES e DEPOIS da
mudança (passou nos dois momentos exigidos pelo briefing); comparer restaurado e `null` da fixture antiga
tratado em `RestoreComparers` (`[MemoryPackOnDeserialized]`). API `DefineName(name, Expression)` e a
conveniência `DefineName(name, formulaText)` (prepende `=` opcional; rejeita referência não-qualificada
via walk com sentinela de sheet `""`, reusando `FormulaWriter.Call` — internalizado — para varrer args de
função; valida nome via `NamedReferences.ValidateName`, reusando `Parser.IsCellReference` internalizado).
Helper novo `Expressions/NamedReferences.cs`: `EvaluateDefinition` (range/union → `ComputedValue.Reference`
para consumidores range-aware; cell/constante/formula → valor escalar) e `TryResolveReference` (unwrap
sintático recursivo name→…→Reference), ambos com a MESMA guarda ThreadStatic `_resolving`
(HashSet<string> OrdinalIgnoreCase, padrão do `_evaluating`; ciclo → `#REF!`/não-resolvível, sem overflow).
`NameReference.Evaluate` resolve LET→DefinedNames→`#NAME?`. Funções que exigem referência sintática passam
a resolver nomes: VLOOKUP/HLOOKUP (table), INDEX, OFFSET (base), ROWS, COLUMNS, AREAS, ISREF — cada uma com
teste de named range. Decisões de design além do briefing: (1) `TryResolveReference` checa o escopo LET
ANTES dos DefinedNames também na via sintática, então um binding LET reference-kind (ex.: range capturado)
resolve e um binding escalar sombreia para não-referência (VLOOKUP → `#REF!`) — shadowing consistente entre
as duas vias; (2) `DefineName(name, Expression)` NÃO exige qualificação (é a via de baixo nível usada pelo
interop e por constantes); só a conveniência string exige — assim `DefineName("Taxa", new NumberValue(.1))`
funciona; (3) interop de Load pula nome cujo `refersTo` não parseia (catch `ParseException`/`ArgumentException`)
em vez de falhar a carga inteira (limite documentado). Interop: `ExcelFile.Load` lê definedNames
workbook-scope (pula `LocalSheetId` e `_xlnm.*`), `SaveAsExcel` escreve `<definedNames>` após `<sheets>`
(ordem de schema) com refersTo por `ToFormula("")` (contexto vazio → tudo qualificado), em ambos os
FormulaMode. Round-trip validado com ClosedXML como oráculo dos dois lados. Suíte core: 544 → **566 verdes**
(22 novos: 21 em `NamedRangeTests` + 1 em `WorkbookSaveLoadTests`); suíte Excel: 16 → **20 verdes** (4 em
`NamedRangeInteropTests`); build da solução 0 warnings. Docs EN: seção "Named ranges" em
`docs/workbook-and-expressions.md`, linha na tabela de mapeamento + bullets de export + limite em
`docs/excel-interop.md`, bullet no README (`docs/pt-BR/` intocado — refresh do release). Release 2.2.0
pendente de merge + aval do usuário; sem push.

---

## Phase 5 (Onda 5 → 2.3.0): Datas e horas — 23/25 (sem TODAY/NOW) 
Status: Complete

Namespaces (pós-Fase R): records novos nascem em `Expressions.Dates` (pasta `Expressions/Dates/`;
`Dates`, não `DateTime` — colisão com `System.DateTime`); o helper `DateSerial` é internal e fica no
núcleo `Expressions`, como os demais helpers.

- [x] Infra `DateSerial` no core (internal): serial ↔ DateTime via OADate; clamp/`#NUM!` para serial
      negativo; fração do dia = hora. Documentar limite 1900 (§A6).
- [x] Construção/extração: `DATE` (com overflow de mês/dia como Excel: `DATE(2020,13,1)`→jan/2021)
      `DATEVALUE` `TIMEVALUE` (parse invariant) `TIME` `YEAR` `MONTH` `DAY` `HOUR` `MINUTE` `SECOND`
- [x] Aritmética de calendário: `DAYS` `DAYS360` (métodos US/EU) `EDATE` `EOMONTH` `WEEKDAY` (return_type
      1/2/3/11-17) `WEEKNUM` `ISOWEEKNUM` `DATEDIF` (6 modos, quirks documentados) `YEARFRAC` (5 bases —
      infra reutilizada pela onda 6)
- [x] Dias úteis (feriados via range): `NETWORKDAYS` `NETWORKDAYS.INTL` `WORKDAY` `WORKDAY.INTL`
      (máscaras de fim de semana 1-17 e "0000011")
- [x] `TODAY`/`NOW` ficam para F1 (voláteis) — registrado ⬜ com nota no README e no reference.
- [x] FormulaWriter.Call + corpus + README ✅ (Date and Time 23/25)

### Verification Plan
- Suítes verdes; DATEDIF/YEARFRAC/DAYS360/WEEKNUM contra oráculo (planilha de golden values citada);
  round-trip xlsx: data serial exportada/lida sem perda.

### Phase Summary
Concluída em 2026-07-02 (branch `feature/functions-wave-5`, TDD por família com golden values de oráculo).
**23 funções novas** (contado por script no `Parser.Functions`: 231 → **254**): 10 construção/extração
(DATE, TIME, DATEVALUE, TIMEVALUE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND) + 9 aritmética de calendário
(DAYS, DAYS360, EDATE, EOMONTH, WEEKDAY, WEEKNUM, ISOWEEKNUM, DATEDIF, YEARFRAC) + 4 dias úteis
(NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL). `TODAY`/`NOW` FORA (voláteis → F1; anotados ⬜
"deferred: volatile" no reference e README). Namespace/pasta NOVOS `Expressions/Dates/` (4 arquivos por
família: `DateConstruction.cs`, `DateComponents.cs`, `CalendarArithmetic.cs`, `WorkdayFunctions.cs` +
`DateTextParser.cs` para o parse invariant do DATEVALUE/TIMEVALUE). Tags MemoryPackUnion 243–265
(append-only; próximo livre = 266; fixture `workbook-pre-namespaces.msgpack.bin` intocada e verde).

**Representação (§A6, decisão fechada):** datas são doubles seriais — zero mudança no `ComputedValue`.
Dois helpers internal NOVOS no núcleo `Expressions`: **`DateSerial`** (serial ↔ DateTime via OADate, base
1899-12-30; `FromComponents` com o overflow do Excel; `TimeOfDaySeconds` arredondando ao segundo mais
próximo) e **`DayCount`** (as 5 bases do YEARFRAC: 0=US/NASD 30/360, 1=actual/actual com denominador =
média do comprimento dos anos-calendário cruzados per MS-OI29500 nota d, 2=actual/360, 3=actual/365,
4=EU 30/360). **`DayCount` é o helper que a ONDA 6 reutiliza** para as financeiras de título (ACCRINT,
PRICE, YIELD, COUP*, …) — vive no núcleo `Danfma.MySheet/Expressions/DayCount.cs`, testável isoladamente,
com `Nasd360Days`/`Euro360Days`/`ActualDays`/`YearFraction(start, end, basis)`. DAYS360 (a função)
implementa a variante US do *DAYS360* (regra do fim-de-mês que rola o fim para o dia 1 do mês seguinte)
inline — é distinta da 30/360 do YEARFRAC basis 0, por isso não compartilha com o DayCount; a europeia
reusa `DayCount.Euro360Days`.

**Limitação documentada (§A6):** seriais 1..59 renderizam 1 dia atrás do Excel (serial 1 = 1899-12-31
aqui vs 1900-01-01 no Excel) e o serial 60 (29-02-1900 fictício do Excel) não é representável — mapeia
para 28-02-1900, colidindo com o 59; a partir do serial 61 (01-03-1900) o mapeamento é exato. Registrado
no doc, no XML do `DateSerial` e num teste-nota (`Serial60_Is1900LeapYearLimitation`).

**Decisões além do briefing:** (1) **Coerção de texto-data**: as funções de data aceitam APENAS serial
numérico (mais texto numérico que o `CoerceToNumber` já parseia); NÃO fazem DATEVALUE implícito — então
`YEAR("2021-01-01")`→`#VALUE!` (divergência do Excel, registrada no reference). Só `DATEVALUE`/`TIMEVALUE`
parseiam string, com formatos invariant fixos e SEM formatos year-less (evita depender do relógio, que é
F1). (2) HOUR/MINUTE/SECOND arredondam ao segundo mais próximo (não truncam) — reproduz o Excel
(`SECOND(TIME(10,30,45))`=45 apesar do ruído IEEE-754). (3) WEEKNUM tipo 21 e ISOWEEKNUM usam
`System.Globalization.ISOWeek`. (4) Máscara de fim de semana: string 7-char "1111111" → NETWORKDAYS.INTL
devolve 0, WORKDAY.INTL devolve `#NUM!` (sem dia útil para pousar); número de weekend inválido → `#NUM!`,
string malformada → `#VALUE!` (conferido contra as páginas .INTL). (5) DATEDIF "MD" reproduz o
comportamento buggy do Excel (a MS avisa que é não-confiável) sem "consertar".

**Golden values de oráculo:** todas as páginas oficiais da Microsoft fetchadas por 3 sub-agentes em
2026-07-02 e citadas por teste (DATE overflow, TIME/HOUR/MINUTE/SECOND, DATEVALUE serials 40777/40685/
40597, TIMEVALUE, DAYS, DAYS360 US 1/360/30, EDATE/EOMONTH, WEEKDAY todos os return_types contra a
quinta-feira 14/02/2008, WEEKNUM 10/11 + ISO 10 contra 09/03/2012, DATEDIF Y=2/D=440/YD=75, YEARFRAC
bases 0/1/3 = 0.58055556/0.57650273/0.57808219, NETWORKDAYS 110/109/107, NETWORKDAYS.INTL 22/-21/22/20 e
máscara string, WORKDAY 30/04/2009 e 05/05/2009, WORKDAY.INTL serial 41013). Onde a página é silente
(DAYS360 europeu, EDATE clamp, DATEDIF M/YM/MD, YEARFRAC bases 2/4, ISO viradas de ano) os valores foram
DERIVADOS da regra documentada e marcados nos comentários dos testes. Suíte core: 566 → **601 verdes**
(35 novos: 14 em `DateConstructionTests`, 13 em `DateCalendarTests`, 8 em `DateWorkdayTests` + as 23
fórmulas novas no corpus do FormulaWriter); suíte Excel intacta (20); build da solução 0 warnings. Docs:
`function-reference.md` → 254 (seção "Date and time" nova com 23 assinaturas, Date and Time 23/25 com
TODAY/NOW ⬜ "deferred: volatile") + contagem 254 espelhada no README (`docs/pt-BR/` intocado — refresh
do release). Release 2.3.0 pendente de merge + aval do usuário; sem push.

---

## Phase 6 (Onda 6 → 2.4.0): Financeiras restantes viáveis (~40 funções)
Status: Complete

Namespaces (pós-Fase R): records novos nascem em `Expressions.Financial`.

Oráculo: `ExcelFinancialFunctions` (já referenciado nos testes) cobre a maior parte — validar cada golden.

- [x] Depreciação: `SLN` `SYD` `DB` `DDB` `VDB` `AMORLINC` `AMORDEGRC` (usam YEARFRAC/datas da onda 5)
- [x] Taxas/valor: `EFFECT` `NOMINAL` `MIRR` `RRI` `PDURATION` `ISPMT` `CUMIPMT` `CUMPRINC` `FVSCHEDULE`
      `DOLLARDE` `DOLLARFR`
- [x] Fluxos com datas: `XNPV` `XIRR` (solver bracketing+bisseção existente do IRR/RATE)
- [x] Títulos (day-count da onda 5): `ACCRINT` `ACCRINTM` `DISC` `DURATION` `MDURATION` `INTRATE` `PRICE`
      `PRICEDISC` `PRICEMAT` `RECEIVED` `YIELD` `YIELDDISC` `YIELDMAT` `TBILLEQ` `TBILLPRICE` `TBILLYIELD`
      `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD`
- [x] `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD`: **ENTRARAM** — o oráculo cobre com precisão
      verificável (fuzz vs oracle: ODDFPRICE/ODDLPRICE/ODDLYIELD exatos, ODDFYIELD ~7e-10 pelo solver).
- [x] FormulaWriter.Call + corpus + README ✅ (Financial **55/55** — categoria COMPLETA)

### Verification Plan
- Suítes verdes; TODAS as financeiras com golden do `ExcelFinancialFunctions` (incluindo o caso stiff de
  mortgage 30 anos — lição do RATE); XIRR com fluxo irregular de datas reais.

### Phase Summary
Concluída em 2026-07-02 (branch `feature/functions-wave-6`, TDD contra o oráculo `ExcelFinancialFunctions`
3.2.0). **46 funções novas** (contado por script no `Parser.Functions`: 254 → **300**): 7 depreciação +
11 taxas/valor + 2 fluxos com datas + 12 títulos simples/desconto/T-bill + 6 agenda de cupom + 4
título periódico + **4 ODD** (todas entraram). **Financial 9 → 55/55 — categoria completa.**

**Descoberta central (registrada em tasks/lessons.md):** o oráculo NÃO implementa a fórmula-textbook de
PRICE (`E = COUPDAYS`); ele usa `dsc = e − a` com `a = DaysBetween(pcd, settlement)` e `e = CoupDays`, e
day-counts de título com uma variante `ModifyStartDate`/`ModifyBothDates` do 30/360-US e um "days in year"
actual/actual próprios que DIVERGEM do `DayCount` do YEARFRAC (onda 5) em casos de fim-de-mês/fevereiro.
Por isso o wave-6 NÃO reusa `DayCount.Nasd360Days` para os títulos: um helper novo **`BondMath`**
(`Expressions/Financial/BondMath.cs`) porta fielmente as convenções do Excel (reusando `DayCount.ActualDays`
/`Euro360Days` onde são idênticos), validado **valor-a-valor por fuzz** contra o oráculo (dezenas de
milhares de casos por função, 5 bases × 3 frequências) ANTES de portar para o codebase — todas exatas
(closed-form maxErr=0; solver-based ~1e-9). As agendas de cupom ANDAM PARA TRÁS a partir do maturity
iterativamente (clamp de fevereiro é "sticky"; fim-de-mês forçado só quando o maturity é fim-de-mês).

Arquivos por família em `Expressions/Financial/`: `Depreciation.cs`, `RatesAndValue.cs`,
`DatedCashFlows.cs`, `CouponSchedule.cs`, `Bonds.cs`, `BondsSimple.cs`, `OddBonds.cs`, mais o helper
`BondMath.cs` e o coercer compartilhado `FinancialArguments.cs` (datas como serial truncado; frequência
1/2/4 e basis 0-4 → `#NUM!`; settlement ≥ maturity → `#NUM!`). Tags MemoryPackUnion **266–311**
(append-only; próximo livre = 312; fixture binária intocada e verde). Solver de YIELD/XIRR/ODDFYIELD =
`TimeValueOfMoney.Solve` (bracketing+bisseção da onda financeira original), com um caso stiff de 35 anos
irregular no XIRR. XNPV/XIRR: primeira data é a âncora, datas fora de ordem permitidas, data antes da
âncora → `#NUM!`, sem mudança de sinal no XIRR → `#NUM!`. Um bug pego pelos testes (não pelo fuzz): o
`DatedFlows.Read` lia `arg[0]/arg[1]` mas no XNPV os values/dates ficam em `arg[1]/arg[2]` — corrigido com
índice paramétrico. Preconditions do ACCRINT seguem o oráculo (`first_interest >= settlement`); ODDFPRICE/
ODDFYIELD exigem o first_coupon alinhado com a agenda vinda do maturity (precondição não-documentada mas
necessária) → `#NUM!` senão. Suíte core: 601 → **676 verdes** (75 novos: 12 depreciação + 20 taxas/valor
+ 7 fluxos datados + 30 títulos/cupom/T-bill + 6 ODD); suíte Excel intacta (20); build da solução (slnx)
0 warnings (rebuild `--no-incremental`). Docs: `function-reference.md` → 300 (seção Financial reescrita
com 55 assinaturas, Financial 55/55, bases/frequência documentadas) + contagem 300 espelhada no README
(bullet financeiro expandido; `docs/pt-BR/` intocado — refresh do release). Release 2.4.0 pendente de
merge + aval do usuário; sem push.

---

## Fases futuras (analisadas, NÃO autorizadas — abrir plano próprio ao ativar)
- **F1 Voláteis**: **CONCLUÍDA 2026-07-02** — ver [`plans/volatile-functions.md`](volatile-functions.md)
  (3 fases Complete). Cache por época + `Recalculate()`; escopo `TODAY/NOW/RAND/RANDBETWEEN` (OFFSET
  não-volátil, INDIRECT adiado). `touch` por célula BARRADO até termos grafo reverso de dependências
  (volatilidade contagiosa desincronizaria dependentes sem ele). Parser 300 → **304**; Date and Time 25/25
  (categoria completa), Math and Trigonometry 74/82; tags MemoryPackUnion 312–315; release aditivo 2.5.0
  (gate do usuário). §F1-DESIGN abaixo é o registro do desenho original.
- **F2 Arrays**: §A2 (`ComputedValueKind.Array`) → `FILTER` `SORT` `SORTBY` `UNIQUE` `SEQUENCE`
  `TRANSPOSE` `MMULT` `MINVERSE` `MDETERM` `MUNIT` `FREQUENCY` `TEXTSPLIT` `TEXTJOIN`-array `TOCOL`
  `TOROW` `WRAPROWS` `WRAPCOLS` `TAKE` `DROP` `CHOOSEROWS` `CHOOSECOLS` `HSTACK` `VSTACK` `EXPAND`
  `RANDARRAY` (também F1) `MODE.MULT` `ARRAYTOTEXT` `PERCENTOF` `AGGREGATE` `TRIMRANGE` + revisitar spill.
  **Evidência K1 no 2.9.0 (usuário, 2026-07-03, `MYSHEET-CALC-DIVERGENCES.md`)** — MySheet 2.9.0 ×
  Aspose.Cells 26.6 no workbook idêntico, pós-fix do port Aspose (display-name → codeName; ~2.063 diffs
  eram limitação do port, não do MySheet): **concordância 99,946% (565.781/566.086)**; as 305 células
  divergentes restantes traçam a DOIS gaps genuínos do MySheet:
  (a) *avaliação implícita de array (CSE)* — DOMINA os 305 diffs (maquinaria de page-break, ex. real:
  `BH25 =IF(BG25="Page Break",IF($BD25="","",INDEX(ROW($A:$A),SMALL(IF($A$2:$A$194="Show",IF(ROW($A$2:
  $A$194)>$BD25,ROW($A$2:$A$194))),$BB$2))),"")` + colunas B*/BK*/BJ* downstream). Repros:
  `=SUM(IF(B2:B5="Show",1,0))` → `#VALUE!` (Excel: 2); `=SMALL(IF(B2:B5="Show",ROW(B2:B5)),1)` →
  `#VALUE!` (Excel: 3); `=INDEX(ROW(B2:B5),1)` → `#REF!` (Excel: 2). Núcleo do caso de negócio do F2.
  (b) *OR/AND retornam `#VALUE!` com qualquer arg de texto* — ex. real: `H23 =IF(OR(Sheet8!A192="Show",
  Sheet8!A208),"*","")` com A208 texto → MySheet `#VALUE!`, Excel/Aspose `"*"`. Para args de
  ARRAY/REFERÊNCIA a regra Excel é documentada (texto/vazio ignorados; `#VALUE!` só se não sobrar nada
  avaliável) — fix seguro, escalar, fora do F2, candidato a 2.9.x/3.0.x. Semântica de texto LITERAL
  direto (`=OR(TRUE,"x")`): o doc do usuário espera TRUE (skip); prior meu diz que o Excel real dá
  `#VALUE!` p/ literal não-coercível — irrelevante p/ o K1, mas decidir por oráculo (Excel real) antes
  de fixar o caso literal e o de literal coercível (`"TRUE"`).
- **F3 LAMBDA**: §A3 → `LAMBDA` `BYROW` `BYCOL` `MAP` `REDUCE` `SCAN` `MAKEARRAY` `ISOMITTED`.
- **F4 Distribuições**: §A4 → NORM/T/CHISQ/F/GAMMA/BETA/BINOM/POISSON/WEIBULL/EXPON/LOGNORM/NEGBINOM/
  HYPGEOM + testes de hipótese + `CONFIDENCE.*` `CRITBINOM` + aliases compat correspondentes +
  LINEST/LOGEST/TREND/GROWTH (também F2) + FORECAST.ETS* (avaliar: pode ficar fora permanente).
- **F5 Database** (`DSUM`...`DVARP`, 12): viáveis com Criteria + headers de range; prioridade baixa.
- **F6 Engineering**: BIN/DEC/HEX/OCT/BIT* triviais sob demanda; `CONVERT` (tabela de unidades grande);
  ERF/BESSEL/IM* (números complexos — nicho; avaliar exclusão permanente ao chegar).

## Final Recap
Roadmap concluído em 2026-07-02. A cobertura saiu de **52 → 300 funções registradas** no `Parser.Functions`
(medido por script a cada onda; nunca estimado), atravessando 6 ondas de features + 2 fases estruturais
(Fase R namespaces, Fase N named ranges), cada onda fechando um release minor lockstep dos 2 pacotes.

Progressão por onda (contagem cumulativa no Parser):
- **Onda 0** (prep): baseline 52, docs completos + tabela de cobertura com as 35 exclusões permanentes.
- **Onda 1 → 1.1.0**: +60 Math & Trigonometria escalar → 112.
- **Onda 2 → 1.2.0**: +43 Logical + Information + Text + Regex → 155.
- **Onda 3 → 1.3.0**: +9 Lookup & Reference escalar → 164.
- **Fase R → 2.0.0** (breaking): nós da AST reorganizados em namespaces semânticos por categoria;
  `ComputedValue`/`Error` para a raiz; fixture binária MemoryPack prova compat do union por tags.
- **Onda 4 → 2.1.0**: +67 condicionais, SUMPRODUCT, estatística descritiva + aliases Compatibility → 231.
- **Fase N → 2.2.0**: named ranges (workbook-scope, interop xlsx) — sem nós novos.
- **Onda 5 → 2.3.0**: +23 Datas e horas (23/25; `TODAY`/`NOW` adiados como voláteis) → 254; entrega os
  helpers `DateSerial` e `DayCount` reusados pela onda 6.
- **Onda 6 → 2.4.0**: +46 financeiras restantes → **300**; Financial 55/55 (categoria completa).

Suítes finais: core **676 verdes**, Excel **20 verdes**, build 0 warnings, fixture binária pré-namespaces
intocada e verde do começo ao fim (append-only respeitado; tags 0–311 alocadas, próximo livre = 312).

Cobertura viável: 300 de ~485 funções viáveis (~62%); ~520 no catálogo bruto da Microsoft menos as 35
exclusões permanentes documentadas (Cubes, Web, RTD/STOCKHISTORY/IMAGE, pivô, UI, locale CJK, tradução).
O restante do catálogo viável está desenhado (NÃO autorizado) nas fases futuras F1–F6: voláteis
(`TODAY/NOW/RAND*/INDIRECT`), arrays (`ComputedValueKind.Array` → FILTER/SORT/UNIQUE/…), LAMBDA,
distribuições estatísticas (NORM/T/CHISQ/… + testes de hipótese), Database (`DSUM`…), Engineering
(BIN/DEC/HEX, CONVERT, complexos). Cada uma abre plano próprio ao ativar.

Lições estruturais registradas em `tasks/lessons.md` que sobreviveram ao roadmap: (1) golden values SEMPRE
de oráculo citado, nunca de cabeça; (2) solver robusto (bracketing+bisseção) validado contra caso stiff,
não Newton ingênuo; (3) contagens recontadas por script antes de publicar; (4) build de verificação com
`--no-incremental` (incremental mascara warnings de analyzer); (5) **onda 6**: quando o oráculo é a fonte
de verdade, portar a lógica DELE (fuzz valor-a-valor) em vez de assumir a fórmula-textbook — o
ExcelFinancialFunctions diverge do PRICE canônico e do 30/360 do YEARFRAC em casos de fim-de-mês.

## Deployment Plan
_(por onda: merge na main → push (aval) → `gh workflow run release.yml` → minor bump lockstep dos 2
pacotes → `git pull`. Nada de publish manual.)_

---

## §F1-DESIGN: Voláteis — desenho detalhado (proposto 2026-07-02, aguardando aval das 3 decisões)

Contexto: cache por célula é tudo-ou-nada (`InvalidateCache()`), sem grafo de dependências. Voláteis do
Excel têm duas propriedades: (1) recalculam sempre; (2) volatilidade é CONTAGIOSA (dependente de volátil
vira volátil); (3) relógio amostrado uma vez por recálculo (todos os `NOW()` de uma passada concordam).

**Mecanismo (comum a todas as decisões):**
- `virtual bool IsVolatile => false` na `Expression`; override `true` nos nós verdadeiramente
  tempo/aleatório. Transitividade sai de graça: flag thread-local "avaliação tocou volátil" no `Workbook`
  (mesmo padrão do detector de ciclos `_evaluating`), setada ao avaliar nó volátil e propagada ao chamador
  porque `GetCellValue` recorre.
- `TimeProvider` injetável no `Workbook` (default `TimeProvider.System`) + RNG semeável — testabilidade e
  controle do host. Amostrados UMA vez por época e mantidos (coerência do `NOW()`/`RAND()` intra-passada).

**Verificação da qualidade da base (feita 2026-07-02, antes de F1):** teste de mutação manual — quebrei
`MOD` (→ `%` do C#) e `SLN` (dropando salvage); ambos pegos na hora; o teste do SLN compara contra o
oráculo `ExcelFinancialFunctions` ao vivo. 1180 asserts / 702 casos, 0 tautologias, 108 citações à MS.
Confiança alta na maioria oráculo-ancorada; média nas ~37 linhas "derived from documented rules"
(page-silent). Opção barata de fechar a lacuna antes de F1: spike Stryker.NET para mutation score real.

**Decisão 1 — política de cache (recomendado: cache por época + set marcado):** célula volátil é cacheada
DENTRO da época e o set de chaves marcadas é limpo ao avançar a época (melhor desempenho; NOW/RAND uma vez
por passada). Alternativa: nunca cachear voláteis (mais simples, recomputa a subárvore a cada leitura).

**Decisão 2 — API de época (recomendado: InvalidateCache + novo `Recalculate()`):** InvalidateCache()
segue limpando tudo (novos inputs); `Recalculate()` novo limpa só as marcadas + re-amostra relógio/RNG
(refresh barato sem recomputar as estáveis). Alternativa: só reusar InvalidateCache().

**Decisão 3 — escopo (recomendado: TODAY/NOW/RAND/RANDBETWEEN):** as 4 puramente tempo/aleatório.
`INDIRECT` FORA (o difícil é resolver referência a partir de texto — feature própria, toca parser/resolvedor).
**`OFFSET` NÃO vira volátil** (divergência consciente do Excel: lá é volátil por segurança de auto-recálculo;
aqui a invalidação é explícita, então marcá-lo contaminaria meia planilha por nada — documentar).

**Verification Plan (quando executar):** TODAY/NOW com TimeProvider fake (data fixa → valor previsível;
avança o provider + Recalculate → novo valor; sem Recalculate → mantém); dependente de volátil também
refresca (contágio); RAND semeável determinístico; célula não-volátil permanece cacheada através de
Recalculate; fixture binária intocada (IsVolatile é comportamento, não serializado). Docs + refresh pt-BR.
