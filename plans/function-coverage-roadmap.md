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
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` (440 após a onda 3);
Excel `.../Danfma.MySheet.Excel.Tests... -c Release` (16 hoje). Build da solução: 0 warnings sempre.

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
Status: In progress
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

- [ ] Mover os records de função para os namespaces por categoria (pastas correspondentes); atualizar
      usings de `Parser`, `FormulaWriter`, testes (preferir `GlobalUsings` nos csproj de teste).
- [x] Decisão do usuário: mover `ComputedValue`/`Error` para a raiz `Danfma.MySheet` — SIM (2026-07-02).
- [ ] **Teste de compat binária MemoryPack**: ANTES do refactor, serializar um workbook representativo
      (fórmulas de todas as categorias) e commitar o blob como fixture; DEPOIS, teste que o deserializa e
      reavalia — prova que o union por tags sobrevive à mudança de namespace. (Tags não mudam: append-only
      continua valendo.)
- [ ] Docs: guia de migração 1.x → 2.0 (tabela de namespaces), atualização dos guias EN + refresh pt-BR.
- [ ] Atualizar os briefings das ondas seguintes: nós novos nascem no namespace da categoria.
- [ ] Commit `refactor!:` (BREAKING CHANGE) → release **2.0.0**.

### Verification Plan
- Suítes completas verdes (core + Excel) + teste de compat binária com a fixture pré-refactor.
- Build 0 warnings; `versionize --dry-run` propõe major.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 4 (Onda 4 → 1.4.0): Condicionais, SUMPRODUCT, estatística descritiva + aliases (~55 funções)
Status: Not started

- [ ] Condicionais (Criteria existente): `AVERAGEIF` `AVERAGEIFS` `MAXIFS` `MINIFS`
- [ ] Variantes A: `AVERAGEA` `MAXA` `MINA`
- [ ] Pairwise multi-range (helper novo `PairwiseRanges` em NumericAggregation): `SUMPRODUCT` `SUMX2MY2`
      `SUMX2PY2` `SUMXMY2` (ranges de shapes diferentes → `#VALUE!`)
- [ ] `SUBTOTAL` (códigos 1-11 = 101-111; ignora SUBTOTAL aninhado de verdade; sem hidden rows — limite de
      modelo documentado, ver §A5)
- [ ] Ordem/posição (helper de sorting sobre `List<double>`): `MEDIAN` `MODE.SNGL` `LARGE` `SMALL`
      `RANK.EQ` `RANK.AVG` `PERCENTILE.INC` `PERCENTILE.EXC` `PERCENTRANK.INC` `PERCENTRANK.EXC`
      `QUARTILE.INC` `QUARTILE.EXC` `TRIMMEAN`
- [ ] Dispersão/momentos: `STDEV.S` `STDEV.P` `STDEVA` `STDEVPA` `VAR.S` `VAR.P` `VARA` `VARPA` `AVEDEV`
      `DEVSQ` `GEOMEAN` `HARMEAN` `SKEW` `SKEW.P` `KURT` `STANDARDIZE`
- [ ] Bivariadas (pairwise): `CORREL` `PEARSON` `COVARIANCE.P` `COVARIANCE.S` `RSQ` `SLOPE` `INTERCEPT`
      `STEYX` `FORECAST.LINEAR`
- [ ] Escalares simples: `FISHER` `FISHERINV` `GAUSS` `PHI` `PERMUT` `PERMUTATIONA` `PROB`
- [ ] Aliases Compatibility dos implementados (mesma factory, nome legado): `MODE` `STDEV` `STDEVP` `VAR`
      `VARP` `RANK` `PERCENTILE` `PERCENTRANK` `QUARTILE` `COVAR` `FORECAST` (⚠️ `CONFIDENCE`/`CRITBINOM`
      e os *DIST/*INV ficam para F4 — dependem de distribuições)
- [ ] FormulaWriter.Call + corpus + README ✅

### Verification Plan
- Suítes verdes; PERCENTILE.INC/EXC e QUARTIS contra oráculo (interpolações diferem entre ferramentas —
  fixar Excel); CORREL/SLOPE contra oráculo com dataset fixo; SUMPRODUCT shape-mismatch → `#VALUE!`.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 5 (Onda 5 → 1.5.0): Datas e horas — 23/25 (sem TODAY/NOW) 
Status: Not started

- [ ] Infra `DateSerial` no core (internal): serial ↔ DateTime via OADate; clamp/`#NUM!` para serial
      negativo; fração do dia = hora. Documentar limite 1900 (§A6).
- [ ] Construção/extração: `DATE` (com overflow de mês/dia como Excel: `DATE(2020,13,1)`→jan/2021)
      `DATEVALUE` `TIMEVALUE` (parse invariant) `TIME` `YEAR` `MONTH` `DAY` `HOUR` `MINUTE` `SECOND`
- [ ] Aritmética de calendário: `DAYS` `DAYS360` (métodos US/EU) `EDATE` `EOMONTH` `WEEKDAY` (return_type
      1/2/3/11-17) `WEEKNUM` `ISOWEEKNUM` `DATEDIF` (6 modos, quirks documentados) `YEARFRAC` (5 bases —
      infra reutilizada pela onda 6)
- [ ] Dias úteis (feriados via range): `NETWORKDAYS` `NETWORKDAYS.INTL` `WORKDAY` `WORKDAY.INTL`
      (máscaras de fim de semana 1-17 e "0000011")
- [ ] `TODAY`/`NOW` ficam para F1 (voláteis) — registrar ⬜ com nota no README.
- [ ] FormulaWriter.Call + corpus + README ✅ (Date and Time 23/25)

### Verification Plan
- Suítes verdes; DATEDIF/YEARFRAC/DAYS360/WEEKNUM contra oráculo (planilha de golden values citada);
  round-trip xlsx: data serial exportada/lida sem perda.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 6 (Onda 6 → 1.6.0): Financeiras restantes viáveis (~40 funções)
Status: Not started

Oráculo: `ExcelFinancialFunctions` (já referenciado nos testes) cobre a maior parte — validar cada golden.

- [ ] Depreciação: `SLN` `SYD` `DB` `DDB` `VDB` `AMORLINC` `AMORDEGRC` (usam YEARFRAC/datas da onda 5)
- [ ] Taxas/valor: `EFFECT` `NOMINAL` `MIRR` `RRI` `PDURATION` `ISPMT` `CUMIPMT` `CUMPRINC` `FVSCHEDULE`
      `DOLLARDE` `DOLLARFR`
- [ ] Fluxos com datas: `XNPV` `XIRR` (solver bracketing+bisseção existente do IRR/RATE)
- [ ] Títulos (day-count da onda 5): `ACCRINT` `ACCRINTM` `DISC` `DURATION` `MDURATION` `INTRATE` `PRICE`
      `PRICEDISC` `PRICEMAT` `RECEIVED` `YIELD` `YIELDDISC` `YIELDMAT` `TBILLEQ` `TBILLPRICE` `TBILLYIELD`
      `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD`
- [ ] `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD`: avaliar custo ao chegar — se o oráculo não cobrir
      com precisão verificável, ficam ⬜ com nota (não entregar número não-validado).
- [ ] FormulaWriter.Call + corpus + README ✅ (Financial ~49-53/55)

### Verification Plan
- Suítes verdes; TODAS as financeiras com golden do `ExcelFinancialFunctions` (incluindo o caso stiff de
  mortgage 30 anos — lição do RATE); XIRR com fluxo irregular de datas reais.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Fases futuras (analisadas, NÃO autorizadas — abrir plano próprio ao ativar)
- **F1 Voláteis**: infra §A1 (IsVolatile + no-cache propagado + TimeProvider) → `TODAY` `NOW` `RAND`
  `RANDBETWEEN` `INDIRECT`.
- **F2 Arrays**: §A2 (`ComputedValueKind.Array`) → `FILTER` `SORT` `SORTBY` `UNIQUE` `SEQUENCE`
  `TRANSPOSE` `MMULT` `MINVERSE` `MDETERM` `MUNIT` `FREQUENCY` `TEXTSPLIT` `TEXTJOIN`-array `TOCOL`
  `TOROW` `WRAPROWS` `WRAPCOLS` `TAKE` `DROP` `CHOOSEROWS` `CHOOSECOLS` `HSTACK` `VSTACK` `EXPAND`
  `RANDARRAY` (também F1) `MODE.MULT` `ARRAYTOTEXT` `PERCENTOF` `AGGREGATE` `TRIMRANGE` + revisitar spill.
- **F3 LAMBDA**: §A3 → `LAMBDA` `BYROW` `BYCOL` `MAP` `REDUCE` `SCAN` `MAKEARRAY` `ISOMITTED`.
- **F4 Distribuições**: §A4 → NORM/T/CHISQ/F/GAMMA/BETA/BINOM/POISSON/WEIBULL/EXPON/LOGNORM/NEGBINOM/
  HYPGEOM + testes de hipótese + `CONFIDENCE.*` `CRITBINOM` + aliases compat correspondentes +
  LINEST/LOGEST/TREND/GROWTH (também F2) + FORECAST.ETS* (avaliar: pode ficar fora permanente).
- **F5 Database** (`DSUM`...`DVARP`, 12): viáveis com Criteria + headers de range; prioridade baixa.
- **F6 Engineering**: BIN/DEC/HEX/OCT/BIT* triviais sob demanda; `CONVERT` (tabela de unidades grande);
  ERF/BESSEL/IM* (números complexos — nicho; avaliar exclusão permanente ao chegar).

## Final Recap
_(escrever quando as ondas 0–6 concluírem)_

## Deployment Plan
_(por onda: merge na main → push (aval) → `gh workflow run release.yml` → minor bump lockstep dos 2
pacotes → `git pull`. Nada de publish manual.)_
