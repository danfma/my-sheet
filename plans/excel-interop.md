# Interop com Excel (.xlsx) via OpenXML — reader primeiro, depois export completo + merge

Nova biblioteca `Danfma.MySheet.Excel` que **lê e escreve** arquivos `.xlsx` (multiplataforma, sem Excel
instalado, via `DocumentFormat.OpenXml`): **reader** (`.xlsx` → nosso `Workbook`, eliminando a dependência de
outras libs para carregar), criação completa (com fórmulas ou só valores, configurável) e merge de valores
computados num `.xlsx` existente (template).

## Context / decisões (fechadas com o usuário)
- **Ordem revertida em 2026-07-01**: o plano original era só-saída ("sem import"); o usuário decidiu **incluir
  o reader (import) e fazê-lo primeiro**, expandindo depois para o writer. Racional: o reader é mais barato —
  reutiliza o `ExpressionParser` que já existe no core (`texto → Expression`); o un-parser (`Expression →
  texto`), a parte difícil, só bloqueia o writer no modo `Formulas`.
- **Reader** = `.xlsx` → `Workbook`: planilhas (nome/ordem), células com **fórmula** (`<f>` →
  `ExpressionParser.Parse("=" + f, sheet)`) ou **literal** (número/texto/bool/erro); shared strings resolvidas.
- **Export completo configurável** (`FormulaMode`): `Formulas` = escreve a fórmula Excel (`<f>`) + o valor
  computado em cache (`<v>`); `ValuesOnly` = só o valor literal (snapshot achatado, sem fórmula).
- **Merge** = grava o **valor literal computado** nas células que temos, **dropando** qualquer fórmula que
  houvesse naquelas células; todo o resto do template (formatação, demais células/planilhas) fica intacto.
- **MVP núcleo**: células (valor + fórmula), múltiplas planilhas, nomes de planilha, referências, tipos
  básicos (número/texto/booleano/erro/blank), shared strings. **Fora do MVP**: estilos, formatos numéricos,
  datas-como-data (entram/saem como número serial), merged cells, larguras, streaming/SAX.
- **Un-parser de fórmula no CORE** (`Danfma.MySheet`): simétrico ao parser, `Expression → texto de fórmula
  Excel` (sem o `=` inicial). Reutilizável, mantém a lib de interop fina. Necessário só a partir da Fase 3.
- **Projeto/API**: `Danfma.MySheet.Excel`; reader como `ExcelFile.Load(path) : Workbook` (espelha
  `Workbook.Load`); escrita como **extension methods** no `Workbook`: `SaveAsExcel(...)`, `MergeIntoExcel(path)`
  (**só in-place** — decisão do usuário em 2026-07-01: merge muta o arquivo dado; criar arquivo é papel do
  `SaveAsExcel`; o fluxo template→relatório é `File.Copy` + merge na cópia, documentado no README e num teste).
- **Padrões assumidos**: tipos xlsx padrão (número/shared-string/bool `t="b"`/erro `t="e"` Display/blank omitido;
  `Reference` como resultado → `#VALUE!`); merge casa planilha por nome (ausente → pula; célula ausente → cria
  na ordem exigida; blank → não escreve); avaliação em `RunWithLargeStack`; DOM do OpenXML; dep só na lib nova.
- **Limitações conhecidas (registradas)**:
  - A AST não guarda marcadores absolutos (`$A$1`). No **reader** isso NÃO é limitação: o tokenizer já aceita
    `$` e a classificação de cell-ref o descarta (`$A$1` → `A1`, verificado por teste). A perda de fidelidade
    é só no **un-parse** (writer), que produzirá referências relativas; fica para fase futura se necessário.
  - **Shared formulas** (xlsx real usa `<f t="shared">` para fórmulas arrastadas): a célula-mestre tem o texto
    (parseia normal); as células-escravas vêm sem texto → fallback para o **valor cacheado** `<v>` como literal.
    Expansão real (shift de referências na AST) fica para fase futura.

## For Future Agents
Marque `- [x]`; ao fechar fase, Status `Complete` + Phase Summary + rode a Verification. TDD por unidade.
Teste (core): `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` (240 verdes hoje).
Teste (Excel): `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release`
— gera fixtures `.xlsx` com **ClosedXML** (oráculo independente, já usado no benchmark) e lê com a nossa lib.

---

## Phase 1: Projeto `Danfma.MySheet.Excel` + reader (`ExcelFile.Load`)
Status: Complete

- [x] Novo `Danfma.MySheet.Excel/Danfma.MySheet.Excel.csproj` (`net10.0`, LangVersion preview), refs
      `Danfma.MySheet` + `DocumentFormat.OpenXml` (3.x). Core permanece sem essa dependência. Adicionar ao slnx.
- [x] Novo `tests/Danfma.MySheet.Excel.Tests` (TUnit, espelha o csproj de testes do core) + `ClosedXML` como
      gerador de fixtures. Adicionar ao slnx.
- [x] `ExcelFile.Load(string path) : Workbook` (+ overload `Stream`): abre `SpreadsheetDocument` read-only;
      para cada `Sheet` do workbook.xml (na ordem, casando `Index`): cria a sheet nossa e popula células:
      - `<f>` com texto → `ExpressionParser.Parse("=" + texto, sheet)` (fórmula real, será reavaliada por nós).
      - `<f>` vazio (shared-formula escrava) → fallback: literal do `<v>` cacheado.
      - Sem fórmula: literal por `t`: `s` (shared string), `b` (bool), `e` (erro → `ErrorValue`), `str`/
        `inlineStr` (texto), `d` (ISO-8601 → serial via `ToOADate`), default número (invariant). Célula sem
        `<v>` → ignorada (blank).
- [x] Testes (fixture ClosedXML → `ExcelFile.Load` → asserções): literais (número/texto/bool) em 2 planilhas
      com nome/ordem; fórmulas (`A1+A2`, `SUM(range)`, cross-sheet `Data!A1*2`, absoluta `$A$1+A2`)
      reavaliadas pelo nosso engine com o resultado esperado; erro literal (`#DIV/0!`); célula vazia → blank.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → verde.
- Suíte do core continua verde (240) e `dotnet build` da solução com 0 Warning(s).

### Phase Summary
TDD (RED 6/6 `NotImplementedException` → GREEN 6/6 → refactor de nulabilidade). Entregues:
`Danfma.MySheet.Excel` (OpenXml 3.*, ref ao core) + `tests/Danfma.MySheet.Excel.Tests` (TUnit + ClosedXML
0.105 como gerador de fixtures — oráculo independente), ambos no slnx. `ExcelFile.Load(path|Stream)` lê:
sheets na ordem do workbook.xml (Index = ordem de aba), fórmulas via `ExpressionParser.Parse("=" + f, sheet)`
(o parser já aceita `$` — refs absolutas normalizam para relativas), literais por tipo (shared strings
achatando rich-text via InnerText; bool 1/TRUE; erro → `new ErrorValue(display)` — record público, sem
precisar de InternalsVisibleTo; `d` ISO → serial `ToOADate`; default número invariant), shared-formula
escrava → literal do `<v>`, célula sem valor → blank. Verificação: solução 0 warnings; Excel 6/6; core
240/240 intacto. Decisão de nome: **`ExcelFile.Load`** (espelha `Workbook.Load`), avisado ao usuário.

---

## Phase 2: Publicação do loader (merge na main + release conjunto)
Status: Complete

Decisão do usuário (2026-07-01): publicar o reader antes de seguir para o writer. A `main` local já contém
todo o trabalho do `ComputedValue` (breaking `feat!` → major); o release sai em **lockstep**: `Danfma.MySheet
1.0.0` + `Danfma.MySheet.Excel 1.0.0` juntos, com o Excel dependendo do core na mesma versão.

- [x] CI (`ci.yml`): passa a rodar também a suíte da lib Excel (`Danfma.MySheet.Excel.Tests`).
- [x] Release (`release.yml`): `versionize` agora roda na **raiz** (lockstep dos 2 pacotes) e o `Pack`
      empacota os dois csproj. `CHANGELOG.md` movido de `Danfma.MySheet/` para a raiz (histórico preservado;
      é onde o versionize passará a escrever).
- [x] Seed de versão do Excel alinhado ao core (`0.2.0`) — o versionize exige versões consistentes
      (**validado por dry-run local**: descobre exatamente os 2 projetos, propõe `0.2.0 → 1.0.0` pelos
      `feat!`, e o changelog já lista o `feat(excel)` do reader).
- [x] Fast-forward `feature/excel-interop` → `main` (90a6d71..771e033).
- [x] `git push origin main` — feito com aval do usuário (b98d7da..771e033, 28 commits).
- [x] Workflow **Release** disparado pelo usuário (2026-07-01): run verde em 35s → `v1.0.0` com
      `Danfma.MySheet.1.0.0.nupkg` + `Danfma.MySheet.Excel.1.0.0.nupkg`; bump commit (6b28424) e tag
      puxados para a main local.

### Verification Plan
- `versionize --dry-run` local: bump `0.2.0 → 1.0.0`, 2 projetos, changelog com `feat(excel)` — **já validado**.
- CI verde na `main` após o push (build + suíte core + suíte Excel).
- Pós-release: `Danfma.MySheet 1.0.0` e `Danfma.MySheet.Excel 1.0.0` no NuGet; o nupkg do Excel depende de
  `Danfma.MySheet >= 1.0.0`.

### Phase Summary
Release **1.0.0 publicado** (lockstep validado em produção): versionize na raiz bumpou os 2 csproj
(0.2.0→1.0.0 pelos `feat!` do ComputedValue), changelog na raiz, tag `v1.0.0`, 2 nupkgs no NuGet + GitHub
Release. Na prática o release saiu DEPOIS das fases 3-5, então o 1.0.0 já incluiu reader + un-parser +
export + merge de uma vez. CI da main verde em todos os pushes.

---

## Phase 3: Un-parser de fórmula no core (`Expression → texto Excel`)
Status: Complete

- [x] `Danfma.MySheet/Parsing/FormulaWriter.cs` (`Expression.ToFormula(contextSheetName)`): renderiza qualquer nó para
      texto de fórmula Excel **sem** o `=` inicial. Cobrir todos os tipos de nó:
      - Valores: `NumberValue` (invariant), `StringValue` (`"..."` com escape de aspas), `BooleanValue`
        (`TRUE`/`FALSE`), `BlankValue` (vazio), `ErrorValue` (`#VALUE!` etc.).
      - Referências: `CellReference` (`A1`, sheet-qualified `Sheet2!A1`, nome com espaço `'My Sheet'!A1`),
        `RangeReference` (`A1:B2`), `UnionReference` (`(A1:A3,C1:C3)`), `NameReference` (o nome).
      - Operadores: `BinaryOperation` com **parentetização por precedência** (aritmética/comparação/`&`),
        `UnaryOperation` (`-x`, `+x`, `x%`).
      - Funções: cada `record : Function` → `NOME(arg1,arg2,…)`; `FunctionCall` custom → `Nome(args)`; `Let`.
- [x] Mapa **tipo-de-nó → nome de função Excel** (espelha o `Functions` do `Parser`): switch central no
      `FormulaWriter` (`Call(Function) → (nome, args)`; única exceção de shape: `Sum.Expressions`).
- [x] Testes em `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`: round-trip
      `Parse(f) → ToFormula → Parse` com um corpus (aritmética com precedência, comparações, `&`, `%`, ranges,
      sheet-qualified, funções variádicas/aninhadas, IF, VLOOKUP, LET, custom `FunctionCall`); asserir
      igualdade estrutural (re-parse) e string onde determinístico.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → verde, incl.
  `FormulaWriterTests`. Corpus prova `Parse(ToFormula(Parse(f))) ≈ Parse(f)` (AST equivalente) para ~20 fórmulas.
- Build do core 0 Warning(s).

### Phase Summary
TDD (RED 34 casos `NotImplementedException` → GREEN 34/34). `FormulaWriter.ToFormula(this Expression, string
contextSheetName)` em `Danfma.MySheet.Parsing`: espelha os binding powers do Pratt parser (10/15/20/30/40,
`%`=44 posfixo, prefixo=45) e emite parênteses mínimos — `Write(nó, minPrecedence)` parentetiza quando a
precedência do nó < mínimo; binários left-assoc apertam a direita (`p+1`), `^` right-assoc aperta a esquerda.
Referências qualificam só fora do contexto (`Sheet2!A1`; aspas simples com escape `''` quando o nome não é
identificador simples); strings escapam `"`→`""`; `SHEET` é a única função cujo nome não é o uppercase do
tipo; `Sum` é o único record cujo parâmetro não se chama `Arguments` (`Expressions`). Cobertura: 30 fórmulas
canônicas (string exata), 3 normalizações (igualdade estrutural via bytes MemoryPack) e as 52 funções
built-in (round-trip estrutural). Core: 274/274; solução 0 warnings.

---

## Phase 4: Export completo (`SaveAsExcel`)
Status: Complete

- [x] `ExcelExportOptions { FormulaMode FormulaMode = ValuesOnly }` (enum `FormulaMode { Formulas, ValuesOnly }`).
- [x] Extension `public static void SaveAsExcel(this Workbook workbook, string path, ExcelExportOptions? options = null)`:
      cria `SpreadsheetDocument`; `WorkbookPart` + uma `WorksheetPart` por planilha (ordem = `Sheet.Index`,
      nome = `Sheet.Name`); `SharedStringTablePart` para textos.
      - Por célula (`Sheet.Cells`): computa o valor (`workbook.GetCellValue(...)`, tudo em `RunWithLargeStack`).
        - `ValuesOnly`: escreve o literal (número / shared-string / bool `t="b"` / erro `t="e"` Display); blank → omite.
        - `Formulas`: se o nó é fórmula (não um literal puro), escreve `<f>ToFormula()</f>` + `<v>valor</v>`;
          se é literal, escreve o literal.
      - Células em ordem de coluna dentro da linha; linhas em ordem (exigência do OpenXML).
- [x] Testes: monta um `Workbook` (literais + fórmulas em 2 planilhas), `SaveAsExcel` num arquivo temp, relê
      com ClosedXML/nosso reader e asseri: valores, tipos, `<f>` presente/ausente conforme o modo, nomes/ordem
      de planilha, shared strings. **Round-trip**: `Load(SaveAsExcel(wb))` ≡ `wb` (células e resultados).

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → verde.
  Um teste abre o `.xlsx` gerado e confirma célula/valor/fórmula nos dois `FormulaMode`; round-trip com o reader.
- Build 0 Warning(s).

### Phase Summary
TDD (RED 4 `NotImplementedException` → GREEN 10/10 na suíte Excel). `ExcelExport.SaveAsExcel(this Workbook,
path|Stream, ExcelExportOptions?)` + `FormulaMode { ValuesOnly (default), Formulas }`. Avalia tudo num único
`RunWithLargeStack` (memoizado) antes de escrever; sheets na ordem de `Index`; linhas/células ordenadas
(exigência OpenXML); shared strings dedupadas (part só criada se usada). Literais: número/bool `t="b"`/texto
shared-string/erro `t="e"`/blank omitido; `Reference` → `#VALUE!`. Modo `Formulas`: nó não-`ValueExpression`
escreve `<f>` via `FormulaWriter.ToFormula(sheetName)` + `<v>` cacheado (texto de fórmula usa `t="str"`, não
shared string — convenção xlsx). Verificado com o nosso reader (round-trip: fórmula volta como árvore e
reavalia igual; ValuesOnly volta achatado) e com ClosedXML como oráculo (FormulaA1 == "A1+A2" e cross-sheet
"Data!A3*2"; CachedValue == 5). Colisões de nome OpenXml×core (Sheet/Workbook/Row/Text) resolvidas com
aliases Xlsx*.

---

## Phase 5: Merge de valores num `.xlsx` existente (`MergeIntoExcel`)
Status: Complete

- [x] `MergeIntoExcel(this Workbook, string path)` — **só in-place** (o overload template→saída foi removido
      após revisão do usuário; a "cópia não-destrutiva" fica a cargo do chamador: `File.Copy` + merge).
- [x] Para cada planilha nossa (casa por nome, case-insensitive; ausente no alvo → pula): para cada
      célula nossa, computa o valor e grava como **literal** na célula correspondente do alvo — **dropa** a
      fórmula que houvesse ali; cria a célula se não existir (mantendo a ordem coluna/linha do OpenXML); blank →
      não escreve. Todo o resto do alvo (formatação, outras células/planilhas, shared strings) é preservado.
- [x] Testes: gera um template `.xlsx` (fórmulas + célula extra como fixture), roda o merge dos
      nossos valores, relê e asseri: valores injetados corretos, fórmula da célula-alvo removida (virou literal),
      células/planilhas não-alvo intactas, planilha ausente no template ignorada.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → verde,
  incl. os testes de merge (não-destrutivo e in-place; template preservado; fórmula-alvo dropada).

### Phase Summary
TDD (RED 2 → GREEN 12/12 na suíte Excel). `ExcelMerge.MergeIntoExcel(this Workbook, path)` (in-place; o
overload template→saída foi removido após revisão do usuário — fluxo template = `File.Copy` + merge,
documentado num teste-receita). Avalia tudo num `RunWithLargeStack`; casa
sheet por nome case-insensitive (ausente → pula); blank não escreve; linha/célula criadas em ordem quando
faltam (`InsertBefore` no primeiro maior); `WriteLiteral` limpa `<f>`/conteúdo mas NÃO toca `StyleIndex`
(formatação do template preservada). Texto vai como **inline string** (`t="inlineStr"`) para não mexer na
shared-string table do alvo — desvio consciente do padrão do export, registrado aqui. `CellId.Parse`
extraído como helper interno compartilhado com o `ExcelExport`. Testes (oráculo ClosedXML): valor injetado
como literal com fórmula-alvo dropada, células novas criadas, célula/planilha não-alvo intactas, sheet
ausente pulada, template original preservado no modo não-destrutivo, e o modo in-place. 12/12; core 274/274;
solução 0 warnings.

---

## Phase 6a: Shared formulas no reader (expansão real)
Status: In progress

Prioridade escolhida dentro da Fase 6 por ser **correção**, não conveniência: arquivos salvos pelo Excel
usam `<f t="shared">` para fórmulas arrastadas; sem expansão, células escravas viram literais estáticos no
`Load` — fórmulas se perdem silenciosamente no caso de uso âncora (servidor com Excel como fonte da verdade).

- [ ] Reader: registrar por worksheet o mapa `si → (id da célula mestre, texto da fórmula)` ao encontrar
      `<f t="shared">` com texto; célula com `<f t="shared" si>` SEM texto → shift do texto da mestre pelo
      delta (linha/coluna) e parse do resultado.
- [ ] Shift no NÍVEL DO TEXTO (não na AST — a AST não guarda `$`, e refs absolutas NÃO devem deslocar):
      scanner que pula strings `"…"` e nomes `'…'`, reconhece refs `($?)LETRAS($?)DÍGITOS` standalone
      (não precedidas/seguidas de caracteres de identificador; não seguidas de `(`), desloca só os
      componentes sem `$`.
- [ ] Testes com fixture OpenXML crua (ClosedXML não escreve shared formulas): mestre + escravas,
      refs relativas deslocam, absolutas (`$A$1`) ficam, refs dentro de strings não mudam.

### Verification Plan
- Suíte Excel verde com os novos testes de shared formula; core intacto; 0 warnings.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 6 (futuro / fora do MVP): estilos, formatos, datas, streaming
Status: Not started

- [ ] Formatos numéricos + datas-como-data (round-trip de `numFmt`/`cellFormats`), estilos básicos, merged
      cells, larguras. Streaming (SAX `OpenXmlWriter`) para arquivos grandes. Fidelidade de `$` (absolutos) na
      AST (writer).

### Verification Plan
- (definir quando/se a fase for priorizada)

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
MVP completo (Fases 1–5, todas TDD sempre-verde, na branch `feature/excel-writer` a partir da 3):
1. **Reader** `ExcelFile.Load(path|Stream) : Workbook` — fórmulas viram `Expression` reais (reavaliadas pelo
   nosso engine), literais tipados, shared strings, `$` normalizado, shared-formula escrava → valor cacheado.
2. **Publicação do loader** — CI roda as 2 suítes; release em lockstep (versionize na raiz + pack duplo);
   main pushada (771e033). Release dispatch = ação do usuário.
3. **Un-parser** `Expression.ToFormula(contextSheet)` no core — parênteses mínimos espelhando o Pratt parser,
   52 funções mapeadas, round-trip provado por string exata + igualdade estrutural MemoryPack.
4. **Export** `SaveAsExcel(path|Stream, options)` — `ValuesOnly` (snapshot achatado, default) ou `Formulas`
   (`<f>` + `<v>` cacheado); shared strings dedupadas; validado pelo nosso reader + oráculo ClosedXML.
5. **Merge** `MergeIntoExcel(path)` (in-place) — injeta valores computados como literais no arquivo,
   dropando fórmulas-alvo e preservando todo o resto (estilos intactos via StyleIndex não tocado). O fluxo
   template→relatório é `File.Copy` + merge na cópia (receita no README e em teste).
Total: suíte Excel 12/12, core 274/274 (34 novos do FormulaWriter), solução 0 warnings.

## Deployment Plan
1. Merge `feature/excel-writer` → `main` (a Fase 2 já publicou a main com o reader; este merge leva
   un-parser + export + merge — commits `feat` → minor bump no próximo release).
2. Push da `main` (aval do usuário) → CI verde (build + 2 suítes).
3. `gh workflow run release.yml` → versionize bumpa os 2 pacotes em lockstep + NuGet + GitHub Release.
   (Se o release do loader (1.0.0) ainda não rodou, um único release cobre tudo.)
4. Pós-release: `git pull` na main local (o workflow empurra commit de bump + tag).
