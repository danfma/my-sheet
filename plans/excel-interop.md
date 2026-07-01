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
  `Workbook.Load`); escrita como **extension methods** no `Workbook`: `SaveAsExcel(...)`, `MergeIntoExcel(...)`
  (dois overloads: template→saída não-destrutivo, e in-place).
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
Status: In progress

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
- [ ] Fast-forward `feature/excel-interop` → `main` (a excel está 1 commit à frente; merge trivial).
- [ ] `git push origin main` — **requer aval do usuário** (leva ~25 commits locais, incluindo o breaking).
- [ ] Disparar o workflow **Release** (workflow_dispatch) → publica os 2 pacotes no NuGet (OIDC) + tag
      `v1.0.0` + GitHub Release.

### Verification Plan
- `versionize --dry-run` local: bump `0.2.0 → 1.0.0`, 2 projetos, changelog com `feat(excel)` — **já validado**.
- CI verde na `main` após o push (build + suíte core + suíte Excel).
- Pós-release: `Danfma.MySheet 1.0.0` e `Danfma.MySheet.Excel 1.0.0` no NuGet; o nupkg do Excel depende de
  `Danfma.MySheet >= 1.0.0`.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Un-parser de fórmula no core (`Expression → texto Excel`)
Status: Not started

- [ ] `Danfma.MySheet/Parsing/FormulaWriter.cs` (ou `Expression.ToFormula()`): renderiza qualquer nó para
      texto de fórmula Excel **sem** o `=` inicial. Cobrir todos os tipos de nó:
      - Valores: `NumberValue` (invariant), `StringValue` (`"..."` com escape de aspas), `BooleanValue`
        (`TRUE`/`FALSE`), `BlankValue` (vazio), `ErrorValue` (`#VALUE!` etc.).
      - Referências: `CellReference` (`A1`, sheet-qualified `Sheet2!A1`, nome com espaço `'My Sheet'!A1`),
        `RangeReference` (`A1:B2`), `UnionReference` (`(A1:A3,C1:C3)`), `NameReference` (o nome).
      - Operadores: `BinaryOperation` com **parentetização por precedência** (aritmética/comparação/`&`),
        `UnaryOperation` (`-x`, `+x`, `x%`).
      - Funções: cada `record : Function` → `NOME(arg1,arg2,…)`; `FunctionCall` custom → `Nome(args)`; `Let`.
- [ ] Mapa **tipo-de-nó → nome de função Excel** (espelha o `Functions` do `Parser`): switch central no
      `FormulaWriter` (preferido, para não engordar cada nó).
- [ ] Testes em `tests/Danfma.MySheet.Tests/Parsing/FormulaWriterTests.cs`: round-trip
      `Parse(f) → ToFormula → Parse` com um corpus (aritmética com precedência, comparações, `&`, `%`, ranges,
      sheet-qualified, funções variádicas/aninhadas, IF, VLOOKUP, LET, custom `FunctionCall`); asserir
      igualdade estrutural (re-parse) e string onde determinístico.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` → verde, incl.
  `FormulaWriterTests`. Corpus prova `Parse(ToFormula(Parse(f))) ≈ Parse(f)` (AST equivalente) para ~20 fórmulas.
- Build do core 0 Warning(s).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 4: Export completo (`SaveAsExcel`)
Status: Not started

- [ ] `ExcelExportOptions { FormulaMode FormulaMode = ValuesOnly }` (enum `FormulaMode { Formulas, ValuesOnly }`).
- [ ] Extension `public static void SaveAsExcel(this Workbook workbook, string path, ExcelExportOptions? options = null)`:
      cria `SpreadsheetDocument`; `WorkbookPart` + uma `WorksheetPart` por planilha (ordem = `Sheet.Index`,
      nome = `Sheet.Name`); `SharedStringTablePart` para textos.
      - Por célula (`Sheet.Cells`): computa o valor (`workbook.GetCellValue(...)`, tudo em `RunWithLargeStack`).
        - `ValuesOnly`: escreve o literal (número / shared-string / bool `t="b"` / erro `t="e"` Display); blank → omite.
        - `Formulas`: se o nó é fórmula (não um literal puro), escreve `<f>ToFormula()</f>` + `<v>valor</v>`;
          se é literal, escreve o literal.
      - Células em ordem de coluna dentro da linha; linhas em ordem (exigência do OpenXML).
- [ ] Testes: monta um `Workbook` (literais + fórmulas em 2 planilhas), `SaveAsExcel` num arquivo temp, relê
      com ClosedXML/nosso reader e asseri: valores, tipos, `<f>` presente/ausente conforme o modo, nomes/ordem
      de planilha, shared strings. **Round-trip**: `Load(SaveAsExcel(wb))` ≡ `wb` (células e resultados).

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → verde.
  Um teste abre o `.xlsx` gerado e confirma célula/valor/fórmula nos dois `FormulaMode`; round-trip com o reader.
- Build 0 Warning(s).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 5: Merge de valores num `.xlsx` existente (`MergeIntoExcel`)
Status: Not started

- [ ] Overloads: `MergeIntoExcel(this Workbook, string templatePath, string outputPath, …)` (não-destrutivo:
      copia o template → escreve no output) e `MergeIntoExcel(this Workbook, string path, …)` (in-place).
- [ ] Para cada planilha nossa (casa por nome, case-insensitive; ausente no alvo → pula e conta): para cada
      célula nossa, computa o valor e grava como **literal** na célula correspondente do alvo — **dropa** a
      fórmula que houvesse ali; cria a célula se não existir (mantendo a ordem coluna/linha do OpenXML); blank →
      não escreve. Todo o resto do alvo (formatação, outras células/planilhas, shared strings) é preservado.
- [ ] Testes: gera um template `.xlsx` (fórmulas + uma célula "formatada"/extra como fixture), roda o merge dos
      nossos valores, relê e asseri: valores injetados corretos, fórmula da célula-alvo removida (virou literal),
      células/planilhas não-alvo intactas, planilha ausente no template ignorada.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release` → verde,
  incl. os testes de merge (não-destrutivo e in-place; template preservado; fórmula-alvo dropada).

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 6 (futuro / fora do MVP): estilos, formatos, datas, streaming, shared formulas
Status: Not started

- [ ] Formatos numéricos + datas-como-data (round-trip de `numFmt`/`cellFormats`), estilos básicos, merged
      cells, larguras. Streaming (SAX `OpenXmlWriter`) para arquivos grandes. Fidelidade de `$` (absolutos) na
      AST. Expansão de shared formulas no reader (shift de referências na AST parseada da célula-mestre).

### Verification Plan
- (definir quando/se a fase for priorizada)

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando as fases 1–5 concluírem)_

## Deployment Plan
_(escrever quando concluir: novo pacote NuGet `Danfma.MySheet.Excel` — versão inicial; depende de
`Danfma.MySheet` + `DocumentFormat.OpenXml`; publicar junto no CI. Core inalterado exceto o `FormulaWriter`.
A publicação inicial (1.0.0, lockstep com o core) acontece na Fase 2; releases seguintes idem via versionize.)_
