# Excel Load Streaming + Save Residual (vs Aspose)

Eliminar a dupla representação em memória do load de .xlsx (DOM OpenXml + modelo MySheet) via leitura streaming, cortar o custo de parse por célula, e reduzir o residual de alocação do save — fechando o gap medido contra o Aspose (load: 2042ms vs 936ms; ambos tempo e memória importam).

> **Nota /real-work**: ao iniciar a implementação, copiar este plano para `plans/excel-load-streaming.md` no repo (artefato durável); manter os checkboxes lá.

## Context

- **Medição externa do usuário** (arquivo perfil K1, ~566k células): load MySheet 2042.55ms × Aspose 935.87ms (2.2x). "Preparação" (popular via API) 166ms × 52ms — fora de escopo direto (Aspose é lazy-parse; comparação por fase é estruturalmente injusta), ganhos só os que vierem de graça do trabalho de parse.
- **Causa raiz #1 (memória e tempo)**: `ExcelFile.Load` é 100% DOM — `SpreadsheetDocument.Open` + `worksheetPart.Worksheet.Descendants<Cell>()` (ExcelFile.cs:57) + `SharedStringTable` DOM (L222). O DOM inteiro + o modelo MySheet coexistem no pico. Export e merge já foram streamados (fases 1–6 de `plans/excel-io-memory-time.md`); **o load é o único caminho DOM restante**.
- **Causa raiz #2 (CPU)**: cada fórmula re-tokeniza + re-parseia (`ExpressionParser.Parse("=" + f, sheet)`, ExcelFile.cs:138); escravas de shared formula sofrem shift TEXTUAL (regex+StringBuilder, SharedFormulaShifter.cs) + reparse completo. Não há cache de AST.
- **Causa raiz #3 (save residual)**: export aloca ~424MB — `double.ToString` por célula (`OpenXmlWriter.WriteString` só aceita string); merge usa `XmlWriter` puro que aceita `WriteChars(char[])` → dá para eliminar a string por célula com `double.TryFormat`.
- **Fatos verificados que moldam o design**:
  - O DOM do SDK é lazy por part: basta nunca tocar `worksheetPart.Worksheet`/`SharedStringTable` e ler via `part.GetStream()` + `XmlReader` — padrão já provado em `ExcelMerge.StreamMergeWorksheet` (ExcelMerge.cs:145).
  - `Sheet.SetCell` não tem side effects durante load (structural index/value store/range cache são lazy e nulos) — o modelo já é barato de popular.
  - AST é imutável (`sealed partial record`), memoização mora no `SheetValueStore` por `(sheet,col,row)` — **compartilhar o mesmo `Expression` entre N células é seguro** (dedup por texto).
  - `CellReference` NÃO tem flags `$` (parser faz `StripDollars`, Parser.cs:670-672) → shift no AST é inviável; mas o `$` sobrevive no TOKEN (Tokenizer.cs:108) → **shift por delta no nível do token** entrega o mesmo ganho sem tocar o wire MemoryPack.
  - `samples/` (JSONs confidenciais + k1.myxl) NÃO está neste clone → **decisão do usuário: gerar fixture sintético** não-confidencial; repo fica auto-suficiente.
  - Aspose local roda em modo avaliação (sem licença) — a referência principal de load são os **936ms externos** do usuário; o Aspose local é indicativo.
- Rejeitados (com justificativa nos designs): parse lazy no set (quebra contrato de ParseException eager + detecção estrutural do SetCell); flags `$` no CellReference (muda wire MemoryPack); cache estático process-wide de parse; System.IO.Packaging puro sem SDK (reimplementar relationships por ganho de µs).

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done,
set its status to `Complete` and write its **Phase Summary** (what was done, key
decisions, anything needed to continue with zero context); run the phase's
**Verification Plan** and record the result before moving on. When all phases are
done, fill in **Final Recap** and **Deployment Plan**.

Regras do repo: commits convencionais SEM referência ao assistente; hooks husky rodam
csharpier check + build (pre-commit) e + testes (pre-push); formatar com
`dotnet csharpier format .` antes de commitar.

## Phase 0: Fixture sintético K1-like + baseline de load
Status: Complete

- [x] Copiar este plano para `plans/excel-load-streaming.md` (repo root)
- [x] Criar `tools/SyntheticK1Builder/` (console, seeded/determinístico): escreve `samples/k1-synthetic.xlsx` **diretamente** via `ZipArchive` + `XmlWriter` (não via ExcelExport — precisamos controlar as shapes que o export não produz). Conteúdo (~566k células, 2+ sheets, perfil K1): coluna(s) de dados numéricos densos; ~40% células de fórmula organizadas em **grupos de shared formula reais** (`<f t="shared" ref si>` master + escravas só com `si`); texto com alta duplicação via shared strings (inclui rich-text run e strings com espaços nas pontas + `xml:space="preserve"`); algumas inline strings, booleans, erros, célula date ISO (`t="d"`)
- [x] Fazer o builder também salvar `samples/k1-synthetic.myxl` (via `ExcelFile.Load` + `workbook.Save`) para os cenários export/merge existentes do harness poderem rodar sem os JSONs confidenciais (fallback: harness tenta `k1.myxl`, senão `k1-synthetic.myxl`)
- [x] `ExcelMemoryHarness`: adicionar **Scenario L** (`ExcelFile.Load` do xlsx sintético, mantendo referência viva para `retained` refletir o modelo) e **Scenario AL** (`new AsposeWorkbook(xlsx)` load puro, ressalva de modo avaliação em comentário)
- [x] Rodar e **registrar baseline** (tempo/allocated/peak-live/retained de L e AL) no Phase Summary

### Verification Plan
- `dotnet run -c Release --project tools/SyntheticK1Builder` → gera os 2 fixtures; imprime contagem de células/fórmulas/strings
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --excel-memory` → Scenarios L e AL reportam números; anotar baseline
- Abrir o xlsx sintético no Excel/LibreOffice se disponível (sanidade manual, opcional)

### Phase Summary
Fixture sintético determinístico (seed fixa) com perfil K1: `tools/SyntheticK1Builder` escreve o xlsx bruto
(ZipArchive+XmlWriter) com 6 grupos de shared formula reais (masters na linha 2, 60k linhas), shared strings
com duplicação alta (208 únicas / 60k refs, incluindo padded + rich text), inline string, boolean, erro,
date ISO e 2 defined names. Sanidade: loader atual releu exatamente 620.012 células (360.001 fórmulas).
Harness ganhou Scenario L (ExcelFile.Load do xlsx sintético) e AL (Aspose load puro) + fallback
k1.myxl → k1-synthetic.myxl.

**BASELINE (2026-07-10, Release):**

| Cenário | tempo | allocated | peak-live | retained |
|---|---:|---:|---:|---:|
| MySheet ExcelFile.Load | **5,72s** | **1.800 MB** | **918 MB** | 256 MB |
| Aspose load (no save, eval mode) | 0,62s | 177 MB | 402 MB | 399 MB |
| MySheet ExcelExport | 2,45s | 269 MB | 369 MB | 272 MB |
| Aspose load+save | 3,10s | 241 MB | 381 MB | 248 MB |
| MySheet ExcelMerge | 1,33s | 308 MB | 432 MB | 277 MB |

Nota: o gap de load neste perfil (9x) é maior que o externo do usuário (2,2x) porque o fixture tem 360k
escravas de shared formula — o pior caminho do loader atual (regex shift textual + reparse por escrava).
Referência externa licenciada: Aspose ~936ms no arquivo real do usuário.

## Phase 1: Shared strings streaming
Status: Not started

- [ ] Criar `Danfma.MySheet.Excel/SharedStringsStreamReader.cs`: `internal static IReadOnlyList<string> Read(SharedStringTablePart? part)` — `part.GetStream()` + `XmlReader` forward-only; `uniqueCount` do `<sst>` como *hint* de presize (nunca limite); caso comum 1 `<t>` → `ReadElementContentAsString` direto; rich text → `StringBuilder` reutilizado concatenando TODOS os `<t>` sob `<si>` (paridade com `InnerText`, incluindo `<rPh>` — anotar em comentário); `<t/>` vazio → `""`; NÃO usar `IgnoreWhitespace` (preserva `xml:space`)
- [ ] `ExcelFile.Load`: trocar `LoadSharedStrings` pelo reader novo; não tocar mais em `.SharedStringTable`

### Verification Plan
- `dotnet build Danfma.MySheet.slnx -c Release` → 0 warnings
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/... -c Release` → 29+ pass
- `--excel-memory` → Scenario L: allocated/peak caem vs baseline (shared strings DOM eliminado)

### Phase Summary
_(write when phase completes)_

## Phase 2: Worksheet streaming (o grosso)
Status: Not started

- [ ] Criar `Danfma.MySheet.Excel/WorksheetStreamLoader.cs`: `internal static void Load(WorksheetPart part, Sheet sheet, IReadOnlyList<string> sharedStrings)` — `part.GetStream()` + `XmlReader`, matching por `LocalName` (imune a prefixo, como o merge); `Skip()` de tudo fora de `sheetData`
- [ ] Scratch struct por célula (`Reference/Type/FormulaText/IsSharedFormula/SharedIndex/Value/InlineText`); ler `f`, `v`, `is` antes de decidir (o `<v>` é necessário para o fallback de escrava órfã)
- [ ] Portar a lógica de decisão 1:1 de `LoadCell`/`LoadLiteral` (mapa de tipos: null/"n"→Number, "s"→SharedString, "b"→Boolean, "e"→Error, "d"→Date ISO→ToOADate senão string, "inlineStr", "str"/desconhecido→string; célula sem v/f/is → não armazenada; cached `<v>` de fórmula ignorado)
- [ ] Suportar células/linhas **sem `@r`** (referência implícita por posição — melhoria sobre o loader atual que as ignora): rastrear `currentRow`/`nextColumn`
- [ ] Shared formulas: dicionário `si → (masterId, texto)` como hoje (document order garante master antes das escravas, ECMA-376 §18.3.1.40); caso patológico → `pendingSlaves` lazy `List<(Id, Si, Expression? CachedLiteral)>` resolvida ao fim do `sheetData` (superset do comportamento atual)
- [ ] `ExcelFile.Load`: substituir loop `Descendants<Cell>` pelo loader novo; remover `LoadCell`/`LoadLiteral`; `LoadDefinedNames` e navegação de `workbook.xml` ficam DOM (part minúscula); atualizar `<summary>` (stream precisa ser readable+seekable — já era exigido)

### Verification Plan
- Suíte completa: `dotnet run --project tests/Danfma.MySheet.Tests/... -c Release` (1044) + Excel.Tests (29+) → 0 fail
- `--excel-memory` → Scenario L: queda grande de peak-live e allocated (DOM de worksheet eliminado); tempo cai; registrar tabela vs baseline

### Phase Summary
_(write when phase completes)_

## Phase 3: Testes de hardening
Status: Not started

- [ ] Novo `tests/Danfma.MySheet.Excel.Tests/StreamingLoadEdgeTests.cs`: `xml:space` (espaços nas pontas, célula só-espaço), shared string rich-text, inline string, `t="str"`, date ISO
- [ ] Fixture manual (escrita bruta de part via `WorksheetPart.GetStream` num builder de teste): células sem `@r`, prefixo `x:`, escrava antes do master, `<f>` vazia com `<v>` cacheado
- [ ] Teste de paridade round-trip: load(export(load(x))) equivalente

### Verification Plan
- Ambas as suítes verdes; novos testes cobrem cada edge listado

### Phase Summary
_(write when phase completes)_

## Phase 4: Parse quick wins
Status: Not started

- [ ] `ExpressionParser.ParseFormulaBody(string, Sheet)` (pula exigência de `=`); usar nos call sites do load (3×) e `Indirect.cs:61` — elimina o concat `"=" + formula` por célula
- [ ] **Dedup de AST por texto de fórmula**: `Dictionary<string, Expression>` POR SHEET dentro do load (AST imutável + memoização externa = sharing seguro; NÃO aplicar a defined names — contexto de sheet diferente); economiza parse E memória em fórmulas repetidas
- [ ] `Parser.Functions` → `FrozenDictionary`
- [ ] `Tokenizer.Tokenize`: capacidade inicial do `List<Token>` (`text.Length/3+4`)

### Verification Plan
- 1044 + 29+ testes verdes
- `--excel-memory` Scenario L: tempo cai (arquivo sintético tem fórmulas duplicadas → dedup mensurável)

### Phase Summary
_(write when phase completes)_

## Phase 5: Shared formulas por delta de token
Status: Not started

- [ ] ANTES: testes congelando paridade do `SharedFormulaShifter` atual (delta negativo além do limite → comportamento atual documentado, `$` misto, ref qualificada por sheet, identificador seguido de `(` ou `!`, open ranges `A:A` NÃO deslocados)
- [ ] Master registra `List<Token>` tokenizada 1× (Token é readonly record struct — reutilizável entre escravas)
- [ ] `Parser` ganha modo delta `(deltaRow, deltaColumn)`: nos 3 pontos onde token vira referência (Parser.cs:566-569, 641, 648-666), `NormalizeCellId` delta-aware lê `($?, letras, $?, dígitos)` do texto do token (que preserva `$`), desloca só componentes relativos
- [ ] Escravas: parse da token list do master com delta — elimina regex+StringBuilder+string do shift E a re-tokenização por escrava
- [ ] Aposentar `SharedFormulaShifter` (testes redirecionados) quando a paridade passar

### Verification Plan
- Testes de paridade da fase (congelados antes) passam no caminho novo; suítes completas verdes
- `--excel-memory` Scenario L: tempo cai de novo (fixture tem grupos de shared formula reais)

### Phase Summary
_(write when phase completes)_

## Phase 6: Presize do CellStore (medido, reversível)
Status: Not started

- [ ] `Danfma.MySheet.csproj`: `+ <InternalsVisibleTo Include="Danfma.MySheet.Excel" />`
- [ ] `CellStore.EnsureDenseCapacity(int)` + `Sheet.EnsureCellCapacity(int)` (internal; `Dictionary.EnsureCapacity` não altera ordem de inserção → wire MemoryPack intacto)
- [ ] `WorksheetStreamLoader`: no `<dimension>`, `EnsureCellCapacity(min(bboxCells, 524_288))` com overflow-check (bbox é bounding box, não contagem real — cap limita desperdício em sheet esparsa a ~25MB liberáveis)
- [ ] MANTER SÓ SE o harness mostrar ganho sem piorar cenário esparso (sanity: fixture 1×1 com dimension gigante)

### Verification Plan
- Suítes verdes; `--excel-memory` compara com/sem presize; decisão documentada no Phase Summary

### Phase Summary
_(write when phase completes)_

## Phase 7: Residual do save
Status: Not started

- [ ] **Merge** (caminho de produção): em `WriteCell`/`WriteNewRow` (ExcelMerge.cs:475-529, 429), trocar `ToString` por `double.TryFormat`/`int.TryFormat` em `char[32]` reutilizado + `writer.WriteChars` (dígitos não precisam escaping). NÃO usar formato "R" explícito — o TryFormat default = shortest round-trip = paridade com ToString atual
- [ ] `SmallNumberStrings` interno compartilhado (cache `string[]` para inteiros 0..1023, fast-path `d == (int)d`) usado por merge, export e índices de shared string
- [ ] **Export**: aplicar `SmallNumberStrings`; trocar o LINQ `Select/OrderBy` com tupla por célula (ExcelExport.cs:158-163) por array pré-ordenado sem iterators
- [ ] Medir; decisão documentada: reescrever export para `XmlWriter` puro (padrão do merge) OU aceitar residual — só reescrever se o número pós-fase ainda incomodar

### Verification Plan
- Suítes verdes (round-trip de merge/export já coberto); `--excel-memory`: allocated de export e merge caem vs fase anterior

### Phase Summary
_(write when phase completes)_

## Phase 8: Fechamento (+ fase opcional gated)
Status: Not started

- [ ] **GATED (só se Scenario L ainda > ~936ms referência externa)**: dieta de alocação do tokenizer — tokens por span `(Type, Start, Length)` + `FrozenDictionary.GetAlternateLookup<ReadOnlySpan<char>>`; risco médio, 1044 testes cobrem
- [ ] Tabela final de números (baseline → cada fase → final) vs referência Aspose no Phase Summary
- [ ] Atualizar CHANGELOG.md, `<summary>` de `ExcelFile`, e `plans/excel-load-streaming.md` (recap)
- [ ] `dotnet csharpier format .` + build + suítes completas + push (release manual via workflow fica a critério do usuário)

### Verification Plan
- Build 0 warnings; 1044 + 29+ + novos testes verdes; `--excel-memory` final registrado
- Comparação final: load MySheet vs 936ms (meta: fechar ou superar; modelo MySheet é mais barato de popular que DOM — plausível)

### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete: summary of the entire piece of work)_

## Deployment Plan
_(write when all phases complete: step-by-step deployment instructions)_
