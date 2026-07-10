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
Status: Complete

- [x] Criar `Danfma.MySheet.Excel/SharedStringsStreamReader.cs`: `internal static IReadOnlyList<string> Read(SharedStringTablePart? part)` — `part.GetStream()` + `XmlReader` forward-only; `uniqueCount` do `<sst>` como *hint* de presize (nunca limite); caso comum 1 `<t>` → `ReadElementContentAsString` direto; rich text → `StringBuilder` reutilizado concatenando TODOS os `<t>` sob `<si>` (paridade com `InnerText`, incluindo `<rPh>` — anotar em comentário); `<t/>` vazio → `""`; NÃO usar `IgnoreWhitespace` (preserva `xml:space`)
- [x] `ExcelFile.Load`: trocar `LoadSharedStrings` pelo reader novo; não tocar mais em `.SharedStringTable`

### Verification Plan
- `dotnet build Danfma.MySheet.slnx -c Release` → 0 warnings
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/... -c Release` → 29+ pass
- `--excel-memory` → Scenario L: allocated/peak caem vs baseline (shared strings DOM eliminado)

### Phase Summary
`SharedStringsStreamReader.Read` streama o sst com XmlReader forward-only (uniqueCount como hint com cap
1<<20; caso comum 1 <t> sem StringBuilder; rich text/rPh concatenados = paridade com InnerText; xml:space
intacto — reader nunca faz trim). `ExcelFile` não toca mais `.SharedStringTable`. Build 0 warnings, 29/29
Excel tests. Scenario L: 5,72s → 5,33s (allocated ~igual: fixture tem só 208 strings únicas — o ganho real
desta fase aparece em arquivos texto-pesados; o grosso do custo é o DOM de worksheet, alvo da Fase 2).

## Phase 2: Worksheet streaming (o grosso)
Status: Complete

- [x] Criar `Danfma.MySheet.Excel/WorksheetStreamLoader.cs`: `internal static void Load(WorksheetPart part, Sheet sheet, IReadOnlyList<string> sharedStrings)` — `part.GetStream()` + `XmlReader`, matching por `LocalName` (imune a prefixo, como o merge); `Skip()` de tudo fora de `sheetData`
- [x] Scratch struct por célula (`Reference/Type/FormulaText/IsSharedFormula/SharedIndex/Value/InlineText`); ler `f`, `v`, `is` antes de decidir (o `<v>` é necessário para o fallback de escrava órfã)
- [x] Portar a lógica de decisão 1:1 de `LoadCell`/`LoadLiteral` (mapa de tipos: null/"n"→Number, "s"→SharedString, "b"→Boolean, "e"→Error, "d"→Date ISO→ToOADate senão string, "inlineStr", "str"/desconhecido→string; célula sem v/f/is → não armazenada; cached `<v>` de fórmula ignorado)
- [x] Suportar células/linhas **sem `@r`** (referência implícita por posição — melhoria sobre o loader atual que as ignora): rastrear `currentRow`/`nextColumn`
- [x] Shared formulas: dicionário `si → (masterId, texto)` como hoje (document order garante master antes das escravas, ECMA-376 §18.3.1.40); caso patológico → `pendingSlaves` lazy `List<(Id, Si, Expression? CachedLiteral)>` resolvida ao fim do `sheetData` (superset do comportamento atual)
- [x] `ExcelFile.Load`: substituir loop `Descendants<Cell>` pelo loader novo; remover `LoadCell`/`LoadLiteral`; `LoadDefinedNames` e navegação de `workbook.xml` ficam DOM (part minúscula); atualizar `<summary>` (stream precisa ser readable+seekable — já era exigido)

### Verification Plan
- Suíte completa: `dotnet run --project tests/Danfma.MySheet.Tests/... -c Release` (1044) + Excel.Tests (29+) → 0 fail
- `--excel-memory` → Scenario L: queda grande de peak-live e allocated (DOM de worksheet eliminado); tempo cai; registrar tabela vs baseline

### Phase Summary
`WorksheetStreamLoader` streama cada worksheet com XmlReader forward-only (contrato consome-e-passa por
elemento; matching por LocalName; Skip() fora de sheetData). Decisão de célula portada 1:1 (fórmula /
shared master / escrava via Shift+parse / literal por @t); `<is>` reaproveita
`SharedStringsStreamReader.ReadFlattenedText` (agora container-agnóstico). Células/linhas sem @r suportadas
(posição implícita via CellId.Format — novo). Escrava-antes-do-master deferida com literal decodificado
eager (superset do DOM). `ExcelFile` perdeu LoadCell/LoadLiteral; summary atualizado (streaming, stream
seekable).

**Resultado (vs baseline)**: 1.044+29 testes verdes; load **5,72s → 1,77s (3,2x)**, allocated
**1.800 → 1.196 MB (−34%)**, peak-live **918 → 495 MB (−46%)**. Paridade de conteúdo PROVADA: MemoryPack
wire byte-idêntico (40.634.922 bytes) entre DOM e streaming.

**Nota sobre `retained`**: o valor do harness subiu (256→470 MB), mas é artefato de contabilidade
(GetTotalMemory reflete regiões comprometidas moldadas pelo padrão de alocação — soltar o workbook inteiro
não muda o número em NENHUMA versão, mesmo com GC.Collect Aggressive). Modelo real idêntico (prova acima).

## Phase 3: Testes de hardening
Status: Complete

- [x] Novo `tests/Danfma.MySheet.Excel.Tests/StreamingLoadEdgeTests.cs`: `xml:space` (espaços nas pontas, célula só-espaço), shared string rich-text, inline string, `t="str"`, date ISO
- [x] Fixture manual (escrita bruta de part via `WorksheetPart.GetStream` num builder de teste): células sem `@r`, prefixo `x:`, escrava antes do master, `<f>` vazia com `<v>` cacheado
- [x] Teste de paridade round-trip: load(export(load(x))) equivalente

### Verification Plan
- Ambas as suítes verdes; novos testes cobrem cada edge listado

### Phase Summary
`StreamingLoadEdgeTests` (10 testes, todos verdes de primeira): fixture bruto via ZipArchive (controla
prefixos/atributos/ordem que um writer DOM normalizaria) + fixtures ClosedXML. Cobertos: posições implícitas
(célula/linha sem @r, inclusive retomada após @r explícito), prefixo x:, escrava-antes-do-master (resolve
via deferral — capacidade nova), escrava sem master → literal cacheado, <f/> vazia não-shared → literal,
inline string com espaços, shared string padded + rich text multi-run, t="str" sem <f>, date ISO → serial,
e round-trip load→export→load. Suíte Excel: 39/39.

## Phase 4: Parse quick wins
Status: Complete

- [x] `ExpressionParser.ParseFormulaBody(string, Sheet)` (pula exigência de `=`); usar nos call sites do load (3×) e `Indirect.cs:61` — elimina o concat `"=" + formula` por célula
- [x] **Dedup de AST por texto de fórmula**: `Dictionary<string, Expression>` POR SHEET dentro do load (AST imutável + memoização externa = sharing seguro; NÃO aplicar a defined names — contexto de sheet diferente); economiza parse E memória em fórmulas repetidas
- [x] `Parser.Functions` → `FrozenDictionary`
- [x] `Tokenizer.Tokenize`: capacidade inicial do `List<Token>` (`text.Length/3+4`)

### Verification Plan
- 1044 + 29+ testes verdes
- `--excel-memory` Scenario L: tempo cai (arquivo sintético tem fórmulas duplicadas → dedup mensurável)

### Phase Summary
`ParseFormulaBody` público (elimina DUAS cópias por fórmula: o concat "="+body no caller e o slice [1..]
de volta no parser); migrados os call sites do load (WorksheetStreamLoader ×3, defined names) e
Indirect.cs. Dedup de AST por texto num `FormulaCache` por sheet (LoadContext novo agrupa o estado do
load); escravas shifted não passam pelo cache (texto único por célula). `Parser.Functions` →
FrozenDictionary; `Tokenizer` presized (len/3+4). 1.044+39 testes verdes. Scenario L: allocated
1.196 → 1.108 MB; tempo estável em 1,77s — esperado: o CPU dominante são as 360k escravas
(shift textual + reparse), alvo da Fase 5. O dedup rende em arquivos com fórmulas literais repetidas
(o fixture quase não tem).

## Phase 5: Shared formulas por delta de token
Status: Complete

- [x] ANTES: testes congelando paridade do `SharedFormulaShifter` atual (delta negativo além do limite → comportamento atual documentado, `$` misto, ref qualificada por sheet, identificador seguido de `(` ou `!`, open ranges `A:A` NÃO deslocados)
- [x] Master registra `List<Token>` tokenizada 1× (Token é readonly record struct — reutilizável entre escravas)
- [x] `Parser` ganha modo delta `(deltaRow, deltaColumn)`: nos 3 pontos onde token vira referência (Parser.cs:566-569, 641, 648-666), `NormalizeCellId` delta-aware lê `($?, letras, $?, dígitos)` do texto do token (que preserva `$`), desloca só componentes relativos
- [x] Escravas: parse da token list do master com delta — elimina regex+StringBuilder+string do shift E a re-tokenização por escrava
- [x] Aposentar `SharedFormulaShifter` (testes redirecionados) quando a paridade passar

### Verification Plan
- Testes de paridade da fase (congelados antes) passam no caminho novo; suítes completas verdes
- `--excel-memory` Scenario L: tempo cai de novo (fixture tem grupos de shared formula reais)

### Phase Summary
Paridade congelada ANTES em `SharedFormulaDeltaTests` (6 testes: todas as combinações de $, grupo 2D,
função/sheet-qualifier nunca deslocados, open range parado, literal de string intacto, range com os dois
endpoints). `Parser` ganhou modo delta (ctor `deltaRow/deltaColumn`); `NormalizeReference` aplica o shift
nos 3 pontos onde token vira referência; `ShiftCellId` é port de paridade exata do shape do shifter
(`$?LLL$?DDD` — fora do shape normaliza SEM shift; row absoluto preserva dígitos verbatim).
`ExpressionParser.TokenizeFormulaBody`/`ParseSharedFormulaBody` internos (InternalsVisibleTo p/ Excel
antecipado da Fase 6). Master tokenizado 1×; escravas re-parseiam a token list com delta — zero shift
textual, zero re-tokenização.

**DESVIO DOCUMENTADO**: `SharedFormulaShifter` NÃO foi aposentado — rebaixado a fallback para delta
negativo (só produtores fora da spec; na spec o master é a primeira célula do ref → deltas ≥ 0, o que
torna underflow impossível no caminho novo). Paridade absoluta com risco zero.

**Resultado**: 1.044+45 testes verdes. Load: 1,77s → **1,21s**; allocated 1.108 → **305 MB (3,6x)**.
Acumulado vs baseline: tempo 5,72 → 1,21s (**4,7x**), allocated 1.800 → 305 MB (**5,9x**).

## Phase 6: Presize do CellStore (medido, reversível)
Status: Complete

- [x] `Danfma.MySheet.csproj`: `+ <InternalsVisibleTo Include="Danfma.MySheet.Excel" />` (antecipado na Fase 5)
- [x] `CellStore.EnsureDenseCapacity(int)` + `Sheet.EnsureCellCapacity(int)` (internal; `Dictionary.EnsureCapacity` não altera ordem de inserção → wire MemoryPack intacto)
- [x] `WorksheetStreamLoader`: no `<dimension>`, `EnsureCellCapacity(min(bboxCells, 524_288))` com overflow-check (bbox é bounding box, não contagem real — cap limita desperdício em sheet esparsa a ~25MB liberáveis)
- [x] MANTER SÓ SE o harness mostrar ganho sem piorar cenário esparso (sanity: fixture 1×1 com dimension gigante)

### Verification Plan
- Suítes verdes; `--excel-memory` compara com/sem presize; decisão documentada no Phase Summary

### Phase Summary
`CellStore.EnsureDenseCapacity` + `Sheet.EnsureCellCapacity` (internos; EnsureCapacity não reordena →
wire intacto). Loader lê `<dimension>` e reserva `min(bboxCells, 512k)` com overflow-check e
FormatException engolida (hint, não contrato). MEDIDO E MANTIDO: allocated 305 → 272 MB, peak-live
515 → 491 MB, tempo estável (2 runs). Sanity de sheet esparsa (dimension A1:XFD1048576 com 1 célula)
coberto por teste. Suíte Excel: 46/46.

## Phase 7: Residual do save
Status: Complete

- [x] **Merge** (caminho de produção): em `WriteCell`/`WriteNewRow` (ExcelMerge.cs:475-529, 429), trocar `ToString` por `double.TryFormat`/`int.TryFormat` em `char[32]` reutilizado + `writer.WriteChars` (dígitos não precisam escaping). NÃO usar formato "R" explícito — o TryFormat default = shortest round-trip = paridade com ToString atual
- [x] `SmallNumberStrings` interno compartilhado (cache `string[]` para inteiros 0..1023, fast-path `d == (int)d`) usado por merge, export e índices de shared string
- [x] **Export**: aplicar `SmallNumberStrings`; trocar o LINQ `Select/OrderBy` com tupla por célula (ExcelExport.cs:158-163) por array pré-ordenado sem iterators
- [x] Medir; decisão documentada: reescrever export para `XmlWriter` puro (padrão do merge) OU aceitar residual — só reescrever se o número pós-fase ainda incomodar

### Verification Plan
- Suítes verdes (round-trip de merge/export já coberto); `--excel-memory`: allocated de export e merge caem vs fase anterior

### Phase Summary
Helper `XlsxNumbers` (nasceu como o "SmallNumberStrings" do plano): cache "0".."1023" + overloads
`Write(XmlWriter, double/int)` com `double.TryFormat` em buffer [ThreadStatic] + `WriteChars` (dígitos
nunca precisam de escaping; TryFormat default = shortest round-trip = paridade byte com ToString).
Merge: número, índice de shared string e @r de linha nova escritos sem string. Export: `Format()` com
cache (OpenXmlWriter continua exigindo string para o resto) + LINQ Select/OrderBy substituído por sort
de chave empacotada (row<<14|col) com payload paralelo — zero iterators/tuple boxing.

**Resultado**: export 269 → 255 MB, merge 308 → 286 MB alocados; round-trip Aspose 0 mismatches
(paridade de formatação PROVADA); 1.044+46 testes verdes. No fixture (decimais aleatórios) o cache de
inteiros rende pouco — em planilhas de negócio (inteiros dominam) rende mais.

**DECISÃO (export → XmlWriter puro): NÃO reescrever agora.** O residual do export é o
`double.ToString` de decimais não-inteiros, inerente ao OpenXmlWriter; export já está em 1,6s / 255 MB
(vs Aspose load+save 2,0s / 241 MB). A reescrita replicaria o merge inteiro por ~30-40 MB — não paga o
risco neste momento. Reavaliar se o export virar gargalo do usuário.

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
