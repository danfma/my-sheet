# Reduzir memória e tempo do save/merge Excel (OpenXML)

Cortar alocações, pico de memória e wall-time de `ExcelExport` (save do zero) e
`ExcelMerge` (in-place) em workbooks grandes (fixture K1, ~566k células),
**permanecendo no `DocumentFormat.OpenXml`** — sem trocar de biblioteca e sem
licença comercial. Decisões do usuário: (1) otimizar OpenXML; (2) atacar Merge e
Export igualmente.

## Baseline medido (k1.myxl, sem licença Aspose)

Harness: `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --excel-memory`

| Cenário | Tempo | Alocado | Pico vivo | Retido |
|---|---|---|---|---|
| ExcelExport (save do zero) | 5,52s | 891 MB | 1.042 MB | 508 MB |
| ExcelMerge (no existente) | 32,34s | 3.865 MB | 1.273 MB | 496 MB |

Notas de baseline:
- O workbook MySheet carregado retém ~500 MB; é o piso de qualquer cenário.
- O baseline Aspose (load+save = 1,75s / 70 MB) é **inválido**: modo avaliação
  trunca linhas e injeta watermark "Evaluation Only", então gravou arquivo
  mutilado. Não usar como meta sem licença.

## Causas-raiz confirmadas

1. **Merge O(n²) [Certain].** `ExcelMerge.GetOrCreateRow`/`GetOrCreateCell` fazem
   varredura linear (`sheetData.Elements<Row>()`, `row.Elements<Cell>()`) desde o
   início para cada célula. Com o alvo já preenchido, achar a linha R custa R
   passos → O(linhas²). Domina os 32s.
2. **Materialização DOM [Certain].** `worksheetPart.Worksheet` materializa a
   planilha inteira como árvore de objetos (cada `Cell`/`Row`/`CellValue` é objeto
   no heap). Tanto Export quanto Merge pagam isso.
3. **Dicionário de valores up-front no Export [Certain].** `SaveAsExcel` monta um
   `Dictionary<(Sheet,Id),ComputedValue>` com TODOS os valores (ExcelExport.cs
   L53-66) além do DOM — terceira cópia de tudo. O Merge já computa on-demand.

## For Future Agents
Marque `- [x]` conforme completa; ao fechar uma fase, defina Status `Complete`,
escreva o **Phase Summary** e rode o **Verification Plan** registrando o
resultado. Ao fim de todas as fases, preencha **Final Recap** e **Deployment
Plan**. O harness `--excel-memory` é a ferramenta de verificação principal; a
suíte `Danfma.MySheet.Excel.Tests` (atualmente 28 testes) é o gate de correção.
NÃO comitar o andaime temporário até a fase de limpeza decidir seu destino.

## Phase 1: Merge — eliminar o O(n²) de lookup de linha/célula
Status: Complete

- [x] Em `ExcelMerge.cs`, construir uma vez por planilha um índice
      `Dictionary<uint, Row>` (RowIndex → Row) a partir do `sheetData` existente,
      e para cada linha um `Dictionary<uint, Cell>` (coluna → Cell) sob demanda.
- [x] Reescrever `GetOrCreateRow` para consultar/inserir via índice, mantendo a
      invariante de ordenação por `RowIndex` (inserção ordenada só quando cria
      linha nova; atualizar o índice).
- [x] Reescrever `GetOrCreateCell` análogo, mantendo ordenação por coluna dentro
      da linha e atualizando o índice ao criar célula nova.
- [x] Garantir que a semântica atual não muda: blanks pulados, estilos
      preservados, calcChain removido, formulas dropadas.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj`
  → todos os testes passam (≥28), sem regressão em `ExcelMergeTests`.
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --excel-memory`
  → tempo do Merge cai de ~32s para < 5s (meta), alocação cai substancialmente.

### Phase Summary
Substituídos os scans lineares `GetOrCreateRow`/`GetOrCreateCell` por duas classes
de índice (`SheetDataIndex` e `RowCells`) que mapeiam RowIndex→Row e coluna→Cell
em `Dictionary`. O caminho quente (writes em ordem ascendente num alvo já
preenchido) vira lookup O(1); a criação de linha/célula nova usa o máximo
rastreado para dar append O(1) no caso comum, com fallback de `InsertBefore`
ordenado só no raro insert fora de ordem.

**Resultado medido (K1):** tempo do Merge 32,34s → **4,97s** (6,5x); alocado
3.865 MB → **1.773 MB** (2,2x). Correção: 28/28 testes Excel passam. O pico vivo
(1.320 MB) permanece — é a materialização do DOM, alvo da Fase 3. Sem mudança de
semântica (estilos/blanks/calcChain preservados; verificado pelos testes de
merge).

## Phase 2: Export — streaming com OpenXmlWriter + remover o dict up-front
Status: Complete

- [x] Remover o `Dictionary<(Sheet,Id),ComputedValue> values` de `SaveAsExcel`;
      computar cada valor sob demanda no thread de large-stack, no momento da
      escrita (espelhando o padrão já usado em `ExcelMerge`).
- [x] Reescrever `BuildSheetData`/loop de escrita para usar `OpenXmlWriter` sobre
      o `WorksheetPart` (`WriteWorksheet`), emitindo `<row>`/`<c>` em streaming,
      sem montar a árvore `SheetData` em memória.
- [x] Manter `SharedStringRegistry` (limitado por strings distintas) e gravar o
      `SharedStringTablePart` ao final (índices já atribuídos durante a escrita).
- [x] Preservar exatamente o output atual: modos ValuesOnly/Formulas, cached
      values, `xml:space="preserve"` (via `XlsxTextFactory`), defined names.

### Verification Plan
- `dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj`
  → todos passam; em especial os round-trips de `ExcelExportTests` e o teste de
  whitespace.
- `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --excel-memory`
  → alocação do Export cai de 891 MB para < ~300 MB (meta) e o pico cai.

### Phase Summary
`SaveAsExcel(stream)` reescrito: todo o corpo roda no thread de large-stack,
computa valores sob demanda (dict up-front removido) e escreve cada planilha via
`OpenXmlWriter` (`WriteWorksheet`), abrindo/fechando `<row>` no boundary e
emitindo cada `<c>` sem reter o `SheetData`. `SharedStringRegistry.WriteTo`
inalterado (roda após as planilhas, quando o registro já está preenchido).

**Resultado medido (K1):** tempo 5,20s → **3,03s** (1,7x); pico vivo 1.044 MB →
**748 MB** (−28%, sem DOM materializado); alocado 890 MB → 730 MB (−18%).
Correção: 28/28 testes Excel passam (round-trips, oráculo ClosedXML, whitespace).

**Meta de alocação (<300 MB) NÃO atingida.** Causa raiz: `BuildCell` ainda aloca
um `Cell`/`CellValue` DOM transitório por célula (~566k). São efêmeros (por isso o
*pico* caiu), logo é pressão de GC, não footprint. Fechar a lacuna exige escrever
os atributos direto no `OpenXmlWriter` (sem construir objetos `Cell`) — registrado
como refinamento opcional na Fase 5.

## Phase 3 (condicional): Merge — read+write em streaming para cortar o DOM
Status: Not started

Executar SOMENTE se, após a Fase 1, o pico de memória do Merge continuar
inaceitável (DOM materializado ~ +777 MB sobre o workbook retido).

- [ ] Ler a planilha existente com `OpenXmlReader` e escrever uma substituta com
      `OpenXmlWriter`, fazendo merge-join dos valores computados em ordem de
      célula; repassar intactos todos os demais elementos (`cols`, `mergeCells`,
      `dataValidations`, `sheetPr`, etc.).
- [ ] Trocar a parte antiga pela nova (via arquivo temporário ou re-link do
      `WorksheetPart`), preservando relacionamentos, estilos e a remoção do
      calcChain.
- [ ] Validar preservação de tudo que não é célula computada (formatação,
      validações, merges).

### Verification Plan
- Testes de `ExcelMergeTests` passam + um novo teste que faz merge num template
  com `mergeCells`/validação e confirma que sobrevivem.
- `--excel-memory` → pico do Merge tende ao piso do workbook retido (~500 MB).

### Phase Summary
_(escrever ao completar)_

## Phase 4: Limpeza do andaime
Status: Not started

- [ ] Decidir destino do `ExcelMemoryHarness.cs` + dispatch em `Program.cs` +
      `ProjectReference` do Excel no benchmark: manter como benchmark
      versionado (documentado) OU remover. Registrar a decisão.
- [ ] `git status` limpo (fora as mudanças intencionais).

### Verification Plan
- `dotnet build -c Release` da solução sem warnings novos.

### Phase Summary
_(escrever ao completar)_

## Phase 5 (opcional): Export — escrita de célula sem DOM transitório
Status: Not started

Fechar a lacuna de alocação (730 MB → alvo < 300 MB) da Fase 2.

- [ ] Reescrever a escrita de célula para emitir `<c>`/`<f>`/`<v>` direto no
      `OpenXmlWriter` via `WriteStartElement(element, atributos)` + `WriteString`,
      sem construir objetos `Cell`/`CellValue` por célula (reutilizar elementos
      "template" como nome de tag, passando atributos explícitos).
- [ ] Preservar exatamente os atributos/tipos atuais (`t="b|s|e|str"`, cached
      values do modo Formulas) e o comportamento de blank (omissão).

### Verification Plan
- Suíte Excel 28/28 passa (round-trips + oráculo ClosedXML detectam qualquer
  divergência de output).
- `--excel-memory` → alocação do Export cai para < ~300 MB.

### Phase Summary
_(escrever ao completar)_

## Final Recap
_(escrever ao completar todas as fases)_

## Deployment Plan
_(escrever ao completar todas as fases)_
