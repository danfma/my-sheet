# Serialização (MemoryPack)

*Tradução do documento canônico em inglês ([serialization.md](../serialization.md)). Em caso de divergência, o inglês prevalece.*

Um `Workbook` é serializado em um formato binário compacto via
[MemoryPack](https://github.com/Cysharp/MemoryPack). Esta é a persistência *nativa* do MySheet — rápida
para escrever, rápida para carregar, e ela preserva no round-trip as árvores de expressão completas (não
apenas os valores). Ela não tem relação com `.xlsx`; para arquivos Excel, veja
[Interop com Excel](excel-interop.md).

## Salvar e carregar

```csharp
using Danfma.MySheet;

workbook.Save("model.mysheet");
Workbook restored = Workbook.Load("model.mysheet");

// Sobrecargas assíncronas:
await workbook.SaveAsync("model.mysheet", cancellationToken);
Workbook restoredAsync = await Workbook.LoadAsync("model.mysheet", cancellationToken);
```

`Load`/`LoadAsync` lançam `InvalidDataException` se o arquivo não contiver um workbook. A extensão do
arquivo é escolha sua — os exemplos usam `.mysheet` por convenção.

## O que é preservado no round-trip — e o que não é

A coluna **frio** é o `Save` padrão; a coluna **aquecido** é um save com
[`IncludeComputedValues`](#warm-start-persistindo-valores-computados) (tudo o que um save frio persiste,
mais os valores memoizados).

| | Frio | Aquecido | Observações |
| --- | --- | --- | --- |
| Planilhas (nome, ordem das abas) | Sim | Sim | A busca de nomes case-insensitive é restaurada na desserialização. |
| Células e árvores de expressão completas | Sim | Sim | Fórmulas continuam sendo fórmulas — um workbook carregado continua recalculando. |
| **Chamadas** de funções personalizadas (nós `FunctionCall`) | Sim | Sim | O nome e as expressões dos argumentos são preservados. |
| **Implementações** de funções personalizadas (delegates) | **Não** | **Não** | Comportamento é código, não dados — registre de novo após carregar. |
| Cache de memoização | **Não** | **Parcial** | O frio recalcula de forma preguiçosa na primeira leitura. O aquecido restaura o cache — exceto células voláteis e do tipo referência (abaixo), que ainda são recalculadas. |

A consequência prática: se o seu workbook usa [funções personalizadas](custom-functions.md), registre-as
novamente após cada `Load`, ou essas chamadas serão avaliadas como `#NAME?`:

```csharp
var restored = Workbook.Load("model.mysheet");

restored.RegisterFunction("CUSTOM", (arguments, wb) =>
{
    var a = arguments[0].Evaluate(wb).AsDouble() ?? 0;
    var b = arguments[1].Evaluate(wb).AsDouble() ?? 0;

    return a + b;
});

double value = restored.GetCellValue("Sheet1", "A1").ToDouble();
```

## Warm-start: persistindo valores computados

Por padrão, um arquivo salvo contém **apenas o modelo** — todo valor é recalculado de forma preguiçosa na
primeira leitura após o carregamento. Passe `WorkbookSaveOptions { IncludeComputedValues = true }` para
também persistir o cache de memoização, de modo que o carregamento comece **aquecido** — o **warm-start**
(inicialização aquecida) — e sirva células já computadas sem reavaliá-las:

```csharp
workbook.Save("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true });
// await workbook.SaveAsync("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true }, ct);

var warm = Workbook.Load("model.mysheet"); // lida de volta com o cache já preenchido
```

`Load`/`LoadAsync` não precisam de nenhuma flag — eles detectam o formato a partir do cabeçalho do arquivo.

### Formato do arquivo

- **Frio** (`Save(path)`, ou `IncludeComputedValues = false`) — o MemoryPack bruto do modelo, byte a byte
  idêntico a toda versão anterior. Este é um contrato permanente, garantido por um teste de regressão.
- **Aquecido** — um pequeno container autodescritivo: o número mágico `MSWM`, 1 byte de versão do formato,
  os **mesmos** bytes do modelo que um save frio escreveria, seguidos por um bloco de valores (o
  MemoryPack dos valores em cache). O `Load` inspeciona os 4 bytes do número mágico: uma correspondência
  indica um container aquecido; qualquer outra coisa é um modelo bruto (frio ou pré-existente), de modo
  que arquivos antigos continuam carregando sem alteração.

Como o modelo e seus valores viajam em um único arquivo, eles nunca podem dessincronizar no carregamento.

### O que o warm-start *não* congela

Dois tipos de valor em cache são deliberadamente **excluídos** do snapshot e são recalculados na primeira
leitura, mesmo a partir de um arquivo aquecido:

- **Células voláteis** — qualquer coisa que tenha envolvido `NOW`/`TODAY`/`RAND`/`RANDBETWEEN` (direta ou
  transitivamente). Persisti-las "congelaria o relógio de ontem"; em vez disso, elas são reamostradas na
  próxima leitura.
- **Resultados do tipo referência** — raros como valor final de célula e baratos de reconstruir.

### Contrato de desatualização

O warm-start persiste valores que você já computou; ele não rastreia edições. O contrato pós-carregamento
é o mesmo de sempre: **após editar células, chame `InvalidateCache()`** (ou `Recalculate()` para uma
atualização apenas das voláteis) antes de ler, ou você lerá valores desatualizados. Um carregamento
aquecido apenas pula a *primeira* recomputação das células inalteradas e não voláteis — isso não muda em
nada como a invalidação funciona depois. E, assim como em um carregamento frio, as [funções
personalizadas](custom-functions.md) ainda precisam ser registradas novamente: células que **não** estavam
em cache no momento do save (ou que você invalidar) reavaliarão suas chamadas e precisarão da
implementação presente.

## Compatibilidade

Os nós de expressão são serializados como uma union do MemoryPack, e as tags da union são **append-only
por política do projeto**: tags existentes nunca são renumeradas, reordenadas ou reutilizadas, e novos
tipos de nó recebem tags novas. Workbooks salvos por uma versão mais antiga permanecem, portanto,
carregáveis por versões mais novas da biblioteca.

Como apenas as tags (nunca os nomes de tipo) vão para o fio, a [reorganização de namespaces da
2.0](migrating-to-2.0.md) não mudou o formato em absolutamente nada: arquivos salvos pela 1.x carregam
na 2.0 sem alteração, garantido por uma fixture binária pré-2.0 congelada na suíte de testes.

## Quando usar cada formato

| Necessidade | Use |
| --- | --- |
| Persistência nativa rápida de um modelo calculado (estilo cache, reinícios de serviço, snapshots entre etapas de processamento) | `Workbook.Save` / `Load` |
| Intercâmbio com pessoas ou outras ferramentas (abrir no Excel, enviar um relatório) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingestão da planilha que é a fonte da verdade | [`ExcelFile.Load`](excel-interop.md) |
