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

| | Persistido? | Observações |
| --- | --- | --- |
| Planilhas (nome, ordem das abas) | Sim | A busca de nomes case-insensitive é restaurada na desserialização. |
| Células e árvores de expressão completas | Sim | Fórmulas continuam sendo fórmulas — um workbook carregado continua recalculando. |
| **Chamadas** de funções personalizadas (nós `FunctionCall`) | Sim | O nome e as expressões dos argumentos são preservados. |
| **Implementações** de funções personalizadas (delegates) | **Não** | Comportamento é código, não dados — registre de novo após carregar. |
| Cache de memoização | **Não** | Os valores são recalculados de forma preguiçosa na primeira leitura após o carregamento. |

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
