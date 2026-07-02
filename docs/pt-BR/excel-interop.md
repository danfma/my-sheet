# Interop com Excel (`Danfma.MySheet.Excel`)

*Tradução do documento canônico em inglês ([excel-interop.md](../excel-interop.md)). Em caso de divergência, o inglês prevalece.*

O `Danfma.MySheet.Excel` conecta a engine MySheet a arquivos `.xlsx` reais através do OpenXML SDK —
multiplataforma, **sem instalação do Excel, sem COM**. Este é o pacote que habilita o cenário central do
MySheet: manter uma pasta de trabalho do Excel em um servidor como fonte da verdade, carregá-la,
reavaliar suas fórmulas em processo e expor ou gravar de volta os resultados.

Ele deliberadamente *não* é uma biblioteca de manipulação de Excel de propósito geral. Ele move **valores
e fórmulas** entre `.xlsx` e a engine; estilos, formatos de número e outros recursos de apresentação
estão fora de escopo no MVP atual (veja [Escopo e limitações](#escopo-e-limitações)). Se você precisa
deles, ClosedXML, EPPlus ou NPOI continuam sendo escolhas excelentes — e combinam bem com o MySheet.

```shell
dotnet add package Danfma.MySheet.Excel
```

## Carregando: `ExcelFile.Load`

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Excel;

Workbook workbook = ExcelFile.Load("model.xlsx");

// ou a partir de qualquer stream legível:
using var stream = File.OpenRead("model.xlsx");
Workbook fromStream = ExcelFile.Load(stream);
```

O resultado é um `Workbook` comum do MySheet. A propriedade-chave: **células de fórmula se tornam árvores
`Expression` de verdade**, interpretadas pelo parser do MySheet e reavaliadas pela engine do MySheet — os
valores em cache armazenados no arquivo são ignorados para células de fórmula. Altere uma célula de
entrada, invalide o cache, e toda fórmula dependente é recalculada:

```csharp
using Danfma.MySheet.Expressions;

var workbook = ExcelFile.Load("model.xlsx");

workbook["Inputs"]["B1"] = new NumberValue(2500);
workbook.InvalidateCache();

double updated = workbook.GetCellValue("Results", "B10").ToDouble();
```

Como o conteúdo do arquivo é mapeado para dentro do workbook:

| No `.xlsx` | No `Workbook` |
| --- | --- |
| Célula de fórmula (`<f>`) | Árvore `Expression` resultante do parse (reavaliada pelo MySheet; o `<v>` em cache é ignorado). |
| Célula numérica | `NumberValue`. |
| String compartilhada / inline | `StringValue` (trechos de rich text achatados em texto puro). |
| Célula booleana | `BooleanValue`. |
| Célula de erro | `ErrorValue` (avaliada como o `Error` correspondente). |
| Célula de data | `NumberValue` com o **número serial** do Excel (datas ISO-8601 em arquivos no modo strict são convertidas via `ToOADate`). |
| Célula vazia / só com estilo | Nada é armazenado — é lida como em branco. |
| "Escrava" de fórmula compartilhada (uma célula de fórmula arrastada que não carrega texto de fórmula) | Expandida em uma fórmula real: o texto da célula mestre é deslocado pelo delta de linha/coluna (referências relativas se movem, componentes ancorados com `$` ficam, texto dentro de literais de string não é tocado) e passa pelo parse como qualquer outra fórmula. |

As planilhas são criadas na ordem das abas do arquivo, então `Sheet.Index` (e a função `SHEET`)
correspondem ao Excel.

Se o arquivo usa funções que o MySheet não implementa, essas células viram nós `FunctionCall` e são
avaliadas como `#NAME?` — a menos que você mesmo forneça o comportamento via
[`RegisterFunction`](custom-functions.md), que é a válvula de escape pretendida.

## Exportando: `SaveAsExcel`

```csharp
using Danfma.MySheet.Excel;

// Padrão: um snapshot achatado — cada célula escrita como seu valor literal calculado.
workbook.SaveAsExcel("snapshot.xlsx");

// Mantenha as fórmulas vivas: escreve o texto da fórmula mais o valor calculado (em cache).
workbook.SaveAsExcel("live.xlsx", new ExcelExportOptions { FormulaMode = FormulaMode.Formulas });

// Sobrecarga com stream:
using var output = File.Create("snapshot.xlsx");
workbook.SaveAsExcel(output);
```

`SaveAsExcel` cria um arquivo **novo** contendo as planilhas e células do workbook. Antes da escrita,
cada célula é avaliada de antemão em uma única thread com pilha grande (`RunWithLargeStack`), com
memoização, para que cadeias de dependência profundas não possam estourar a pilha no meio da escrita.

`ExcelExportOptions.FormulaMode` controla as células de fórmula:

| Modo | As células de fórmula se tornam | Use quando |
| --- | --- | --- |
| `FormulaMode.ValuesOnly` (padrão) | Seu valor literal calculado — um snapshot achatado, sem fórmulas. | O destinatário deve ver resultados, não a lógica (relatórios, repasse de dados). |
| `FormulaMode.Formulas` | A fórmula do Excel (`<f>`, renderizada pelo `FormulaWriter`) **mais** seu valor calculado como o `<v>` em cache. | O arquivo deve continuar recalculando ao ser aberto no Excel. |

Detalhes que vale a pena conhecer:

- Resultados em branco são omitidos por completo (como nos próprios arquivos do Excel).
- Literais de texto são deduplicados por meio de uma tabela de strings compartilhadas (shared strings);
  texto produzido *por uma fórmula* é escrito como a string em cache da fórmula, conforme a convenção do
  `.xlsx`.
- Uma célula cujo resultado é uma referência pura (por exemplo, um `OFFSET` multicélula usado como
  escalar) é escrita como `#VALUE!`, espelhando como a engine a trata.
- No modo `Formulas`, chamadas a [funções personalizadas](custom-functions.md) são escritas com o nome
  registrado — o Excel mostrará o valor em cache e sinalizará a função desconhecida, o que é esperado.

## Mesclando em um template: `MergeIntoExcel`

```csharp
workbook.MergeIntoExcel("report.xlsx");
```

`MergeIntoExcel` edita um arquivo **existente**, **no próprio arquivo (in place)**: cada célula mantida
pelo workbook do MySheet é escrita no destino como seu valor literal calculado, enquanto *todo o resto*
do arquivo — estilos, formatos de número, outras células, outras planilhas, gráficos — permanece
intocado. Esta é a ferramenta para o clássico fluxo de relatório "template bonito, números calculados":

```csharp
// A receita template→relatório: copie o template intacto e mescle na cópia.
File.Copy("template.xlsx", "report.xlsx", overwrite: true);
workbook.MergeIntoExcel("report.xlsx");
```

O design somente in place é deliberado: a mesclagem *muta o arquivo que você fornece*, e criar arquivos é
trabalho do `SaveAsExcel`. Intencionalmente não existe uma sobrecarga `MergeIntoExcel(template, output)`
— o passo com `File.Copy` mantém o template intacto e torna a mutação explícita.

Semântica da mesclagem:

- As planilhas são casadas **pelo nome, sem diferenciar maiúsculas de minúsculas**. Planilhas do workbook
  ausentes no destino são ignoradas.
- Cada célula escrita recebe o **valor literal calculado** — qualquer fórmula que a célula de destino
  tivesse é descartada (o arquivo mesclado mostra os números da sua engine, não o recálculo do Excel).
- **Valores em branco não são escritos**, deixando a célula de destino exatamente como estava.
- A formatação da célula é preservada: apenas o conteúdo é substituído; a referência de estilo da célula
  não é tocada.
- O texto é escrito como string inline, então a tabela de strings compartilhadas do destino não é
  modificada.
- Linhas/células ausentes são criadas na ordem correta do OpenXML conforme necessário.
- Assim como no `SaveAsExcel`, todos os valores são calculados de antemão via `RunWithLargeStack`, com
  memoização.

## Escopo e limitações

Sendo honestos sobre o que o MVP de interop **não** faz:

- **Sem estilos nem apresentação**: fontes, cores, formatos de número, larguras de coluna, células
  mescladas, gráficos e afins não são modelados. `Load` os ignora; `SaveAsExcel` não os produz;
  `MergeIntoExcel` *preserva* a formatação existente no destino, mas não consegue criá-la.
- **Datas são números seriais**: elas entram e saem como `double`s (a própria representação do Excel).
  Aplique a formatação de data no template (fluxo de mesclagem) ou converta com `DateTime.FromOADate` no
  seu código.
- **Marcadores absolutos não são preservados na escrita**: `$A$1` passa pelo parse sem problema (ele
  identifica a mesma célula), mas é reescrito como `A1` — uma perda de fidelidade apenas em exportações
  com `FormulaMode.Formulas`, e apenas cosmética, a menos que você planeje copiar/preencher fórmulas no
  Excel depois.
- **Fórmulas compartilhadas SÃO expandidas**: células escravas de uma fórmula arrastada (que não carregam
  texto de fórmula no arquivo) são reconstruídas a partir do texto da célula mestre, deslocando as
  referências relativas pelo delta de células e mantendo fixos os componentes ancorados com `$` — assim
  elas continuam sendo fórmulas reais, que reagem a mudanças de entrada. Uma célula escrava cuja mestre
  está ausente do arquivo recorre ao seu valor literal em cache.
- **A cobertura de funções não é total** (veja a contagem e a lista atuais de nativas, além das suas funções personalizadas, na
  [referência de funções](function-reference.md). Fórmulas que usam outras funções carregam como nós
  `FunctionCall` e são avaliadas como `#NAME?`, a menos que registradas.

## Veja também

- [Primeiros passos](getting-started.md) — o fluxo de ponta a ponta em miniatura.
- [Funções personalizadas](custom-functions.md) — fornecendo o comportamento das funções que a engine não
  tem.
- [Desempenho](performance.md) — por que as exportações avaliam via `RunWithLargeStack`.
