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

O carregamento é **totalmente streaming**: shared strings e worksheets são lidos com um `XmlReader`
forward-only, então o DOM do OpenXML nunca é materializado e a única representação completa em memória é
o próprio modelo MySheet (num arquivo de ~620 mil células: ~4,8x mais rápido e ~6,6x menos alocação que
um load baseado em DOM). Duas notas práticas: o overload de `Stream` exige um stream legível **e
seekable** (requisito do leitor de pacote subjacente), e células ou linhas sem o atributo `r` — posições
implícitas, emitidas por alguns writers minimalistas — carregam na posição implícita definida pela spec.

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
| "Escrava" de fórmula compartilhada (uma célula de fórmula arrastada que não carrega texto de fórmula) | Um nó leve que compartilha a árvore já interpretada da mestre (veja [Fórmulas compartilhadas](#fórmulas-compartilhadas-uma-árvore-mestre-compartilhada-com-deltas-por-escrava) abaixo) quando a forma da mestre é suportada; caso contrário, é expandida em uma fórmula independente exatamente como antes. |
| Nome definido com escopo de workbook (`<definedName>`) | Uma entrada em [`Workbook.DefinedNames`](workbook-and-expressions.md#intervalos-nomeados): o texto `refersTo` passa pelo parse como uma fórmula. Nomes **com escopo de planilha** (aqueles com `localSheetId`) e os nomes **nativos `_xlnm.*`** do Excel (`Print_Area`, `Print_Titles`, `_FilterDatabase`, …) são ignorados. |

### Fórmulas compartilhadas: uma árvore mestre compartilhada com deltas por escrava

Uma fórmula do Excel arrastada é armazenada uma única vez no arquivo (a "mestre"), mais um intervalo de
células "escravas" que não carregam texto de fórmula próprio. Historicamente o MySheet dava a cada escrava
sua própria árvore `Expression` interpretada de forma independente — correto, mas em um workbook com
centenas de milhares de células arrastadas isso significava centenas de milhares de árvores alocadas
separadamente para o que, estruturalmente, é uma única fórmula.

Quando a forma da mestre é totalmente suportada — referências de célula simples, ranges limitados,
operadores aritméticos, nomes definidos e chamadas de função sobre argumentos que também sejam suportados —
toda escrava do grupo agora compartilha **uma única** árvore interpretada. A mestre é interpretada uma vez
em um modo ancorado que preserva seus marcadores `$` em vez de descartá-los, e cada escrava se torna um nó
pequeno `SharedFormulaSlave(Master, ΔRow, ΔColumn)`. Avaliar uma escrava propaga seu próprio `(ΔRow,
ΔColumn)` para dentro da árvore mestre compartilhada: as referências relativas dentro dela se deslocam por
esse delta, as ancoradas com `$` não — o mesmo resultado que uma árvore totalmente expandida produziria. O
deslocamento é memoizado por época de avaliação como qualquer outra célula, então reler uma escrava
repetidamente não repete a aritmética do delta (veja [Desempenho →
Memoização](performance.md#memoização) para o mecanismo completo e o custo de cálculo medido).

Na forma usada para medir isso (um grupo de fórmulas arrastadas sobre ~360 mil células), compartilhar a
árvore mestre reduziu o benchmark de carregamento em **-46% de tempo, -59% de memória alocada e -50% de
coletas Gen1+Gen2** em comparação com expandir cada escrava em sua própria árvore.

Nem toda forma arrastada pode ser compartilhada dessa maneira. Uma mestre cuja fórmula envolva um
**range aberto** (referência de coluna/linha inteira), uma **união** (`(A1:A3, C1:C3)`) ou um **endpoint de
range cross-sheet encadeado** recai honestamente: o grupo inteiro é expandido por escrava em fórmulas
independentes, exatamente como sempre foi. Não há aproximação silenciosa — uma forma não suportada é
detectada antes de qualquer escrava ser construída, e o grupo segue o caminho mais lento e mais seguro.

> **Nota de compatibilidade futura.** Um arquivo nativo em MemoryPack (veja
> [Serialização](serialization.md)) salvo a partir de um workbook que contenha nós de fórmula compartilhada
> nesta representação compartilhada não pode ser aberto por versões desta biblioteca **anteriores** à que a
> introduziu — os novos tipos de nó registram tags de fio que essas versões não conhecem. Arquivos salvos
> por versões antigas continuam carregando sem alteração. Veja [Serialização →
> Compatibilidade](serialization.md#compatibilidade) para o contrato completo, incluindo a ressalva honesta
> de que essa otimização economiza **memória e pressão de GC, não espaço em disco** — o arquivo em si
> continua serializando a árvore da mestre uma vez por escrava.

As planilhas são criadas na ordem das abas do arquivo, então `Sheet.Index` (e a função `SHEET`)
correspondem ao Excel.

Se o arquivo usa funções que o MySheet não implementa, essas células viram nós `FunctionCall` e são
avaliadas como `#NAME?` — a menos que você mesmo forneça o comportamento via
[`RegisterFunction`](custom-functions.md), que é a válvula de escape pretendida.

Alguns problemas de carregamento (um nome definido inválido, um literal de data que falha ao ser
interpretado) são ignorados em vez de falhar o carregamento inteiro — por padrão, silenciosamente, como
sempre foi. Passe `ExcelLoadOptions` com um callback `OnWarning` para `Load` para observá-los:

```csharp
var warnings = new List<ExcelLoadWarning>();
var workbook = ExcelFile.Load("model.xlsx", new ExcelLoadOptions { OnWarning = warnings.Add });
```

Cada `ExcelLoadWarning` carrega um `Kind` (`InvalidDefinedName` ou `UnparsableDateLiteral`), um `Subject`
(o nome definido, ou o id da célula) e uma string `Detail`. O callback é um `Action<T>` simples em vez de
uma lista acumulada, então quem hospeda decide se registra, coleta ou ignora cada aviso.

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
- Os [intervalos nomeados](workbook-and-expressions.md#intervalos-nomeados) em `Workbook.DefinedNames`
  são escritos como entradas `<definedName>` com escopo de workbook, em **ambos** os modos de fórmula. O
  texto `refersTo` é totalmente qualificado (`FormulaWriter` com um contexto vazio, então toda referência
  mantém seu prefixo `Sheet!`).

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
- **Fórmulas compartilhadas continuam sendo fórmulas reais**: células escravas de uma fórmula arrastada
  (que não carregam texto de fórmula no arquivo) são reconstruídas a partir da mestre, aplicando o delta de
  células para que referências relativas se movam enquanto os componentes ancorados com `$` ficam fixos —
  assim elas continuam reagindo a mudanças de entrada. Quando a forma da mestre é suportada, toda escrava
  compartilha uma única árvore interpretada em vez de ganhar a sua própria (veja [Fórmulas
  compartilhadas](#fórmulas-compartilhadas-uma-árvore-mestre-compartilhada-com-deltas-por-escrava) acima);
  caso contrário, o grupo recai na expansão completa por escrava. Uma célula escrava cuja mestre está
  ausente do arquivo recorre ao seu valor literal em cache.
- **A cobertura de funções não é total** (veja a contagem e a lista atuais de nativas, além das suas funções personalizadas, na
  [referência de funções](function-reference.md). Fórmulas que usam outras funções carregam como nós
  `FunctionCall` e são avaliadas como `#NAME?`, a menos que registradas.
- **Só nomes definidos com escopo de workbook atravessam**: nomes com escopo de planilha (com
  `localSheetId`) e os nomes nativos `_xlnm.*` (áreas de impressão, bancos de filtro, …) são ignorados no
  carregamento, e o MySheet só escreve nomes com escopo de workbook. Um nome definido cujo `refersTo` não
  pode ser interpretado é ignorado em vez de falhar o carregamento.

## Veja também

- [Primeiros passos](getting-started.md) — o fluxo de ponta a ponta em miniatura.
- [Funções personalizadas](custom-functions.md) — fornecendo o comportamento das funções que a engine não
  tem.
- [Desempenho](performance.md) — por que as exportações avaliam via `RunWithLargeStack`.
