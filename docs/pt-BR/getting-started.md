# Primeiros passos

*Tradução do documento canônico em inglês ([getting-started.md](../getting-started.md)). Em caso de divergência, o inglês prevalece.*

O MySheet é uma engine de fórmulas de planilha em memória. Você monta (ou carrega) um `Workbook`, coloca
valores e fórmulas em suas `Sheet`s e pede à engine os resultados calculados — sem instalação do Excel,
sem COM e sem interface de planilha envolvida.

## Instalação

```shell
dotnet add package Danfma.MySheet
dotnet add package Danfma.MySheet.Excel   # apenas se você precisar de interop com .xlsx
```

Ambos os pacotes têm como alvo o **.NET 10** e são sempre publicados juntos, com a mesma versão.

## Seu primeiro workbook

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

// Valores literais também são nós de Expression:
sheet["A1"] = new NumberValue(10);
sheet["A2"] = new NumberValue(32);

// Fórmulas passam pelo parse e viram árvores de expressão de verdade:
sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
sheet["A4"] = ExpressionParser.Parse("=IF(A3>40, \"big\", \"small\")", sheet);
```

`ExpressionParser.Parse` segue a convenção de entrada de célula que você conhece do Excel: uma string que
começa com `=` é uma fórmula; qualquer outra coisa é um literal (número, booleano ou texto). Veja
[Workbook, planilhas e expressões](workbook-and-expressions.md) para as regras completas.

## Avaliando

A única API de avaliação é `Evaluate`, que retorna um [`ComputedValue`](computed-value.md) — uma união em
tipo de valor de número / booleano / texto / em branco / erro / referência que nunca faz boxing de
números:

```csharp
ComputedValue total = sheet["A3"].Evaluate(workbook);
ComputedValue label = sheet["A4"].Evaluate(workbook);

double sum = total.ToDouble();     // 42.0 — estrito: lança exceção se o resultado não for um número
string text = label.ToText();      // "big"
```

Para ler células, prefira `Workbook.GetCellValue` — ele retorna o mesmo `ComputedValue`, mas o memoiza
por célula, de modo que uma célula referenciada por muitas fórmulas é calculada uma única vez:

```csharp
ComputedValue cached = workbook.GetCellValue("Sheet1", "A3");
```

A extração é explícita e estrita — não há coerção escondida na superfície da API:

```csharp
var value = workbook.GetCellValue("Sheet1", "A3");

if (value.TryGetNumber(out var number))
{
    Console.WriteLine($"Número: {number}");
}
else if (value.TryGetError(out var error))
{
    Console.WriteLine($"A fórmula falhou: {error}");   // imprime, por exemplo, "#DIV/0!"
}

double? maybe = value.AsDouble();   // null quando o tipo (kind) não corresponde
object? boxed = value.AsObject();   // ponte para código baseado em object? (faz boxing de números)
```

## Editando células e o cache

O cache de memoização **não** é invalidado automaticamente. Depois de editar células, chame
`InvalidateCache()` antes de ler novamente:

```csharp
sheet["A1"] = new NumberValue(100);
workbook.InvalidateCache();

double updated = workbook.GetCellValue("Sheet1", "A3").ToDouble();   // 132.0
```

Esse é um design deliberado para o caso de uso principal — extração com muitas leituras e mutações raras
e em lote. Veja [Desempenho](performance.md) para a justificativa.

## Cadeias de dependência profundas

A avaliação é recursiva. Se o seu workbook tem cadeias de dependência muito longas (por exemplo, uma
coluna cumulativa com milhares de linhas), envolva todo o lote de extração em `RunWithLargeStack`, que
executa o trabalho em uma thread com uma pilha grande reservada:

```csharp
var value = Workbook.RunWithLargeStack(() => workbook.GetCellValue("Sheet1", "A20000"));
```

## Carregando um arquivo Excel

Com o pacote `Danfma.MySheet.Excel`, um arquivo `.xlsx` se torna um `Workbook` comum, cujas fórmulas são
árvores de expressão de verdade, reavaliadas pelo MySheet:

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Excel;

Workbook workbook = ExcelFile.Load("model.xlsx");
double result = workbook.GetCellValue("Data", "B10").ToDouble();
```

Veja [Interop com Excel](excel-interop.md) para exportar workbooks e mesclar valores calculados em
templates.

## Próximos passos

- [Workbook, planilhas e expressões](workbook-and-expressions.md) — o modelo de objetos em profundidade.
- [ComputedValue e erros](computed-value.md) — tudo sobre a leitura de resultados.
- [Funções personalizadas](custom-functions.md) — adicione suas próprias funções à engine.
- [Referência de funções](function-reference.md) — as 155 funções nativas.
