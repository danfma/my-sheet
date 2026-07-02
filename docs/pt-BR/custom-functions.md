# Funções personalizadas

*Tradução do documento canônico em inglês ([custom-functions.md](../custom-functions.md)). Em caso de divergência, o inglês prevalece.*

O conjunto de funções do MySheet é extensível: registre um delegate .NET sob um nome, e as fórmulas
poderão chamá-lo como qualquer função nativa — `=DOUBLE(A1)`, `=RISK_SCORE(B1:B10, 0.95)`, e assim por
diante. Este é o principal ponto de extensão da engine, e ele foi projetado em torno de duas ideias:
**os argumentos chegam sem avaliação** (assim a sua função controla a avaliação, permitindo avaliação
preguiçosa (lazy) e curto-circuito), e **retornar escalares é sem atrito** (as conversões implícitas
fazem o empacotamento).

## Registrando

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

workbook.RegisterFunction("DOUBLE", (arguments, wb) =>
{
    var value = arguments[0].Evaluate(wb);

    return value.TryGetNumber(out var number)
        ? number * 2
        : ComputedValue.Error(Error.Value);
});

sheet["A1"] = ExpressionParser.Parse("=DOUBLE(21)", sheet);
double result = workbook.GetCellValue("Sheet1", "A1").ToDouble();   // 42.0
```

A assinatura do delegate é:

```csharp
public delegate ComputedValue CustomFunction(Expression[] arguments, Workbook workbook);
```

- `arguments` são os nós de expressão **crus, não avaliados**, vindos do ponto de chamada.
- `workbook` é o contexto de avaliação — passe-o para `arguments[i].Evaluate(workbook)` quando quiser um
  valor.
- O tipo de retorno é [`ComputedValue`](computed-value.md).

Nomes de função:

- Os nomes são **case-insensitive** (`=double(21)` funciona) e podem conter underscores
  (`=RISK_SCORE(...)`).
- O Excel armazena funções mais novas com um prefixo `_xlfn.`; o parser o remove na normalização, então
  uma função personalizada chamada `MYFN` também corresponde a `=_xlfn.MYFN()` em um arquivo `.xlsx`
  carregado.
- **Funções nativas não podem ser sobrescritas.** Os 164 nomes na tabela de funções do parser sempre
  passam pelo parse como seus nós nativos; o registro de funções personalizadas só é consultado para os
  demais nomes.
- Chamar um nome que nunca foi registrado não é uma exceção — a chamada é avaliada como `#NAME?`, como
  no Excel.

## Argumentos lazy e curto-circuito

Como os argumentos chegam sem avaliação, a sua função decide *o que* é avaliado e *quando*. Ramos caros
ou com efeitos colaterais simplesmente nunca são tocados, a menos que você os avalie:

```csharp
// Retorna o primeiro argumento que é avaliado como um valor não em branco.
// Os argumentos seguintes nunca são avaliados depois que uma correspondência é encontrada.
workbook.RegisterFunction("FIRSTNONBLANK", (arguments, wb) =>
{
    foreach (var argument in arguments)
    {
        var value = argument.Evaluate(wb);

        if (value.Kind != ComputedValueKind.Blank)
        {
            return value;
        }
    }

    return ComputedValue.Blank;
});
```

Este é o mesmo mecanismo que permite ao `IF` avaliar apenas o ramo escolhido.

## Retornando valores

Escalares convertem implicitamente, então os casos comuns não exigem cerimônia:

```csharp
workbook.RegisterFunction("PI2", (_, _) => 6.283185307179586);   // double → Number
workbook.RegisterFunction("YES", (_, _) => true);                // bool → Boolean
workbook.RegisterFunction("GREET", (_, _) => "hello");           // string → Text (null → Blank)
```

Para todo o resto, use os métodos de fábrica:

```csharp
workbook.RegisterFunction("SAFEDIV", (arguments, wb) =>
{
    var left = arguments[0].Evaluate(wb);
    var right = arguments[1].Evaluate(wb);

    if (!left.TryGetNumber(out var numerator) || !right.TryGetNumber(out var denominator))
    {
        return ComputedValue.Error(Error.Value);   // tipos de argumento errados → #VALUE!
    }

    if (denominator == 0)
    {
        return ComputedValue.Blank;                // escolha desta função: em branco em vez de #DIV/0!
    }

    return numerator / denominator;
});
```

Diretrizes:

- Sinalize falhas visíveis ao usuário **retornando** `ComputedValue.Error(...)` — não lançando exceção.
  Os erros se propagam pelas fórmulas do mesmo jeito que os erros do Excel (e `IFERROR`/`IFNA` podem
  capturá-los).
- Reserve exceções para bugs de verdade; elas se propagarão para fora de `Evaluate` como exceções.

## Aceitando intervalos e referências

Um argumento de intervalo (`=MYFN(A1:A10)`) chega como um nó `RangeReference` — avaliá-lo diretamente
produz `#VALUE!` (um intervalo puro não tem valor escalar). Para consumir as células, envolva a
referência em um `ComputedValue` e enumere seus valores através do cache memoizado:

```csharp
workbook.RegisterFunction("PRODUCT", (arguments, wb) =>
{
    var product = 1.0;

    foreach (var argument in arguments)
    {
        IEnumerable<ComputedValue> values = argument is Reference reference
            ? ComputedValue.Reference(reference).EnumerateValues(wb)
            : [argument.Evaluate(wb)];

        foreach (var value in values)
        {
            if (value.TryGetError(out var error))
            {
                return error;                       // propaga erros, ao estilo do Excel
            }

            if (value.TryGetNumber(out var number))
            {
                product *= number;
            }
        }
    }

    return product;
});

sheet["B1"] = ExpressionParser.Parse("=PRODUCT(A1:A3, 2)", sheet);
```

`Reference` cobre células únicas (`CellReference`), retângulos (`RangeReference`) e uniões
(`UnionReference`), então o mesmo ramo trata `=PRODUCT(A1)`, `=PRODUCT(A1:A10)` e
`=PRODUCT((A1:A3, C1:C3))`.

## Interação com a memoização

As funções personalizadas participam do cache por célula como todo o resto: uma vez calculado o valor de
uma célula, `GetCellValue` o serve a partir do cache e o delegate **não** é invocado de novo até que você
chame `workbook.InvalidateCache()`. Não conte com a reexecução de uma função personalizada a cada leitura
— se ela lê estado externo (hora, banco de dados, …), o valor observado é o da primeira avaliação, até
que o cache seja invalidado. Veja [Desempenho](performance.md).

## Serialização e arquivos Excel

Uma *chamada* de função personalizada é um nó comum de AST (`FunctionCall`) que guarda o nome e as
expressões dos argumentos. Esse nó:

- **é preservado no round-trip via MemoryPack** (`Workbook.Save`/`Load`),
- **passa pelo parse a partir de arquivos `.xlsx`** carregados com `ExcelFile.Load`, e
- **é reconvertido em texto** via `ToFormula` (então exportações com `FormulaMode.Formulas` mantêm o
  texto da chamada).

O **delegate em si nunca é serializado** — comportamento é código, não dados. Depois de carregar,
registre novamente a implementação antes de avaliar, ou a chamada será avaliada como `#NAME?`:

```csharp
workbook.Save("model.mysheet");

var restored = Workbook.Load("model.mysheet");

// Sem isto, qualquer célula que chame CUSTOM(...) é avaliada como #NAME?.
restored.RegisterFunction("CUSTOM", (arguments, wb) =>
{
    var a = arguments[0].Evaluate(wb).AsDouble() ?? 0;
    var b = arguments[1].Evaluate(wb).AsDouble() ?? 0;

    return a + b;
});
```

Um bom padrão para aplicações é um único método `RegisterAll(Workbook workbook)` que instala todas as
funções que as suas fórmulas usam, chamado tanto depois de `new Workbook()` quanto depois de cada
`Workbook.Load` / `ExcelFile.Load`.

## Veja também

- [ComputedValue e erros](computed-value.md) — os valores que a sua função recebe e retorna.
- [Serialização](serialization.md) — o que persiste e o que precisa ser registrado de novo.
