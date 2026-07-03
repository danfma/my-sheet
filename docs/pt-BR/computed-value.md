# ComputedValue e erros

*Tradução do documento canônico em inglês ([computed-value.md](../computed-value.md)). Em caso de divergência, o inglês prevalece.*

Avaliar uma expressão retorna um `ComputedValue` — o único tipo de resultado do MySheet. É uma
`readonly struct` que emula uma união discriminada: um campo `double` (números, booleanos, códigos de
erro), um campo `object?` (texto, referências) e uma tag de um byte. Produzir um número, booleano, em
branco ou erro **não aloca nada**; texto e referências apenas carregam uma referência que já existia.

```csharp
using Danfma.MySheet; // Workbook, ComputedValue, ComputedValueKind, Error

ComputedValue value = workbook.GetCellValue("Sheet1", "A3");
```

## Tipos (kinds)

`value.Kind` é um `ComputedValueKind`:

| Kind | Significado | Alocação |
| --- | --- | --- |
| `Blank` | Uma célula vazia / um argumento omitido. | nenhuma |
| `Number` | Um `double`. Datas também são números (datas seriais do Excel). | nenhuma |
| `Boolean` | `TRUE`/`FALSE`. | nenhuma |
| `Text` | Uma `string`. | carrega a string já existente |
| `Error` | Um erro do Excel (`#DIV/0!`, `#N/A`, …) como uma struct [`Error`](#a-struct-error). | nenhuma |
| `Reference` | Uma referência produzida por uma função como `OFFSET` (veja [Referências](#referências-e-enumeratevalues)). | carrega a referência já existente |

> **Resultados de fórmula nunca são em branco na borda da célula (paridade com o Excel).** `Blank` é o
> que você obtém de uma célula verdadeiramente vazia (expressão `BlankValue`) ou de um argumento omitido.
> Mas uma célula que TEM conteúdo cuja fórmula avalia para em branco é coagida: `GetCellValue` retorna
> `Number(0)`, exatamente como o Excel — por exemplo, `=Sheet2!F10` com `F10` vazio, ou `=IF(TRUE, F10)`,
> é lido como `0`, não em branco. A coerção é da **célula**, não da expressão — `Evaluate` mantém o
> branco internamente (em branco ainda se compara como `""`/`0`/`FALSE` dentro de uma expressão). Veja
> [Resultados de fórmula nunca são em branco](workbook-and-expressions.md#resultados-de-fórmula-nunca-são-em-branco-paridade-com-o-excel).

## Extraindo valores

A extração é **explícita e estrita por tipo exato** — não há coerção ao estilo do Excel nesta superfície
(a coerção é interna à engine). Em particular, `Number` e `Boolean` não se cruzam: `TryGetNumber` em um
booleano retorna `false`.

Três estilos, um para cada necessidade ergonômica:

### `TryGet*` — seguro, out + bool

```csharp
if (value.TryGetNumber(out double number)) { /* ... */ }
if (value.TryGetBoolean(out bool flag)) { /* ... */ }
if (value.TryGetText(out string? text)) { /* ... */ }
if (value.TryGetError(out Error error)) { /* ex.: error == Error.DivZero */ }
if (value.TryGetReference(out Reference? reference)) { /* ... */ }
```

### `As*` — açúcar sintático com nullable

```csharp
double? number = value.AsDouble();    // null, a menos que Kind == Number
bool? flag = value.AsBoolean();       // null, a menos que Kind == Boolean
string? text = value.AsString();      // null, a menos que Kind == Text
```

### `To*` — asserções estritas (lançam exceção quando o tipo não corresponde)

```csharp
double number = value.ToDouble();     // lança InvalidOperationException, a menos que Kind == Number
bool flag = value.ToBoolean();        // lança, a menos que Kind == Boolean
string text = value.ToText();         // lança, a menos que Kind == Text
```

Use `To*` quando o contrato do workbook garante o tipo (por exemplo, uma célula de resultado numérico em
um modelo que você controla) e uma divergência é um bug de verdade.

### `AsObject()` — a ponte para `object?`

Para interop com código fracamente tipado, `AsObject()` faz boxing do valor:

| Kind | `AsObject()` retorna |
| --- | --- |
| `Blank` | `null` |
| `Number` | `double` com boxing |
| `Boolean` | `bool` com boxing |
| `Text` | a `string` |
| `Error` | o nó de AST `ErrorValue` correspondente (ex.: `ErrorValue.DivByZero`) |
| `Reference` | a expressão `Reference` |

Esta é uma válvula de escape permanente, não o caminho principal — ela reintroduz o boxing que a struct
existe para evitar; portanto, mantenha-a fora de laços críticos de desempenho.

## Construindo valores

Você constrói `ComputedValue`s principalmente ao escrever [funções personalizadas](custom-functions.md).
Métodos de fábrica mais conversões implícitas **apenas na entrada** (a extração nunca é implícita):

```csharp
ComputedValue a = ComputedValue.Number(42);
ComputedValue b = ComputedValue.Boolean(true);
ComputedValue c = ComputedValue.Text("hello");
ComputedValue d = ComputedValue.Blank;
ComputedValue e = ComputedValue.Error(Error.Num);

// Conversões implícitas de double / bool / string / Error:
ComputedValue f = 42.0;
ComputedValue g = true;
ComputedValue h = "hello";       // uma string null converte para Blank
ComputedValue i = Error.Value;
```

## A struct `Error`

`Error` é uma struct sem alocações, no estilo smart-enum, que encapsula um código `int` — ela cabe
inteiramente dentro de um `ComputedValue`. Os erros conhecidos do Excel são instâncias estáticas
nomeadas:

| Instância | `Display` |
| --- | --- |
| `Error.Null` | `#NULL!` |
| `Error.DivZero` | `#DIV/0!` |
| `Error.Value` | `#VALUE!` |
| `Error.Ref` | `#REF!` |
| `Error.Name` | `#NAME?` |
| `Error.Num` | `#NUM!` |
| `Error.NA` | `#N/A` |

`ToString()` imprime o texto de exibição, e a igualdade é por código:

```csharp
if (value.TryGetError(out var error))
{
    Console.WriteLine(error);                    // "#DIV/0!"
    Console.WriteLine(error.Display);            // "#DIV/0!"

    if (error == Error.NA)
    {
        // um lookup não encontrou nada — muitas vezes é razoável tratar como "sem dados"
    }
}
```

Observações:

- Uma **referência circular** é detectada pela camada de memoização e aparece como `Error.Ref` (`#REF!`)
  em vez de um estouro de pilha.
- Uma **chamada a uma função não registrada** é avaliada como `Error.Name` (`#NAME?`).
- `Error` é diferente de `ErrorValue`, que é o *nó* serializável de AST para um erro literal armazenado
  em uma célula (por exemplo, carregado de um arquivo `.xlsx`). `AsObject()` mapeia um `Error` de volta
  para o singleton `ErrorValue` correspondente; avaliar um `ErrorValue` produz o `Error` correspondente.
- O registro de códigos de erro personalizados está deliberadamente fora de escopo por enquanto.

## Referências e `EnumerateValues`

Algumas funções (atualmente `OFFSET`) são avaliadas como uma *referência*, em vez de um escalar —
`Kind == ComputedValueKind.Reference`. `EnumerateValues` percorre as células referenciadas e produz seus
**valores calculados** (através do cache de memoização):

```csharp
var offset = ExpressionParser.Parse("=OFFSET(A1, 0, 0, 3, 1)", sheet);
ComputedValue reference = offset.Evaluate(workbook);

foreach (ComputedValue cell in reference.EnumerateValues(workbook))
{
    if (cell.TryGetNumber(out var number))
    {
        Console.WriteLine(number);
    }
}
```

Em qualquer valor que não seja `Reference`, `EnumerateValues` não produz nada. Note que um *intervalo
puro* em uma fórmula (`=A1:B2` usado como escalar) não produz um valor `Reference` — ele é avaliado como
`#VALUE!`, como no Excel; intervalos são consumidos pelas funções que os aceitam. Para enumerar um
intervalo a partir de uma função personalizada, veja
[Funções personalizadas — argumentos de intervalo](custom-functions.md#aceitando-intervalos-e-referências).

## Veja também

- [Funções personalizadas](custom-functions.md) — produzindo seus próprios `ComputedValue`s.
- [Desempenho](performance.md) — por que o design de união existe, com números medidos.
