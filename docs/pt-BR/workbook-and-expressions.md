# Workbook, planilhas e expressões

*Tradução do documento canônico em inglês ([workbook-and-expressions.md](../workbook-and-expressions.md)). Em caso de divergência, o inglês prevalece.*

Este guia cobre o modelo de objetos do MySheet — `Workbook`, `Sheet` e a árvore de `Expression` — além
das regras de parsing, do conjunto de operadores, das referências e de como transformar uma expressão de
volta em texto de fórmula.

## Workbook

Um `Workbook` é o objeto raiz: um conjunto de planilhas nomeadas mais os serviços de avaliação (cache de
memoização, registro de funções personalizadas, serialização).

```csharp
using Danfma.MySheet;

var workbook = new Workbook();

var sheet = workbook.Sheets.Add("Sheet1");   // cria (ou retorna) uma planilha pelo nome
var same = workbook["Sheet1"];               // acesso por indexador
```

- **Os nomes de planilha são case-insensitive**, como no Excel: `workbook["sheet1"]` e
  `workbook["SHEET1"]` chegam à mesma planilha.
- `Sheets` é um `ConcurrentDictionary<string, Sheet>`, seguro para leitores concorrentes (o cenário
  pretendido de extração em segundo plano).
- `Sheets.Add(name)` atribui à planilha um `Index` igual à sua ordem de inserção — é isso que a função
  `SHEET` reporta, e o que define a ordem das abas ao exportar para Excel.

Principais membros de `Workbook`:

| Membro | Propósito |
| --- | --- |
| `Sheets` / `this[string]` | Acessa planilhas pelo nome (case-insensitive). |
| `GetCellValue(sheetName, id)` | Avaliação memoizada de uma célula → `ComputedValue`. |
| `InvalidateCache()` | Esvazia explicitamente o cache de memoização (obrigatório após edições). |
| `RegisterFunction(name, fn)` / `TryGetFunction(name, out fn)` | Registro de funções personalizadas ([guia](custom-functions.md)). |
| `Save(path)` / `SaveAsync(path)` / `Load(path)` / `LoadAsync(path)` | Serialização MemoryPack ([guia](serialization.md)). |
| `RunWithLargeStack(work)` (estático) | Executa um lote de avaliação em uma thread com pilha grande ([guia](performance.md)). |

## Sheet

Uma `Sheet` mapeia ids de célula (`"A1"`, `"B12"`, …) para nós `Expression`:

```csharp
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

sheet["A1"] = new NumberValue(1);                       // set: armazena a expressão
sheet["B1"] = ExpressionParser.Parse("=A1*2", sheet);

Expression cell = sheet["A1"];      // get: nunca lança — uma célula ausente é lida como BlankValue.Instance
bool exists = sheet.ContainsKey("C1");                  // false
bool found = sheet.TryGetValue("A1", out var stored);   // true

foreach (var (id, expression) in sheet) { /* itera as células armazenadas */ }
```

- O **getter nunca lança exceção**: ler um id que nunca foi definido retorna `BlankValue.Instance`, que
  é avaliado como em branco — exatamente como o Excel trata uma célula vazia.
- `Keys`, `Values` e `Count` expõem apenas as células que foram de fato armazenadas.
- Ids de célula são strings simples no estilo A1. O parser as normaliza para maiúsculas e remove os
  marcadores absolutos (`$A$1` → `A1`); ao definir células diretamente pelo indexador, use a forma
  normalizada (`"A1"`, não `"a1"`).

## Expressões

Toda célula guarda uma `Expression` — um record imutável. Literais, referências e operadores vivem em
`Danfma.MySheet.Expressions`; os nós das funções nativas vivem em namespaces filhos por categoria
(`Danfma.MySheet.Expressions.Mathematics`, `.Logical`, `.Statistical`, `.Text`, `.Information`,
`.Lookup`, `.Financial` — veja [Migrando para a 2.0](migrating-to-2.0.md)). Juntos, eles formam uma
árvore.

### Parsing

`ExpressionParser.Parse(text, sheet)` converte uma entrada de célula em uma expressão, usando a planilha
como contexto para as referências não qualificadas:

```csharp
using Danfma.MySheet.Parsing;

var formula = ExpressionParser.Parse("=SUM(A1:A10) * 1.1", sheet);   // árvore de expressão
var number = ExpressionParser.Parse("42.5", sheet);                  // NumberValue
var flag = ExpressionParser.Parse("true", sheet);                    // BooleanValue
var text = ExpressionParser.Parse("hello", sheet);                   // StringValue
var blank = ExpressionParser.Parse("", sheet);                       // BlankValue
```

Regras:

- Entradas que começam com `=` passam pelo parse como fórmulas (um parser Pratt / top-down por
  precedência de operadores).
- Qualquer outra coisa é um literal: número, se puder ser interpretado como tal (cultura invariante),
  depois booleano (`true`/`false`); caso contrário, texto.
- **Erros de sintaxe lançam `ParseException`** (com uma propriedade `Position` apontando para o token
  problemático). Funções nativas também validam a quantidade de argumentos em tempo de parse —
  `=ROUND(1)` lança exceção, assim como o Excel rejeitaria a fórmula na digitação.
- **Erros semânticos não lançam exceção** — uma função desconhecida é avaliada como `#NAME?`, uma
  referência inválida como `#REF!`, e assim por diante, na forma de erros de `ComputedValue`.

### Construindo árvores em código

Você pode construir expressões diretamente — útil para workbooks programáticos e testes:

```csharp
using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

sheet["A1"] = Number(10);
sheet["A2"] = Number(20);
sheet["A3"] = Sum(Cell("A1", sheet), Cell("A2", sheet));
sheet["A4"] = Add(Cell("A3", sheet), Number(5));
sheet["A5"] = Sum(Range("A1", "A2", sheet));
```

A classe base `Expression` fornece métodos de fábrica (`Number`, `String`, `Cell`, `Range`, `Sum`,
`Average`, `Min`, `Max`, `Count`, `Add`, `Subtract`, `Divide`, `Power`, `GreaterThan`, `Negate`,
`Plus`), e cada tipo de nó é um record público que você pode instanciar diretamente com `new`
(`new NumberValue(1)`, `new BinaryOperation(BinaryOperator.Multiply, left, right)`, …). Para
instanciar um record de função diretamente com `new`, importe o namespace da sua categoria (por
exemplo, `using Danfma.MySheet.Expressions.Mathematics;` para `new Sum(…)`).

### Avaliando

`Evaluate` é o único contrato de avaliação. Ele retorna um [`ComputedValue`](computed-value.md), sem
boxing para resultados numéricos:

```csharp
ComputedValue direct = sheet["A3"].Evaluate(workbook);          // avalia a árvore
ComputedValue cached = workbook.GetCellValue("Sheet1", "A3");   // memoizado por célula
```

Prefira `GetCellValue` ao ler células: ele armazena o resultado em cache, e qualquer `CellReference`
dentro de uma fórmula passa pelo mesmo cache, de modo que células compartilhadas são calculadas uma única
vez. `Evaluate` em uma instância de expressão é a ferramenta certa para expressões ad hoc que não estão
armazenadas em uma célula:

```csharp
var adHoc = ExpressionParser.Parse("=AVERAGE(A1:A2) > 10", sheet);
bool isHigh = adHoc.Evaluate(workbook).ToBoolean();
```

Não existe outra API de avaliação: quem precisa de um `object?` fracamente tipado chama `.AsObject()` no
resultado.

## Operadores

O MySheet faz o parse do conjunto de operadores do Excel. Forças de ligação (precedência) da mais fraca
para a mais forte:

| Precedência | Operadores | Observações |
| --- | --- | --- |
| 1 (mais fraca) | `=` `<>` `<` `>` `<=` `>=` | Comparações, com a ordenação entre tipos do Excel (números < texto < lógicos). |
| 2 | `&` | Concatenação de texto. |
| 3 | `+` `-` | Adição, subtração. |
| 4 | `*` `/` | Multiplicação, divisão. |
| 5 | `^` | Exponenciação (parse associativo à direita). |
| 6 | `%` | Percentual pós-fixado: `50%` é `0.5`. |
| 7 | `-` `+` unários | O prefixo unário liga mais forte que `^`, então `-2^2` é `(-2)^2 = 4`, como no Excel. |
| 8 (mais forte) | `:` | Construção de intervalo. |

Mais o agrupamento com `( )`. Divisão por zero produz `#DIV/0!`; incompatibilidades de tipo produzem
`#VALUE!`.

## Referências

```csharp
ExpressionParser.Parse("=A1", sheet);                 // célula na mesma planilha
ExpressionParser.Parse("=$A$1+A2", sheet);            // marcadores absolutos aceitos (e normalizados)
ExpressionParser.Parse("=Sheet2!A1", sheet);          // qualificada por planilha
ExpressionParser.Parse("='My Sheet'!A1:B2", sheet);   // nome de planilha entre aspas, intervalo
ExpressionParser.Parse("=SUM((A1:A3, C1:C3))", sheet); // união de referências (entre parênteses)
```

- Referências não qualificadas resolvem contra a planilha passada a `Parse`.
- Um intervalo (`A1:B2`) exige referências de célula em ambos os lados e vive na planilha da célula
  inicial (`Sheet2!A1:B2` está inteiramente em `Sheet2`).
- Os marcadores `$` identificam a mesma célula — o MySheet não faz copiar/preencher, então absoluto vs.
  relativo não tem efeito comportamental e o marcador não é preservado.
- Um intervalo puro usado onde se espera um escalar (por exemplo, `=A1:B2` sozinho) é avaliado como
  `#VALUE!`, como no Excel; intervalos são consumidos por funções (`SUM`, `COUNT`, lookups, …).
- Um nome puro que não é um id de célula (por exemplo, `=total`) é um `NameReference` — ele resolve
  contra as vinculações de `LET` e os [intervalos nomeados](#intervalos-nomeados) (*named ranges*) do
  workbook em tempo de avaliação, e produz `#NAME?` se não estiver vinculado.

## Intervalos nomeados

Um workbook pode definir **nomes** que representam uma expressão — geralmente um intervalo ou célula
qualificados por planilha, mas qualquer expressão (uma constante, uma fórmula, outro nome) é permitida.
Os nomes são de nível de workbook e **case-insensitive**, exatamente como no Excel.

```csharp
var workbook = new Workbook();
var data = workbook.Sheets.Add("Data");
data["A1"] = new NumberValue(10);
data["A2"] = new NumberValue(20);
data["A3"] = new NumberValue(30);

// Sobrecarga de conveniência: faz o parse do texto. As referências DEVEM ser qualificadas por planilha
// (nomes não têm planilha implícita); um '=' inicial e marcadores '$' são opcionais.
workbook.DefineName("Sales", "Data!A1:A3");

// Sobrecarga de expressão: um nome pode apontar para qualquer expressão, por exemplo uma constante.
workbook.DefineName("Rate", new NumberValue(0.1));

var main = workbook.Sheets.Add("Main");
ExpressionParser.Parse("=SUM(Sales)", main).Evaluate(workbook);   // 60
ExpressionParser.Parse("=Rate*100", main).Evaluate(workbook);     // 10
```

**Definição.** `Workbook.DefinedNames` é o mapa `nome → Expression`. Defina por meio de
`DefineName(string, Expression)` ou da sobrecarga de conveniência `DefineName(string, string)`, que faz o
parse do texto e **exige que toda referência seja qualificada por planilha** — uma referência não
qualificada (por exemplo, `A1:A3`) lança `ArgumentException`, já que um nome de nível de workbook não tem
planilha implícita. Um nome vazio, ou um que colida com o formato de uma referência de célula (`A1`) ou
com um literal booleano, também é rejeitado.

**Ordem de resolução.** Um `NameReference` resolve nesta ordem:

1. **Primeiro o escopo de `LET`** (*shadowing*) — uma vinculação de `LET` com o mesmo nome vence, então
   `LET(Sales, 5, Sales+1)` é `6`, não uma soma sobre o intervalo.
2. **`Workbook.DefinedNames`** — a expressão do nome é avaliada. Um intervalo/união permanece um valor de
   *referência*, então funções que aceitam intervalos o expandem (`SUM(Sales)`); uma única célula ou
   constante é avaliada para seu escalar. As funções que exigem uma referência sintática —
   `VLOOKUP`/`HLOOKUP` (tabela), `INDEX`, `OFFSET`, `ROWS`, `COLUMNS`, `AREAS`, `ISREF` — aceitam um nome
   que representa um intervalo (por exemplo, `VLOOKUP(2, Sales, 2)`).
3. Caso contrário, `#NAME?`.

**Ciclos.** Um nome que se refere a si mesmo, diretamente ou por meio de uma cadeia (`A → B → A`), é
detectado por um rastreamento thread-local e produz `#REF!` em vez de estourar a pilha.

## Da expressão de volta ao texto de fórmula

O `FormulaWriter` é o inverso do parser — ele renderiza uma expressão como texto de fórmula do Excel
(sem o `=` inicial), emitindo o mínimo de parênteses que, ao passar pelo parse de novo, reproduz a mesma
árvore:

```csharp
using Danfma.MySheet.Parsing;

var expression = ExpressionParser.Parse("=SUM(A1:A2)*Sheet2!B1", sheet);
string formula = expression.ToFormula(sheet.Name);   // "SUM(A1:A2)*Sheet2!B1"
```

O argumento `contextSheetName` controla a qualificação: referências dessa planilha ficam sem qualificação
(`A1`); referências a outras planilhas são qualificadas (`Sheet2!A1`, entre aspas quando o nome exigir).
É isso que o exportador de Excel usa no `FormulaMode.Formulas` ([Interop com Excel](excel-interop.md)).

## Veja também

- [ComputedValue e erros](computed-value.md) — leitura dos resultados da avaliação.
- [Funções personalizadas](custom-functions.md) — estendendo o conjunto de funções.
- [Referência de funções](function-reference.md) — as 164 funções nativas.
