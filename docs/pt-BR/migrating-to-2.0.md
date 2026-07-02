# Migrando para a 2.0

*Tradução do documento canônico em inglês ([migrating-to-2.0.md](../migrating-to-2.0.md)). Em caso de divergência, o inglês prevalece.*

O MySheet 2.0 é uma reorganização da AST **somente de namespaces**. Nenhum comportamento de fórmula
mudou, nenhuma API foi renomeada ou removida, e **workbooks serializados são 100% compatíveis** nos dois
sentidos da movimentação de tipos — a quebra é estritamente em tempo de compilação: suas diretivas
`using`.

## Por quê

`Danfma.MySheet.Expressions` havia crescido para ~190 tipos públicos em um único namespace plano. A 2.0
o divide seguindo as mesmas linhas de categoria da [referência de funções](function-reference.md) (que
espelha o próprio catálogo do Excel), e promove os dois tipos de *resultado* da avaliação para o
namespace raiz, ao lado de `Workbook`.

## Mapa tipo → namespace

| Tipos | Namespace na 1.x | Namespace na 2.0 |
| --- | --- | --- |
| `ComputedValue`, `ComputedValueKind`, `Error` | `Danfma.MySheet.Expressions` | `Danfma.MySheet` |
| AST principal: `Expression`, `ValueExpression`, `Function`, `Reference`, `FunctionCall`, `EvaluationContext`, `NumberValue`, `StringValue`, `BooleanValue`, `BlankValue`, `ErrorValue`, `CellReference`, `RangeReference`, `UnionReference`, `NameReference`, `BinaryOperation`, `UnaryOperation` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions` (sem mudança) |
| Funções lógicas (12): `If`, `Ifs`, `Switch`, `And`, `Or`, `Xor`, `Not`, `IfError`, `IfNa`, `Let`, `TrueFunction`, `FalseFunction` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Logical` |
| Funções de matemática e trigonometria (67): `Sum`, `SumIf`, `SumIfs`, `Product`, `Abs`, `Int`, `Round`/`RoundUp`/`RoundDown`, `Sqrt`, `Power`, `Mod`, `Sin` … `Acoth`, `Ceiling*`/`Floor*`, `Fact`, `Combin`, `Base`, `Roman`, … | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Mathematics` |
| Funções estatísticas (8): `Average`, `Count`, `CountA`, `CountBlank`, `CountIf`, `CountIfs`, `Max`, `Min` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Statistical` |
| Funções de texto (34): `Upper`, `Lower`, `Trim`, `Len`, `Left`, `Mid`, `Right`, `Value`, `Text`, `T`, `Concat`, `Concatenate`, `TextJoin`, `Find`, `Search`, `Substitute`, `RegexTest`/`RegexExtract`/`RegexReplace`, `Fixed`, `Dollar`, … | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Text` |
| Funções de informação (18): `IsNumber`, `IsBlank`, `IsError` e a família `Is*`, `N`, `Na`, `TypeFunction`, `ErrorType`, `SheetNumber`, `SheetsCount` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Information` |
| Funções de pesquisa e referência (16): `VLookup`, `HLookup`, `XLookup`, `Lookup`, `Match`, `XMatch`, `Index`, `Choose`, `Offset`, `Row`, `Rows`, `Column`, `Columns`, `Address`, `Areas`, `FormulaText` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Lookup` |
| Funções financeiras (9): `Pmt`, `Pv`, `Fv`, `Nper`, `Ipmt`, `Ppmt`, `Npv`, `Rate`, `Irr` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Financial` |

As pastas espelham os namespaces, então o arquivo-fonte de um tipo está sempre onde o namespace indica.

Observações sobre o desenho:

- A categoria de uma função é a sua seção na [referência de funções](function-reference.md), que espelha
  o catálogo do Excel. É por isso que `SUM`/`SUMIF(S)` são **Mathematics** enquanto
  `AVERAGE`/`COUNT*`/`MAX`/`MIN` são **Statistical**, e por isso que `T` é **Text** (o Excel a lista lá),
  embora tenha sido lançada junto com as funções de informação.
- O namespace é `Mathematics`, e não `Math`, para que `Math.Sqrt(…)` continue resolvendo para
  `System.Math` dentro e ao redor desses tipos.
- `FunctionCall` (o nó da AST para funções personalizadas) permanece no namespace principal
  `Danfma.MySheet.Expressions`: ele é o nó de extensibilidade, não uma categoria.

## Atualizando o seu código

A maioria dos pontos de chamada apenas faz parse e avalia — eles tocam `Workbook`, `ExpressionParser`,
`ComputedValue` e os nós de valor. Para esses, nada muda se você já importa `Danfma.MySheet` (onde
`Workbook` vive — `ComputedValue` e `Error` agora vêm junto):

```csharp
// 1.x
using Danfma.MySheet;             // Workbook, Sheet
using Danfma.MySheet.Expressions; // NumberValue, ComputedValue, Error
using Danfma.MySheet.Parsing;     // ExpressionParser

// 2.0 — as mesmas diretivas continuam compilando; ComputedValue/Error simplesmente passam a vir de Danfma.MySheet.
using Danfma.MySheet;             // Workbook, Sheet, ComputedValue, ComputedValueKind, Error
using Danfma.MySheet.Expressions; // Expression, NumberValue, FunctionCall, …
using Danfma.MySheet.Parsing;     // ExpressionParser, FormulaWriter
```

Código que nomeia os records de função diretamente (construindo árvores à mão, fazendo pattern matching
nos nós) adiciona os namespaces de categoria que utiliza:

```csharp
// 1.x
using Danfma.MySheet.Expressions;

Expression tree = new Sum([new CellReference("A1", "Sheet1"), new NumberValue(2)]);
bool isLookup = tree is VLookup or XLookup;

// 2.0
using Danfma.MySheet.Expressions;             // CellReference, NumberValue
using Danfma.MySheet.Expressions.Lookup;      // VLookup, XLookup
using Danfma.MySheet.Expressions.Mathematics; // Sum

Expression tree = new Sum([new CellReference("A1", "Sheet1"), new NumberValue(2)]);
bool isLookup = tree is VLookup or XLookup;
```

Atenção a duas colisões de nomes ao importar os novos namespaces:

- `Danfma.MySheet.Expressions.Lookup.Index` vs. `System.Index` — qualifique (`Lookup.Index`) ou crie um
  alias (`using IndexFunction = Danfma.MySheet.Expressions.Lookup.Index;`) onde os dois estiverem em
  escopo.
- `Text` e `Lookup` são, ao mesmo tempo, um namespace **e** um record dentro dele (`…Text.Text` é a
  função `TEXT`, `…Lookup.Lookup` é `LOOKUP`); código que está dentro do próprio namespace
  `Danfma.MySheet.Expressions` precisa qualificar esses dois records.

## Compatibilidade de serialização

O formato MemoryPack está intocado. Os nós de expressão são serializados através de uma
`[MemoryPackUnion]` baseada em tags, e **nenhuma tag mudou** — namespaces não fazem parte do formato de
fio. Workbooks salvos pela 1.x carregam na 2.0 (e vice-versa) byte a byte. Isso é protegido por um teste
de regressão: uma fixture binária serializada com o layout da 1.x
(`tests/Danfma.MySheet.Tests/Fixtures/workbook-pre-namespaces.msgpack.bin`) é carregada e reavaliada a
cada execução.

## Comportamento

Nenhuma das 164 funções nativas mudou de semântica, o parser e o `FormulaWriter` aceitam e emitem
exatamente o mesmo texto, e os resultados da avaliação são idênticos. Se um upgrade 1.x → 2.0 mudar
qualquer valor calculado, isso é um bug — por favor, reporte.
