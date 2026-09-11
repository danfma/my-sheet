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
- O indexador `this[string]` **lança `KeyNotFoundException`** para um nome sem planilha correspondente —
  é uma consulta direta ao host, como em um dicionário. Use `TryGetSheet(name, out sheet)` para sondar
  sem `try`/`catch`. (Uma planilha ausente referenciada *dentro de uma fórmula* é outra história: ela
  resolve para `#REF!` em vez de lançar exceção — veja [semântica de erros de avaliação](#parsing).)
- `Sheets` é um `ConcurrentDictionary<string, Sheet>`, seguro para leitores concorrentes (o cenário
  pretendido de extração em segundo plano).
- `Sheets.Add(name)` atribui à planilha um `Index` igual à sua ordem de inserção — é isso que a função
  `SHEET` reporta, e o que define a ordem das abas ao exportar para Excel.

Principais membros de `Workbook`:

| Membro | Propósito |
| --- | --- |
| `Sheets` / `this[string]` | Acessa planilhas pelo nome (case-insensitive); o indexador lança `KeyNotFoundException` para um nome sem planilha correspondente. |
| `TryGetSheet(name, out sheet)` | Consulta de planilha sem lançar exceção (case-insensitive) → `bool`; a contraparte do indexador que lança exceção. |
| `GetCellValue(sheetName, id)` | Avaliação memoizada de uma célula → `ComputedValue`; uma referência a uma planilha ausente resolve para `#REF!` (nunca lança exceção). |
| `Sheet.CellAddresses` / `Sheet.EnumerateCells()` | Enumeração de células populadas (struct enumerators, sem boxing de interface): `CellAddresses` produz `(Column, Row)` sem alocação (só células canônicas); `EnumerateCells()` produz `(Id, Column, Row)` derivando o id canônico uma vez por célula (ids overflow incluídos com `0,0`). |
| `CellRef.TryFormat(col, row, span, out written)` / `CellRef.Format(col, row)` | Formatação de id A1 a partir de endereços numéricos; a forma span (`TryFormat`) não aloca nada (18 chars sempre bastam). |
| `GetValueReader(sheetName)` | Leitor em massa por endereço numérico para uma planilha → `SheetValueReader`; `GetValue(column, row)` serve valores memoizados sem string de id, sem parse A1 e sem hash do nome da planilha por célula — misses avaliam sob demanda, idêntico ao `GetCellValue`. Veja [Leituras em massa](#leituras-em-massa-getvaluereader). |
| `ComputeAll()` | Avalia avidamente (eagerly) todas as células (a contraparte "calcular agora" do `GetCellValue` preguiçoso), preenchendo o cache para que um salvamento subsequente a quente carregue os valores computados. Roda em uma large stack para cadeias profundas; uma segunda chamada é toda hits. Após edições, chame `InvalidateCache()` primeiro. |
| `InvalidateCache()` | Esvazia explicitamente **todo** o cache de memoização (obrigatório após edições); também reinicia a época volátil. |
| `Recalculate()` | Atualiza apenas as células voláteis (veja [Funções voláteis](#funções-voláteis)); mantém em cache toda célula estável. |
| `TimeProvider` | Relógio injetável para `NOW`/`TODAY` (padrão `TimeProvider.System`, lido em horário local). |
| `RandomSeed` | Semente `int?` opcional para `RAND`/`RANDBETWEEN` (valor fixo → execuções reproduzíveis). |
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
bool removed = sheet.Remove("B1");                       // exclui uma célula → true se ela existia

foreach (var (id, expression) in sheet) { /* itera as células armazenadas */ }
```

- O **getter nunca lança exceção**: ler um id que nunca foi definido retorna `BlankValue.Instance`, que
  é avaliado como em branco — exatamente como o Excel trata uma célula vazia.
- **A escrita passa pelo `set` do indexador, e a exclusão passa por `Remove`** — os dois, e únicos,
  caminhos que alteram as células de uma planilha. `Remove(id)` retorna `true` quando havia uma célula
  ali e `false` para uma operação sem efeito. Assim como o `set`, o `Remove` não limpa o cache de
  memoização por conta própria: depois de editar (escrever ou remover), chame `workbook.InvalidateCache()`
  antes de ler de novo para que a mudança seja observada.
- **`Cells` é uma visão somente leitura** (`IReadOnlyDictionary<string, Expression>`) das células
  armazenadas — enumerável e indexável para leitura (`sheet.Cells["A1"]`, `sheet.Cells.Count`), mas não
  mutável; mute através do indexador e de `Remove`. (Para quem está migrando da 2.x e mutava `Cells`
  diretamente: veja [Migrando para a 3.0](migrating-to-3.0.md).)
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
  referência inválida como `#REF!`, e assim por diante, na forma de erros de `ComputedValue`. Uma
  **referência a uma planilha que não existe** (`=Ghost!A1`, `SUM(Ghost!A:A)`) é uma dessas referências
  inválidas: ela resolve para `#REF!` — nunca uma `KeyNotFoundException` lançada — de forma que uma única
  referência cruzada pendente não consegue abortar um lote de workbook inteiro. Uma planilha ausente é um
  erro *estrutural*, então ele se propaga por **todas** as funções consumidoras — as agregações, a família
  `COUNT` que ignora erros (`COUNT`/`COUNTA`/`COUNTIF` sobre uma planilha fantasma são `#REF!`, não `0`),
  os lookups (`VLOOKUP`/`MATCH`/`XLOOKUP`/`INDEX`), `SUMPRODUCT` e os pares estatísticos, as funções
  financeiras de fluxo de caixa (`NPV`/`IRR`/`XIRR`/`MIRR`), as junções de texto (`CONCAT`/`TEXTJOIN`) e as
  demais. Um erro de valor *dentro* de uma célula de uma planilha existente mantém sua política habitual
  por função (`COUNT` o ignora, `SUM` o propaga), e um intervalo vazio sobre uma planilha *existente*
  continua sendo um resultado de valor (`MATCH` sobre ele é `#N/A`, não `#REF!`).

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

### Leituras em massa: `GetValueReader`

Um laço de extração que constrói ids — `GetCellValue(sheetName, "C" + r)` — paga três custos por célula
que o resultado não precisa: a alocação da string do id, o parse A1 e um lookup por hash do nome da
planilha. `Workbook.GetValueReader(sheetName)` resolve o handle da planilha uma única vez e lê por
endereço numérico (1-based):

```csharp
var reader = workbook.GetValueReader("Results");

for (var row = 2; row <= 60_001; row++)
{
    if (reader.GetValue(column: 2, row).TryGetNumber(out var value))
    {
        total += value;
    }
}
```

A semântica é idêntica à do `GetCellValue`: um hit é uma leitura direta do value store paginado; um miss
avalia sob demanda (memoização, guarda de ciclos), então literais e fórmulas nunca computadas também são
servidos. `InvalidateCache()` se aplica da mesma forma, e a instância do reader permanece válida entre
invalidações. Medido numa extração de 360 mil células: `29,8 ms / 24,2 MB` alocados com ids por célula →
`6,9 ms / 0 bytes` com o reader.

Para dirigir o laço a partir da própria planilha (em vez de limites conhecidos), enumere as células
populadas. O pipeline totalmente livre de alocação — incluindo o texto do id, por exemplo para um campo
JSON — combina `CellAddresses`, `CellRef.TryFormat` e o reader:

```csharp
var reader = workbook.GetValueReader("Results");
Span<char> id = stackalloc char[18];

foreach (var (column, row) in sheet.CellAddresses)      // struct enumerator, zero alocação
{
    CellRef.TryFormat(column, row, id, out var length); // texto do id sem string
    jsonWriter.WriteString(id[..length], reader.GetValue(column, row).ToDouble());
}
```

Quando você precisa do id como `string` de qualquer forma, `sheet.EnumerateCells()` produz
`(Id, Column, Row)` — a lib deriva o id canônico uma vez por célula (uma string cada), e o endereço
numérico casa direto com o reader. A ordem de enumeração é a de inserção, não row-major; ordene se
precisar de saída determinística.

### Resultados de fórmula nunca são em branco (paridade com o Excel)

Na **borda da célula** — `GetCellValue` — o resultado de uma fórmula nunca é em branco, exatamente como
no Excel: quando uma célula que TEM conteúdo (sua expressão não é o `BlankValue` vazio) avalia para em
branco, `GetCellValue` retorna `ComputedValue.Number(0)`, e é esse `0` coagido que entra no cache. Uma
célula verdadeiramente vazia (sua expressão é `BlankValue.Instance`, por exemplo um id que nunca foi
definido) permanece em branco.

```csharp
sheet["A1"] = Cell("F10", sheet);                       // =F10, F10 vazio
workbook.GetCellValue("Sheet1", "A1").ToDouble();        // 0   (resultado de fórmula coagido)
workbook.GetCellValue("Sheet1", "F10").Kind;             // Blank (célula verdadeiramente vazia)

sheet["A2"] = ExpressionParser.Parse("=IF(TRUE, F10)", sheet);
workbook.GetCellValue("Sheet1", "A2").ToDouble();        // 0   (ramo em branco coagido)
```

A coerção pertence à **célula**, não à expressão: `Evaluate` mantém o valor em branco INTERNAMENTE,
então em branco ainda se compara como `""`/`0`/`FALSE` dentro de uma expressão. Isso preserva a
semântica interna enquanto casa com o Excel na borda de exibição:

```csharp
sheet["A3"] = ExpressionParser.Parse("=IF(F10=\"\",1,2)", sheet);
workbook.GetCellValue("Sheet1", "A3").ToDouble();        // 1   (F10 vazio ainda é igual a "" internamente)
sheet["A4"] = ExpressionParser.Parse("=F10&\"\"", sheet);
workbook.GetCellValue("Sheet1", "A4").ToText();          // ""  (resultado é texto, não em branco → não coagido)
```

Os efeitos de paridade se propagam, todos batendo com o Excel: `ISBLANK(A1)` com `A1 = "=F10"` é
**FALSE** (A1 agora é 0), `COUNT` conta uma célula formula-vazia (0 é um número) enquanto `COUNTBLANK`
não conta mais, e a exportação `SaveAsExcel` `ValuesOnly` grava `0` para uma célula formula-vazia em vez
de omiti-la.

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
- Um intervalo puro é consumido pelas funções que o aceitam (`SUM`, `COUNT`, lookups, …); ele não tem valor
  escalar próprio. Avaliado diretamente (`Parse("=A1:B2", sheet).Evaluate(workbook)`) é `#VALUE!`, mas a
  mesma fórmula armazenada em uma **célula** sofre
  [interseção implícita](#interseção-implícita-na-fronteira-da-célula) com a linha e a coluna dessa célula.
- Um nome puro que não é um id de célula (por exemplo, `=total`) é um `NameReference` — ele resolve
  contra as vinculações de `LET` e os [intervalos nomeados](#intervalos-nomeados) (*named ranges*) do
  workbook em tempo de avaliação, e produz `#NAME?` se não estiver vinculado.

## Referências de coluna e linha inteira

O MySheet suporta referências que são **abertas** (ilimitadas) em pelo menos um lado: uma coluna inteira,
uma linha inteira, e as formas mistas de um lado só.

```csharp
ExpressionParser.Parse("=SUM(A:A)", sheet);     // coluna inteira A
ExpressionParser.Parse("=SUM(A:C)", sheet);     // colunas A..C
ExpressionParser.Parse("=SUM(1:1)", sheet);     // linha inteira 1
ExpressionParser.Parse("=SUM(1:5)", sheet);     // linhas 1..5
ExpressionParser.Parse("=SUM(A2:A)", sheet);    // coluna A a partir da linha 2 para baixo
ExpressionParser.Parse("=SUM(A:A10)", sheet);   // coluna A até a linha 10
ExpressionParser.Parse("=SUM(A1:C)", sheet);    // colunas A..C a partir da linha 1 para baixo
ExpressionParser.Parse("=SUM(Data!A:A)", main); // qualificada por planilha; $A:$A é aceito e normalizado
```

Isso é convertido em um único `OpenRangeReference(int? ColMin, int? ColMax, int? RowMin, int? RowMax,
string SheetName)`; cada limite é `null` no lado que estiver aberto. O endpoint da **esquerda** dá os
limites inferiores (`ColMin`/`RowMin`), o da **direita** os limites superiores (`ColMax`/`RowMax`); um
endpoint que nomeia só uma coluna não informa linha nenhuma (e vice-versa), então esse eixo permanece
aberto naquele lado. Quando os quatro limites são conhecidos, o parser produz um `RangeReference` comum
em vez disso.

**Semântica — células populadas.** Um intervalo aberto significa *as células populadas dentro dos
limites*, não uma grade fixa. A enumeração varre as células armazenadas da planilha e mantém aquelas cuja
(coluna, linha) caem dentro dos limites não nulos; células em branco contribuem `0`, então `SUM(A:A)`
corresponde ao Excel sem nunca materializar a coluna de 1.048.576 linhas. Uma coluna vazia soma `0`. Isso
mantém a agregação de coluna/linha inteira barata no modelo esparso que o MySheet usa.

**`:` força semântica de referência.** Um endpoint só de letras adjacente a `:` é uma **coluna**, e um
endpoint inteiro uma **linha**, mesmo quando existe um intervalo nomeado com a mesma grafia — então
`Sales:Sales` é a coluna `SALES`, não o intervalo nomeado. (Essa borda só importa se você nomear algo com
o rótulo de uma coluna.)

**`ROWS` / `COLUMNS` — uma divergência documentada em relação ao Excel.** O Excel reporta o tamanho fixo
da grade (`ROWS(A:A)` = 1.048.576). Um modelo sem grade não tem essa grade, então o MySheet usa a
**extensão populada** num eixo aberto e a **contagem estrutural exata** num eixo limitado:

| Fórmula        | Resultado no MySheet                                    | Excel      |
| -------------- | --------------------------------------------------------- | ---------- |
| `ROWS(A:A)`    | linha populada máxima − linha populada mínima + 1 (0 se vazia) | 1.048.576 |
| `COLUMNS(A:C)` | `3` (estrutural, exato)                                    | `3`        |
| `COLUMNS(A:A)` | `1`                                                         | `1`        |
| `ROWS(1:5)`    | `5` (estrutural)                                            | `5`        |

**Consumidores de referência.** Onde um intervalo concreto é exigido — `VLOOKUP`/`HLOOKUP` (tabela),
`INDEX`, `OFFSET` (base) — um intervalo aberto resolve para a **caixa delimitadora populada** dentro de
seus limites; `AREAS` conta como uma área e `ISREF` reporta `true`. Assim, `VLOOKUP(2, A:B, 2)` e
`INDEX(A:A, 3)` funcionam.

**Fora de escopo.** Interseção espacial de dois intervalos abertos não é modelada.

## Interseção implícita na fronteira da célula

Uma fórmula cujo valor **final** é uma referência multicélula não é um erro dentro de uma célula. O MySheet
aplica a interseção implícita do Excel — o operador `@` — contra a linha e a coluna da própria célula da
fórmula:

| A fórmula da célula denota | A célula mostra |
| --- | --- |
| um intervalo de **uma única coluna** que abrange a **linha** da fórmula | a célula dessa linha na coluna — `=A1:A3` em `C3` é `A3` |
| um intervalo de **uma única linha** que abrange a **coluna** da fórmula | a célula dessa coluna na linha — `=A1:C1` em `B5` é `B1` |
| um intervalo **1x1** | essa célula, onde quer que a fórmula esteja — `=A1:A1` em `Z99` é `A1` |
| um intervalo do qual a linha/coluna da fórmula está **fora** | `#VALUE!` — `=A1:A3` em `C9` |
| um intervalo maior que uma célula em **ambos** os eixos | `#VALUE!` — `=A1:C3` em `B2` |
| uma **união** de áreas | `#VALUE!` — `=(A1:A3,B1:B3)` não tem um único eixo de linha/coluna para intersectar |

Detalhes:

- **Limites declarados, não os populados.** Um [intervalo aberto](#referências-de-coluna-e-linha-inteira)
  usa os limites que declara, então `=A:A` na linha 7 é `A7` mesmo que `A7` esteja vazia (o branco então
  vira `0` pela [regra do nunca-em-branco](#resultados-de-fórmula-nunca-são-em-branco-paridade-com-o-excel)),
  `=1:1` em `B7` é `B1`, e `=A2:A` é `#VALUE!` na linha 1, mas `A3` na linha 3. Isso deliberadamente **não**
  é a extensão populada que `ROWS`/`COLUMNS` usam.
- **Posicional e independente de planilha.** Apenas o número da linha e da coluna da célula da fórmula entram
  na regra: `=Sheet1!A1:A3` digitado em `Sheet2!C2` é `Sheet1!A2`. *(Inferência, não medição: o Excel define
  o `@` puramente em termos de linha e coluna, sem nenhum termo de planilha; esse caso entre planilhas não
  foi verificado no Excel.)*
- **Tudo o que denota uma referência segue a mesma tabela**, não apenas um intervalo literal — `=MyName`,
  `=INDIRECT("MyName")`, `=OFFSET(A1,0,0,3,1)`, `=CHOOSE(1,A1:A3)`, `=+A1:A3` e `=LET(x,A1:A3,x)` todos
  sofrem a interseção. Antes desta regra eles armazenavam um valor do tipo referência que todo acessador
  tipado lia de volta como branco.
- **Dentro de uma fórmula nada muda.** `=SUM(A1:A3)` continua sendo uma soma sobre três células: quem decide
  o que uma referência multicélula significa é o *consumidor*, não a célula. Só uma referência que sobrevive
  como valor final da célula sofre a interseção.
- **O caminho direto de `Expression.Evaluate` continua produzindo `#VALUE!`.**
  `ExpressionParser.Parse("=A1:A3", sheet).Evaluate(workbook)` não tem célula de fórmula com a qual
  intersectar. A regra vive em `Workbook.EvaluateCell`, que é o ponto de estrangulamento único de toda
  leitura de célula (`GetCellValue`, o [leitor de valores](#leituras-em-massa-getvaluereader), o snapshot de
  warm start, a exportação `.xlsx`) — por isso um `ComputedValueKind.Reference` nunca pode ser o valor de
  uma célula.
- **Uma célula que intersecta a si mesma é `#REF!`.** `=A1:A3` em `A2` desreferencia `A2`, a célula que já
  está na pilha de avaliação, então a guarda de ciclos responde `#REF!` (o Excel, em vez disso, levanta a
  caixa de diálogo de referência circular).
- **Sem *spill*.** A célula intersectada é o resultado inteiro — o MySheet nunca escreve nas células
  vizinhas.

**Divergências que vale conhecer.** Um intervalo 2-D responde `#VALUE!` mesmo quando a célula da fórmula
está dentro do retângulo. Isso segue a regra do `@` como documentada, mas é o ponto mais provável de
divergência em relação a um Excel real e **não** foi medido; a leitura alternativa — a (coluna, linha) da
própria célula da fórmula quando o retângulo a contém — é um ramo a mais em `ImplicitIntersection`. Uma
união também fica deliberadamente em `#VALUE!`. Um **array computado** nunca chega a esta regra — um
`=A1:C3*E1:E3` ou um `=LEN(A1:A3)` puro em uma célula é o `#VALUE!` do próprio operador ou da própria função,
e não uma interseção — e a resposta do Excel para a forma digitada desses casos está medida e registrada em
[argumentos implícitos de array](#argumentos-implícitos-de-array); reconciliar as duas é a metade de array
desta regra.

## Argumentos implícitos de array

Algumas funções avaliam um **argumento com valor de array** elemento a elemento, reproduzindo a semântica
implícita (CSE) do Excel — sem `Ctrl+Shift+Enter`, sem *spilling* e sem um valor de array público: o vetor
vive apenas dentro da avaliação da função consumidora. Isso fecha os idiomas comuns `SUM(IF(range=…))` /
`SMALL(IF(range=…, ROW(range)))`.

```csharp
ExpressionParser.Parse("=SUM(IF(B2:B5=\"Show\",1,0))", sheet);            // → 2 (contagem de correspondências)
ExpressionParser.Parse("=SMALL(IF(B2:B5=\"Show\",ROW(B2:B5)),1)", sheet); // → 3 (primeira linha correspondente)
ExpressionParser.Parse("=INDEX(ROW(B2:B5),1)", sheet);                    // → 2 (vetor de linhas, indexado)
ExpressionParser.Parse("=INDEX(ROW($A:$A),4)", sheet);                    // → 4 (identidade: n-ésima linha)
```

**De qual modo de entrada vêm os números do Excel nesta página.** O Excel responde a estas formas de maneira
*diferente* conforme o modo como a fórmula foi inserida. Digitada normalmente, ele aplica a interseção
implícita legada **dentro** do argumento; inserida como array (`Ctrl+Shift+Enter`), ele avalia o array
inteiro. Por isso `SUM(A1:C3*E1:E3)` sobre a fixture abaixo é `#VALUE!` digitada e **108** inserida como
array, e `SUM(ROW(A1:C3))` é **1** digitada — ali o `ROW` de um retângulo responde a sua única linha
superior — contra **6** inserida como array. O MySheet não tem nenhum conceito de entrada como array: uma
fórmula é uma fórmula, e todos os consumidores descritos aqui implementam a regra da entrada **como array**.
Portanto, todo número do Excel citado nesta seção e na lista de divergências dela é o da entrada como array,
a não ser que a linha diga *digitada* (tudo medido no Aspose.Cells 26.6.0, em 2026-09-10). Se você digitar
uma destas fórmulas em um Excel de verdade e comparar, espere a resposta da forma digitada, e não a nossa; o
único ponto em que a resposta do próprio MySheet não segue nenhuma das duas é
[a fronteira da célula](#interseção-implícita-na-fronteira-da-célula), abaixo.

**Suportado.** Os consumidores são os agregadores numéricos (`SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX`,
`PRODUCT` e — através da mesma dobra — `MEDIAN`, a família `STDEV`/`VAR`, `SMALL`, `LARGE`, os percentis e
quartis), `INDEX`, `ROWS`/`COLUMNS` (que informam a *extensão* do array, e não os valores dele),
`COUNTA`/`CONCAT`/`TEXTJOIN` (que transmitem os elementos dele), `SUMPRODUCT` e a **forma-array** do
[`AGGREGATE`](function-reference.md) (`function_num` 14-19), onde a opção 6 descarta os elementos
`#DIV/0!` que fazem um `SMALL` simples sobre o mesmo vetor falhar. O `SUBTOTAL` e a forma-*referência* do
`AGGREGATE` (1-13) deliberadamente **não** são consumidores: os argumentos deles são `ref`, e o Excel
rejeita um array computado em um deles — `SUBTOTAL(9,ROW(A1:A3))` e `AGGREGATE(9,4,ROW(A1:A3))` são
`#VALUE!` ali, medido no Aspose.Cells 26.6.0, e é exatamente por isso que o AGGREGATE documenta uma
segunda sintaxe para arrays. Um argumento é avaliado como um array quando é uma comparação de **intervalo
fechado** (`B2:B5="Show"`), um `IF` cuja condição é um array assim (com ou sem ramo `else`),
`ROW`/`COLUMN` sobre um retângulo, uma das duas formas **elevadas** (*lifted*) descritas mais abaixo, ou um
dos quatro [produtores de array dinâmico](#produtores-de-array-dinâmico)
(`FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`).
Esse retângulo pode estar escrito literalmente — `ROW(A1:C3)` é a coluna 3x1 `[1,2,3]` e `COLUMN(A1:C3)` a
linha 1x3 `[1,2,3]`, um número por posição de linha ou coluna e não um por célula, de modo que
`SUM(ROW(A1:C3))` e `SUM(COLUMN(A1:C3))` são ambos 6 e o `COUNT` de qualquer um deles é 3 (inseridos como
array; digitados, os dois somam 1) — ou apenas ser *denotado* pelo argumento — um
[nome definido](#intervalos-nomeados) (`SUM(ROW(MyName))` = 6 e `COUNT(ROW(MyName))` = 3 para um nome sobre
três linhas, enquanto `COUNT(MyName)` conta os valores das próprias células) ou um intervalo `:` com
extremidades que retornam referências (`SUM(ROW(INDEX(A1:A3,1,1):A3))` = 6). Um `IF` sem ramo produz um
lógico `FALSE` onde a condição é falsa, e os
agregadores ignoram lógicos/texto (exatamente por isso `SMALL(IF(…))` pula as linhas sem correspondência).
O primeiro erro por elemento prevalece, como no Excel.

**Como dois arrays de formatos diferentes se combinam (propagação, *broadcasting*).** Um escalar é propagado
para todas as posições — e um **intervalo 1x1** também, porque ele igualmente é um vetor:
`SUM(A1:C3*E1:E1)` = 45. Fora disso, cada eixo é decidido por conta própria, e uma extensão igual a 1 cede à
do outro lado: uma **coluna** Nx1 se repete em todas as colunas (`SUM(A1:C3*E1:E3)` = 108), uma **linha** 1xM
se repete em todas as linhas (`SUM(A1:C3*E5:G5)` = 960), e uma Nx1 contra uma 1xM é o **produto externo** NxM
(`SUM(E1:E3*E5:G5)` = 360). O `IF` dobra a condição e *os dois* ramos em uma única extensão pela mesma regra
(`SUM(IF(E1:E3>1,A1:C3,0))` = 39, `SUM(IF(E1:E3>1,E5:G5,0))` = 120). Quando as duas extensões passam de 1 e
diferem, o resultado assume a **maior** delas e toda posição que o operando mais curto não cobre é
**`#N/A`** — um erro por elemento, não da expressão inteira, então um consumidor que descarta ou captura
erros ainda responde: `SUM(A1:C3*H1:H2)` é `#N/A`, enquanto `COUNT(A1:C3*H1:H2)` = 6 e
`AGGREGATE(15,6,A1:C3*H1:H2,6)` = 12. O `SUMPRODUCT` mantém a própria regra, mais estrita, e apenas para a
**sua** lista de argumentos: eles precisam bater exatamente, inclusive um 1x1 — `SUMPRODUCT(A1:C3,E1:E3)` e
`SUMPRODUCT(A1:C3,E1:E1)` são ambos `#VALUE!`, onde o operador teria propagado os dois —, enquanto uma
propagação escrita *dentro* de um argumento é consumida como o array que ela computa
(`SUMPRODUCT(A1:C3*E1:E3)` = 108). Fixture: `A1:C3` = 1…9 por linhas, `E1:E3` = 1, 2 e 3, `E5:G5` = 10, 20 e
30, `H1:H2` = 1 e 2. Todos os números daqui foram medidos no Aspose.Cells 26.6.0, em 2026-09-10, com entrada
como array — digitadas, todas as formas com `SUM` acima dão `#VALUE!` ali e `COUNT(A1:C3*H1:H2)` dá 0,
enquanto as formas com `AGGREGATE` e `SUMPRODUCT` respondem o mesmo nos dois modos (são nativas de array em
qualquer entrada) — e estão fixados por `VectorBroadcastingTests`, `MiniCseConsumerTests` e
`MathAggregateTests`.

**Operadores unários e funções escalares elevados (*lifted*).** Duas outras formas passam a produzir um
array onde qualquer um dos consumidores acima pede um, aplicando um corpo escalar elemento a elemento:

- um `-` ou `%` unário sobre um array — com `A1:A3` = 1, 2 e 3: `SUM(-A1:A3)` = -6, `SUM(A1:A3%)` = 0.06,
  `SUM(-(A1:A3>1))` = -2 e o idioma da dupla negação `SUMPRODUCT(--(A1:A3>1))` = 2;
- qualquer **função nativa puramente escalar** com ao menos um argumento de array — `SUM(LEN(D7:F9))` = 7
  sobre um retângulo 3x3 contendo `"abc"`, `"def"` e um espaço, `COUNT(LEN(D7:F9))` = 9 (nove comprimentos,
  incluindo os das células vazias), elevações aninhadas (`SUM(LEN(TRIM(D7:F9)))` = 6),
  `SUM(ABS(A1:A3*-1))` = 6, `SUM(ROUND(A1:A3,0))` = 6, `SUM(ISNUMBER(A1:A3)*1)` = 3,
  `SUM(IFERROR(A1:A3,0))` = 6 e o idioma completo de planilha
  `IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))>0,"Show","Hide")`.

**180 das 310 funções nativas registradas** podem ser elevadas: as puramente escalares (texto, matemática,
financeiras, datas, informação, as auxiliares estatísticas escalares (`FISHER`, `PERMUT`, `PHI`, `STANDARDIZE`, …),
`IFERROR`/`IFNA`/`IFS`/`NOT`/`SWITCH`, `ADDRESS`). As outras 130 são **cientes de intervalos** e o MySheet nunca as
eleva, porque já consomem intervalos ou arrays por conta própria — ou, no caso dos quatro
[produtores de array dinâmico](#produtores-de-array-dinâmico), porque *produzem* um — `SUM`, `COUNT`, `INDEX`,
`ROW`, `COLUMN`, `ROWS`,
`COLUMNS`, `AREAS`, `SUMPRODUCT`, `SUBTOTAL`, `AGGREGATE`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`,
`VLOOKUP`, `MATCH`, `OFFSET`, `INDIRECT`, `IF`, `LET`,
`RANDBETWEEN`, `AND`/`OR`/`XOR`, as séries de fluxo de caixa (`NPV`, `IRR`, …), as estatísticas de população inteira e
de arrays pareados (`RANK`, `MODE`, `CORREL`, `SUMXMY2`, …) e a família de critérios. Essa é a regra do MySheet,
**não** a do Excel: o Excel também eleva uma função ciente de intervalos, sobre os slots que recebem um *escalar*,
enquanto continua consumindo o intervalo no slot que recebe um — a última das divergências conhecidas abaixo, com os
doze casos medidos. Uma [função personalizada](custom-functions.md) também nunca é elevada — ela não tem entrada no
registro, então permanece um escalar avaliado uma única vez.

Dentro de uma chamada elevada:

- **Argumentos escalares são propagados** (*broadcast*) e cada um é avaliado exatamente **uma vez** por
  avaliação, e não uma vez por elemento: `SUM(ROUND(A1:A3,0))` lê o `0` uma vez, e uma volátil ou uma função
  do host em um slot escalar é chamada uma única vez para todo o vetor.
- **Dois argumentos de array são propagados eixo a eixo**, pela mesma regra que os operandos de um operador
  seguem (o parágrafo sobre propagação acima). Formatos iguais são pareados posição a posição
  (`SUM(ROUND(A1:A3,B1:B3))` = 6 sobre 1, 2, 3 e 10, 20, 30); uma extensão igual a 1 cede à do outro lado,
  então `SUM(ROUND(A1:C3,E1:E3))` — 3x3 contra 3x1 — é 45; e onde as duas extensões passam de 1 e diferem, as
  posições não cobertas são `#N/A`, então `SUM(ROUND(A1:C3,H1:H2))` é `#N/A` enquanto
  `COUNT(ROUND(A1:C3,H1:H2))` = 6. Uma versão anterior entregava ao corpo um marcador `#VALUE!` para *cada*
  elemento; isso acabou, o que muda o que um corpo que *consome* erros vê —
  `SUM(LEN(IFERROR(LEFT(D7:F9,E6:E8),"zz")))` é 0, e não 18, porque a coluna em branco propagada torna cada
  elemento `LEFT(x,0)` = `""` e não sobra nada para o `IFERROR` recuperar (tudo inserido como array, medido no
  Aspose.Cells 26.6.0).
- **Um argumento omitido mantém o padrão da própria função.** `FIXED(A1:A3,,TRUE)` continua formatando duas
  casas decimais por elemento: um slot omitido permanece um literal em branco na árvore, em vez de virar um
  slot por elemento, então o ramo "argumento não fornecido" da função ainda é acionado.
- **Erros se propagam por elemento**, prevalecendo o primeiro na ordem de varredura por linhas:
  `SUM(ABS(1/(A1:A3-2)))` é `#DIV/0!` e `SUM(LEN(Ghost!A1:A3))` é `#REF!` — o caminho elevado lê células, e
  portanto a regra de planilha ausente se aplica (ao contrário de `SUM(ROW(Ghost!A1:A3))`, a divergência
  abaixo). Texto onde se espera um número torna aquele elemento `#VALUE!` (`SUM(ABS(B1:B3))` com
  `B2` = `"x"`), e um elemento em branco é convertido para `0`, então `SUM(LEN(A1:A3))` sobre três células
  vazias é `0` enquanto `COUNT(LEN(A1:A3))` é `3`.

### Produtores de array dinâmico

`FILTER`, `SORT`, `UNIQUE` e `SEQUENCE` são a outra direção da mesma maquinaria: em vez de transformar um
intervalo em array, eles **produzem** um. Em qualquer lugar onde um dos consumidores acima aceita um array,
uma dessas quatro pode ficar no lugar dele, e elas compõem entre si e com todas as formas descritas acima.

```csharp
// A1:A3 = 5, 0, 9;  B1:B3 = 1, 2, 3;  Q1:Q4 = 9, 5, 9, 0
ExpressionParser.Parse("=SUM(FILTER(A1:A3,A1:A3>0))", sheet);   // → 14  (o 0 não passa no predicado)
ExpressionParser.Parse("=ROWS(FILTER(A1:A3,A1:A3>0))", sheet);  // → 2   (quantas linhas casaram)
ExpressionParser.Parse("=INDEX(SORT(A1:A3,1,-1),1)", sheet);    // → 9   (o maior)
ExpressionParser.Parse("=COUNTA(UNIQUE(Q1:Q4))", sheet);        // → 3   (contagem de distintos)
ExpressionParser.Parse("=SUM(SEQUENCE(5))", sheet);             // → 15  (1+2+3+4+5)
ExpressionParser.Parse("=INDEX(SEQUENCE(2,3),2,2)", sheet);     // → 5   (preenchimento por linha)
```

Os argumentos, padrões, conversões e regras de erro de cada função estão na referência de funções —
`FILTER`/`SORT`/`UNIQUE` em [Pesquisa e referência](function-reference.md#pesquisa-e-referência-20),
`SEQUENCE` em [Matemática e trigonometria](function-reference.md#matemática-e-trigonometria-76). O que esta
seção cobre é como elas se comportam *como arrays*.

**Todo consumidor as lê, sem nenhum caso especial por função.** Elas são operandos na mesma árvore preguiçosa,
então toda a lista de **Suportado** acima se aplica sem alteração. Medido neste motor sobre `A1:A3` = 5, 0, 9:
`SUM(FILTER(A1:A3,A1:A3>0))` = 14, `COUNT` 2, `AVERAGE` 7, `MIN` 5, `MAX` 9, `PRODUCT` 45,
`SMALL(…,1)` 5, `LARGE(…,1)` 9, `MEDIAN(SEQUENCE(5))` = 3, `PERCENTILE(SEQUENCE(5),0.5)` = 3,
`INDEX(SORT(A1:A3),1)` = 0, `ROWS(FILTER(A1:A3,A1:A3>0))` = 2, `COLUMNS(SEQUENCE(2,3))` = 3,
`COUNTA(UNIQUE(Q1:Q4))` = 3, `CONCAT(SEQUENCE(3))` = `"123"`,
`TEXTJOIN(",",TRUE,FILTER(A1:A3,A1:A3>0))` = `"5,9"`, `SUMPRODUCT(SEQUENCE(3))` = 6 e
`AGGREGATE(15,6,FILTER(A1:A3,A1:A3>0),1)` = 5.

Dois desses consumidores são recém-chegados, e agora respondem para **todo** array computado, e não apenas
para um produtor: `ROWS`/`COLUMNS` informam a extensão do resultado de um operador ou de uma chamada elevada
também — `ROWS(A1:C3*2)`, `ROWS(A1:C3*H1:H2)` (a extensão do *broadcast*), `ROWS(LEN(A1:A3))`, `ROWS(-A1:A3)`
e `ROWS(ROW(A1:C3))` são todos 3, onde antes eram `#VALUE!` ou `1` — e `COUNTA`/`CONCAT`/`TEXTJOIN`
transmitem os elementos de um (`COUNTA(IF(A1:A3>0,A1:A3))` = 3, `CONCAT(A1:A3*2)` = `"10018"`). O **erro**
1x1 próprio de um produtor é informado como ele mesmo, e não como uma forma de 1
(`ROWS(FILTER(A1:A3,A1:A3>100))` = `#CALC!`, `ROWS(SEQUENCE(-1))` = `#VALUE!`), enquanto um *elemento* de erro
dentro de um retângulo não esconde a forma (`ROWS(SORT(E1:E3))` = 3).

**Dois membros da família de achatamento são exceções, por dois motivos diferentes**, e importa qual é qual:

- **O `CONCATENATE` toma o elemento superior esquerdo, ele não expande.** Ele junta escalares, então um
  argumento produtor chega à resposta-de-célula do próprio produtor: `CONCATENATE(SEQUENCE(3))` é `"1"`,
  enquanto `CONCAT(SEQUENCE(3))` é `"123"`. Essa é a resposta do Excel também, nos dois modos de entrada
  (medido no Aspose.Cells 26.6.0, 2026-09-10).
- **O `COUNTBLANK` recusa um array computado de saída**, com `#REF!`, antes de qualquer transmissão:
  `COUNTBLANK(SEQUENCE(3))` é `#REF!` — a mesma regra e o mesmo erro que a família de critérios aplica (o item
  em **Não suportado** abaixo), e a mesma resposta que o Excel dá nos dois modos.

**Elas compõem, e a composição não precisou de código próprio** — um produtor sob um operador, uma função
elevada, um `IF`, ou sob outro produtor, é alcançado pelo mesmo construtor recursivo. Todos medidos neste
motor e iguais ao Excel inserido como array: `SUM(SORT(FILTER(A1:A3,A1:A3>0)))` = 14,
`SUM(FILTER(A1:A3,A1:A3>0)*2)` = 28, `SUM(LEN(FILTER(A1:A3,A1:A3>0)))` = 2,
`SUM(FILTER(SEQUENCE(5),SEQUENCE(5)>2))` = 12, `ROWS(UNIQUE(FILTER(A1:A3,A1:A3>0)))` = 2 e
`SUM(IF(TRUE,SEQUENCE(3),0))` = 6. O broadcast se aplica como em qualquer outro lugar, então um produtor cuja
extensão é menor do que aquilo com que é combinado deixa `#N/A` nas posições que não cobre:
`SUM(FILTER(A1:A3,A1:A3>0)*B1:B3)` é `#N/A` — uma seleção 2x1 contra um intervalo 3x1 — enquanto o `COUNT` da
mesma expressão é 2.

**Um resultado vazio é `#CALC!`, nunca um array vazio.** `SUM(FILTER(A1:A3,A1:A3>100))` e
`ROWS(FILTER(A1:A3,A1:A3>100))` são [`#CALC!`](computed-value.md), o erro de array vazio do Excel, com
`ERROR.TYPE` **14**; informe o terceiro argumento do `FILTER` para responder outra coisa
(`SUM(FILTER(A1:A3,A1:A3>100,0))` = 0). Um resultado 1x1 — inclusive um vazio — faz broadcast como um escalar,
então `SUM(FILTER(A1:A3,A1:A3>100,7)*A1:A3)` é 98, o 7 contra cada célula. Como sempre, `COUNT` e `COUNTA`
*descartam* erros em vez de propagá-los, então `COUNT(FILTER(A1:A3,A1:A3>100))` é `0` — que é a resposta do
Excel ali também.

**Brancos sobrevivem à seleção como brancos; nada é normalizado para zero.** Uma célula de origem em branco
continua em branco através dos três seletores, então `COUNTA(FILTER(A5:A8,A5:A8<>"zzz"))` = `COUNTA(A5:A8)` = 3
sobre 7, branco, `"t"`, 7, o `ISBLANK` de um branco mantido é `TRUE`, o `UNIQUE` trata um branco como chave
**própria** (igual a nem `0`, nem `""`, nem `FALSE`, então `ROWS(UNIQUE(A5:A8))` = 3) e o `SORT` põe os brancos
**por último** nas duas direções. Cada um desses casos bate com o Excel nos dois modos de entrada, medido em
2026-09-10 — um argumento de INTERVALO e um resultado de PRODUTOR concordam, e não há costura entre eles.

**Sozinho em uma célula: a regra do `@` — e o MySheet não derrama (*spill*).**

Uma célula guarda um único valor escalar, então um produtor escrito como a fórmula inteira de uma célula mostra
o **elemento superior esquerdo** do array e nada é escrito nas células ao redor. Essa é a regra `@`-sobre-array
do Excel (Microsoft, "Implicit intersection operator: @": para um array o Excel "picks the top-left value"), e é
exatamente o que o próprio Excel escreve quando converte uma fórmula legada em `=@FILTER(...)`. Medido no
Aspose.Cells 26.6.0, 2026-09-10, uma fórmula por pasta de trabalho, em uma célula cuja linha está *dentro* do
intervalo de origem (`D2`) e em outra *fora* dele (`D5`), com entrada digitada e inserida como array
concordando em todas as linhas — e o MySheet responde o mesmo nas duas células:

| Em uma célula | Excel e MySheet |
| --- | --- |
| `=SEQUENCE(5)` | `1` |
| `=SEQUENCE(2,3,7,1)` | `7` |
| `=FILTER(A1:A3,A1:A3>0)` | `5` |
| `=SORT(A1:A3)` | `0` (o superior esquerdo já ordenado) |
| `=UNIQUE(A1:A3)` | `5` |
| `=SEQUENCE(2,3)*10` | `10` |
| `=FILTER(A1:A3,A1:A3>100)` | `#CALC!` |

Ao contrário de um intervalo sozinho, a resposta **não** depende de onde a fórmula está: um produtor não tem
posição na planilha com que se interseccionar, e é isso que distingue a metade-array da regra da
[metade-intervalo](#interseção-implícita-na-fronteira-da-célula).

**Portanto `=SEQUENCE(5)` mostra `1` em vez de preencher cinco células com 1..5, e `=FILTER(A:A,B:B>0)` mostra
um único valor em vez de uma lista.** Essa é a consequência mais surpreendente deste recurso e vale dizê-la
claramente em vez de deixá-la ser descoberta: o MySheet **não tem modelo de spill** — não há como uma fórmula
escrever em células que não são dela — então o superior esquerdo é a resposta inteira. Se você quer a lista,
consuma-a: `ROWS(FILTER(...))` para a contagem, `INDEX(FILTER(...),k)` para a *k*-ésima correspondência,
`TEXTJOIN(",",TRUE,FILTER(...))` para todas elas em uma célula.

**Os outros desvios, na íntegra.** Cada um é fixado por um teste — com o número do próprio Excel no teste onde
os dois motores divergem — então fechar um é sempre uma edição deliberada.

- **Sem spill**, como acima. É o único desvio estrutural, e não uma escolha: uma célula é um escalar. Note que
  o *valor exibido* concorda com o Excel nos dois modos de entrada — o que difere é que o Excel moderno também
  preencheria as células vizinhas.
- **Um argumento de intervalo aberto ou de coluna inteira é recusado.** A guarda de custo que recusa um
  intervalo aberto em posição de array vale também para os argumentos de um produtor, então
  `SUM(FILTER(A:A,A:A>0))` é `#VALUE!` aqui, onde o Excel responde 14 (medido em 2026-09-10, nos dois modos de
  entrada, sobre uma coluna A contendo apenas `A1:A3` = 5, 0, 9). A recusa tem uma **metade silenciosa** que
  vale conhecer: o `COUNT` descarta o canal de erro, então `COUNT(FILTER(A:A,A:A>0))` é `0` aqui, contra o 2 do
  Excel nessa mesma configuração — um número errado em vez de um erro. Limite o intervalo (`A1:A100000`) e
  funciona. Fazer a forma aberta funcionar exige uma regra de limites compartilhados, porque `array` e
  `include` seriam limitados de forma independente e poderiam discordar na contagem de linhas.
- **O `SEQUENCE` tem um limite de tamanho que o Excel não tem.** `rows > 1048576`, `columns > 16384` ou
  `rows * columns > 1048576` responde `#NUM!`; exatamente no limite é permitido. O Excel não tem limite algum
  em posição consumida — `ROWS(SEQUENCE(1048577))` é 1048577 e `COLUMNS(SEQUENCE(1,16385))` é 16385 lá (medido
  em 2026-09-10, nos dois modos) — mas o fluxo de elementos é preguiçoso enquanto todo consumidor percorre
  todos os elementos, então sem o limite `SUM(SEQUENCE(1000000,10000))` travaria em vez de responder.
- **`UNIQUE(…, exactly_once)` segue a página da Microsoft, e não o oráculo medido.** Sobre `Q1:Q4` = 9, 5, 9, 0
  as linhas que ocorrem exatamente uma vez são 5 e 0, e é isso que o MySheet devolve (`ROWS` 2, `SUM` 5). O
  oráculo preserva a forma da contagem de *distintos* e a completa repetindo o último valor mantido — `ROWS`
  **3** com `SUM` **5**, isto é, as linhas 5, 0, 0 — de modo que o resultado de `UNIQUE` dele contém uma
  duplicata, o que a própria contagem de linhas dele contradiz. Onde o oráculo se contradiz, a regra
  documentada vence; a medição fica registrada ao lado do teste para que a decisão possa ser revista.
- **O `AVERAGE` sobre o `UNIQUE` também segue a página, pelo mesmo motivo.** Sobre esse mesmo `Q1:Q4` o
  MySheet responde 14/3, que é o `SUM` dele sobre o `COUNT` dele. O oráculo informa `SUM` **14**, `COUNT` **3**
  e `AVERAGE` **0** para a mesma expressão — três respostas que não podem estar todas certas — e a página do
  `AVERAGE` da Microsoft é explícita: a média é a soma sobre a contagem, com os zeros incluídos.
- **Um produtor vinculado por `LET`, passado por `CHOOSE` ou por um `+` unário colapsa para o superior
  esquerdo dele.** `SUM(LET(x,FILTER(A1:A3,A1:A3>0),x))` é `5` aqui, onde o Excel responde **14** nos dois
  modos de entrada (medido em 2026-09-10; lá `ROWS(LET(x,FILTER(…),x))` = 2,
  `SUM(CHOOSE(1,FILTER(…)))` e `SUM(+FILTER(…))` = 14). Um vínculo é capturado como *valor* antes que a
  avaliação elemento-a-elemento possa vê-lo, então o que é vinculado é a resposta-de-célula do produtor. Use o
  produtor diretamente no slot de argumento do consumidor. A instância mais alta é um produtor vinculado por
  `LET` em um slot de critérios — `LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,">0"))` é `1` aqui, o `COUNTIF`
  sobre o único elemento colapsado, contra o `#REF!` do Excel nos dois modos de entrada, com `SUM(f)` 5 contra
  14 e `ROWS(f)` 1 contra 2 na mesma forma. Essa linha é o único pin deliberadamente vermelho deste motor,
  vermelho para que a correção o torne verde em vez de ser descoberta por acidente.

**Qual fábrica uma nova função nativa usa (para quem contribui).** A classificação é um sinalizador
explícito por entrada em [`FunctionRegistry`](../../Danfma.MySheet/Parsing/FunctionRegistry.cs):
`Entry<T>(…)` registra uma função que consome intervalos/arrays por conta própria e nunca é elevada, e
`Elementwise<T>(…)` uma puramente escalar que o mini-CSE pode elevar. **O padrão é `Entry<T>` — negar** —
porque os dois erros não são simétricos: escrever `Entry<T>` onde cabia `Elementwise<T>` apenas perde a
otimização, enquanto escrever `Elementwise<T>` onde cabia `Entry<T>` faz a função responder a partir de um
único elemento do retângulo que deveria consumir inteiro, ficando **silenciosamente errada**, sem erro
nenhum para alguém notar. E como `Entry<T>` é também o que uma entrada recebe quando ninguém escolhe,
*esquecer* o sinalizador cai no lado seguro por construção — o erro perigoso é o deliberado.

Os testes de guarda são precisos sobre qual desses dois erros cada um pega:

- Um sinalizador **esquecido** é seguro por *construção*, não por um teste. `Consumes` é o valor zero do
  enum, então uma nova função nativa registrada pela fábrica padrão `Entry<T>` nunca é elevada, faça ela o
  que fizer com os argumentos — o erro custa apenas a otimização. Um teste ainda exige que essa entrada seja
  *vigiada*: toda entrada ciente de intervalos precisa ser visível para a sonda abaixo ou ser nomeada à mão,
  e uma que não seja nenhuma das duas quebra a suíte carregando o próprio nome.
- Um sinalizador **errado** — `Elementwise<T>` em uma função ciente de intervalos, o erro que publica um
  número silenciosamente errado — é pego pelo nome. O conjunto exato dos 180 nomes `Elementwise` está
  registrado como uma lista ordenada, então acrescentar um nome quebra a suíte nomeando o recém-chegado e
  remover um quebra nomeando a perda. Uma contagem não serviria: ela sobrevive a uma troca compensada e
  sobrevive à edição de aparência honesta de virar a fábrica e ajustar o número. Medido: essa edição deixava
  a suíte inteira verde.
- O mesmo sinalizador errado é *também* pego com um diagnóstico onde a sonda consegue vê-lo, e essa sonda
  reexecuta a derivação a cada build. Ela varre cada posição de argumento de cada aridade de `MinArgs` até
  `MinArgs+3`, preenchendo os slots restantes com um número, um texto, um lógico e um intervalo de três
  células por vez, e entrega à entrada três retângulos que diferem em posição, formato e conteúdo. Um corpo
  puramente escalar responde de forma idêntica para os três; um ciente de intervalos não, e a falha nomeia a
  chamada que os distinguiu. A varredura ainda é cega para **22** das 130 entradas cientes de intervalos —
  as que respondem a mesma coisa para todo retângulo: os testes de formato e de referência (`AREAS`,
  `ISREF`, `ISFORMULA`, `FORMULATEXT`, `SHEET`, `TYPE`), `OFFSET`/`INDIRECT`, as exclusões de projeto (`IF`,
  `LET`, `RANDBETWEEN`), as reduções que erram de forma idêntica nos três (`AND`, `OR`, `IRR`, `MIRR`,
  `XNPV`, `PROB`, `FORECAST`, `FORECAST.LINEAR`, `PERCENTILE.EXC`, `TRIMMEAN`) e o `SEQUENCE`, que não recebe
  intervalo algum — os argumentos dele são um tamanho, um início e um passo, então não há retângulo para lhe
  entregar e a varredura é cega para ele em definitivo. Essas 22 têm a lista
  registrada e a lista à mão como única defesa, então o próprio conjunto cego é fixado pelo nome e ganhar um
  membro também quebra a suíte.

**Não suportado (por design).**

- Uma **célula seca** cuja fórmula inteira é um array é `#VALUE!` **a menos que o array venha de um dos quatro
  [produtores](#produtores-de-array-dinâmico)**, que respondem o elemento superior esquerdo deles. A regra da
  fronteira é, portanto, por tipo de nó, e há três casos. (1) Um **produtor** — `=FILTER(...)`, `=SORT(...)`,
  `=UNIQUE(...)`, `=SEQUENCE(...)` e qualquer expressão construída sobre um deles — devolve o valor superior
  esquerdo do array, a regra `@`-sobre-array do Excel, a mesma resposta em toda célula e nos dois modos de
  entrada do Excel: veja a tabela naquela seção. (2) Um **operando de intervalo sob um operador ou uma função
  elevada** mantém `#VALUE!` aqui — `=A1:A3*2`, `=LEN(A1:A3)`, `=ROUND(A1:A3,0)`, `=-A1:A3` — porque a
  elevação acontece dentro dos *consumidores* e a fronteira da célula não é um deles: a célula vê o corpo
  escalar comum do `LEN` recebendo um intervalo. O Excel faz ali algo diferente de qualquer um dos dois
  motores, e isso é uma lacuna genuína, não uma regra: digitada, o Excel aplica interseção implícita a cada
  operando de intervalo *antes* do operador, usando a linha da própria célula da fórmula, então com
  `A1:A3` = 5, 0, 9 um `=-A1:A3` sozinho é `-5` em `C1`, `0` em `C2`, `-9` em `C3` e `#VALUE!` em `C5`, e
  `=ROUND(A1:A3,0)` é 5, 0, 9 e `#VALUE!` nessas mesmas células; inserida como array, ele toma o superior
  esquerdo em todas elas (`-5`, `5`). As duas colunas medidas no Aspose.Cells 26.6.0, 2026-09-10. O `#VALUE!`
  de hoje está fixado por teste para que fechar essa lacuna seja deliberado — e note que o comentário do
  próprio teste que a fixa ainda afirma que a forma digitada é `#VALUE!` em qualquer lugar, o que a medição
  acima contradiz para uma fórmula em linha *dentro* do intervalo. (3) Um `IF(range…)` ou uma comparação de
  intervalo sozinhos são `#VALUE!` pelo mesmo motivo do caso (2) — `=IF(B2:B5="Show",1,0)` e
  `=IF(TRUE,A1:A3,B1)` sozinhas são erros — uma inconsistência conhecida com o caso (1) ao lado. Em todos os
  casos, envolver a expressão em um consumidor funciona: `=SUM(LEN(A1:A3))` nessa mesma célula é `3` para
  `A1:A3` = 5, 0 e 9 (um caractere cada). Arrays continuam existindo apenas como *argumentos* e como o
  superior esquerdo colapsado de um produtor, nunca como o valor multicélula de uma célula: o cache por célula
  permanece estritamente escalar e não há spill. Isso **não** contradiz a
  [interseção implícita na fronteira da célula](#interseção-implícita-na-fronteira-da-célula): aquela regra
  intersecta uma *referência* — o `=A1:A3` puro ao lado destas é `A3` em `C3` — enquanto um array computado não
  tem posição na planilha, e é por isso que a resposta do produtor não depende da posição.
- **Um produto propagado em uma célula pura também é `#VALUE!`**, e essa lacuna vive na fronteira, não na
  regra de propagação: `=A1:C3*E1:E3` digitada em uma célula nunca entra na avaliação elemento a elemento,
  então é o `#VALUE!` do próprio operador de multiplicação, enquanto `=SUM(A1:C3*E1:E3)` nessa mesma célula é
  108. A regra do Excel para a forma digitada não é uma regra de array: ele aplica a interseção implícita a
  **cada operando de intervalo separadamente, antes do operador**, usando a linha e a coluna da própria
  célula da fórmula — por isso a mesma fórmula responde coisas diferentes em células diferentes:
  `=A1:A3*E1:E3` é **21** em `J3` (`A3`×`E3`) e **1** em `J1`, `=E1:E3*10` é **20** em `L2`, **30** em `L3` e
  `#VALUE!` em `L5` (a linha 5 não alcança `E1:E3`), e um operando 2-D nunca intersecta, então `=A1:C3*E1:E3`
  é `#VALUE!` ali também — o único jeito de obter o elemento do canto superior esquerdo do array computado é
  pedir por ele, `=INDEX(E1:E3*10,1,1)` = **10** (tudo com entrada digitada, medido no Aspose.Cells 26.6.0,
  em 2026-09-10). Fechar isso é assunto da regra da
  [fronteira da célula](#interseção-implícita-na-fronteira-da-célula) — o comportamento por operando acima é
  com o que a metade de array do `@` precisa ser reconciliada —, e não dos consumidores descritos aqui; o
  `#VALUE!` de hoje está fixado por `CellBoundaryIntersectionTests` para que a mudança seja deliberada.
- **O `+` unário deliberadamente não é elevado.** Ele é o no-op do Excel que preserva referências, então
  `+A1:A3` continua sendo uma *referência* e o consumidor a dobra pelo caminho comum de intervalo:
  `SUM(+A1:A3)` = 6 para `A1:A3` = 1, 2 e 3, exatamente como `SUM(A1:A3)`, e sem mudança alguma com a
  elevação. O custo de mantê-lo opaco é que um `-` sobre ele não tem o que elevar: `SUM(-(+A1:A3))` é
  `#VALUE!` onde o Excel responde -6 (Aspose.Cells 26.6.0, `Ctrl+Shift+Enter`, medido em 2026-09-09).
  Escreva `SUM(-A1:A3)` em vez disso.
- Uma **função que retorna referência** como argumento de `ROW`/`COLUMN` permanece escalar:
  `SUM(ROW(INDEX(A1:A3,1,1)))` é `1`, a linha superior da referência resolvida, e não o vetor `[1,2,3]`.
  Descobrir o formato dela resolveria o argumento uma segunda vez e sortearia uma volátil duas vezes, então
  ali o formato de array é deliberadamente adiado.
- A família de **critérios / varredura posicional** não lê um array computado *onde consegue reconhecer um* —
  ela o **rejeita** com `#REF!`. `SUMIF`/`SUMIFS`, `COUNTIF`/`COUNTIFS`, `AVERAGEIF`/`AVERAGEIFS` e
  `MAXIFS`/`MINIFS` percorrem seus argumentos posição a posição, e todo slot de intervalo que elas recebem —
  intervalo de critérios e intervalo de soma/média/máximo/mínimo igualmente — exige uma *referência*: um
  argumento que não é um nó de referência e que a avaliação elemento a elemento transmitiria é recusado
  antes de a varredura abrir, em todos os slots e em todas as aridades. `COUNTIF(A1:A3*1,">0")`,
  `SUMIF(A1:A3*1,">0")`, `COUNTIFS(A1:A3*1,">0")`, `AVERAGEIF(A1:A3*1,">0")`, `SUMIFS(B1:B3,A1:A3*1,">0")`,
  `SUMIFS(A1:A3*1,B1:B3,">0")`, `SUMIF(A1:A3,">0",B1:B3*1)`, `COUNTIFS(A1:A3,">0",B1:B3*1,">1")`,
  `COUNTIF(ROW(A1:A3),">1")`, `COUNTIF(-A1:A3,"<0")` e `COUNTIF(LEN(A1:A3),">0")` são todos `#REF!`. O Excel
  recusa a família do mesmo jeito: com um argumento **computado** (`SUMIFS((A1:A3)*1,A1:A3,">0")` e as sete
  irmãs) ele responde `#VALUE!` na digitação normal e `#REF!` quando a fórmula é inserida como array, e com
  um **produtor** de array dinâmico no slot ele responde `#REF!` nos *dois* modos de entrada —
  `COUNTIF(FILTER(A1:A3,A1:A3>0),">5")`, `COUNTIF(SEQUENCE(5),">3")`, `SUMIF(SORT(A1:A3),">0")` e
  `COUNTIF(UNIQUE(A1:A3),">0")`, tudo medido no Aspose.Cells 26.6.0, em 2026-09-10, e respondido da mesma
  forma aqui agora que [as quatro existem](#produtores-de-array-dinâmico). O `#REF!` é, portanto, ao
  mesmo tempo a resposta do modo array que esta seção reproduz e a única resposta em que as duas colunas do
  produtor concordam, e é por isso que a regra é `#REF!` e não `#VALUE!`. Fixado por
  `CriteriaComputedArgumentTests` e `MathAggregateTests.CriteriaFamily_RejectsAComputedArrayWithRef`. Um
  argumento **propagado** ou composto recebe a mesma rejeição — a família nunca entra na avaliação elemento a
  elemento —, então `SUMIF(A1:C3*H1:H2,">0")`, `COUNTIF(A1:C3*H1:H2,">0")` e
  `SUMIFS(A1:C3,A1:C3*H1:H2,">0")` são `#REF!` aqui, coincidindo com a coluna do oráculo inserida como array
  (`#VALUE!` digitado; medido em 2026-09-10, fixado por
  `MiniCseConsumerTests.CriteriaFamily_OverABroadcastArray_IsRef`), e um argumento **elevado** também:
  `SUMIFS(LEN(A1:A3),A1:A3,">0")` é `#REF!`, com a mesma divisão `#VALUE!` na digitação normal / `#REF!` como
  array no oráculo (fixado por `MiniCseConsumerTests.CriteriaFamily_OverALiftedFunction_IsRef`). O que **não**
  é rejeitado é tudo o que já é uma referência ou não é elegível a array: uma função que retorna referência
  (`CHOOSE`, `OFFSET`, `INDEX`), um nome definido, uma célula única e uma coluna inteira continuam sendo
  intervalos, então `COUNTIF(CHOOSE(1,A1:A3,B1:B3),">0")` e `COUNTIF(OFFSET(A1,0,0,3,1),">0")` dão `2`, como
  no oráculo nos dois modos. Quatro formas são **desvios deliberados**, cada uma fixada como tal em
  `CriteriaComputedArgumentTests` — três deixadas para a varredura de compatibilidade e a quarta o limite
  permanente do `LET`:
  `COUNTIF(IF(TRUE,A1:A3,B1:B3),">0")` dá `0` aqui, onde o oráculo responde `2` nos *dois* modos de entrada —
  um `IF` de condição escalar aqui é um escalar opaco em vez da referência do seu ramo, e fechar isso é item
  da própria varredura, deliberadamente fora desta regra; `COUNTIF(5,">0")` e `COUNTIF(A1*1,">0")` dão `1`
  onde o oráculo responde `#REF!` nos dois modos (um *escalar* puro em slot de intervalo, forma que nenhum
  produtor de array assume); e `SUMIF(A:A*1,">0")` dá `0` onde o oráculo responde `#REF!` nos dois modos (a
  guarda de custo recusa um operando de coluna inteira, então o argumento nunca é elegível a array e a
  comporta nunca o vê); e um `LET` dá `0` dos **dois** lados da vinculação —
  `COUNTIF(LET(r,A1:A3,r*1),">0")` e `LET(r,A1:A3*1,COUNTIF(r,">0"))`, com os gêmeos de `SUMIF` também —
  onde o oráculo responde `#REF!` inserido como array (`#VALUE!` digitado), porque um nó `LET` é um escalar
  opaco para a sondagem de formato, enquanto um nome vinculado por `LET` *é* um nó de referência cuja
  vinculação já foi reduzida a escalar na captura, então a comporta não vê array de jeito nenhum. Essa última
  é **pré-existente** (medida idêntica antes de a regra chegar) e é um limite permanente, não parte desta
  regra: `LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,">0"))` é exatamente essa forma e é o único pin
  deliberadamente vermelho da suíte — veja o item sobre `LET` em
  [produtores de array dinâmico](#produtores-de-array-dinâmico). O `SUMPRODUCT` é o único membro dessa família que optou
  por aceitar arrays computados — `SUMPRODUCT((A1:A3<>0)*1)` = 2 e `SUMPRODUCT(A1:A3*1,B1:B3)` = 32,
  coincidindo com o oráculo nos dois modos — e os consumidores de dobra listados em **Suportado**
  acima (`SUM(IF(…))` e companhia) sempre os aceitaram. O `SUBTOTAL` e a forma-referência do
  `AGGREGATE` não seguem nem um caminho nem o outro — eles rejeitam um array computado de saída,
  inclusive um elevado (`SUBTOTAL(9,LEN(A1:A3))` e `AGGREGATE(9,4,LEN(A1:A3))` são `#VALUE!` nos
  dois motores); quem o consome é a forma-array do `AGGREGATE`, elevações incluídas
  (`AGGREGATE(15,6,LEN(A1:A3),1)` = 1, medido nos dois).
- Um intervalo **aberto/de coluna inteira** em posição de array é recusado e o consumidor permanece em seu
  caminho escalar/de intervalo comum — a única exceção é a identidade `INDEX(ROW($A:$A), n)` acima, que
  retorna `n` sem materializar a coluna. `SMALL(IF(A:A=…, ROW(A:A)), k)` sobre uma coluna *aberta* portanto
  não é avaliado como array. Uma chamada elevada sobre uma coluna aberta é recusada da mesma forma, e a
  recusa é *tolerada* em vez de fatal: a chamada colapsa para um único escalar opaco avaliado uma vez, então
  `SUM(LEN(A:A))` é `#VALUE!` (o `LEN` escalar de um intervalo) enquanto uma expressão de array que a
  envolva continua funcionando — `SUM(IF(A1:A3>0,1,LEN(B:B)))` continua sendo 3. O Excel, em vez disso,
  dobra a coluna aberta (`SUM(LEN(A:A))` = 3 sobre três células de um caractere, Aspose.Cells 26.6.0
  inserido como array, 2026-09-09); uma fórmula que funciona sobre `A1:A3` e depois é arrastada para uma
  coluna inteira recupera o antigo `#VALUE!`, sem nenhum outro aviso.
- Uma condição **escalar** mantém o curto-circuito nativo do `IF`: só o ramo tomado é avaliado, e apenas uma
  *condição* de array conduz o zip. O ramo tomado, porém, continua sendo lido como array quando *é* um — um
  produtor, uma chamada elevada ou o resultado de um operador — então `SUM(IF(TRUE,SEQUENCE(3),0))` é 6 e
  `ROWS(IF(TRUE,SEQUENCE(3),0))` é 3, coincidindo com o Excel nos *dois* modos de entrada, e
  `SUM(IF(TRUE,A1:C3*2,0))` é 90 sobre `A1:C3` = 1…9, coincidindo com a coluna dele inserida como array
  (`#VALUE!` digitado). A exceção é um ramo que é um **intervalo puro**: `SUM(IF(TRUE,A1:C3,0))` é `#VALUE!`
  aqui, onde o Excel responde 45 nos dois modos, deliberadamente intocado porque mexer nisso responderia
  "o `IF` devolve uma referência?" para essa única forma enquanto as irmãs dela ficam sem resposta. Está
  registrado para a varredura de compatibilidade e fixado como lacuna. Tudo medido no Aspose.Cells 26.6.0,
  2026-09-10.

**Divergências conhecidas.** Cada uma delas está fixada por teste como uma *lacuna*, e não afirmada como a
regra do Excel, de modo que fechar qualquer uma é sempre uma edição deliberada; a única entrada sem teste que
a fixe diz isso com as próprias palavras. Excel aqui significa
Aspose.Cells 26.6.0, a versão contra a qual este projeto mede, com a fórmula inserida como array
(`Ctrl+Shift+Enter`) — a forma de entrada cuja semântica esta avaliação elemento a elemento reproduz sem a
combinação de teclas — e todo número tirado da forma digitada vem rotulado como *digitada* onde aparece.

- **Não coberto nos DOIS eixos ao mesmo tempo é `#N/A` aqui, e o oráculo não tem resposta a igualar.** A
  regra de propagação acima é por eixo, então um 2x2 contra um 3x3 deixa cinco das nove posições não cobertas
  no eixo das linhas, no das colunas ou em ambos, e cada uma é `#N/A`, enquanto as quatro cobertas calculam:
  `SUM(A1:B2*A1:C3)` é `#N/A` e `COUNT(A1:B2*A1:C3)` = 4. O Excel também responde `#N/A` nas faltas de
  cobertura em **um** eixo medidas aqui — `INDEX(A1:B3*A1:C2,3,1)`, um 3x2 contra um 2x3, é `#N/A` nos dois
  motores (inserida como array e digitada), e o mesmo vale para as formas fixadas aqui em que o *vetor* é o
  operando mais curto (`SUM(A1:C3*H1:H2)` e `SUM(A1:C3*E5:F5)` são `#N/A` com `COUNT` 6 nos dois, inseridas
  como array) —
  mas essa concordância **não** é universal: um retângulo 2-D mais curto que um vetor LINHA é um
  contraexemplo medido, a entrada logo abaixo. Já para a forma duplamente não coberta, ele responde
  **`#REF!`** pelo `INDEX` — tanto digitada
  quanto inserida como array — e a calculadora dele nunca retorna para um `SUM` ou um `COUNT` sobre esse mesmo
  array (sem resposta depois de dez minutos aqui, e mais de 200 s em cada modo de entrada quando a fase topou
  com isso pela primeira vez, enquanto `ROWS`/`COLUMNS` sobre ele ainda informam 3 e 3 imediatamente).
  Tudo medido no Aspose.Cells 26.6.0, em 2026-09-10. É por isso que esta divergência *não* está na varredura
  de compatibilidade: um `#REF!` em dois eixos contra um `#N/A` em um só não é uma regra a copiar, e um cálculo
  que não termina não é um comportamento a reproduzir. Fixada por
  `VectorBroadcastingTests.UncoveredOnBothAxes_StaysNotAvailable_WhereTheOracleIsSelfInconsistent`, cujo
  comentário carrega a medição.
- **Um RETÂNGULO 2-D mais curto que um vetor LINHA é `#N/A` aqui, e o oráculo preenche a coluna não coberta
  com `0`.** Esta é a única falta de cobertura em *um* eixo encontrada até agora em que os dois motores
  discordam, e é a direção que a Fase 10 nunca fixturou: as outras formas de vetor desencontradas da fase
  têm todas o VETOR como operando mais curto, então esta classe ficou sem teste. Sobre a fixture de
  propagação (`A1:C3` = 1..9 em ordem de linha, `E5:G5` = 10,20,30), um retângulo 3x2 contra esse vetor
  linha 1x3 assume uma extensão 3x3 cuja terceira coluna o retângulo não cobre, e ali `SUM(A1:B3*E5:G5)` é
  `#N/A` aqui contra **420** no oráculo, `COUNT(A1:B3*E5:G5)` = 6 contra **9**, e `INDEX(A1:B3*E5:G5,1,3)` é
  `#N/A` contra **0** (Aspose.Cells 26.6.0, inseridas como array, medidas em 2026-09-10). A *extensão*
  concorda — `INDEX(…,4,1)` e `INDEX(…,1,4)` são `#REF!` nos dois — e a aritmética das posições cobertas
  também, já que `SUM(IFERROR(A1:B3*E5:G5,0))` é 420 nos dois motores nesse modo; toda a diferença está no
  que preenche a coluna não coberta. O MySheet mantém o `#N/A` dele porque o oráculo não é consistente
  consigo mesmo ali, e as três verificações são todas inseridas como array: reduza o retângulo para duas
  linhas e os agregados voltam a `#N/A`, com `COUNT(A1:B2*E5:G5)` = 4, mas `INDEX(A1:B2*E5:G5,1,3)` continua
  **0** — uma posição que, portanto, o `COUNT` não conta; e o espelho com vetor COLUNA nunca preenche, pois
  um retângulo 2x3 contra o 3x1 `E1:E3` é `#N/A` com `COUNT` 6 e `INDEX(A1:C2*E1:E3,3,1)` `#N/A` nos
  **dois** motores, assim como um retângulo 1x2 contra `E5:G5` (`#N/A`, `COUNT` 2, nos dois). Registrado
  para a varredura de compatibilidade com o Excel já planejada. Fixada por
  `VectorBroadcastingTests.RectangleShorterThanARowVector_StaysNotAvailable_WhereTheOracleFillsWithZero`,
  cujo comentário carrega todos os números acima.
- **`SUM(ROW(Ghost!A1:A3))`** — um retângulo escrito *literalmente* sobre uma planilha que não existe, em
  posição de array — responde `6`, os números de linha `1+2+3`, enquanto o Excel responde `#REF!`. O
  `ROW(Ghost!A1:A3)` escalar na mesma pasta de trabalho já é `#REF!`, assim como o caminho de array sobre um
  nome que representa o mesmo intervalo (`SUM(ROW(GhostName))`): a divergência está apenas no retângulo
  escrito por extenso, cujo caminho rápido sintático vai direto a um vetor de linhas ou colunas e nunca
  resolve a referência, de modo que a guarda de planilha ausente — executada por todo caminho que resolve —
  não tem sobre o que atuar. Uma função *elevada* sobre o mesmo retângulo fantasma lê células e portanto
  responde `#REF!` (`SUM(LEN(Ghost!A1:A3))`), e é por isso que as duas formas vizinhas discordam. Fixado por
  `MiniCseConsumerTests.Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap`.
- **`INDEX(<array computado>, 0)`** é `#REF!` aqui, enquanto o Excel intersecta o vetor inteiro e responde o
  primeiro elemento dele — `INDEX(LEN(A1:A3),0)` e `INDEX(ROW(A1:A3),0)` são ambos **1** lá (medido em
  2026-09-09, tanto na digitação normal quanto como array). O MySheet rejeita de saída um `row_num` ou um
  `column_num` menor que 1.
- **Um slot de argumento vazio mantém o padrão documentado da função**, enquanto o Excel o lê como um `0`
  fornecido: `FIXED(A1,,TRUE)` é `1.00` aqui e **`1`** lá, `DOLLAR(A1,)` é `$1.00` aqui e **`$1`** lá, para
  `A1` = 1 (medido em 2026-09-09). Com o slot totalmente ausente os dois motores concordam — `FIXED(A1)` e
  `DOLLAR(A1)` são `1.00` e `$1.00` em cada um — então a divergência é o slot *vazio*, e não o padrão, e ela
  vale igualmente para a chamada escalar e para o `FIXED(A1:A3,,TRUE)` elevado.
- **Uma chamada elevada sob um `+` unário também não é elevada.** O `+` é o no-op do Excel que preserva
  referências, e o MySheet mantém toda a expressão com `+` opaca, o que esconde do mini-CSE o que está
  *dentro* dela: com `A1:A3` = 1, 22 e 333, `SUM(+LEN(A1:A3))` é `#VALUE!` aqui e **6** no Excel (medido em
  2026-09-09). É a irmã do caso `SUM(-(+A1:A3))` = -6 acima — o mesmo `+` opaco, com uma *função* elevada
  dentro em vez de um operador unário — e `SUM(LEN(+A1:A3))`, com o `+` do lado de dentro, é o mesmo
  `#VALUE!` aqui contra os mesmos **6** lá. Escreva `SUM(LEN(A1:A3))`. Fixado por
  `ElementwiseLiftingTests.LiftedCall_UnderAnOpaqueUnaryPlus_IsNotLifted_KnownDivergence`.
- **Um NOME definido em posição de array é aquilo a que ele está vinculado** — a regra em si é *concordância*,
  e o que esta entrada registra são as três formas ainda recusadas. Um nome vinculado a um retângulo é
  elegível a array em uma posição de array **aninhada** e responde exatamente o que o retângulo escrito por
  extenso responde. Para um `MyName` vinculado a 1, 22 e 333 e um `Rng` vinculado a `A1:A3` = 5, 0 e 9, tudo
  medido no Aspose.Cells 26.6.0 inserido como array (2026-09-10): `SUM(LEN(MyName))` = **6**, `SUM(-MyName)` =
  **-356**, `SUM(MyName%)` = **3.56**, `SUM(MyName*2)` = **712**, `SUM((MyName>1)*1)` =
  `SUM(IF(MyName>1,1,0))` = `SUMPRODUCT(--(MyName>1))` = **2**, `COUNT((Rng<>"")*1)` = `COUNT(Rng*1)` = **3**,
  `SUM((Rng<>0)*1)` = **2**, `SMALL(IF(Rng>0,Rng),1)` = **5** e `INDEX(Rng*2,3)` = **18** — os mesmos valores
  dos gêmeos literais, que é a regra enunciada como teste. Um nome em uma **planilha inexistente** transmite
  o `#REF!` por elemento do literal do mesmo modo: `SUM((GhostName<>0)*1)` é `#REF!` e
  `COUNT((GhostName<>"")*1)` é `0`, nos dois modos do oráculo. Ler o nome em si não é afetado
  (`SUM(MyName)` = 356 e `SUM(ROW(MyName))` = 6 nos dois motores) e, no **nível superior** de um consumidor,
  um nome puro continua sendo uma *referência* que mantém o caminho de referência, exatamente como um
  intervalo literal puro — `SUBTOTAL(9,Rng)` = 14, `AGGREGATE(9,4,Rng)` = 14, `SUM(A1:INDEX(Rng,3))` = 14 e
  `ISREF(INDEX(Rng,2))` = `TRUE` —, porque esse caminho carrega o que um fluxo elemento a elemento não
  carrega: o salto do `SUBTOTAL` aninhado, a varredura do primeiro erro em ordem de coluna do motor e um
  `INDEX` que retorna referência. Fixado por `DefinedNameArrayEligibilityTests` e
  `ElementwiseLiftingTests.LiftedShapes_OverADefinedName_AreLifted`. Três formas continuam recusadas, cada uma
  um desvio deliberado fixado como tal: um nome de **intervalo aberto** encontra a guarda de custo, então
  `SUM((MyCol<>0)*1)` com `MyCol` = `$A:$A` dá `1` aqui — o valor de referência verdadeiro — contra os **2**
  do oráculo inserido como array (`0` digitado); um nome de **união** resolve para um escalar, o que torna a
  expressão inteira apenas escalar, então `SUM((UnN<>0)*1)` dá `1` contra **2** inserido como array
  (`#VALUE!` digitado), e o gêmeo de união *literal* é `#VALUE!` aqui, o que faz desta a única linha em que
  um nome não coincide com seu literal; e um nó `LET` no próprio slot de argumento de um consumidor continua
  opaco, porque a sondagem de formato não olha para dentro dele, então `SUM(LET(r,Rng,(r<>0)*1))` dá `1`
  contra **2** nos dois modos de entrada — o mesmo limite permanente do `LET` que a
  [seção dos produtores](#produtores-de-array-dinâmico) registra, ainda aberto. Um
  nome vinculado por `LET` *dentro* de uma posição de array, por outro lado, resolve **quando o nome está
  vinculado a um intervalo**, através do escopo do `LET` que a
  [resolução de nomes](#intervalos-nomeados) consulta primeiro: `LET(r,A1:A3,SUM((r<>0)*1))` = **2**,
  `LET(r,A1:A3,COUNT(r*1))` = **3** e `LET(r,A1:A3,INDEX(r*2,3))` = **18**, coincidindo com o oráculo nos
  dois modos de entrada, onde antes desta regra eram `1`, `0` e `#REF!`. Um nome vinculado a um **array
  computado** não resolve: a vinculação é avaliada como escalar no momento da captura, então
  `LET(r,A1:A3*1,COUNT(r*1))` dá `0`, `LET(r,A1:A3*1,SUM(r*1))` dá `#VALUE!` e
  `LET(r,A1:A3*1,INDEX(r*2,3))` dá `#REF!` contra os **3**, **14** e **18** do oráculo nos dois modos de
  entrada — inalterado por esta regra (medido no Aspose.Cells 26.6.0, em 2026-09-10, e no motor antes e
  depois da regra), e é o mesmo limite permanente do `LET` — um produtor vinculado por `LET` colapsa
  exatamente por esse motivo.
- **Uma função ciente de intervalos nunca é elevada sobre os slots ESCALARES dela.** O Excel também eleva
  uma função ciente de intervalos: ele consome o intervalo no slot que recebe um e repete a *chamada
  inteira* por elemento de um retângulo entregue a qualquer outro slot. A classificação do MySheet é por
  *função*, e não por slot, então um retângulo em um slot escalar continua um intervalo e a chamada responde
  uma única vez. Com `A1:A3` = 1, 2 e 3 e `B1:B3` = 10, 20 e 30, com a resposta do Excel primeiro e a do
  MySheet entre colchetes, todos medidos em 2026-09-09: `SUM(MATCH(A1:A3,A1:A3,0))` **6** [`#N/A`],
  `SUM(VLOOKUP(A1:A3,A1:B3,2,FALSE))` **60** [`#VALUE!`], `SUM(CHOOSE(A1:A3,10,20,30))` **60** [`#VALUE!`],
  `SUM(LARGE(A1:A3,A1:A3))` **6** [`#VALUE!`], `SUM(COUNTIF(A1:A3,A1:A3))` **3** [`0`],
  `SUM(INDEX(B1:B3,A1:A3))` **60** [`#VALUE!`], `SUM(RANK(A1:A3,A1:A3))` **6** [`#VALUE!`],
  `SUM(WORKDAY(A1:A3,1))` **9** [`#VALUE!`], `SUM(NETWORKDAYS.INTL(A1:A3,4))` **9** [`#VALUE!`],
  `SUM(NPV(A1:A3/10,10,20,30))` **120.92** [`#VALUE!`], `SUM(TYPE(A1:A3))` **3** [`16`] e
  `SUM(RANDBETWEEN(A1:A3,A1:A3))` **6** [`#VALUE!`]. Duas das respostas do MySheet são **silenciosas** em
  vez de erros: o `0` do `COUNTIF` (o retângulo está no slot de *critério* dele, e não em um slot de
  intervalo, então a rejeição com `#REF!` acima não o alcança e o argumento colapsado não corresponde a
  critério nenhum) e o `16` do `TYPE` (o código de tipo do `#VALUE!` que ele recebeu). O
  `NETWORKDAYS` simples é o único membro da família que o Excel *não* eleva — `SUM(NETWORKDAYS(A1:A3,B1:B3))`
  é **8** lá, que é `NETWORKDAYS(A1,B1)` sozinho, uma interseção implícita ao primeiro elemento e não uma
  elevação por elemento, e `#VALUE!` aqui. Fixado por
  `ElementwiseLiftingTests.AConsumesFunction_IsNotLiftedOverItsScalarSlots_KnownDivergence`.

Todas elas são deliberadamente mantidas como estão por enquanto, e todas menos duas estão registradas para
uma varredura de compatibilidade com o Excel já planejada: a divergência dos dois eixos não está, porque ali
o oráculo não oferece resposta a igualar, e o retângulo em planilha inexistente também não, porque a causa
dele é um atalho sintático e não uma regra.

Subexpressões voláteis dentro do array se comportam como qualquer outra volátil: um `RAND()` (propagado,
ou em uma célula de intervalo que a comparação lê) contamina a célula consumidora, então
[`Recalculate()`](#o-modelo-de-época) a atualiza enquanto uma fórmula de array não volátil permanece em
cache.

## Intervalos nomeados

Um workbook pode definir **nomes** que representam uma expressão — geralmente um intervalo ou célula
qualificados por planilha, mas qualquer expressão (uma constante, uma fórmula, outro nome) é permitida.
Os nomes são de nível de workbook e **case-insensitive**, exatamente como no Excel.

> Um intervalo nomeado **não** é uma **Tabela** do Excel (um ListObject). Um nome é um apelido estático para
> uma expressão; uma tabela é uma região nomeada com colunas nomeadas, linha de totais, um intervalo que
> cresce conforme linhas são adicionadas e uma sintaxe de referência própria (`Tabela1[Valor]`, `[@Valor]`).
> O MySheet modela os nomes *e* o modelo da tabela ([Tabelas](#tabelas) — nome, intervalo, flags de
> cabeçalho/totais e nomes de coluna), mas nem a sintaxe de referência estruturada, nem o intervalo que
> cresce por conta própria: um redimensionamento é uma segunda chamada de `DefineTable`. Veja
> [Interop com Excel → Escopo e limitações](excel-interop.md#escopo-e-limitações).

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
com um literal booleano, também é rejeitado — e um nome já tomado por uma tabela também é, porque nomes e
tabelas compartilham um único namespace ([Tabelas](#tabelas)).

**Ordem de resolução.** Um `NameReference` resolve nesta ordem:

1. **Primeiro o escopo de `LET`** (*shadowing*) — uma vinculação de `LET` com o mesmo nome vence, então
   `LET(Sales, 5, Sales+1)` é `6`, não uma soma sobre o intervalo.
2. **`Workbook.DefinedNames`** — a expressão do nome é avaliada. Um intervalo/união permanece um valor de
   *referência*, então funções que aceitam intervalos o expandem (`SUM(Sales)`); uma única célula ou
   constante é avaliada para seu escalar. As funções que exigem uma referência sintática —
   `VLOOKUP`/`HLOOKUP` (tabela), `INDEX`, `OFFSET`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `AREAS`, `ISREF` —
   aceitam um nome que representa um intervalo (por exemplo, `VLOOKUP(2, Sales, 2)`).
3. Caso contrário, `#NAME?`.

Um nome usado **puro em uma célula** (`=Sales`) também não é um erro: a referência que ele representa sofre
[interseção implícita](#interseção-implícita-na-fronteira-da-célula) com a linha e a coluna da célula da
fórmula, então `=Sales` sobre `Data!A1:A3` mostra `Data!A3` quando digitado na linha 3.

**Ciclos.** Um nome que se refere a si mesmo, diretamente ou por meio de uma cadeia (`A → B → A`), é
detectado por um rastreamento thread-local e produz `#REF!` em vez de estourar a pilha.

## Tabelas

Um workbook também pode registrar **tabelas** — o modelo por trás da parte `<table>` do Excel (um
ListObject): um retângulo nomeado, ancorado numa planilha, com colunas nomeadas. `Workbook.Tables` é o
registro `nome → Table` somente para leitura, e `DefineTable` é seu único escritor.

> **O que é modelado, e o que não é.** O registro guarda o *modelo* da tabela — o nome, o intervalo, as flags
> de cabeçalho/totais e os nomes das colunas — e ele sobrevive ao `Save`/`Load`. A **sintaxe** de referência
> estruturada ainda não está implementada: `=SUM(Tabela1[Valor])` lança `ParseException: Unexpected character
> '[' (at position 11).` (medido em 2026-09-10 na versão que introduz o registro — o ponto final faz parte da
> mensagem), e o `ExcelFile.Load` também não preenche o registro a partir de uma parte `<table>` do xlsx
> ([Interop com Excel → Escopo e limitações](excel-interop.md#escopo-e-limitações)). Ou seja, nada no avaliador
> lê uma tabela ainda: você registra uma para preservar o modelo num round-trip e para dar à sintaxe de
> referência algo contra o que resolver quando ela chegar.

```csharp
var workbook = new Workbook();
workbook.Sheets.Add("Data");

// A sobrecarga A1: `reference` é exatamente a string <table ref="…"> do xlsx, então abrange a tabela TODA.
workbook.DefineTable(
    "Tabela1", "Data", "C2:E9", ["Produto", "Qtd", "Total"],
    hasHeaderRow: true, hasTotalsRow: true);

var table = workbook.Tables["tabela1"];   // case-insensitive, como no Excel

table.HeaderRow;     // 2  — FirstRow, porque HasHeaderRow
table.FirstDataRow;  // 3
table.LastDataRow;   // 8
table.TotalsRow;     // 9  — LastRow, porque HasTotalsRow
table.DataRowCount;  // 6
table.FirstColumn;   // 3  (C)
table.LastColumn;    // 5  (E) — derivado de ColumnNames.Count

table.TryGetColumnIndex("QTD", out var ordinal);      // true, ordinal = 1
table.TryGetColumnRange("qtd", out var column, out var firstRow, out var lastRow);
                                                      // true, 4, 3, 8 — apenas a faixa [#Data]
```

**Geometria.** `FirstRow`/`LastRow` abrangem o intervalo inteiro, **incluindo** as linhas de cabeçalho e de
totais, exatamente como faz o atributo `ref` do xlsx; `FirstColumn` é a coluna mais à esquerda na planilha, e
a última decorre de `ColumnNames.Count`. Todo o resto — `HeaderRow`, `TotalsRow`, `FirstDataRow`,
`LastDataRow`, `DataRowCount`, `LastColumn` — é **derivado**, então as faixas de cabeçalho, dados e totais
nunca podem se contradizer, e nada disso vai para o fio.

**Nomes.** Os nomes de tabela seguem a **regra de nome de tabela do Excel**, não a de nome definido: começar
com uma letra ou `_`, depois apenas letras, dígitos, `.` e `_`, no máximo 255 caracteres, e nada que o Excel
leria como uma referência — uma célula dentro da grade na forma A1 (`A1`, `T1`, `XFD1048576`), a forma `R1C1`
(`R1C1`, `r2c3`) ou as letras isoladas reservadas `C`/`R` (em qualquer caixa). O MySheet também rejeita
`TRUE`/`FALSE`, que seu tokenizador lê como booleanos, e uma barra invertida em qualquer posição do nome (o
Excel permite uma inicial, mas o tokenizador nunca a leria dentro de um identificador). Os nomes padrão do
próprio Excel **são** válidos (`Tabela1`, `Table1`: a sequência de letras passa do limite de três letras de
uma coluna, então não são células), assim como um nome letras-seguidas-de-dígitos fora da grade (`XFE1`,
`A1048577`).

**Um único namespace com os nomes definidos.** O Gerenciador de Nomes do Excel mantém tabelas e nomes juntos
e recusa uma duplicata; o `Workbook` impõe isso simetricamente — `DefineTable` lança `ArgumentException` para
um nome que um nome definido já ocupa, e `DefineName` lança para um que uma tabela ocupa —, de modo que a
invariante nunca depende de qual veio primeiro.

**Colunas.** Os nomes de coluna são escritos por humanos e *não* estão sujeitos à regra de nomes, mas têm de
ser não vazios, ao menos um, e **únicos ignorando a caixa** (o Excel resolve `Table1[col]` para
`Table1[Col]`), que é também como `TryGetColumnIndex`/`TryGetColumnRange` os comparam. Na sobrecarga A1, a
largura do intervalo tem de ser igual a `columnNames.Count`.

**Redefinir substitui.** Uma segunda chamada de `DefineTable` com o mesmo nome substitui a entrada — é assim
que uma tabela é redimensionada ou tem suas colunas alteradas — e o registro mantém exatamente uma entrada
por nome.

**A planilha não precisa existir.** O registro nunca toca em `Sheets`, exatamente como um nome definido: uma
tabela pode nomear uma planilha que não foi adicionada (ou que foi removida). Uma planilha ausente é uma
questão de tempo de avaliação — uma referência a ela resolve para `#REF!` em vez de lançar (veja
[`GetCellValue`](#workbook)).

**Zero linhas de dados é válido.** Uma tabela só de cabeçalho (`ref="A1:A1"` com linha de cabeçalho) é um
estado de modelo legítimo, não um erro: `DataRowCount` é `0` e `TryGetColumnRange` devolve `false` para uma
coluna *conhecida*, deixando a decisão entre `#REF!` e vazio para quem chama, em vez de devolver um intervalo
invertido.

**Redefinição e valores desatualizados.** Nem `DefineTable` nem `DefineName` removem nada do cache de
memoização — uma (re)definição não muda nenhuma célula —, então uma fórmula que já leu a definição antiga
continua servindo seu valor memoizado. Duas formas de tornar a mudança observável:

- Chamar `InvalidateCache()` você mesmo: o cache inteiro vai embora e tudo é recomputado sob demanda.
- Ou deixar o motor de recálculo incremental fazer isso. `Workbook.CreateRecalculationEngine()` devolve um
  `RecalculationEngine` — um grafo de dependências reverso que, dadas as células que você reporta como
  editadas, remove do cache apenas o cone afetado em vez do cache inteiro. Ele rastreia mudanças de definição
  separadamente (elas não tocam em nenhuma planilha, então nenhuma versão estrutural de planilha se move) e
  responde a uma delas com uma invalidação **completa**, não com uma remoção parcial: o
  `RecalculationEngine.Recalculate(...)` seguinte chama `InvalidateCache()` por você e relata
  `Mode = RecalculationMode.FullFallback` com `DirtyCellCount = -1`. Seu correspondente de planejar antes de
  agir, o `RecalculationEngine.EstimateImpact(...)`, relata `RecommendFull = true` (`ConeSize = -1`) pelo
  mesmo motivo e **não** consome o sinal, então a ordem das duas chamadas não importa: estimar primeiro ainda
  deixa a remoção para o `Recalculate` seguinte.

> `RecalculationEngine.Recalculate(edited)` **não** é `Workbook.Recalculate()`. O método do workbook atualiza
> apenas as células voláteis e avança a [época volátil](#o-modelo-de-época); o método do motor é a passagem
> incremental guiada por edições descrita acima. Os dois nomes não têm relação além da palavra.

## Funções voláteis

Cinco funções são **voláteis** — seu resultado não é determinado apenas pelas células que leem. Quatro
dependem do relógio ou de um sorteio aleatório: `NOW()`, `TODAY()`, `RAND()` e `RANDBETWEEN(bottom, top)`.
A quinta, `INDIRECT(ref_text, [a1])`, depende de *quais células ela lê*: a referência é montada a partir
de texto em tempo de avaliação, então nenhuma dependência estática é conhecível. O MySheet oferece a elas
os dois comportamentos que o Excel define para esse caso — *recalcular sob demanda* e *volatilidade
contagiosa* — sem um grafo de dependências, por meio de um **modelo de cache por época** (*epoch*).

### O modelo de época

Dentro de uma mesma época, uma volátil é calculada **uma única vez** e fica em cache, de modo que todo
`NOW()`/`TODAY()` de uma mesma passagem concorda no mesmo instante e uma célula `RAND()` lida duas vezes
retorna o mesmo valor. Uma célula que toca uma volátil — diretamente (`=NOW()`) ou transitivamente
(`=A1+1`, em que `A1=NOW()`) — fica em cache **e marcada**; a marca aproveita a mesma propagação
thread-local que o detector de ciclos usa, então a volatilidade se espalha para os dependentes de graça.

- **`Recalculate()`** avança a época: descarta **apenas** as células marcadas (tocadas por uma volátil) e
  lê o relógio/sorteia o RNG de novo, mantendo em cache toda célula estável. Os valores são atualizados de
  forma **preguiçosa** (*lazy*) — a próxima leitura os recalcula. É a chamada barata para "me dê a hora
  atual / um novo sorteio aleatório".
- **`InvalidateCache()`** continua limpando **tudo** (use-a após editar entradas de célula) e também
  reinicia a época.

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");
sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);
sheet["B1"] = ExpressionParser.Parse("=A1+1", sheet);   // volátil transitivamente

ComputedValue first = workbook.GetCellValue("Sheet1", "A1");
ComputedValue again = workbook.GetCellValue("Sheet1", "A1");   // mesma época → idêntico

workbook.Recalculate();                                        // avança a época
ComputedValue later = workbook.GetCellValue("Sheet1", "A1");   // sorteado de novo → mais recente
ComputedValue b = workbook.GetCellValue("Sheet1", "B1");       // B1 também atualizada (contágio)
```

O relógio é lido de forma **preguiçosa** — na primeira leitura volátil de uma época, não dentro de
`Recalculate()` — então `NOW()` reflete o instante em que o valor foi de fato produzido, e nada é lido se
nenhuma volátil for consultada.

### Injetando o relógio e o RNG

`NOW`/`TODAY` leem `Workbook.TimeProvider` (padrão `TimeProvider.System`) em **horário local**, como o
Excel. Atribua qualquer `TimeProvider` para congelar o tempo durante um lote, ou para tornar os testes
determinísticos independentemente do relógio e do fuso da máquina. `RAND`/`RANDBETWEEN` sorteiam a partir
de um RNG persistente; defina `Workbook.RandomSeed` (um `int?`) **antes da primeira leitura volátil** para
tornar toda a execução reproduzível, ou deixe-o `null` (padrão) para um RNG semeado (*seeded*) pelo
relógio.

```csharp
workbook.TimeProvider = TimeProvider.System;   // o padrão; troque por um fake para controlar o relógio
workbook.RandomSeed = 12345;                    // RAND/RANDBETWEEN reproduzíveis
```

O RNG avança entre épocas e nunca é semeado de novo, então passagens sucessivas de `Recalculate()`
produzem sorteios diferentes enquanto uma única célula permanece estável dentro da sua época. Nem
`TimeProvider` nem `RandomSeed` são serializados (são configuração de tempo de execução): um workbook
carregado começa a partir de `TimeProvider.System` e um RNG não semeado.

### Limites (por design)

- **Sem atualização por célula.** É possível atualizar *todas* as voláteis (`Recalculate()`), não apenas
  uma. Atualizar somente `A1=NOW()` deixando uma `B1=A1+1` em cache desatualizada exigiria um grafo de
  dependências reverso, que a engine deliberadamente não mantém — então a atualização grosseira, porém
  correta, é a que é oferecida.
- **`OFFSET` não é volátil.** O Excel marca `OFFSET` como volátil como uma rede de segurança para o
  recálculo automático; aqui a invalidação é explícita, então marcá-la contaminaria metade de uma planilha
  sem necessidade — uma divergência consciente.

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
