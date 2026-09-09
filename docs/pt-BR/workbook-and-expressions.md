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
união também fica deliberadamente em `#VALUE!`.

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

**Suportado.** Os consumidores são os agregadores numéricos (`SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX` e —
através da mesma dobra — `SMALL`, `LARGE`, os percentis), `INDEX` e `SUMPRODUCT`. Um argumento é avaliado
como um array quando é uma comparação de **intervalo fechado** (`B2:B5="Show"`), um `IF` cuja condição é um
array assim (com ou sem ramo `else`), ou `ROW`/`COLUMN` sobre um retângulo. Esse retângulo pode estar
escrito literalmente (`SUM(ROW(A1:C3))` = 18, `SUM(COLUMN(A1:C3))` = 18) ou apenas ser *denotado* pelo
argumento — um [nome definido](#intervalos-nomeados) (`SUM(ROW(MyName))` = 6 e `COUNT(ROW(MyName))` = 3 para
um nome sobre três linhas, enquanto `COUNT(MyName)` conta os valores das próprias células) ou um intervalo
`:` com extremidades que retornam referências (`SUM(ROW(INDEX(A1:A3,1,1):A3))` = 6). Escalares são
propagados (*broadcast*) por todo o vetor. Um `IF` sem ramo produz um lógico `FALSE` onde a condição é
falsa, e os agregadores ignoram lógicos/texto (exatamente por isso `SMALL(IF(…))` pula as linhas sem
correspondência). O primeiro erro por elemento prevalece, como no Excel.

**Não suportado (por design).**

- Uma **célula seca** cuja fórmula inteira é o array mantém `#VALUE!` — `=IF(B2:B5="Show",1,0)` sozinha
  ainda é um erro. Arrays existem apenas como *argumentos* dentro dos consumidores acima, nunca como o
  valor de uma célula (o cache por célula permanece estritamente escalar). Isso **não** contradiz a
  [interseção implícita na fronteira da célula](#interseção-implícita-na-fronteira-da-célula): aquela regra
  intersecta uma *referência*, e um array computado não é uma — então `=IF(TRUE,A1:A3,B1)` em uma célula
  continua sendo `#VALUE!`, enquanto o `=A1:A3` puro ao lado dela é `A3`.
- Uma **função que retorna referência** como argumento de `ROW`/`COLUMN` permanece escalar:
  `SUM(ROW(INDEX(A1:A3,1,1)))` é `1`, a linha superior da referência resolvida, e não o vetor `[1,2,3]`.
  Descobrir o formato dela resolveria o argumento uma segunda vez e sortearia uma volátil duas vezes, então
  ali o formato de array é deliberadamente adiado.
- A família de **critérios / varredura posicional** não lê um array computado. `SUMIF`/`SUMIFS`,
  `COUNTIF`/`COUNTIFS`, `AVERAGEIF`/`AVERAGEIFS` e `MAXIFS`/`MINIFS` percorrem seus argumentos posição a
  posição, e um array em uma dessas posições simplesmente não é um intervalo — ele colapsa para uma
  sequência de um único elemento contendo o `#VALUE!` de um intervalo em uma operação aritmética. O que
  cada função faz com esse elemento solitário assume **quatro** formas, todas fixadas por teste: as formas
  **pareadas** (`SUMIFS`/`AVERAGEIFS`/`MAXIFS`/`MINIFS`), que têm um intervalo de critérios real ao lado do
  argumento colapsado, veem 1 elemento contra 3 e levantam a diferença de comprimento da varredura —
  `#VALUE!`, que é também a resposta do Excel; `SUMIF((A1:A3)*1, ">0")` não tem contra o que divergir, então
  o `#VALUE!` solitário não corresponde a critério nenhum e a varredura volta vazia — `0`;
  `COUNTIF`/`COUNTIFS` igualmente contam essa varredura vazia como `0`; e `AVERAGEIF` a divide por uma
  contagem zero — `#DIV/0!`. As três últimas são respostas **silenciosas**, não erros. O `SUMPRODUCT` é o
  único membro dessa família que optou por aceitar arrays computados; os consumidores de dobra listados em
  **Suportado** acima (`SUM(IF(…))` e companhia) sempre os aceitaram.
- Um intervalo **aberto/de coluna inteira** em posição de array é recusado e o consumidor permanece em seu
  caminho escalar/de intervalo comum — a única exceção é a identidade `INDEX(ROW($A:$A), n)` acima, que
  retorna `n` sem materializar a coluna. `SMALL(IF(A:A=…, ROW(A:A)), k)` sobre uma coluna *aberta* portanto
  não é avaliado como array.
- Uma condição **escalar** mantém o curto-circuito nativo do `IF` — apenas uma condição de array conduz o
  zip.

**Uma divergência conhecida.** `SUM(ROW(Ghost!A1:A3))` — um retângulo escrito *literalmente* sobre uma
planilha que não existe, em posição de array — responde `6`, os números de linha `1+2+3`, enquanto o Excel
responde `#REF!`. O `ROW(Ghost!A1:A3)` escalar na mesma pasta de trabalho já é `#REF!`, assim como o
caminho de array sobre um nome que representa o mesmo intervalo (`SUM(ROW(GhostName))`): a divergência está
apenas no retângulo escrito por extenso, cujo caminho rápido sintático vai direto a um vetor de linhas ou
colunas e nunca resolve a referência, de modo que a guarda de planilha ausente — executada por todo caminho
que resolve — não tem sobre o que atuar. Isso está fixado como uma lacuna, e não como uma regra, por
`MiniCseConsumerTests.Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap`, para que fechá-la seja
uma edição deliberada.

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
> O MySheet modela o primeiro, e não a segunda — veja
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
com um literal booleano, também é rejeitado.

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

## Funções voláteis

Quatro funções são **voláteis** — seu resultado depende do relógio ou de um sorteio aleatório, não apenas
das células que leem: `NOW()`, `TODAY()`, `RAND()` e `RANDBETWEEN(bottom, top)`. O MySheet oferece a elas
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
- **`INDIRECT` não está implementado** (resolver uma referência a partir de texto é um recurso à parte).

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
