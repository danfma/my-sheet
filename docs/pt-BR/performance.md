# Desempenho

*Tradução do documento canônico em inglês ([performance.md](../performance.md)). Em caso de divergência, o inglês prevalece.*

A meta de design do MySheet é uma engine de fórmulas barata o suficiente para rodar no caminho crítico de
um servidor: carregar um workbook uma vez e, então, avaliar e reavaliar muitas células com alocação
mínima e sem custo de preparação por leitura. Este guia explica os três mecanismos que fazem isso
funcionar — a união `ComputedValue` sem alocações, a memoização por célula e o `RunWithLargeStack` — e
reporta os números realmente medidos neste repositório.

## Avaliação sem alocações

`Expression.Evaluate` retorna um [`ComputedValue`](computed-value.md): uma união em `readonly struct` com
um `double`, um `object?` e uma tag de um byte. Números, booleanos, em branco e erros são carregados
inteiramente dentro da struct — **sem boxing, sem alocação no heap**. Texto e referências apenas carregam
um ponteiro para uma string ou um nó que já existia.

Escolhas de apoio no mesmo espírito:

- `EvaluationContext` é uma `readonly struct`, então propagar o contexto pela avaliação recursiva não
  aloca nada (apenas as vinculações de nome de `LET`, quando usadas, vivem no heap).
- O cache de memoização armazena a struct `ComputedValue` **inline** no dicionário — sem uma caixa (box)
  de vida longa por célula, que era o que antes gerava pressão de GC Gen1.
- Funções que consomem intervalos (`SUM(A1:A1000)`, lookups, …) enumeram os valores das células através
  do cache com o mesmo tipo sem boxing.

### Números medidos

O design do `ComputedValue` foi adotado com base em um experimento com BenchmarkDotNet neste repositório
(`plans/cellvalue-boxing-experiment/`, spike executável em `benchmarks/Danfma.MySheet.Benchmark`),
comparando o avaliador anterior, que retornava `object?`, com a união em struct. Ambiente: Apple M1 Pro,
.NET 10, BenchmarkDotNet ShortRun. Destaques (razões em relação à linha de base com `object?` e boxing):

| Carga de trabalho | `object?` com boxing | `ComputedValue` | Alocação |
| --- | --- | --- | --- |
| Cadeia de dependências cumulativa, 3.000 células | 44,4 μs, 144 KB | 18,4 μs (**0,42×** do tempo) | **0 B** (antes, 144 KB) |
| Fold de `SUM` sobre 100.000 valores | 476,9 μs, 2,4 MB | 111,8 μs (**0,23×** do tempo) | **0 B** (antes, 2,4 MB) |
| Aritmética mista, 100.000 operações | 611,7 μs, 2,28 MB | 202,4 μs (**0,33×** do tempo) | **0 B** (antes, 2,28 MB) |
| Grafo intensivo em cache, 100.000 células | 5,64 ms, 2,4 MB | 4,95 ms (**0,88×** do tempo) | **0 B** (antes, 2,4 MB) |

Em resumo: **2,3–4,3× mais rápido em caminhos com aritmética pesada, 4–12% mais rápido em caminhos
intensivos em cache, e zero coletas de GC em todos os casos**. Duas notas de honestidade: esses números
comparam alternativas de design do próprio MySheet, não outras bibliotecas; e as tabelas brutas
(incluindo execuções com 1.000 células) estão em `plans/cellvalue-boxing-experiment/results-*.md`, caso
você queira os dados completos. A eliminação de alocações — zero atividade Gen0/Gen1 em um extrator em
segundo plano — foi o argumento decisivo, mais do que a taxa de transferência bruta.

## Memoização

`Workbook.GetCellValue(sheetName, id)` armazena em cache o `ComputedValue` de cada célula sob
`(planilha, id)`. Toda `CellReference` dentro de uma fórmula resolve através do mesmo cache, então uma
célula referenciada por muitas fórmulas — ou alcançada tanto diretamente quanto por uma expansão de
intervalo — é **calculada exatamente uma vez**:

```csharp
sheet["A1"] = ExpressionParser.Parse("=EXPENSIVE()", sheet);
sheet["B1"] = ExpressionParser.Parse("=A1+A1", sheet);
sheet["C1"] = ExpressionParser.Parse("=SUM(A1:A1)+A1", sheet);

workbook.GetCellValue("Sheet1", "B1");   // EXPENSIVE() roda uma vez
workbook.GetCellValue("Sheet1", "C1");   // ...e não roda de novo aqui
```

### A invalidação é explícita

O cache **nunca é invalidado automaticamente**. Depois de mutar células, chame `InvalidateCache()`
(uma limpeza completa) antes de ler novamente:

```csharp
sheet["A1"] = new NumberValue(20);
// Sem isto, as leituras ainda servem os valores previamente armazenados em cache:
workbook.InvalidateCache();
```

Por que esse design (decidido e documentado em `plans/memoization.md`):

- A carga de trabalho alvo é **extração com muitas leituras e mutações raras e em lote** — edite,
  invalide uma vez e depois leia bastante.
- A invalidação cirúrgica por grafo de dependências foi avaliada e rejeitada: intervalos tornam o grafo
  reverso caro (`SUM(A1:A1000)` sozinho são 1.000 arestas) para pouco ganho sob mutações raras.
- O trade-off conhecido: cada referência adiciona uma consulta de dicionário, um custo pequeno em células
  triviais que se paga sempre que há células compartilhadas ou caras.

O cache é um `ConcurrentDictionary`, seguro para leitores concorrentes em segundo plano. Ele não é
serializado — um workbook carregado começa frio e vai sendo preenchido de forma preguiçosa.

### Referências circulares

A camada de memoização rastreia as células em avaliação na thread atual. Um ciclo (`A1=B1`, `B1=A1`) é
detectado e retorna `#REF!` (`Error.Ref`), em vez de estourar a pilha. O rastreamento é thread-local,
então a avaliação concorrente da mesma célula em threads diferentes não é um falso ciclo.

## Referências de coluna inteira em escala

Fórmulas que consomem uma **referência de coluna inteira** — `MATCH(x, A:A)`, `VLOOKUP(x, A:B, 2, FALSE)`,
`SUMIF(A:A, k)`, `SMALL(A:A, k)`, `SUM(A:A)` — são o caso patológico para uma engine ingênua. Cada fórmula
dessas varre toda célula populada da coluna, então uma planilha com *F* fórmulas de coluna inteira sobre
uma coluna de *N* células custa **O(F·N)**. Em um workbook real (~506 mil células em uma coluna, ~400 mil
fórmulas referenciando-a) isso é ~2×10¹¹ visitas — aproximadamente **57 minutos** de varredura pura em uma
passagem de carregar-uma-vez / ler-uma-vez. Dois caches internos reduzem isso a segundos. Ambos são
limitados, descartáveis e não adicionam **nenhuma API pública** — o modelo esparso e o caminho rápido de
célula única permanecem intocados.

### Camada 1 — o índice estrutural

Por planilha, um índice preguiçoso mapeia `coluna → ids de célula ordenados por linha` (e o simétrico
`linha → ids` sob demanda). Ele responde "quais células `A:A` cobre, em ordem?" sem revarrer o dicionário
de células. A resolução de limites de range (a busca em tabela que `VLOOKUP`/`INDEX`/`OFFSET` fazem em
toda avaliação) lê a caixa delimitadora diretamente das listas ordenadas por busca binária.

É isso que torna **colunas pequenas em uma planilha grande** baratas assim que o índice existe: antes do
índice, uma fórmula referenciando uma coluna de 16 células ainda percorria todas as ~1,5 milhão de
chaves da planilha para encontrar essas 16; depois dele, apenas as 16 são visitadas.

**Admissão na segunda leitura de uso.** Construir o índice é, em si, uma passagem O(N) (mais um bucket
por coluna ou linha populada). Isso só compensa quando uma planilha é lida através de ranges abertos
**mais de uma vez em uma época**. Uma planilha lida *exatamente uma vez* — o formato "`InvalidateCache()`
depois uma leitura de coluna inteira por época" — pagaria a construção de um índice da planilha inteira
só para servir uma coluna que nunca reutiliza. Então, exatamente como o snapshot de range abaixo, o
índice é **admitido na segunda leitura de range aberto de uma planilha, não na primeira**: a primeira
leitura se serve de uma **varredura direta de chaves** (o caminho pré-índice), coletando e ordenando
*apenas os ids correspondentes* na mesma ordem determinística que o índice produz; a segunda leitura
constrói o índice e toda leitura seguinte o reutiliza. Em uma passagem de reuso intenso (milhares de
fórmulas sobre `A:A`) a única varredura extra na primeira fórmula é imperceptível; em uma passagem de
uma leitura por época, ela permanece na paridade pré-índice (2.6.1) em vez de reconstruir o índice a
cada época.

**Ordenação preguiçosa por bucket.** Construir o índice organiza os ids em buckets, mas **não** os
ordena; a lista de cada coluna é ordenada (por linha) apenas no primeiro acesso, então uma leitura que
toca apenas uma coluna estreita de uma planilha larga ordena somente essa lista, não toda coluna.

### Camada 2 — caches de range por época

Para um range populado acima de um limiar de tamanho que é lido **mais de uma vez em uma época**, a
engine materializa um **snapshot** — os `ComputedValue`s do range, em ordem de enumeração, lidos
exatamente uma vez através da Camada 1 + o cache de células memoizado. Cada acelerador é então
construído **de forma preguiçosa, sob demanda, e compartilhado por toda fórmula da época**:

- um **hash de valor exato** (`valor → primeira posição`) → `MATCH(…,0)`, `XLOOKUP`/`XMATCH` exatos,
  `VLOOKUP(…,FALSE)` em O(1);
- um **índice ordenado** com posições de máximo por prefixo/sufixo → `MATCH` aproximado tipo 1/-1 e
  `VLOOKUP(…,TRUE)` por busca binária, reproduzindo o desempate "último dos empatados" do Excel para
  qualquer ordem de entrada;
- uma **visão numérica ordenada** → `SMALL`/`LARGE`/`MEDIAN`/`PERCENTILE`/`QUARTILE` por indexação
  direta;
- um **mapa de igualdade numérica** (`valor → (soma, contagem)`) → `SUMIF`/`COUNTIF`/`AVERAGEIF` de
  igualdade em O(1);
- um **memo de agregado** → `SUM`/`COUNT`/`MAX`/`MIN`/`AVERAGE` de um range puro repetido dobrado uma
  única vez.

O custo geral cai de O(F·N) para **O(N + F·log N)**. A semântica é preservada bit a bit: quando um valor
não se encaixa de forma limpa num índice (buscas equivalentes a branco, critérios de wildcard/comparação,
o `XLOOKUP` aproximado "mais próximo"), o consumidor cai para uma **varredura linear sobre o snapshot em
cache** — ainda servida pelo cache, então a releitura O(N) dos valores das células desaparece mesmo no
caminho de fallback.

**Admissão na segunda leitura de uso.** Materializar um snapshot só compensa quando um range é lido
repetidamente. Um range lido *exatamente uma vez* por época — uma janela deslizante (`SUM(A$1:A500)`,
`SUM(A$1:A501)`, …), uma busca limitada de uso único, um loop com muita invalidação — pagaria uma
construção O(N) que nunca reutiliza. Então um range é **admitido na sua segunda leitura, não na
primeira**: a primeira leitura apenas registra um marcador leve e segue o caminho linear; o snapshot é
construído na segunda leitura e compartilhado por toda leitura seguinte. Em uma passagem de reuso
intenso (milhares de fórmulas sobre `A:A`) isso custa uma varredura linear extra na primeiríssima
fórmula e é imperceptível; em uma passagem de uso único, isso remove por completo a materialização
desperdiçada. (Um marcador é descartado junto com o cache ao final da época; um limite defensivo de 64
mil marcadores impede que uma inundação adversarial de ranges distintos de uso único faça o conjunto de
marcadores crescer sem limite.)

**O limiar de 256 células.** Um range com menos de 256 células populadas não é colocado em cache — nem
sequer marcado: uma varredura linear já vence ali, e rastrear todo range minúsculo só inundaria o
dicionário. A verificação usa uma estimativa barata de limite superior (área do retângulo, ou a soma
das listas do índice estrutural cobertas) — ela nunca materializa um snapshot só para decidir.

### Ciclo de vida: quando os caches são construídos e descartados

Ambos os caches são por workbook, baseados em `ConcurrentDictionary`, criados livres de corrida no
primeiro uso, e **não serializados** (um workbook carregado começa frio). Eles diferem apenas na
invalidação, porque o índice estrutural descreve *estrutura* e o snapshot do range carrega *valores*:

| Cache | Construído | `Recalculate()` | `InvalidateCache()` |
| --- | --- | --- | --- |
| Índice estrutural (Camada 1) | **segunda** leitura de range aberto de uma planilha | **sobrevive** (estrutura ≠ valores) | descartado |
| Snapshots de range (Camada 2) | **segunda** leitura de range cacheável | **descartado** (valores podem estar contaminados por volatilidade) | descartado |

```csharp
// Carregue uma vez, depois leia bastante: os caches se preenchem de forma preguiçosa na primeira
// passagem e toda leitura posterior é servida a partir deles.
var total = workbook.GetCellValue("Calc", "A1");   // 1ª leitura de A:A → marca, caminho linear
var next  = workbook.GetCellValue("Calc", "A2");   // 2ª leitura → constrói o snapshot compartilhado
var more  = workbook.GetCellValue("Calc", "A3");   // 3ª+ leitura → servida pelo snapshot compartilhado

// Depois de editar células, limpe tudo antes de ler de novo:
sheet["A1"] = new NumberValue(20);
workbook.InvalidateCache();                          // descarta ambas as camadas

// Uma atualização volátil (NOW/RAND) descarta apenas os snapshots de valor; o índice estrutural sobrevive:
workbook.Recalculate();
```

### Números medidos

Benchmark sintético de coluna inteira (`benchmarks/Danfma.MySheet.Benchmark`, `--whole-column-scale`),
Apple Silicon, .NET 10. Cada linha é um `InvalidateCache()` + uma passagem completa de avaliação (o ciclo
carregar-uma-vez / ler-uma-vez). "Antes" é a engine pré-cache; "Depois" é a engine de duas camadas.

**Reduzido (50 mil células de dados × 10 mil fórmulas), bloco de fórmula sobre a coluna grande, wall-clock:**

| Fórmula | Antes | Depois |
| --- | --- | --- |
| `MATCH(…,1)` | 80,5 s | 141 ms |
| `MATCH(…,0)` | 73,6 s | 95 ms |
| `VLOOKUP(…,FALSE)` | 8,1 s | 155 ms |
| `SUMIF` (igualdade) | 70,5 s | 68 ms |
| `SMALL` | 59,5 s | 75 ms |
| `SUM` (repetido) | 55,2 s | 89 ms |

**Escala plena (500 mil células de dados ≈ 1,5 milhão de células na planilha × 100 mil fórmulas), medido,
não extrapolado:** qualquer bloco único de 100 mil fórmulas de coluna inteira sobre a coluna grande roda
em **≤1,15 s** (contra uma estimativa de **~4,2 horas** na engine pré-cache — ~18.000× nas funções de
varredura pura). Os 14 blocos (7 funções × 2 alvos, 1,4 milhão de avaliações) totalizam **7,7 s**. A
memória extra que a avaliação retém (memo de célula + índice estrutural + snapshot de range) é de
~75–122 MB para uma coluna grande de 500 mil células, transiente e descartada na invalidação; o pico do
processo é dominado pelo próprio workbook, não pelos caches.

> **Meça, não extrapole.** O custo pré-cache era verdadeiramente O(F·N), então cronometrar 1 mil fórmulas
> e multiplicar por 100 era uma estimativa válida. Com os caches o custo é O(N + F·log N): a construção do
> snapshot é um custo único O(N) amortizado por todo o bloco, então uma amostra de 1 mil × 100 multiplica
> essa construção única por 100 e superestima muito. O número da escala plena acima é medido sobre as
> 100 mil fórmulas reais.

### Comparação com o ClosedXML

A mesma carga de trabalho reduzida de coluna inteira (50 mil × 10 mil), cada engine em seu próprio
processo:

| Fórmula | MySheet | ClosedXML |
| --- | --- | --- |
| `MATCH(…,1)` (aproximado) | 123 ms | 67.294 ms |
| `MATCH(…,0)` (exato) | 105 ms | 1.821 ms |
| `VLOOKUP(…,FALSE)` | 154 ms | 1.463 ms |
| `SUMIF` (igualdade) | 95 ms | 16.874 ms |
| `COUNTIF` (igualdade) | 97 ms | 16.970 ms |
| `SUM` (repetido) | 113 ms | 20.192 ms |
| `SMALL` | ~78 ms | **`#NAME?` — não implementado** |

Os resultados são idênticos onde ambas as engines calculam. O MySheet é mais rápido em toda função
suportada (9,5× a 547×) **e** mantém um working set de pico menor (~152 MB vs. ~156 MB) mesmo carregando
os caches por época — e ele responde `SMALL`/`LARGE` sobre uma coluna inteira, o que o ClosedXML não
avalia de forma alguma.

## Cadeias profundas: `RunWithLargeStack`

A avaliação é recursiva, e o risco não são *fórmulas* profundas, e sim **cadeias de dependência entre
células** longas — por exemplo, uma coluna cumulativa (`B2=B1+A2`, `B3=B2+A3`, …) com milhares de linhas.
Calcular a última célula percorre recursivamente a cadeia inteira, e a memoização não ajuda no *primeiro*
cálculo. Com a pilha padrão de thread, de ~1 MB, isso estoura depois de alguns milhares de frames — e o
.NET não consegue capturar uma `StackOverflowException`.

A solução é executar o lote de avaliação em uma thread com uma pilha grande reservada:

```csharp
var value = Workbook.RunWithLargeStack(() => workbook.GetCellValue("Sheet1", "A20000"));

// Ou envolva um lote de extração inteiro — o custo da thread é pago uma vez, não por célula:
var totals = Workbook.RunWithLargeStack(() =>
{
    var results = new Dictionary<string, ComputedValue>();

    foreach (var id in workbook["Sheet1"].Keys)
    {
        results[id] = workbook.GetCellValue("Sheet1", id);
    }

    return results;
});
```

- O tamanho de pilha padrão é 256 MB, e trata-se de uma **reserva**: a memória física cresce apenas com a
  profundidade realmente alcançada. Um segundo parâmetro opcional sobrescreve o tamanho.
- Exceções lançadas dentro do trabalho são capturadas e relançadas no chamador com o stack trace
  original.
- Uma cadeia cumulativa de 20.000 células — que estoura a pilha padrão — é coberta pela suíte de testes
  por meio dessa API.
- O exportador e o mesclador de Excel fazem isso por você: `SaveAsExcel` e `MergeIntoExcel` avaliam todas
  as células de antemão dentro de uma única chamada de `RunWithLargeStack`.

## Checklist prático

1. Leia células através de `GetCellValue` (memoizado), não reavaliando expressões em um laço.
2. Agrupe as mutações em lote e, então, chame `InvalidateCache()` uma vez.
3. Envolva grandes lotes de extração em uma única chamada de `Workbook.RunWithLargeStack(...)`.
4. Extraia resultados com `TryGet*`/`To*`; mantenha `AsObject()` (que faz boxing) fora de laços críticos.
5. Funções personalizadas são armazenadas em cache por célula, como as nativas — não dependa de
   reexecução a cada leitura ([detalhes](custom-functions.md#interação-com-a-memoização)).

## Benchmarks no repositório

`benchmarks/Danfma.MySheet.Benchmark` é um projeto BenchmarkDotNet com os benchmarks da engine (parsing,
construção e avaliação de workbook em memória, com o ClosedXML como ponto de referência independente em
memória) e o spike do `ComputedValue`:

```shell
dotnet run --project benchmarks/Danfma.MySheet.Benchmark -c Release
```

Os números de escala de coluna inteira acima vêm de um harness dedicado de wall-clock (a linha de base
O(F·N) é lenta demais para iterar sob o BenchmarkDotNet):

```shell
# Escala reduzida (50k × 10k), roda em segundos:
dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale
# + a escala plena medida (500k × 100k), com memória por bloco:
dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale --full
```
