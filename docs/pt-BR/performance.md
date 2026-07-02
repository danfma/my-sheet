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
