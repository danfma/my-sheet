# Migrando para a 3.0

*Tradução do documento canônico em inglês ([migrating-to-3.0.md](../migrating-to-3.0.md)). Em caso de divergência, o inglês prevalece.*

O MySheet 3.0 **encapsula o armazenamento de células da `Sheet`**. O dicionário que costumava ser uma
propriedade pública, mutável e com `init` agora é exposto como uma **visão somente leitura**, e toda
mutação passa a fluir por dois caminhos: o `set` do indexador (inalterado) e um novo `Remove`. A quebra é
estritamente **em tempo de compilação e restrita** — a leitura permanece intocada, a escrita via
`sheet["A1"] = …` permanece intocada, e **workbooks serializados são 100% compatíveis** nos dois sentidos
(o formato de fio é byte a byte idêntico).

## Por quê

A 3.0 transforma o armazenamento de células em um único **ponto de estrangulamento** de escrita.
Canalizar toda inserção, sobrescrita e remoção por um único lugar é a base sobre a qual o trabalho
posterior da série 3.x se constrói — um índice estrutural mantido na escrita (mantido atualizado na
escrita, em vez de reconstruído a cada época de cache) e, mais adiante, um grafo de dependências reverso
para recálculo incremental. Nada dessa infraestrutura é lançada na 3.0; o release é o encapsulamento que
torna isso possível sem tocar de novo nos chamadores.

## O que mudou

O tipo de `Sheet.Cells` mudou, e dois membros foram adicionados:

| Membro | 2.x | 3.0 |
| --- | --- | --- |
| `Sheet.Cells` | `public Dictionary<string, Expression> { get; init; }` | `public IReadOnlyDictionary<string, Expression> { get; }` |
| `Sheet.SetCell(string id, Expression expr)` | — | `internal` — o único caminho de escrita para o qual o `set` do indexador delega |
| `Sheet.Remove(string id)` | — | `public bool` — remove uma célula, retorna se ela existia |

Todo o resto em `Sheet` permanece inalterado: o indexador de leitura `sheet["A1"]` (em branco para uma
célula ausente), o indexador de escrita `sheet["A1"] = expr`, `Count`, `Keys`, `Values`, `ContainsKey`,
`TryGetValue`, e a enumeração (`foreach (var (id, expr) in sheet)`) mantêm exatamente o comportamento da
2.x.

## Atualizando o seu código

### Leitura — nada a fazer

Todo caminho de leitura é compatível em nível de código-fonte. `Cells` continua enumerável e indexável, e
`Count`/`Keys`/`Values` continuam retornando as mesmas coisas:

```csharp
foreach (var (id, expr) in sheet) { /* … */ }   // inalterado
var n = sheet.Count;                             // inalterado
if (sheet.ContainsKey("A1")) { /* … */ }         // inalterado
var expr = sheet["A1"];                           // inalterado (em branco para uma célula ausente)
var only = sheet.Cells["A1"];                     // continua funcionando — indexador somente leitura
```

As únicas leituras que quebram são as que **mutavam através de `Cells`**, por exemplo
`sheet.Cells["A1"] = expr`, `sheet.Cells.Remove("A1")` ou `sheet.Cells.Clear()`. Essas nunca precisaram
passar pela propriedade — use os membros de `Sheet` abaixo.

### Escrita — use o indexador

O indexador de escrita permanece inalterado e continua sendo o caminho de escrita pretendido. Se você
estava mutando o dicionário subjacente diretamente, troque para ele:

```csharp
// 2.x — alcançando o dicionário (então mutável) diretamente
sheet.Cells["A1"] = new NumberValue(10);

// 3.0
sheet["A1"] = new NumberValue(10);   // delega para o ponto de estrangulamento SetCell
```

### Exclusão — use `Remove`

Excluir uma célula costumava significar `sheet.Cells.Remove("A1")`. Isso agora é `Sheet.Remove`:

```csharp
// 2.x
sheet.Cells.Remove("A1");

// 3.0
bool existed = sheet.Remove("A1");   // true se havia uma célula ali, false para uma operação sem efeito
```

`Remove` compartilha a **mesma semântica de invalidação explícita de uma escrita**: ele não limpa valores
memoizados por conta própria. Remover uma célula muda o resultado de toda fórmula que a lia, então —
exatamente como após uma escrita — chame `workbook.InvalidateCache()` para que a mudança seja observada
na próxima leitura.

### Inicializadores de objeto — construa e depois popule

Como `Cells` perdeu seu acessor `init`, um inicializador de objeto que semeava células não compila mais.
Construa a planilha e depois popule através do indexador:

```csharp
// 2.x
var sheet = new Sheet
{
    Name = "Sheet1",
    Cells = { ["A1"] = new NumberValue(10), ["A2"] = new NumberValue(20) },
};

// 3.0
var sheet = new Sheet { Name = "Sheet1" };
sheet["A1"] = new NumberValue(10);
sheet["A2"] = new NumberValue(20);
```

Na prática a maior parte do código já cria planilhas com `workbook.Sheets.Add("Sheet1")` e as preenche
através do indexador, então não há nada para mudar ali.

## Compatibilidade de serialização

**Nenhuma ação necessária.** O dicionário de células passou de uma propriedade pública para um campo
privado `[MemoryPackInclude]` **exatamente na mesma posição de declaração**. O MemoryPack ordena os
membros pela declaração, então o esquema de fio — membro nº 3, um `Dictionary<string, Expression>` — é
byte a byte idêntico. Workbooks salvos pela 2.x (e pela 1.x) carregam na 3.0 e vice-versa, sem alteração.
Isso é protegido pela mesma fixture binária congelada que cobre o formato desde a 2.0
(`tests/Danfma.MySheet.Tests/Fixtures/workbook-pre-namespaces.msgpack.bin`), carregada e reavaliada a
cada execução.

## Desempenho — nenhuma ação necessária

O ponto de estrangulamento de escrita permite que a 3.0 torne o índice estrutural de coluna inteira
**mantido na escrita**. Na 2.x esse índice era reconstruído a cada época de cache (em uma leitura de
range aberto após cada `InvalidateCache()`); na 3.0, a `Sheet` mantém o índice atualizado à medida que
células são escritas e removidas, então ele é construído uma vez por planilha e **sobrevive ao
`InvalidateCache()`**. O efeito visível é que o formato "carregar uma vez, depois reler uma coluna inteira
a cada época" deixa de pagar uma reconstrução do índice por época — seu custo passa a acompanhar a
coluna, não o total da planilha (veja
[Desempenho](performance.md#leituras-repetidas-de-coluna-inteira-escalam-com-a-coluna-não-com-a-planilha)).
Isso é transparente: os mesmos resultados, a mesma sequência de chamadas (ainda
`editar → InvalidateCache() → ler`), nada para mudar no seu código. A única consequência visível para o
chamador é o encapsulamento descrito acima — escritas e remoções precisam passar pelo indexador e por
`Remove` (é assim que o índice se mantém correto), exatamente como documentado.

## Comportamento

Nenhuma semântica de fórmula, parsing, resultado de avaliação ou formato de gravação mudou. A quebra fica
restrita à superfície de API de `Sheet` descrita acima. Se um upgrade 2.x → 3.0 mudar qualquer valor
calculado ou qualquer byte salvo, isso é um bug — por favor, reporte.
