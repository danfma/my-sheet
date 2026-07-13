# Serialização (MemoryPack)

*Tradução do documento canônico em inglês ([serialization.md](../serialization.md)). Em caso de divergência, o inglês prevalece.*

Um `Workbook` é serializado em um formato binário compacto via
[MemoryPack](https://github.com/Cysharp/MemoryPack). Esta é a persistência *nativa* do MySheet — rápida
para escrever, rápida para carregar, e ela preserva no round-trip as árvores de expressão completas (não
apenas os valores). Ela não tem relação com `.xlsx`; para arquivos Excel, veja
[Interop com Excel](excel-interop.md).

## Salvar e carregar

```csharp
using Danfma.MySheet;

workbook.Save("model.mysheet");
Workbook restored = Workbook.Load("model.mysheet");

// Sobrecargas assíncronas:
await workbook.SaveAsync("model.mysheet", cancellationToken);
Workbook restoredAsync = await Workbook.LoadAsync("model.mysheet", cancellationToken);
```

`Load`/`LoadAsync` lançam `InvalidDataException` se o arquivo não contiver um workbook. A extensão do
arquivo é escolha sua — os exemplos usam `.mysheet` por convenção.

## Opções de salvamento

As sobrecargas `Save(path, WorkbookSaveOptions)` / `SaveAsync` recebem dois switches **ortogonais**. O
`Load` não precisa de nenhuma flag correspondente — ele detecta o formato (bruto vs. container,
descomprimido vs. Brotli) a partir do cabeçalho do arquivo.

| Opção | Tipo | Padrão | Efeito |
| --- | --- | --- | --- |
| [`IncludeComputedValues`](#warm-start-persistindo-valores-computados) | `bool` | `false` | Persiste o cache de memoização junto com o modelo, de modo que um carregamento comece **aquecido** (pula a recomputação). |
| [`Compression`](#compressão) | `WorkbookCompression` | `None` | `Brotli` reduz o arquivo usando o Brotli da BCL. |
| [`CompressionLevel`](#compressão) | `CompressionLevel` | `Optimal` | Qualidade do Brotli ao comprimir. `Fastest` reduz sensivelmente o tempo de salvamento em workbooks grandes em troca de um arquivo maior; é um knob só de escrita — `Load` lê qualquer nível. |

Com os dois primeiros em seus padrões, `Save(path, options)` é byte a byte idêntico a `Save(path)`.

## O que é preservado no round-trip — e o que não é

A coluna **frio** é o `Save` padrão; a coluna **aquecido** é um save com
[`IncludeComputedValues`](#warm-start-persistindo-valores-computados) (tudo o que um save frio persiste,
mais os valores memoizados).

| | Frio | Aquecido | Observações |
| --- | --- | --- | --- |
| Planilhas (nome, ordem das abas) | Sim | Sim | A busca de nomes case-insensitive é restaurada na desserialização. |
| Células e árvores de expressão completas | Sim | Sim | Fórmulas continuam sendo fórmulas — um workbook carregado continua recalculando. |
| **Chamadas** de funções personalizadas (nós `FunctionCall`) | Sim | Sim | O nome e as expressões dos argumentos são preservados. |
| **Implementações** de funções personalizadas (delegates) | **Não** | **Não** | Comportamento é código, não dados — registre de novo após carregar. |
| Cache de memoização | **Não** | **Parcial** | O frio recalcula de forma preguiçosa na primeira leitura. O aquecido restaura o cache — exceto células voláteis e do tipo referência (abaixo), que ainda são recalculadas. |

A consequência prática: se o seu workbook usa [funções personalizadas](custom-functions.md), registre-as
novamente após cada `Load`, ou essas chamadas serão avaliadas como `#NAME?`:

```csharp
var restored = Workbook.Load("model.mysheet");

restored.RegisterFunction("CUSTOM", (arguments, wb) =>
{
    var a = arguments[0].Evaluate(wb).AsDouble() ?? 0;
    var b = arguments[1].Evaluate(wb).AsDouble() ?? 0;

    return a + b;
});

double value = restored.GetCellValue("Sheet1", "A1").ToDouble();
```

## Warm-start: persistindo valores computados

Por padrão, um arquivo salvo contém **apenas o modelo** — todo valor é recalculado de forma preguiçosa na
primeira leitura após o carregamento. Passe `WorkbookSaveOptions { IncludeComputedValues = true }` para
também persistir o cache de memoização, de modo que o carregamento comece **aquecido** — o **warm-start**
(inicialização aquecida) — e sirva células já computadas sem reavaliá-las:

```csharp
workbook.Save("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true });
// await workbook.SaveAsync("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true }, ct);

var warm = Workbook.Load("model.mysheet"); // lida de volta com o cache já preenchido
```

`Load`/`LoadAsync` não precisam de nenhuma flag — eles detectam o formato a partir do cabeçalho do arquivo.

### Formato do arquivo

- **Frio, descomprimido** (`Save(path)`, ou `IncludeComputedValues = false` com `Compression = None`) — o
  MemoryPack bruto do modelo, byte a byte idêntico a toda versão anterior. Este é um contrato permanente,
  garantido por um teste de regressão.
- **Container** — qualquer outra combinação é um pequeno container autodescritivo: o número mágico `MSWM`,
  1 byte de versão do formato, o tamanho do modelo descomprimido (int32 LE), e então o corpo. O `Load`
  inspeciona os 4 bytes do número mágico: uma correspondência indica um container; qualquer outra coisa é
  um modelo bruto (frio ou pré-existente), de modo que arquivos antigos continuam carregando sem
  alteração. O byte de versão seleciona a codificação do corpo:
  - **v1 (aquecido descomprimido)** — os **mesmos** bytes do modelo que um save frio escreveria, seguidos
    por um bloco de valores (o MemoryPack dos valores em cache). Arquivos de warm-start escritos antes de
    a compressão existir são exatamente isso.
  - **v2 (Brotli, legado, somente leitura)** — o modelo e o bloco de valores concatenados e comprimidos com
    Brotli como um *único* stream, escrito como duas chamadas `BrotliStream.Write` de buffer inteiro. Nada
    na biblioteca escreve mais v2 (substituído pelo v3, abaixo) — arquivos v2 antigos continuam carregando
    sem alteração, para sempre, a mesma política de mão única de uma tag de union append-only.
  - **v3 (Brotli, em chunks — o padrão desde a versão atual)** — o modelo e o bloco de valores concatenados
    e comprimidos com Brotli como um único stream, escrito como uma sequência de **escritas de exatamente
    64KB** (a última mais curta) no `BrotliStream`, em vez de duas escritas de buffer inteiro. Esse
    chunking fixo é parte da DEFINIÇÃO do formato, não um detalhe de implementação: todo mecanismo de
    escrita (`WorkbookIoBuffering.Pooled` ou `.Pipelines`, `Save` ou `SaveAsync`) reagrupa os bytes que
    produz nessas mesmas janelas de 64KB antes de chegarem ao Brotli, então as quatro combinações produzem
    arquivos v3 byte a byte idênticos — o v2 não conseguia fazer essa promessa (dividir os mesmos bytes em
    várias chamadas `Write` pequenas muda mensuravelmente a saída comprimida do Brotli, então todo escritor
    v2 precisava recorrer a materializar o modelo e o bloco de valores como dois buffers inteiros para se
    manter byte a byte idêntico). Veja [Tamanhos medidos](#tamanhos-medidos) para o ganho de taxa de
    compressão que esse esquema de chunking trouxe num workbook real grande.

Como o modelo e seus valores viajam em um único arquivo, eles nunca podem dessincronizar no carregamento.

### O que o warm-start *não* congela

Dois tipos de valor em cache são deliberadamente **excluídos** do snapshot e são recalculados na primeira
leitura, mesmo a partir de um arquivo aquecido:

- **Células voláteis** — qualquer coisa que tenha envolvido `NOW`/`TODAY`/`RAND`/`RANDBETWEEN` (direta ou
  transitivamente). Persisti-las "congelaria o relógio de ontem"; em vez disso, elas são reamostradas na
  próxima leitura.
- **Resultados do tipo referência** — raros como valor final de célula e baratos de reconstruir.

### Contrato de desatualização

O warm-start persiste valores que você já computou; ele não rastreia edições. O contrato pós-carregamento
é o mesmo de sempre: **após editar células, chame `InvalidateCache()`** (ou `Recalculate()` para uma
atualização apenas das voláteis) antes de ler, ou você lerá valores desatualizados. Um carregamento
aquecido apenas pula a *primeira* recomputação das células inalteradas e não voláteis — isso não muda em
nada como a invalidação funciona depois. E, assim como em um carregamento frio, as [funções
personalizadas](custom-functions.md) ainda precisam ser registradas novamente: células que **não** estavam
em cache no momento do save (ou que você invalidar) reavaliarão suas chamadas e precisarão da
implementação presente.

## Compressão

O MemoryPack otimiza para velocidade, então seu layout é de largura fixa e redundante — o que significa
que ele comprime extremamente bem. Passe `WorkbookCompression.Brotli` para reduzir o arquivo salvo com o
Brotli da BCL (`CompressionLevel.Optimal`); nenhuma dependência de terceiros é adicionada.

```csharp
workbook.Save("model.mysheet.br", new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli });

var restored = Workbook.Load("model.mysheet.br"); // detecta e descomprime de forma transparente
```

A compressão é ortogonal ao warm-start — combine as duas para persistir um cache aquecido em um arquivo
comprimido:

```csharp
workbook.Save("model.mysheet.br", new WorkbookSaveOptions
{
    IncludeComputedValues = true,
    Compression = WorkbookCompression.Brotli,
});
```

### Tamanhos medidos

Brotli no nível `Optimal` sobre os bytes de MemoryPack de produção, três workbooks representativos (Apple
M1 Pro, .NET 10). As porcentagens são o tamanho comprimido como fração do arquivo MemoryPack bruto:

| Workbook | Células | MemoryPack bruto | Brotli | Fração |
| --- | ---: | ---: | ---: | ---: |
| Pequeno (estilo fixture) | 20 | 1.147 B | 289 B | ~25% |
| Médio (valores + fórmulas) | 7.500 | 348.035 B | 33.626 B | ~10% |
| Grande (modelo de coluna inteira) | 302.048 | 7.935.568 B | 1.090.808 B | ~14% |

Quanto maior e mais repetitivo o modelo, maior o ganho — um workbook real tipicamente cai para bem menos
da metade do seu tamanho bruto. A compressão troca CPU no momento do save/load por esse espaço; deixe em
`None` quando você salva com frequência em um disco local rápido e o tamanho do arquivo não é uma
preocupação.

O chunking fixo de 64KB do v3 (veja [Formato do arquivo](#formato-do-arquivo)) não é só um contrato de
determinismo/identidade de bytes — ele também comprime *melhor* do que a escrita de buffer inteiro do v2.
Num workbook real grande, rico em fórmulas (~680 mil células, `CompressionLevel.Optimal`): o v2 produziu um
arquivo de 3.616.055 bytes; o mesmo workbook pelo v3 produziu 3.137.719 bytes — **~13% menor**, de graça,
só por escrever o stream comprimido em janelas fixas em vez de de uma vez só.

### Convenção de nomenclatura de arquivo

A biblioteca **nunca** renomeia o arquivo que você passa — um save comprimido escreve exatamente o caminho
que você fornecer, sem nenhuma extensão anexada. Como o container `MSWM` é autodescritivo, o `Load` não
depende do nome para decidir se deve descomprimir. Se você quiser que arquivos comprimidos sejam
reconhecíveis, adote uma convenção de sufixo em seu próprio código (um sufixo `.br`, como nos exemplos
acima, é a escolha comum).

## Compatibilidade

Os nós de expressão são serializados como uma union do MemoryPack, e as tags da union são **append-only
por política do projeto**: tags existentes nunca são renumeradas, reordenadas ou reutilizadas, e novos
tipos de nó recebem tags novas. Workbooks salvos por uma versão mais antiga permanecem, portanto,
carregáveis por versões mais novas da biblioteca.

Como apenas as tags (nunca os nomes de tipo) vão para o fio, a [reorganização de namespaces da
2.0](migrating-to-2.0.md) não mudou o formato em absolutamente nada: arquivos salvos pela 1.x carregam
na 2.0 sem alteração, garantido por uma fixture binária pré-2.0 congelada na suíte de testes.

### Compatibilidade futura: nós de delta de fórmula compartilhada (tags 319-321)

Fórmulas compartilhadas (fórmulas do Excel arrastadas) agora podem ser representadas por três tipos de nó
adicionais — `AnchoredCellReference` (319), `AnchoredRangeReference` (320) e `SharedFormulaSlave` (321) —
que permitem que toda célula escrava de um grupo suportado compartilhe uma única árvore de expressão mestre
em vez de manter uma árvore própria totalmente expandida (veja [Interop com Excel → Fórmulas
compartilhadas](excel-interop.md#fórmulas-compartilhadas-uma-árvore-mestre-compartilhada-com-deltas-por-escrava)
para o que torna um grupo "suportado" e o ganho de tempo de carregamento medido).

Este é um limite de compatibilidade em **uma única direção**, como qualquer adição de tag append-only:

- Um arquivo salvo por esta versão da biblioteca **ou por uma posterior** — seja produzido por
  `Workbook.Save` ou por `ExcelFile.Load` seguido de um save — pode conter células usando as tags 319-321
  sempre que o workbook tiver um grupo de fórmula compartilhada suportado. Esse arquivo **não pode ser
  aberto por uma versão da biblioteca anterior à que introduziu essas tags**: a union do MemoryPack mais
  antiga não as reconhece e a desserialização falha.
- Um arquivo salvo por uma versão **mais antiga** da biblioteca nunca contém essas tags e continua
  carregando sem alteração nesta e em toda versão posterior, exatamente como garante a política
  append-only acima.

**Nota honesta: isto é uma otimização de RAM/GC, não de tamanho em disco.** Em memória, toda escrava de um
grupo suportado compartilha uma única instância de `Expression` para sua árvore mestre — é daí que vem o
ganho de alocação e de GC. No fio, o MemoryPack serializa os dados de cada nó de forma independente e
**não** faz rastreamento de referências (reference-tracking) nem deduplicação estrutural: um
`SharedFormulaSlave` ainda escreve sua própria cópia dos bytes serializados da árvore mestre, uma vez por
escrava. Um workbook com um grande grupo de fórmula compartilhada, portanto, não diminui em disco só por
causa desta mudança — apenas sua pegada em memória após o carregamento diminui.

### Compatibilidade futura: container v3 (Brotli em chunks)

`WorkbookCompression.Brotli` agora escreve o container versão 3 por padrão (veja [Formato do
arquivo](#formato-do-arquivo)) em vez do v2. Este também é um limite de **mão única**:

- Um arquivo salvo por esta versão da biblioteca **ou por uma posterior** com `Compression = Brotli` é um
  container v3. Esse arquivo **não pode ser aberto por uma versão da biblioteca anterior à que introduziu
  o v3**: o switch de versão de container do leitor mais antigo não reconhece a tag 3, e `Load` lança
  `InvalidDataException`.
- Um arquivo salvo por uma versão **mais antiga** da biblioteca (v1 ou v2) continua carregando sem
  alteração nesta e em toda versão posterior — o `Load` ainda reconhece e descomprime containers v2, só
  não os escreve mais.

Não há opção para voltar a escrever v2 — a mesma política em espírito append-only das tags acima: um
formato de escrita substituído é descartado, não mantido como um knob, enquanto o leitor o mantém para
sempre.

## Quando usar cada formato

| Necessidade | Use |
| --- | --- |
| Persistência nativa rápida de um modelo calculado (estilo cache, reinícios de serviço, snapshots entre etapas de processamento) | `Workbook.Save` / `Load` |
| Intercâmbio com pessoas ou outras ferramentas (abrir no Excel, enviar um relatório) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingestão da planilha que é a fonte da verdade | [`ExcelFile.Load`](excel-interop.md) |
