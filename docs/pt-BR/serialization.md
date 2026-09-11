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
| Cache de memoização | **Não** | **Parcial** | O frio recalcula de forma preguiçosa na primeira leitura. O aquecido restaura o cache — exceto células voláteis (abaixo), que ainda são recalculadas. |

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
  MemoryPack bruto do modelo, **sem cabeçalho de container de nenhuma espécie**. Esse *formato de escrita* é
  o contrato permanente, e ele é determinístico: o mesmo modelo sempre é serializado nos mesmos bytes,
  garantido por goldens em base64 congeladas na suíte de testes
  (`CellStoreTests.Wire_IsByteIdentical_AfterNumericKeys` e
  `SheetNameInterningTests.Wire_IsByteIdentical_AfterInterning`). O que **não** é prometido é identidade de
  bytes entre *versões* da biblioteca: acrescentar um membro serializado a `Workbook` desloca todo arquivo
  salvo, porque o MemoryPack escreve a contagem de membros no cabeçalho do objeto. Isso já aconteceu duas
  vezes — `0x01` → `0x02` quando `DefinedNames` foi acrescentado (a fixture congelada
  `workbook-pre-namespaces.msgpack.bin` ainda carrega `0x01`) e `0x02` → `0x03` para o registro de tabelas
  (veja [Compatibilidade futura: o registro de
  tabelas](#compatibilidade-futura-o-registro-de-tabelas-um-terceiro-membro-de-workbook)). Os *arquivos*
  antigos continuam carregando nos dois casos; os *leitores* antigos, não.
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

Um tipo de valor em cache é deliberadamente **excluído** do snapshot e é recalculado na primeira
leitura, mesmo a partir de um arquivo aquecido:

- **Células voláteis** — qualquer coisa que tenha envolvido `NOW`/`TODAY`/`RAND`/`RANDBETWEEN` (direta ou
  transitivamente). Persisti-las "congelaria o relógio de ontem"; em vez disso, elas são reamostradas na
  próxima leitura.

O surrogate também recusa um valor **do tipo referência**, mas nenhuma célula consegue mais produzir um: a
fronteira da célula aplica a
[interseção implícita](workbook-and-expressions.md#interseção-implícita-na-fronteira-da-célula) antes de o
valor ser armazenado, então `=MyName` sobre um intervalo é persistido como seu valor intersectado, como
qualquer outro escalar. A recusa permanece como defesa em profundidade sobre a API pública
`ComputedValue.Reference`.

### Contrato de desatualização

O warm-start persiste valores que você já computou; ele não rastreia edições. O contrato pós-carregamento
é o mesmo de sempre: **após editar células, chame `InvalidateCache()`** (ou `Recalculate()` para uma
atualização apenas das voláteis) antes de ler, ou você lerá valores desatualizados. Um carregamento
aquecido apenas pula a *primeira* recomputação das células inalteradas e não voláteis — isso não muda em
nada como a invalidação funciona depois. E, assim como em um carregamento frio, as [funções
personalizadas](custom-functions.md) ainda precisam ser registradas novamente: células que **não** estavam
em cache no momento do save (ou que você invalidar) reavaliarão suas chamadas e precisarão da
implementação presente.

**O epoch de datas do Excel (3.17.0).** O formato do fio não muda: uma data é um `double` de `NumberValue` e
continua sendo isso, sem nenhuma tag nova de union. O que muda é o que um serial do início de 1900
SIGNIFICA. Da 3.17.0 em diante, o serial 1 é 1900-01-01, o serial 0 é o dia zero do Excel (1900-01-00) e o
serial 60 é o 1900-02-29 fantasma do Excel; antes da 3.17.0 o motor lia esses seriais um dia de calendário
mais cedo (o serial 1 era 1899-12-31). Os seriais a partir de 61 (1900-03-01) — toda data que uma planilha
real guarda — não são afetados. Um snapshot escrito pela 3.16.x, portanto, recarrega byte a byte igual, mas o
resultado em cache de uma fórmula sobre a janela `[0, 61)` agora está um dia deslocado: `=DATE(1900,1,1)`
guardado em cache como `2` é lido de volta como `2` até um `InvalidateCache()`. Se um workbook computa sobre
datas do início de 1900, invalide o cache uma vez depois de atualizar.

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

Versões que mudaram como um valor salvo é *interpretado* sem tocar no formato:

- **3.17.0, a época de datas** — nenhuma mudança de formato e nenhuma tag nova *para essa parte da versão*.
  Os seriais de data do início de 1900 (`[0, 61)`) mudam de SIGNIFICADO em um dia de calendário; resultados em
  cache sobre essa janela precisam de um `InvalidateCache()`. Outras partes da 3.17.0 *adicionam* tags de
  union e um membro de `Workbook` — veja as subseções de compatibilidade futura abaixo.

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

### Compatibilidade futura: o nó `AGGREGATE` (tag 322)

O `AGGREGATE` é um novo tipo de nó de expressão e toma a próxima tag append-only da union, a **322** (veja
a [Referência de funções](function-reference.md) para o que a função faz). Uma célula cuja fórmula o
chama é serializada sob essa tag.

Este é um limite de compatibilidade em **uma única direção**, como qualquer adição de tag append-only:

- Um arquivo salvo por esta versão da biblioteca **ou por uma posterior** — seja produzido por
  `Workbook.Save` ou por `ExcelFile.Load` seguido de um save — pode conter células usando a tag 322 sempre
  que a fórmula de uma célula for uma chamada de `AGGREGATE`. Esse arquivo **não pode ser aberto por uma
  versão da biblioteca anterior à que introduziu essa tag**: a union do MemoryPack mais antiga não a
  reconhece e a desserialização falha.
- Um arquivo salvo por uma versão **mais antiga** da biblioteca nunca contém essa tag e continua
  carregando sem alteração nesta e em toda versão posterior, exatamente como garante a política
  append-only acima.

**O que uma tag nova *não* quebra.** A tag é escrita por nó, não por arquivo, então um workbook que não usa
o `AGGREGATE` é serializado exatamente nos mesmos bytes de antes: as goldens binárias congeladas da suíte
de testes — o snapshot em base64 do formato de fio do armazenamento de células e a fixture pré-2.0
`.msgpack.bin` — continuam válidas e não precisam ser regeneradas. Só um novo **membro no próprio
`Workbook`** mudaria o formato de todo arquivo salvo e obrigaria a isso.

### Compatibilidade futura: os nós produtores de array dinâmico (tags 323-326)

`FILTER`, `SORT`, `UNIQUE` e `SEQUENCE` são quatro novos tipos de nó de expressão e tomam as quatro próximas
tags append-only da union, em uma única edição coordenada (veja [argumentos implícitos de
array](workbook-and-expressions.md#produtores-de-array-dinâmico) para o que elas fazem):

| Tag | Nó |
| --: | --- |
| 323 | `Lookup.Filter` |
| 324 | `Lookup.Sort` |
| 325 | `Lookup.Unique` |
| 326 | `Mathematics.Sequence` |

Uma célula cuja fórmula chama uma delas é serializada sob a tag correspondente. A próxima tag livre é a
**327**.

Este é um limite de compatibilidade em **uma única direção**, como qualquer adição de tag append-only:

- Um arquivo salvo por esta versão da biblioteca **ou por uma posterior** — seja produzido por
  `Workbook.Save` ou por `ExcelFile.Load` seguido de um save — pode conter células usando as tags 323-326
  sempre que a fórmula de uma célula chamar uma das quatro. Esse arquivo **não pode ser aberto por uma versão
  da biblioteca anterior à que introduziu essas tags**: a union do MemoryPack mais antiga não as reconhece e a
  desserialização falha.
- Um arquivo salvo por uma versão **mais antiga** da biblioteca nunca contém essas tags e continua carregando
  sem alteração nesta e em toda versão posterior, exatamente como garante a política append-only acima.
- Como no `AGGREGATE`, as tags são escritas por nó, então um workbook que não usa nenhuma das quatro é
  serializado exatamente nos mesmos bytes de antes e as goldens binárias congeladas não precisam ser
  regeneradas.

**Um snapshot de warm-start agora pode carregar o código de erro 7, `#CALC!`.** Essas quatro funções
introduzem o erro de array vazio do Excel como o oitavo código de [`Error`](computed-value.md), então o bloco
de valores de um warm-start pode guardar um `CachedCellValue` cujo `ErrorCode` é `7` — por exemplo o resultado
em cache de `=SUM(FILTER(A1:A3,A1:A3>100))`. Isso é um **valor**, não uma mudança de formato: o formato do
bloco de valores não muda e nenhuma tag está envolvida.

A degradação é graciosa nas duas direções. O que está *fixado por testes* é o mecanismo, nesta versão:

- Um código além do fim da tabela de erros é exibido como o marcador de desconhecido **`#ERR?`** em vez de
  lançar exceção — `Error.FromCode(8).Display` é `#ERR?` —, que é a regra que uma versão **anterior à 3.17**
  aplica ao código 7, cuja tabela parava no `#N/A`. Então uma versão mais antiga lendo esse snapshot mostra um
  texto de erro errado; ela não quebra e não deixa de carregar. (A versão antiga em si não é exercitada pela
  suíte; o que a suíte fixa é que a regra do marcador existe e que o 7 já não está além da tabela.)
- Um texto de erro que o motor não conhece é dobrado para `#VALUE!` em vez de lançar exceção —
  `Error.FromDisplay("#SPILL!")` é `#VALUE!` —, que é também o que torna o `#CALC!` demonstravelmente um
  código *real* agora, em vez de um que a dobra conjura.
- O surrogate de warm-start faz o round-trip exato do código 7: `CachedCellValue.ErrorCode` é `7` na ida e
  `Error.Calc` na volta.

Esses três estão fixados por `ErrorTests.Calc_IsTheEighthError_AndRoundTripsExactly` e por
`ErrorTests.AnUnknownDisplay_StillFoldsOntoValue_AndCalcIsNoLongerUnknown`. O round-trip de `.xlsx` também
fecha — o exportador escreve `#CALC!` como uma célula `t="e"` comum e o carregador o reconhece ao lado dos
outros códigos com singleton —, mas isso é sustentado pelo código (os dois switches espelhados de
canonicalização), e não por uma fixture dedicada na suíte do Excel.

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

### Compatibilidade futura: o registro de tabelas (um terceiro membro de `Workbook`)

O `Workbook` agora serializa um **terceiro** membro: o registro de tabelas que `Workbook.DefineTable`
escreve e `Workbook.Tables` expõe (veja [Workbook, planilhas e expressões →
Tabelas](workbook-and-expressions.md#tabelas)). Diferente de uma tag nova de union, um membro novo no
próprio `Workbook` muda o formato de **todo** arquivo salvo — o cabeçalho do objeto carrega a contagem de
membros —, então este limite se aplica haja ou não uma única tabela no workbook, e tanto para o arquivo
vindo de `Workbook.Save`/`SaveAsync` quanto para um save feito depois de `ExcelFile.Load`. (O
`ExcelFile.Load` ainda não preenche o registro a partir da parte `<table>` do xlsx — veja [Interop com Excel
→ Escopo e limitações](excel-interop.md#escopo-e-limitações) —, mas um workbook produzido por ele também
recebe o cabeçalho novo quando você o salva.)

Este é um limite de compatibilidade em **uma única direção**:

- **Nenhuma tag nova de union.** O registro vive no `Workbook`, não na union de expressões, e nada na
  linguagem de fórmulas lê uma tabela ainda: uma referência estruturada (`Tabela1[Valor]`) não passa pelo
  parser. O nó que representará uma delas, e a tag de union que ele tomará (a próxima tag livre é a **323** —
  as de 0 a 322 estão ocupadas), pertencem ao trabalho de semântica de referências; essa metade do limite
  ainda não existe.
- **O cabeçalho do objeto vai de `0x02` para `0x03`, e um registro vazio custa quatro bytes.** MEDIDO em
  2026-09-10 no branch que introduz o registro (MemoryPack 1.21.4, `Save` frio e descomprimido): um
  `Workbook` vazio é serializado em **13** bytes (`03` + três mapas de comprimento zero), onde a versão
  anterior escrevia **9** (`02` + dois); o byte 0 é o único byte que muda, e os quatro bytes acrescentados
  são o registro vazio no final. Num workbook populado, a golden congelada do armazenamento de células vai
  de **726 para 730** bytes, com todo byte depois do cabeçalho ainda no offset antigo — fixado por
  `CellStoreTests.Wire_NewGolden_IsPreTablesGoldenPlusEmptyTablesMember` e por seu gêmeo em
  `SheetNameInterningTests` (429 → 433 bytes).
- **Um arquivo salvo por esta versão ou por uma posterior não pode ser aberto por uma anterior.** O
  `Workbook` antigo declara dois membros e o MemoryPack recusa um cabeçalho que anuncia três. MEDIDO em
  2026-09-10, com um leitor da versão anterior sobre um arquivo escrito por esta versão (tanto um
  `MemoryPackSerializer.Deserialize` cru quanto `Workbook.Load` lançam, e tanto para um workbook vazio
  quanto para um com uma tabela): `MemoryPackSerializationException: Danfma.MySheet.Workbook property count
  is 2 but binary's header maked as 3, can't deserialize about versioning.` — "maked" é um erro de digitação
  do próprio MemoryPack, e é a busca por esse texto literal que deve levar o usuário a esta subseção. Mesma
  forma de mão única de uma tag de union acrescentada, mas aplicada a **todo** arquivo, não apenas aos que
  usam o recurso novo.
- **Arquivos salvos por uma versão mais antiga continuam carregando — para sempre, com o registro vazio.** O
  MemoryPack tolera uma contagem de cabeçalho *abaixo* da contagem de membros declarada e deixa os membros
  ausentes como `null`, e o hook `[MemoryPackOnDeserialized]` do `Workbook` transforma esse `null` num
  registro vazio (case-insensitive). O precedente já está congelado na suíte:
  `workbook-pre-namespaces.msgpack.bin` começa com `01 02 00 00 00 …` — um `Workbook` de UM membro, escrito
  antes de `DefinedNames` existir — e continua carregando
  (`MemoryPackCompatibilityTests.PreNamespaceFixture_LoadsAndReevaluates`), assim como a golden de dois
  membros (`0x02`) do armazenamento de células
  (`CellStoreTests.PreTablesGolden_StillLoads_WithEmptyTables`) e o container v2 aquecido cujo corpo
  comprimido é um modelo de dois membros
  (`ContainerVersionCompatibilityTests.GoldenV2Fixture_IsVersion2_AndLoadsForever`) — os três chegam com
  `Tables.Count == 0`.

**O que isto quebra e uma tag nova não.** A nota do `AGGREGATE` acima termina dizendo que só um membro novo
no próprio `Workbook` mudaria o formato de todo arquivo salvo e obrigaria a regenerar as goldens congeladas.
Este é esse caso: as duas goldens de fio foram regeneradas por causa dele, e *mecanicamente* — a constante
nova é `[0x03] + antigo[1..] + quatro bytes zero`, nunca um print de depuração capturado —, então as
próprias constantes registram exatamente o que se moveu.

## Quando usar cada formato

| Necessidade | Use |
| --- | --- |
| Persistência nativa rápida de um modelo calculado (estilo cache, reinícios de serviço, snapshots entre etapas de processamento) | `Workbook.Save` / `Load` |
| Intercâmbio com pessoas ou outras ferramentas (abrir no Excel, enviar um relatório) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingestão da planilha que é a fonte da verdade | [`ExcelFile.Load`](excel-interop.md) |
