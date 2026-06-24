# Memorização (cache) de valores de célula

Cache lazy por célula no `Workbook` para evitar recomputar células-folha compartilhadas em carga
read-heavy de extração em background. NÃO persiste (cache fora da serialização). Invalidação por
flush total, disparada explicitamente. **Status: F1–F4 prontos.** (F5 benchmark já medido.)
F2 = ranges roteados pelo cache (`RangeReference.ExpandValues` via `GetCellValue`, usado por
`NumericAggregation`/`ArgumentFlattening`). F4 = detecção de ciclo via `[ThreadStatic]` set de células
"em avaliação" no `GetCellValue` → `A1=B1, B1=A1` retorna `#REF!` em vez de StackOverflow.
Pendência menor: acesso por `CellAt` (INDEX/VLOOKUP/OFFSET) ainda computa direto, fora do cache.

### Implementado (149/149 verde, 0 warnings)
- `Workbook`: `[MemoryPackIgnore] ConcurrentDictionary<(string Sheet, string Id), object?> _cache` (lazy),
  `GetCellValue(sheet, id)` (TryGetValue → computa fora → grava) e `InvalidateCache()` (flush total explícito).
- `CellReference.Compute` delega a `GetCellValue` → célula referenciada por N fórmulas computa 1×.
- **Otimização**: `EvaluationContext` virou `readonly struct` → threading do contexto NÃO aloca; toda
  computação caiu para **24 B** (só o resultado boxed; era 72 B). Benchmark confirmado.
- Testes: `MemoizationTests` (computa-1× via contador; `InvalidateCache` atualiza após mutação).
- Trade-off medido: cache adiciona um lookup por referência (custo de tempo para células triviais),
  compensado por não recomputar células caras/compartilhadas (o caso de uso). `null` (blank) é valor
  cacheado válido; `ConcurrentDictionary` cobre o uso concorrente em background.

## Context
Uso: extrair/processar dados de planilha em background, leitura-pesada, mutações raras após a montagem.
Avaliação de design feita por sub-agente. Decisões do usuário:
- Estratégia: cache por célula + **FLUSH TOTAL**. Invalidação cirúrgica por dependências (grafo reverso)
  descartada — não compensa dado mutações raras + execução rápida; ranges tornariam o grafo reverso caro
  (`SUM(A1:A1000)` = 1000 arestas).
- Invalidação: **explícita** via `Workbook.InvalidateCache()` (consumidor chama após editar). Sem hook
  automático no setter — o `Sheet` não tem back-reference ao `Workbook`.

Fora de escopo: grafo reverso/invalidação cirúrgica, `EvaluationContext` separado, funções voláteis,
persistência do pré-cálculo.

## Design recomendado
- Cache vive no `Workbook` (já é o contexto que atravessa toda a recursão `Compute(workbook)`); NÃO nos
  nós `Expression` (imutáveis, compartilhados via singletons como `BlankValue.Instance`, valor depende do
  workbook). `Compute(Workbook)` fica **inalterado** → não quebra testes/benchmark (~30 call sites).
- Granularidade por CÉLULA, chave `(SheetName, Id)`. Subexpressão = baixo valor/alto custo (nós-valor são
  singletons compartilhados → corromperiam um cache por instância).
- `[MemoryPackIgnore] ConcurrentDictionary` no `Workbook` + `Workbook.GetCellValue(sheet, id)` lazy.
  Padrão `TryGetValue` → computar fora → `TryAdd` (NÃO `GetOrAdd` com factory recursiva, que reentra).
- `CellReference.Compute` delega a `GetCellValue` (cobre a recursão interna, onde está o ganho:
  célula-folha lida por várias fórmulas). `null` (blank) é valor válido cacheado — distinga
  "chave ausente = não computado" de "valor null = blank".

### Gotchas confirmados (correções às suposições iniciais)
1. `Sheet` não alcança o cache do `Workbook` → invalidação **explícita** (`InvalidateCache()`), não hook no setter.
2. Ranges não entram no cache de graça: `RangeReference.Expand` devolve `Expression` e `NumericAggregation`
   chama `.Compute` direto — precisa rotear por `GetCellValue` (ex.: um `Expand` por id/coordenada).
3. Detecção de ciclo (bônus): o marcador "em avaliação" deve ser **thread-local** (não no dicionário
   compartilhado), senão recomputação concorrente benigna vira falso ciclo.

## Phases (futuro)
- Fase 1: cache lazy + `GetCellValue` + redirecionar `CellReference.Compute`. Verif.: suíte verde;
  round-trip `WorkbookTests` intacto (cache `[MemoryPackIgnore]`); teste provando célula-folha computada 1x.
- Fase 2: cobrir ranges (`Expand` por id → `GetCellValue`). Verif.: SUM sobre range memoizado.
- Fase 3: `Workbook.InvalidateCache()` (flush total). Verif.: ler → mutar → `InvalidateCache()` → ler atualizado;
  sem invalidar → valor estável.
- Fase 4 (opcional): detecção de ciclo via stack thread-local → `ErrorValue.Reference` (`#REF!`); hoje
  `A1=B1, B1=A1` é StackOverflow.
- Fase 5 (opcional): atualizar benchmark para medir o caminho `GetCellValue`.

## Consideração: profundidade de recursão (StackOverflow) — futuro
A avaliação é recursiva (`CellReference.Compute` → `Compute` da célula referenciada). O risco NÃO são
fórmulas profundas (rasas na prática), e sim **cadeias longas de dependência entre células** — ex.: coluna
cumulativa `B2=B1+A2`, `B3=B2+A3`, … por milhares de linhas. Computar a última célula recursiona a cadeia
inteira → StackOverflow (~poucos milhares de frames na stack default de ~1MB).
- **Memorização sozinha NÃO resolve**: a primeira computação ainda recursiona toda a profundidade antes de
  o cache encher.
- **"Extrair o executor do nó" é necessário mas não suficiente**: se o avaliador externo continuar recursivo,
  o problema persiste. A solução é avaliação **iterativa (stack explícita)** ou **topológica** (avaliar
  células folha-primeiro, em ordem de dependência).
- **Ganho barato que pega carona na memorização**: preencher o cache em **ordem topológica** das dependências.
  Aí cada fórmula avalia com suas dependências já no cache → a recursão entre células fica limitada à
  profundidade (rasa) da AST de cada fórmula. Resolve a cadeia longa sem reescrever o avaliador.
- Reservar a extração completa de um avaliador iterativo para SE/QUANDO dados reais estourarem a stack.

## Arquivos-chave
- `MySheet/Workbook.cs` (cache + `GetCellValue` + `InvalidateCache`)
- `MySheet/Expressions/CellReference.cs` (delegar ao cache)
- `MySheet/Expressions/RangeReference.cs` + `NumericAggregation.cs` (rotear ranges pelo cache)
