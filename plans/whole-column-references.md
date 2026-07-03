# Referências de coluna inteira (A:A) e linha inteira (1:1)

Suportar `A:A`/`A:C` (colunas inteiras) e `1:1`/`1:5` (linhas inteiras), com a semântica "todas as células
**populadas** naquelas colunas/linhas" — mais limpa que o Excel (que tem grade fixa de 1.048.576 linhas).
Estratégia decidida por benchmark (`plans/whole-column-spike.md`): **NaiveScan no-alloc** (273 µs por scan
em 100k células, 0 alocação, 1× por época via memoização; sem estado novo). Release aditivo **2.6.0**.

## Decisões travadas (com o usuário / pelo benchmark, 2026-07-02)
- **NaiveScan no-alloc**, não Tabular (regride o hot path universal em +46–72%) nem bounds/min-máx
  (236.725× pior em dados esparsos). LazyColumnIndex fica como evolução futura condicionada (ver Apêndice).
- **Semântica = células populadas**: `SUM(A:A)` soma as células populadas da coluna A (blanks contribuem 0,
  então o resultado bate com o Excel). Modelo esparso, sem varrer linhas vazias.

## Semântica DIVERGENTE do Excel — TRAVADA (usuário, 2026-07-02): extensão populada
`ROWS(A:A)` no Excel = **1.048.576** (grade fixa). No nosso modelo gridless isso é sem sentido. **Decisão:
`ROWS`/`COLUMNS` sobre referência ilimitada usam a EXTENSÃO POPULADA** (dentro dos limites do ref):
- `ROWS(A:A)` = (maior linha populada − menor linha populada + 1); coluna vazia → 0.
- `COLUMNS(A:C)` = 3 (estrutural e exato). `COLUMNS(A:A)` = 1. `ROWS(1:5)` = 5.
- Simétrico p/ linha inteira. Divergência do Excel documentada no doc.

## Escopo — refs MISTAS INCLUÍDAS (usuário, 2026-07-02)
Suportar, além de coluna:coluna e linha:linha puras, as **abertas de um lado**: `A1:A` (col A, linha ≥ 1),
`A5:A` (col A, linha ≥ 5), `A:A10` (col A, linha ≤ 10), `A1:C` (cols A–C, linha ≥ 1). Semântica: o endpoint
ESQUERDO fornece os limites INFERIORES (colMin/rowMin), o DIREITO os SUPERIORES (colMax/rowMax); um endpoint
só-coluna não informa linha, só-linha não informa coluna → aquele eixo fica **aberto** (null) desse lado.
Não há linha < 1, então rowMin aberto ≡ 1.
- Sheet-qualified (`Data!A:A`) e absolutos (`$A:$A` — tokenizer engole `$`, normaliza pra relativo).
- **FORA (futuro ⬜)**: interseção por espaço.

## Design: UM tipo unificado (o "elegante" que o usuário pediu)
`OpenRangeReference(int? ColMin, int? ColMax, int? RowMin, int? RowMax, string SheetName) : Reference` —
null = aberto naquele lado/eixo. Cobre TODOS os casos com um só tipo, uma enumeração, um caminho de
ROWS/COLUMNS: `A:A`=(col A,A / row null,null); `1:5`=(col null,null / row 1,5); `A1:A`=(A,A / 1,null);
`A:A10`=(A,A / null,10); `A1:C`=(A,C / 1,null). (Quando os 4 limites são não-null vira um `RangeReference`
normal — o parser só produz `OpenRangeReference` quando ≥1 limite é aberto.) Tag MemoryPackUnion 316.

## For Future Agents
Marque `- [x]`; ao fechar fase: Status `Complete` + Phase Summary + rode a Verification. TDD rigoroso
(RED antes de GREEN). **Verificação SEMPRE com `dotnet build ... --no-incremental`** (lição: builds
incrementais mascaram warnings de analyzer). Fixture `workbook-pre-namespaces.msgpack.bin` e
`MemoryPackCompatibilityTests` INTOCÁVEIS. Suítes hoje: core **698**, Excel **20**; build **0 warnings**.
Commits inglês, curto + corpo, semantic (`feat(refs): ...`), SEM atribuição a IA. NÃO push (gate do usuário).

Contexto de código:
- `Parser.ParseRange` (Parsing/Parser.cs) hoje só aceita `CellReference:CellReference`; `A` sozinho vira
  `NameReference("A")`, `1` vira `NumberValue(1)`. `IsCellReference` exige letras+dígitos.
  `ParseQualifiedReference` trata `Sheet!ref`.
- `Reference` (abstract record) é a base de `CellReference`/`RangeReference`/`UnionReference`. Funções que
  consomem range vão por `ArgumentFlattening.ExpandComputedValues` → `Expand`/`ExpandComputedValues` do ref.
  Consumidores SINTÁTICOS (VLOOKUP/HLOOKUP table, INDEX, ROWS, COLUMNS, OFFSET base, AREAS, ISREF) resolvem
  a um `Reference` via `NamedReferences.TryResolveReference`.
- `CellAddress.Parse` ALOCA (substring) — NÃO usar no scan. Criar extrator no-alloc de coluna/linha a partir
  do id (varre chars da string sem alocar), conforme o benchmark (273 µs vs 1.303 µs).
- Tag MemoryPackUnion: maior atual = 315 → append 316 (`OpenRangeReference`, tipo único).

---

## Phase 1: Parser + OpenRangeReference + enumeração (o caso de agregação)
Status: Not started

- [ ] `OpenRangeReference(int? ColMin, int? ColMax, int? RowMin, int? RowMax, string SheetName) : Reference`
      (record MemoryPackable, tag 316). `Evaluate` = `#VALUE!` (como `RangeReference` bare). Normaliza
      cantos (min ≤ max quando ambos não-null).
- [ ] Extrator no-alloc em `CellAddress` (ou helper): `TryGetColumnRow(string id, out int col, out int row)`
      sem alocar substring (o benchmark: 273 µs vs 1.303 µs com substring).
- [ ] `Expand(context)` / `ExpandComputedValues(context)`: NaiveScan — iterar `sheet.Cells.Keys`, extrair
      (col,row) no-alloc, manter as células dentro dos limites NÃO-null (null = passa sempre), `yield` o
      valor/`GetCellValue`. Ordem determinística (col, depois row) para testes.
- [ ] `Parser.ParseRange` + `ParseQualifiedReference`: montar `OpenRangeReference` a partir dos endpoints.
      Cada endpoint informa o que sabe — `CellReference`→(col,row); coluna-pura (`NameReference` só-letras)→
      col; linha-pura (`NumberValue` inteiro)→row. Esquerdo dá limites inferiores, direito superiores; eixo
      não-informado por um lado fica aberto (null) desse lado. **O `:` força semântica de referência**: um
      nome só-letras adjacente a `:` é COLUNA, mesmo que exista defined name homônimo (documentar).
      `Data!A:A` → OpenRangeReference qualificado. Se TODOS os 4 limites resultarem não-null → produzir um
      `RangeReference` normal (não regredir o caminho existente).
- [ ] Testes (agregação — caminho limpo): `SUM(A:A)`, `AVERAGE(A:A)`, `COUNTA(A:A)`, `MAX(A:C)`, `SUM(1:1)`,
      `SUM(1:5)`, `SUMIF(A:A,">5")`, `SUM(Data!A:A)` cross-sheet; **mistas**: `SUM(A2:A)` (ignora A1),
      `SUM(A:A10)` (só até linha 10), `SUM(A1:C)`; coluna esparsa (A1 e A100 → 2 células, 98 vazias
      intocadas); coluna vazia → 0; `$A:$A` normaliza.

### Verification Plan
- `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` → 0 warnings.
- Suíte core verde + os novos; `MemoryPackCompatibilityTests` verde (tags append). Excel (20) intacta.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 2: Consumidores sintáticos + ROWS/COLUMNS
Status: Not started

- [ ] `ToBoundedRange(context)` em `OpenRangeReference`: resolve à bounding-box POPULADA dentro dos limites —
      cols/rows abertos preenchidos pela extensão populada (min/max das células que casam). Vazio → range
      degenerado (`RowCount`=0). Usado onde um `RangeReference` concreto é exigido.
- [ ] `NamedReferences.TryResolveReference` (ou os call-sites): `OpenRangeReference` resolve via
      `ToBoundedRange` para VLOOKUP/HLOOKUP (table), INDEX, OFFSET (base), AREAS (conta 1), ISREF (true).
      Testar `VLOOKUP(x, A:B, 2)` e `INDEX(A:A, 3)`.
- [ ] `ROWS`/`COLUMNS`: reconhecer `OpenRangeReference` e usar a EXTENSÃO POPULADA (decisão travada) em vez
      de `RangeReference.RowCount`. Testes: `ROWS(A:A)`, `COLUMNS(A:C)`, `ROWS(1:5)`, coluna/linha vazia → 0.

### Verification Plan
- Build `--no-incremental` 0 warnings; suíte verde incl. VLOOKUP/INDEX/ROWS/COLUMNS sobre coluna inteira.

### Phase Summary
_(escrever quando a fase concluir)_

---

## Phase 3: Un-parse + interop xlsx + documentação
Status: Not started

**Documentação: usar a skill `code-documentation-doc-generate`** (extrair do código → atualizar → validar
exemplos por compilação).

- [ ] `FormulaWriter.Write`: emitir `A:A`, `A:C`, `1:1`, `1:5`, as mistas (`A2:A`, `A:A10`, `A1:C`) e
      sheet-qualified fora do contexto. Corpus exaustivo de round-trip em `FormulaWriterTests`.
- [ ] Interop: `ExcelFile.Load` parseia `<f>` com `A:A` (já cai no ParseRange novo); `SaveAsExcel`/defined
      names emitem `A:A` corretamente. Teste round-trip com ClosedXML (que suporta full-column refs).
- [ ] `docs/function-reference.md` ou `workbook-and-expressions.md`: seção "Whole-column and whole-row
      references" — sintaxe, semântica "células populadas", a divergência do `ROWS`/`COLUMNS`, os limites
      (misto fora de escopo). `README.md`: bullet. NÃO tocar `docs/pt-BR/` (refresh no deploy).
- [ ] Plano: fases Complete + Phase Summary + Final Recap.

### Verification Plan
- Build `--no-incremental` 0 warnings; ambas as suítes verdes; round-trip `Parse(ToFormula(Parse("A:A")))`
  estrutural; round-trip xlsx com oráculo ClosedXML.

### Phase Summary
_(escrever quando a fase concluir)_

## Final Recap
_(escrever quando as fases 1–3 concluírem)_

## Deployment Plan
_(mesmo ritual: verificação independente minha com rebuild forçado → merge `feature/whole-column-refs` →
push → `gh workflow run release.yml` → **2.6.0** lockstep → `git pull` → refresh `docs/pt-BR/` via Sonnet.)_

---

## Apêndice — Índice estrutural write-time (pergunta do usuário, 2026-07-02)

**Pergunta:** vale criar um índice reverso atualizado no momento em que a célula é definida (coluna → células)?

**Veredito: sim, é a evolução CORRETA — e melhor que o LazyColumnIndex do benchmark — mas DEFERIDA até
haver gatilho. Não construir junto da fase 1.** Raciocínio (não precisa de benchmark novo; a decisão é
arquitetural e o spike já deu os números adjacentes):

**Por que é melhor que o LazyColumnIndex (o insight que o benchmark NÃO capturou):** um índice estrutural
(quais células existem na coluna A) é **ortogonal ao cache de valores**. `InvalidateCache()` limpa VALORES
computados, mas NÃO muda quais células estão populadas. O LazyColumnIndex, por ser read-time e amarrado à
época, é DESTRUÍDO a cada `InvalidateCache()` — o benchmark mediu 171 MB de churn no cenário
invalidate-every-1. Já o índice write-time SOBREVIVE ao `InvalidateCache()` inteiro: só muda quando células
são ADICIONADAS/REMOVIDAS. No caso âncora do MySheet (carrega em massa → seta valores de input → Invalidate
→ lê, repetido por batch), a ESTRUTURA de células populadas é ESTÁVEL entre batches — muda o valor do input,
não quais células existem. Então o índice é construído 1× (no load) e reusado por milhares de batches.
Custo de memória: o spike mediu +1,7 MB / 100k células (0,15× aditivo) — barato.

**O bloqueio (por que não agora):** `Sheet.Cells` é um `Dictionary` **público e mutável**. Um índice
write-time correto exige que TODA escrita passe por um caminho encapsulado (o indexer `sheet[id] = expr`);
uma mutação direta `sheet.Cells[...] =` corromperia o índice silenciosamente. Além disso o MemoryPack
reconstrói `Cells` no `Load` bypassando o indexer → o índice teria de ser reconstruído no
`[MemoryPackOnDeserialized]` (padrão que já usamos). Encapsular `Sheet.Cells` é uma mudança de contrato do
modelo (hoje o contrato é invalidação explícita, não change-tracking) — fazê-la por um ganho não-provado é
prematuro, e o NaiveScan a 273 µs/época já é desprezível ao lado do resto de um batch.

**Gatilho para construir (qualquer um):** (a) profiling real mostrar coluna-inteira como hot path; (b)
encapsularmos `Sheet.Cells` por OUTRO motivo (ex.: o grafo reverso de dependências para touch-por-célula de
voláteis / recálculo incremental — que exige o MESMO encapsulamento de escrita). Se (b) acontecer, o índice
de coluna vira quase de graça, sob a mesma refatoração. Recomendo esperar (b) e fazer os dois juntos.
