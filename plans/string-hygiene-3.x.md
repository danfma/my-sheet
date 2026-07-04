# Higiene de strings 3.x — interning de SheetName e chaves numéricas do _cells

Capturar os dois levers NÃO-breaking do veredito do spike v4 (~59MB do modelo residente K1), com a
condição do dono: mudança pequena/controlada, e **o Load coberto** (interning que não sobrevive à
desserialização não vale — lembrete explícito do dono, 2026-07-04).

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-04** ("Se não for uma alteração grande...
> lembre-se do Load tb"). Fase 1 é pequena e sai; Fase 2 tem gate explícito de PARE-se-invasivo.

## Evidência (spike v4, `--v4-resident`/`--v4-hotpath`, verificado)
- `CellReference.SheetName`: 800k instâncias de "Data" no modelo K1 ≈ **24,4MB** → interned = 1 instância.
- `_cells` `Dictionary<string,Expression>`: 52,9MB → `Dictionary<(int,int),Expression>` 18,0MB =
  **34,9MB (2,94×)**.
- `SetCell` int-keyed medido NÃO é mais rápido (17,6 string vs 20,4 int ms @ 600k) — o lever é MEMÓRIA.
- AST numérica (breaking) refutada: capturaria só ~24MB extras (~6%).

## Restrições duras
1. **Wire format INTOCADO**: `_cells` é o membro serializado nº 3 do `Sheet` (a fixture
   `workbook-pre-namespaces.msgpack.bin` + `MemoryPackCompatibilityTests` são O JUIZ). A troca da chave
   em memória exige serialização wire-preserving (formatter/surrogate que grava/lê o mapa string-keyed
   idêntico, convertendo nas fronteiras: `TryGetColumnRow` no Load ~6ms/600k; `CellAddress.ToId` no
   Save). `SheetName` idem: o wire continua string; o interning acontece no formatter de LEITURA.
2. **Load coberto nos dois levers** (o lembrete do dono): pós-`Load`, SheetName compartilha instância e
   `_cells` está int-keyed — testes de round-trip provam ambos.
3. **Ids não-A1**: o indexer público do `Sheet` aceita qualquer string — chaves que não parseiam como A1
   vão para um overflow string-keyed (mesmo padrão do dense store), preservando comportamento e
   serialização.
4. **Superfície pública do `Sheet` inalterada**: enumeração continua `KeyValuePair<string,Expression>`
   (id derivado na enumeração — caminho frio), `Keys`/`TryGetValue`/`ContainsKey`/indexer por string.
5. Interning: pool próprio pequeno (nomes de sheet são dezenas) OU `InternStringFormatter` do MemoryPack
   se existir/couber — investigar; NUNCA andar na AST pós-Load reconstruindo nós.

## For Future Agents
TDD; build `--no-incremental` antes de `--no-build`; baseline: core **934**, Excel **24**; TUnit/.NET 10.
Gates permanentes: fixture byte-intocada E VERDE, round-trip novo→velho comportamentalmente idêntico,
`--k1-endtoend` agregado 27143285713, harnesses nas bandas. Git append-only, sem push, branch nomeada;
commits `perf(...)`/`fix(...)` (patch — SEM api nova; se criar API, avisar o orquestrador ANTES: tipo de
commit define a versão, lição). SEM atribuição a IA. Verificação e integração em blocos separados.

## Phase 1: Interning de SheetName (parse + Load)
Status: Complete
- [x] Parse: qualificador `Nome!` resolve para instância do pool (referência local já reusa `sheet.Name`
      — estender a cortesia às cross-sheet; pool por parse/global pequeno, design justificado).
- [x] Load: formatter de leitura interna o `SheetName` dos nós de referência (wire inalterado);
      investigar `InternStringFormatter` nativo vs formatter próprio com pool (String.Intern é
      process-lifetime — aceitável para nomes de sheet, mas justifique a escolha).
- [x] Testes: pós-parse, N referências cross-sheet compartilham a MESMA instância (ReferenceEquals);
      pós-Load idem; fixture verde; probe `--v4-resident` re-rodado mostra o colapso (~24MB @ K1-shape).
### Verification Plan
- ReferenceEquals nos dois caminhos; fixture byte-intocada; suítes verdes; probe antes/depois.
### Phase Summary
Entregue em `perf/sheetname-interning` (`a92a39f`, merged; +21/−4 de produção). Parse: `string.Intern`
no ponto único de entrada do qualificador (`ParseQualifiedReference`). Load: **`[InternStringFormatter]`
NATIVO do MemoryPack por membro** nos três records de referência — decompilado pelo agente para provar:
Serialize = `WriteString` idêntico ao default (wire byte-igual), Deserialize = `ReadString`+`Intern`;
parse e Load convergem na MESMA instância (`ReferenceEquals` provado). Construtor rejeitado com razão
(o ctor de `RangeReference` está no hot path do `ToBoundedRange`). **Caixa: exata/Ordinal, sem
canonicalizar** — o `FormulaWriter` ecoa o texto verbatim; canonicalizar mudaria o round-trip textual
(teste trava `data`≠`DATA` distintos + resolução case-insensitive intacta). **Colapso medido MAIOR que
o previsto: ~56MB no modelo K1** (421→365MB; o probe subestimava 2× — string de 4 chars retém 64B, não
32B; lever direto ~49MB). Golden de wire byte-idêntico (429 bytes) embutido como teste. Verificação
independente: core **942** (934+8), Excel 24, fixture verde, 0 warnings, K1 agregado idêntico.

## Phase 2: `_cells` com chave `(int,int)` wire-preserving
Status: Not started
- [ ] `Sheet._cells` int-keyed em memória + overflow não-A1; serialização wire-preserving (formatter/
      surrogate gravando o mapa string-keyed byte-idêntico — a fixture decide); superfície pública
      derivando strings (enumeração fria).
- [ ] Consumidores internos (`SetCell`/`Remove`/indexer/`GetCellValueDense` miss/estruturas) adaptados —
      o miss do dense deixa de materializar id para o lookup (bônus).
- [ ] **GATE PARE-SE-INVASIVO**: se a serialização wire-preserving exigir gambiarra frágil OU o diff
      de produção passar de ~500 linhas, PARAR e reportar análise (o lever é 35MB de memória; não vale
      fragilidade de schema).
- [ ] Testes: fixture byte-intocada e verde; round-trip salva bytes IDÊNTICOS aos de antes da mudança
      (workbook igual → arquivo igual); ids não-A1 round-trip; probe re-rodado (~35MB @ K1-shape).
### Verification Plan
- Fixture + bytes de save idênticos; suítes verdes; k1 agregado idêntico; probe antes/depois.
### Phase Summary
_(write when phase completes)_

## Phase 3: Release (patch) + refresh pt-BR
Status: Not started
- [ ] Ritual completo (merge-base em chamada separada; versionize deriva patch dos `perf:`); nota curta
      em `performance.md`; refresh pt-BR.
### Verification Plan
- Release publicado; espelho em paridade; re-run do dono quando quiser.
### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
