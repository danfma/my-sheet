# Compressão opcional no Save (sequência do spike MessagePack)

Decisão do usuário (2026-07-03): não migrar de formato; adicionar compressão OPCIONAL sobre o MemoryPack.
Antes: medir **MemoryPack+ZStandard** (via `ZstdSharp.Port` — managed puro, portável .NET 8/10) contra o
Brotli/GZip já medidos. A escolha do codec tem um trade central: **Brotli é BCL (zero dependência nova no
core)**; ZstdSharp é dependência nova no pacote publicado — só se justifica com vitória CLARA.

## Números de referência (spike, payload grande 302k células, fração do MemoryPack cru)
MsgPack+LZ4 35,0% · MemoryPack+GZip 15,4% · **MemoryPack+Brotli 13,7%**.

## Gate de decisão — RESOLVIDO por decisão do usuário (2026-07-03)
O gate foi resolvido **fora da medição**: o usuário decidiu que **não entra dependência nova no core**. O
ZStandard nativo do .NET só chega na **versão 11** e nós targetamos **net10.0**; portanto o codec é
**Brotli** (BCL, zero-dep) independentemente de qualquer número. O gate original (abaixo, travado) fica
registrado como contexto histórico, mas nenhum dos seus braços foi aplicado — a decisão é do usuário, não da
medição.

**Medição ZstdSharp: ABORTADA por decisão do usuário** — ZStd nativo é .NET 11+, sem dependência nova no
core → Brotli direto. O pacote `ZstdSharp.Port` **não** foi adicionado ao csproj de benchmark; nada a
reverter no repositório. (Reavaliar ZStd apenas numa futura migração ao .NET 11, quando o codec for BCL.)

Gate original (travado, NÃO aplicado):
- Se ZStd (níveis padrão e alto) ficar **dentro de ~±15% do Brotli em tamanho** sem vantagem dramática de
  velocidade → **implementar Brotli** (BCL, zero-dep), sem consultar de novo.
- Se ZStd **vencer com folga** (>15% menor OU >2× mais rápido comprimindo com tamanho ≤ Brotli) → PARAR
  após a medição e reportar (dependência nova no core é gate do usuário).

## Design da feature (defaults; usuário pode vetar)
- `WorkbookSaveOptions.Compression` enum `{ None (default), Brotli }` — ortogonal ao
  `IncludeComputedValues` (frio ou warm, ambos compressíveis).
- **Mecanismo autodescritivo, não extensão**: Brotli NÃO tem magic próprio → o container `MSWM` ganha um
  byte de flags (bump de versão) e passa a ser usado também para "frio comprimido". `Load` sniffa MSWM →
  lê flags → descomprime. `Save(path)` default segue byte-idêntico (raw, teste permanente).
- **Extensão**: convenção DOCUMENTADA (recomendar sufixo `.br` aplicado PELO CHAMADOR); a lib NÃO
  auto-anexa extensão ao path (surpreendente/mágico). — Registrado para veto do usuário, que sugeriu
  "salvar com uma extensão adicional".
- Nível: `CompressionLevel.Optimal` default (medir Fastest como referência no relatório).

## Fases
1. **Medição ZStd** — **ABORTADA** (decisão do usuário; ver seção do gate). Status: Complete (cancelada).
2. **Feature Brotli** — Status: Complete. Opção + container versionado + testes + docs.

## Phase 1 Summary — Complete (medição cancelada)
Medição ZstdSharp abortada por decisão do usuário em 2026-07-03: ZStd nativo é .NET 11+ e targetamos
net10.0, então não entra dependência nova no core — codec Brotli direto. Nenhum pacote adicionado, nada
medido, nada a reverter. Os números de referência de Brotli/GZip-sobre-MemoryPack já vivem em
`plans/messagepack-spike.md` (spike), reusados na doc.

## Phase 2 Summary — Complete (feature Brotli)
Entregue na branch `feat/save-compression` (a partir do HEAD `24e5703`), um commit de feature, sem
push/amend.

**Desenho do container (decisão: bump de VERSÃO, não byte de flags).** O header de 9 bytes do warm-start
(`magic "MSWM"(4) | version(1) | modelLength int32 LE(4)`) é **preservado byte-a-byte**; só o valor da
versão e a codificação do corpo mudam. Isso é mais limpo que inserir um byte de flags, que deslocaria os
offsets e quebraria os testes de warm-start existentes (eles afirmam `version==1` e leem `modelLength` em
offset 5). Encodings:
- **v1 (warm não comprimido)** — corpo = `model || value-block`, inalterado. Arquivos warm gravados pela
  2.8.0 continuam lendo pelo mesmo caminho.
- **v2 (Brotli)** — corpo = `Brotli(model || value-block)` como **um único stream** (comprime melhor que
  dois blocos). Usado para qualquer save comprimido, frio ou warm; frio comprimido leva um value-block
  vazio. `modelLength` continua sendo o tamanho do modelo **descomprimido**, usado para fatiar model×values
  após descomprimir. "Frio comprimido" passa pelo container (Brotli cru não tem magic) e o `Load` sniffa o
  `MSWM` → versão → descomprime. Default `Save(path)` (frio, `Compression=None`) segue raw MemoryPack,
  byte-idêntico (contrato de regressão permanente intacto).

**API.** `WorkbookCompression { None, Brotli }` (novo enum) + `WorkbookSaveOptions.Compression` (default
`None`), ortogonal a `IncludeComputedValues`. `CompressionLevel.Optimal`. Sem auto-anexar extensão;
convenção `.br` documentada pelo chamador.

**Tamanhos medidos** (Brotli Optimal sobre os bytes MemoryPack de produção; M1 Pro, .NET 10 — os mesmos
números do spike, conferidos nesta execução):

| Payload | Células | MemoryPack cru | Brotli | Fração |
|---------|--------:|---------------:|-------:|-------:|
| Small   |      20 |      1.147 B   |  289 B |  ~25%  |
| Medium  |   7.500 |    348.035 B   | 33.626 B | ~10% |
| Large   | 302.048 |  7.935.568 B   | 1.090.808 B | ~14% |

**Suítes:** core **841** (833 baseline + 8 novos em `WorkbookCompressionTests`), Excel **24**, todas verdes;
`dotnet build -c Release --no-incremental` → 0 warnings / 0 errors. Fixture `workbook-pre-namespaces.
msgpack.bin` e `MemoryPackCompatibilityTests` intocados e verdes (carregam pelo ramo raw). Os testes de
warm-start existentes passam **inalterados** (o header v1 não mudou).

**Provas (`WorkbookCompressionTests`):** default `Compression=None`+frio == `Save(path)` (byte-idêntico);
frio comprimido é container v2 e round-trip; frio comprimido < 50% do raw num workbook de 5k células; warm
comprimido preserva todos os Kinds; warm comprimido serve célula cacheada com **zero recompute** (função
`TICK()` com contador em 0); voláteis excluídos re-amostram pós-load; round-trip async comprimido; container
Brotli corrompido lança `InvalidDataException`.

**Arquivos:** `Danfma.MySheet/WorkbookCompression.cs` (novo), `Danfma.MySheet/WorkbookSaveOptions.cs`
(prop `Compression`), `Danfma.MySheet/Workbook.cs` (SerializeToBytes/BuildContainer/BrotliCompress +
DeserializeContainer/BrotliDecompress versionados), `tests/Danfma.MySheet.Tests/WorkbookCompressionTests.cs`
(novo), `docs/serialization.md`, `README.md`.

_(Regras de sempre respeitadas: TDD, `--no-incremental` 0 warnings, fixture intocável, commits semantic em
inglês sem atribuição a IA, sem push/amend. `docs/pt-BR/` não tocado.)_
