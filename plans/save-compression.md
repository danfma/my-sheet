# Compressão opcional no Save (sequência do spike MessagePack)

Decisão do usuário (2026-07-03): não migrar de formato; adicionar compressão OPCIONAL sobre o MemoryPack.
Antes: medir **MemoryPack+ZStandard** (via `ZstdSharp.Port` — managed puro, portável .NET 8/10) contra o
Brotli/GZip já medidos. A escolha do codec tem um trade central: **Brotli é BCL (zero dependência nova no
core)**; ZstdSharp é dependência nova no pacote publicado — só se justifica com vitória CLARA.

## Números de referência (spike, payload grande 302k células, fração do MemoryPack cru)
MsgPack+LZ4 35,0% · MemoryPack+GZip 15,4% · **MemoryPack+Brotli 13,7%**.

## Gate de decisão (travado)
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
1. **Medição ZStd** (ZstdSharp SÓ no csproj de benchmark): tamanho+velocidade nos 3 payloads, níveis 3 e
   ~19, tabela no plano ao lado dos números do spike. Aplicar o gate.
2. **Feature Brotli** (se o gate liberar): opção + container flags + testes (round-trip frio/warm
   comprimido; default intacto byte-idêntico; legado raw carrega; fixture verde) + docs
   (`serialization.md` via skill de docs) — release minor.

_(Resultados e Phase Summaries ao concluir; mesmas regras de sempre: TDD, `--no-incremental`, fixture
intocável, commits semantic, sem push/amend.)_
