# Spike: MessagePack (indexed keys) como formato de serialização — validar ANTES de decidir

Hipótese do usuário (2026-07-03): MemoryPack é rapidíssimo mas gera arquivos maiores; MessagePack-CSharp
com **chaves inteiras indexadas** (`[Key(n)]`), source-gen e `[Union]` polimórfico (mesma lógica de tags
append-only que já usamos) pode reduzir tamanho mantendo velocidade aceitável. **Quebra o formato** →
spike primeiro, números na mesa, usuário decide o merge (seria 3.0 com leitura dual-format p/ migração).

## O que o spike DEVE responder (critérios de decisão)
1. **Tamanho**: workbooks representativos (pequeno; médio com fórmulas de todas as categorias; grande
   ~100k células) serializados em MemoryPack vs MessagePack indexed vs MessagePack+LZ4
   (`MessagePackCompression.Lz4BlockArray`). Tabela de bytes + %.
2. **Velocidade**: serialize/deserialize (BenchmarkDotNet, MemoryDiagnoser) nos mesmos payloads.
   Quanto perdemos vs MemoryPack? (Hipótese honesta: MemoryPack vence em CPU; a questão é a margem.)
3. **Viabilidade técnica** (prototipar, não migrar): os ~320 tipos do union — `[Union(tag, typeof(...))]`
   na base `Expression` (gerável por script a partir das tags MemoryPackUnion — MESMOS números);
   records posicionais C# (construtor + `[Key]`) funcionam com o source generator do MessagePack v3?
   `[MemoryPackOnDeserialized]` → `IMessagePackSerializationCallbackReceiver`; dicionários com comparer
   (Sheets case-insensitive) → como restaurar. Prototipar num SUBCONJUNTO representativo (~30 nós cobrindo
   o workbook de teste) + estimar o esforço da migração completa.
4. **Migração/compat**: estratégia dual-read (sniff de prefixo: LZ4/msgpack vs MemoryPack) para 3.0 ler
   arquivos 2.x; custo de manter os dois writers OU só reader legado.
5. **Interação com o warm-start** (plans/warm-start-values.md): o container MSWM é agnóstico de formato —
   confirmar que a troca do bloco interno não muda o desenho.

## Regras do spike
- Branch `spike/messagepack-format`; protótipo em `benchmarks/.../Spike/MessagePackFormat/` + tipos
  espelho — **ZERO mudança em código de produção**. Sem push. Relatório com números REAIS no próprio
  plano (tabelas), recomendação fundamentada e estimativa de esforço para a migração completa.
- Lições aplicáveis: medir, não extrapolar; verificar capability da lib antes (records/source-gen);
  commits semantic; sem amend.

## Resultado

Spike executado na branch `spike/messagepack-format` (2026-07-03). Zero mudança em código de produção: todo
o protótipo vive em `benchmarks/Danfma.MySheet.Benchmark/Spike/MessagePackFormat/` (33 nós espelho + Workbook/
Sheet espelho + conversor + harness de bytes + benchmarks). MessagePack-CSharp **3.1.7** adicionado só ao
`.csproj` do benchmark. Máquina: Apple M1 Pro, .NET SDK 10.0.301 / runtime 10.0.9 (Arm64 RyuJIT).

### 1+3. Viabilidade técnica (capability confirmada ANTES de medir)

Probe isolado (`scratchpad`, descartado) e depois o protótipo real confirmam no MessagePack v3 + .NET 10:

- **Records posicionais + `[property: Key(n)]` inteiros funcionam** (`[MessagePackObject]` no record, `Key`
  em cada parâmetro do construtor). Sem pegadinha — não precisou de classe espelho por incompatibilidade,
  os espelhos existem só para não tocar produção.
- **`[Union(tag, typeof(...))]` polimórfico funciona** com **os MESMOS números** do `[MemoryPackUnion]`
  (reusei tags 0,1,2,…,316 do `Expression.cs` no subconjunto). A tabela de `[Union]` é gerável por script a
  partir das tags MemoryPack existentes — migração append-only preservada.
- **`[MemoryPackOnDeserialized]` → `IMessagePackSerializationCallbackReceiver.OnAfterDeserialize`**: restaurei
  os comparers `OrdinalIgnoreCase` de `Sheets`/`DefinedNames` exatamente como o `RestoreComparers` de hoje.
  Dicionário case-insensitive volta correto (verificado: lookup `a1` acha `A1`).
- **Runtime resolver E source-gen resolver funcionam** no .NET 10. O caminho default (dynamic resolver,
  Reflection.Emit) roda direto; o `[GeneratedMessagePackResolver]` (AOT-safe) também roundtripa. Para NativeAOT
  usar-se-ia o source-gen; nas medições abaixo usei o default.
- Round-trip verificado célula-a-célula nos 3 payloads (contagem + tipo de cada nó reconstruído) — os números
  de tamanho descrevem uma codificação FIEL, não lossy.

Subconjunto medido: 33 dos ~320 tipos (valores, referências Cell/Range/Open/Union/Name, Binary/Unary, SUM/
AVERAGE/MIN/MAX/COUNT, IF/IFERROR/LET, VLOOKUP/MATCH/INDEX/CHOOSE, COUNTIF/SUMIF/AVERAGEIF/SMALL, CONCAT/UPPER/
LEFT, ROUND/ABS, FunctionCall). O conversor lança em nó não-mapeado, então os payloads não escapam do subconjunto.

### 1. Tamanho (bytes) — `dotnet run -c Release -- --messagepack-size`

| Payload | Células | MemoryPack | MsgPack indexed | MsgPack+LZ4 | MsgPack+GZip | **MemoryPack+GZip** | **MemoryPack+Brotli** |
|---------|--------:|-----------:|----------------:|------------:|-------------:|--------------------:|----------------------:|
| Small   |      20 |    1.147 B |   587 B (51,2%) | 359 (31,3%) |  276 (24,1%) |         308 (26,9%) |       **289 (25,2%)** |
| Medium  |   7.500 |  348.035 B | 190.835 (54,8%) | 75,6k (21,7%)| 44,6k (12,8%)|       44,2k (12,7%) |      **33,6k (9,7%)** |
| Large   | 302.048 | 7.935.568 B| 5.559.160 (70,1%)| 2,78M (35,0%)| 1,54M (19,4%)|       1,23M (15,4%) |     **1,09M (13,7%)** |

Percentuais = fração do MemoryPack cru.

**Leitura honesta do tamanho:**
- **RAW (sem compressão), MessagePack indexed ganha de verdade: 51–70% do MemoryPack.** MemoryPack é layout
  fixo (double = 8 B sempre, sem varint); MessagePack usa varint/small-int e strings compactas. É um ganho
  real SE o arquivo é gravado/transmitido sem compressão.
- **Com QUALQUER compressor genérico, a vantagem inverte.** `MemoryPack+Brotli` é o MENOR em todos os payloads
  (13,7% no grande) e bate `MsgPack+LZ4` (35%) por ~2,5×, e até `MsgPack+GZip` (19,4%). Motivo: LZ4 é
  compressor fraco/rápido; o layout redundante do MemoryPack "some" sob um entropy coder (Brotli/GZip) e ainda
  fica menor que MessagePack+LZ4. `MemoryPack+GZip` (15,4%) também ganha do MsgPack+LZ4.
- **Tamanho importa para quê?** Arquivo em disco / payload de rede. Nesses cenários você quase sempre já quer
  compressão — e aí a resposta certa é **uma linha de `BrotliStream`/`GZipStream` sobre os bytes MemoryPack que
  já produzimos**, não quebrar o formato. O único cenário onde MessagePack-indexed cru vence é "gravar sem
  compressão e otimizar o cru" — nicho.

### 2. Velocidade — BenchmarkDotNet ShortRun + MemoryDiagnoser (`--filter *MessagePackBenchmarks*`)

Ratio = relativo ao MemoryPack (baseline 1,00). Alloc Ratio idem. (ShortRun em macOS: barras de erro largas em
alguns pontos — ler a tendência central, não o dígito.)

**Serialize**

| Payload | MemoryPack (base) | MessagePack indexed | MessagePack + LZ4 |
|---------|------------------:|--------------------:|------------------:|
| Small   | 2,18 µs / 2.200 B | 2,27 µs (1,04×) / 616 B (0,28×) | 4,03 µs (1,85×) / 384 B (0,17×) |
| Medium  | 665 µs / 698 KB   | 772 µs (1,16×) / 191 KB (0,27×) | 894 µs (1,34×) / 75,6 KB (0,11×)|
| Large   | 11,0 ms / 7,94 MB | 12,1 ms (1,10×) / 5,56 MB (0,70×) | 23,4 ms (2,12×) / 2,78 MB (0,35×)|

**Deserialize**

| Payload | MemoryPack (base) | MessagePack indexed | MessagePack + LZ4 |
|---------|------------------:|--------------------:|------------------:|
| Small   | 4,15 µs / 8.368 B | 3,70 µs (0,90×) / 6.208 B (0,74×) | 4,22 µs (1,02×) / 6.208 B |
| Medium  | 1,94 ms / 1,40 MB | 1,99 ms (1,03×) / 1,40 MB (1,00×) | 2,08 ms (1,07×) / 1,40 MB |
| Large   | 63,2 ms / 32,6 MB | 62,6 ms (0,99×) / 32,6 MB (1,00×) | 80,3 ms (1,27×) / 32,6 MB |

**Leitura honesta da velocidade** (hipótese "MemoryPack vence em CPU; a questão é a margem" — CONFIRMADA, margem
modesta):
- **Serialize: MessagePack indexed é 4–16% mais lento**, mas aloca 3,7×–1,4× menos (0,27–0,70×). Nada de ordem
  de grandeza.
- **Deserialize: praticamente empate** (0,90–1,03×) — no small MessagePack chega a ganhar.
- **LZ4 cobra CPU real**: até 2,12× no serialize grande e 1,27× no deserialize grande. LZ4 troca ~65% de bytes
  por ~2× de tempo de escrita — mau negócio vs Brotli-sobre-MemoryPack, que dá 86% de redução.
- A conversão Expression→espelho NÃO está no número quente (uma migração real serializaria os tipos de produção
  direto). Medi formato-vs-formato, não overhead de conversor.

### 4. Migração / compat (dual-read para 3.0)

Estratégia certa NÃO é sniff de MemoryPack-vs-MessagePack (frágil: distribuições de byte inicial se sobrepõem).
O **plano warm-start já define o container `MSWM` (magic + byte de versão)** — `Load`/`LoadAsync` já vão diferenciar
por prefixo. Trocar o bloco interno de MemoryPack→MessagePack vira **um bump de versão no header**, não um novo
mecanismo. Frio legado (raw MemoryPack, sem magic) continua lido pelo reader legado. Custo dual-read: manter o
**reader** MemoryPack legado (leve — o writer novo emite só o formato novo). O container é agnóstico de formato,
então o desenho do warm-start (Q5) **não muda**.

### Estimativa de esforço da migração completa (~320 tipos)

Mecânico e amplo, não profundo:
- **~320 nós**: `[MemoryPackable]`→`[MessagePackObject]`, `[Union]` (scriptável das tags atuais), e `[property:
  Key(n)]` em cada parâmetro de construtor. O `Key(n)` por-parâmetro é o item chato — dá pra roteirizar com
  Roslyn (contar params e emitir os índices), mas exige revisão. `[MemoryPackIgnore]`→`[IgnoreMember]`.
  Estimo **2–3 dias** com script + revisão.
- **Workbook/Sheet**: callbacks `IMessagePackSerializationCallbackReceiver`, comparers, `Save/Load/Async`,
  container `MSWM` e dual-read. **1–2 dias**.
- **Testes + fixture**: `WorkbookTests` round-trip, **regenerar a fixture binária**, e os
  `MemoryPackCompatibilityTests` viram testes de "3.0 lê arquivo 2.x" (reader legado). **1–2 dias**.
- **Risco de terceiros**: MessagePack 3.1.x carrega **advisories NuGet (NU1903 alta, NU1902 médias)** — precisa
  avaliar/fixar versão antes de produção. MemoryPack não adiciona esse eixo.
- **Total: ~5–7 dias de engenharia** + a quebra de formato (major 3.0, com dual-read).

### Recomendação

**Não migrar para MessagePack.** O único ganho sólido — 30–49% de bytes no cru — evapora assim que entra
qualquer compressão, e a via barata (`MemoryPack + Brotli`, uma linha de `BrotliStream`) entrega arquivo MENOR
(13,7% vs 35% do MsgPack+LZ4 no payload grande) sem quebrar formato, sem 5–7 dias de migração, sem dual-read,
sem advisory de segurança de dependência. Em CPU o MemoryPack ainda vence por margem modesta (serialize
+4–16%, deserialize empate), então trocar também custaria um pouco de velocidade. **Alternativa mais barata e
recomendada:** se/quando tamanho em disco/rede virar requisito, adicionar compressão opcional (Brotli) por cima
do MemoryPack atual — idealmente como flag no container `MSWM` do warm-start. Reavaliar MessagePack só se
aparecer um requisito de *cru sem compressão* (ex.: mmap/zero-copy de arquivo não comprimido), que hoje não existe.
