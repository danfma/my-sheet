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
_(tabelas + recomendação ao concluir — decisão de merge é do usuário)_
