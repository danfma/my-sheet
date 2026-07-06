# Multi-targeting .NET 8 — Danfma.MySheet + Danfma.MySheet.Excel

Pedido do dono (2026-07-06): os pacotes devem suportar **net8.0** além de net10.0 ("mesmo que seja
depois"). Fila: APÓS o ciclo de pesca de alocações (`plans/function-allocation-fishing.md`).

> **STATUS: NA FILA (autorizado, não iniciado).**

## Escopo
- `<TargetFrameworks>net8.0;net10.0</TargetFrameworks>` nos dois projetos de pacote (testes/benchmarks
  podem permanecer net10.0).
- Auditoria de APIs pós-net8 em uso (candidatos a checar: `System.Threading.Lock` se usado em vez de
  `object`; APIs novas de BCL em spans/collections; features de linguagem são ok com `LangVersion`
  se a BCL existir). MemoryPack/TUnit: conferir floors de framework das versões pinadas.
- CI/release: matriz de build cobre os dois TFMs; NuGet empacota ambos; suítes rodam no net10 (testar
  a lib net8 via um projeto de teste secundário OU `dotnet build -f net8.0` + smoke mínimo — decidir).
- Gates: 0 warnings nos dois TFMs; fixture/suítes verdes; benchmarks inalterados (rodam no net10).
- Release: minor (`feat`: novo TFM é feature de pacote).

## Fases
1. Spike de compilação: ligar o TFM duplo e listar TUDO que não compila no net8 (relatório antes de
   adaptar — pode exigir #if/polyfills pontuais).
2. Adaptação + CI + testes net8.
3. Release + docs (nota de frameworks suportados no README/getting-started) + pt-BR.
