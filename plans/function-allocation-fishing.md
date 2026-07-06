# Pesca de alocações por função — A/B MySheet × Aspose com MemoryDiagnoser

Eliminar, função a função, as alocações transientes da avaliação em fórmulas reais (arrays, criteria,
texto) — o eixo que o pipeline fundido expôs: **4,1GB × 2,58GB por execução** (1,59×), com Gen0/Gen1
proporcionais. O hot path escalar já é zero-alloc; o array/criteria/texto NÃO é. Método do dono:
benchmark por função (par MySheet × Aspose), medir → otimizar pontualmente (ArrayPool, stackalloc,
buffers reusáveis) → re-medir, "pescando cada ponto".

> **STATUS: EXECUÇÃO AUTORIZADA pelo usuário em 2026-07-06.** Iterações de pesca uma a uma, cada uma com
> verificação independente. Alvo inicial apontado pelo dono: `SUMIFS` e `SUMPRODUCT` (listas temporárias).

## Evidência
- Fundido preciso (12 iter, dados reais, GC default): MySheet 4,00s/4,1GB × Aspose 2,57s/2,58GB.
- Suspeitos por família (estrutura da carga real): fórmulas-array (mini-CSE materializa vetores
  `ComputedValue[]` por avaliação — `ArrayEvaluation`), criteria (`CriteriaPairs.Expand`/
  `ConditionalAggregates` — listas), `SUMPRODUCT` (`SumProducts` — listas pareadas), `COUNTIF` sobre
  coluna aberta (snapshots/enumeração), funções de texto (strings novas por op), `XLOOKUP` em colunas.
- Pré-dimensionamento (3.2.1) removeu a DOBRA das listas, não a alocação por avaliação.

## Regras do ciclo
1. **Nada de otimização sem par de benchmark antes/depois** — o A/B com Aspose ancora o alvo (paridade
   ou melhor por função) e o MemoryDiagnoser é o juiz de alocação.
2. Gates por iteração: alocação da função ↓ SEM regressão de tempo (±5%); suítes completas verdes
   (core **957** baseline / Excel **24**); fixture intocada; build `--no-incremental` 0 warnings.
3. `ArrayPool`/buffers reusáveis: atenção a exception-safety (return no finally), a NUNCA vazar buffer
   alugado para resultado retido, e ao contrato de concorrência (avaliação concorrente existe — pool é
   thread-safe, mas buffer alugado é por-avaliação). `stackalloc` só para tamanhos pequenos limitados
   (criteria de N pares, não ranges de 100k).
4. Ao final do ciclo: re-rodar o pipeline real (fases + fundido) e registrar o efeito agregado; release
   perf (patch) com o ritual completo.

## Phase 1: Suíte FunctionBenchmarks (A/B por função)
Status: In progress
- [ ] Nova classe em `benchmarks/Danfma.MySheet.Benchmark`: pares `{Função}_MySheet` × `{Função}_Aspose`
      (baseline) por categoria, `[MemoryDiagnoser]`, formas sintéticas pequenas E grandes inspiradas
      nas famílias reais: SUMIFS multi-criteria, SUMPRODUCT, COUNTIF coluna aberta, fórmula-array
      INDEX/SMALL/IF/ROW, IF-passthrough, OR com muitos disjuntos, XLOOKUP em colunas, texto
      (UPPER/TRIM/LEFT/FIND/concat), MATCH/INT, controle escalar (aritmética).
- [ ] Rodar baseline completa e registrar a tabela (tempo + Allocated por função, razões vs Aspose) no
      plano — a ORDEM DA PESCA nasce dela (pior alocador primeiro).
### Verification Plan
- Suíte compila/roda nos dois motores; valores conferem entre motores em cada par (asserção de sanidade);
  tabela baseline registrada.
### Phase Summary
_(write when phase completes)_

## Phase 2+: Iterações de pesca (uma função/família por vez)
Status: Not started
- [ ] Iteração N: pegar o pior alocador da tabela → otimizar pontualmente → antes/depois no par →
      suítes → registrar no plano → próximo. Alvos iniciais do dono: SUMIFS, SUMPRODUCT.
### Verification Plan
- Por iteração: alocação ↓, tempo ±5%, suítes/fixture verdes.
### Phase Summary
_(write when phase completes)_

## Final: efeito agregado + release
Status: Not started
- [ ] Pipeline real re-medido (fases + fundido); release perf patch; refresh pt-BR se docs mudarem.
### Phase Summary
_(write when phase completes)_
