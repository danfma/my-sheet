# Catálogo de cobertura de funções do Excel no README

Levantar todas as funções do Excel (catálogo oficial da Microsoft, por categoria) e publicar no
`README.md` uma seção mostrando o que o MySheet já implementa e o que falta, por categoria.

## Context
Pedido do usuário: "levantar mais funções e comportamentos do Excel que podemos portar". Depois de duas
rodadas de perguntas sobre prioridade/arquitetura (funções voláteis vs. cache, dynamic arrays vs. modelo
célula-única), o usuário simplificou o pedido para algo mais concreto e auto-contido: um catálogo
completo, sem comprometer ainda a implementação de nenhuma categoria.

Decisões fechadas:
- **Fonte da verdade para "o que já existe"**: extraído direto do dicionário `Functions` em
  `Danfma.MySheet/Parsing/Parser.cs` via grep (não confiar em resumo de agente) — 52 nomes únicos.
- **Fonte da verdade para "o que existe no Excel"**: página oficial da Microsoft, "Excel functions (by
  category)" — ~520 entradas em 14 categorias (algumas funções aparecem em mais de uma categoria, ex.
  `CONCATENATE` em Text e Compatibility, `LET` em Logical e Math).
- **Geração via script** (`scratchpad/gen_coverage.py`), não transcrição manual — risco de erro/omissão
  alto demais para fazer à mão em ~520 itens. Script cruza as duas listas e confirma que todos os 52 nomes
  implementados batem com algum nome oficial (sem nomes implementados "órfãos").
- **Formato no README**: `<details>` colapsável por categoria (Financial, Logical, Lookup and Reference,
  Math and Trigonometry, Statistical, Text, Information, Date and Time, Compatibility, Engineering,
  Database, Cubes, Web, User Defined), com lista inline `✅`/`⬜` em vez de tabela linha-a-linha (uma
  tabela de ~520 linhas seria ilegível). Categorias com 0% de cobertura ficam fechadas por padrão;
  categorias com pelo menos 1 função implementada ficam abertas.
- Esta rodada é só **documentação/levantamento** — não decide nem implementa nenhuma função nova. A
  priorização de qual categoria portar a seguir fica para uma iteração futura (ver "Próximos passos").

## For Future Agents
Este plano tem fase única, já concluída. Se uma função nova for implementada, repetir
`python3 scratchpad/gen_coverage.py`-style (recriar a lista `implemented` a partir de
`grep -oE '\["[A-Za-z0-9_.]+"\]' Danfma.MySheet/Parsing/Parser.cs`) e regenerar a seção do README para
não deixá-la desatualizada.

---

## Phase 1: Levantamento e atualização do README
Status: Complete

- [x] Buscar a lista oficial de funções do Excel por categoria (Microsoft Learn/Support).
- [x] Extrair a lista real de funções registradas em `Parser.cs` (grep, não resumo de agente).
- [x] Cruzar as duas listas via script (`scratchpad/gen_coverage.py`), confirmando 0 nomes implementados
      órfãos (todos batem com algum nome oficial).
- [x] Gerar markdown compacto (`<details>` + listas inline ✅/⬜) por categoria.
- [x] Inserir a seção "Excel function coverage" no `README.md`, entre `## Features` e `## Quick start`.
- [x] Validar balanceamento de tags HTML (`<details>`/`</details>`) e integridade do restante do arquivo.

### Verification Plan
- `grep -c "^<details" README.md` e `grep -c "^</details>" README.md` → ambos retornam o mesmo número
  (14, uma por categoria).
- Conferência visual: `## Quick start` e `## License` permanecem intactos após a seção nova.
- Nenhum nome em `✅` deveria estar ausente do `grep -oE '\["[A-Za-z0-9_.]+"\]' Parser.cs` atual (e
  vice-versa) — checado pelo script (`unmatched` vazio).

### Phase Summary
Seção "Excel function coverage" publicada no `README.md` (linhas 24–169 aprox.), com 14 categorias
colapsáveis. Cobertura atual do MySheet: **52 funções únicas** (Financial 9/55, Logical 7/19, Lookup and
Reference 7/40, Math and Trigonometry 7/82, Statistical 8/111, Text 11/49, Information 3/22, Compatibility
1/41 — só `CONCATENATE`) contra **0/25** em Date and Time, **0/54** em Engineering, **0/12** em Database,
**0/7** em Cubes, **0/3** em Web e **0/3** em User Defined.

**Maiores lacunas por volume e provável relevância de domínio** (não decidido para implementação, só
observado): Date and Time (zero funções — bloqueia `DATEDIF`/`EDATE`/`NETWORKDAYS`, comuns em planilhas
financeiras), Statistical (8/111 — faltam `STDEV.*`/`VAR.*`/`MEDIAN`/`PERCENTILE`/`RANK`/`LARGE`/`SMALL`,
úteis para análise de risco/retorno de portfólio), Financial avançadas (`XNPV`/`XIRR`/`MIRR`/depreciação
`SLN`/`DDB`/`SYD` — fluxo de caixa em datas irregulares).

## Final Recap
README passou a documentar de forma exaustiva e auditável (gerada por script, não transcrita à mão) a
cobertura de funções do Excel pelo MySheet, por categoria oficial da Microsoft. Nenhuma função nova foi
implementada nesta rodada — escopo era só levantamento/documentação.

## Deployment Plan
Mudança é só em `README.md` (documentação) — sem build/release necessário. Basta commitar e dar push;
o NuGet README só atualiza no próximo `dotnet pack`/publish de uma nova versão (não retroage em versões
já publicadas).
