# Lessons

Padrões aprendidos com correções e descobertas, para não repetir erros.

## MySheet — funções financeiras (2026-06-29)

- **Não confiar em golden values calculados de cabeça.** Eu cravei `IRR{-10000,3000,4200,6800} ≈ 0.201`
  no plano; o valor real é `0.16341`. E "corrigi" um PV correto para um errado. Regra: usar um oráculo
  (`ExcelFinancialFunctions`) e verificar antes de afirmar números.
- **Verificar a capacidade de uma lib antes de depender dela.** ClosedXML só computa 3 das 9 financeiras
  (PV/NPER/PPMT/NPV/RATE/IRR → `#NAME?`). Spike rápido evitou apostar numa estratégia de teste furada.
- **"Replicar o Excel com Newton ingênuo (~20 iter)" é uma armadilha para RATE/IRR.** Newton e o secante
  do POI/LibreOffice falham em `RATE(360,-600,100000)` (financiamento de 30 anos) — caso trivial que o
  Excel resolve. `(1+guess)^360` domina o resíduo e o método local diverge ou converge falso. O Excel usa
  solver robusto; bracketing + bisseção (convergência na *taxa*, não no resíduo em moeda) é o que de fato
  bate com o Excel. Lição: quando o objetivo é "igual ao Excel", validar contra um caso stiff (mortgage)
  antes de fixar o algoritmo.
- **Checar o framework de teste antes de rodar.** O projeto usa **TUnit** (`[Test]`, `await Assert.That().IsEqualTo().Within()`),
  não xUnit. `dotnet test` falha no .NET 10 (VSTest removido); o comando é
  `dotnet run --project tests/...Tests.csproj -c Release -- --treenode-filter "/*/*/Classe/*"`.
- **`.Within(tolerance)` do TUnit exige `double` (não `double?`).** Extrair o número com um helper que vira
  `NaN` quando vier `ErrorValue`, para o assert falhar limpo no RED em vez de lançar exceção no cast.

## MySheet — design de API pública / refactor de representação (2026-07-01)

- **Separar mudança de representação interna da forma da API pública.** No experimento `ComputedValue`
  (trocar `object?` por struct para matar boxing), meu esboço de migração assumiu de cara "manter `object?`
  público / remover a ponte `AsObject`" — um default conservador apresentado como se fosse obviamente certo.
  O usuário queria justamente o oposto: **retornar** o novo tipo publicamente, com helpers ergonômicos. Lição:
  o ganho de GC vem do interno (cache + nós passando o struct); a forma da API pública é decisão de produto
  do usuário — apresentar como opções (A: preservar / B: opt-in / C: trocar), não cravar a conservadora.
- **Validar a semântica do domínio de valor no código antes de opinar.** Sobre "qual o valor de uma range?":
  o código já separa duas camadas — `RangeReference.Compute` → `#VALUE!`, `Expand()` → `IEnumerable<Expression>`
  (referência), `ExpandValues()` → valores via cache. `IEnumerable<Expression>` NÃO sai do `Compute`; é camada
  de referência. Ler `RangeReference.cs` antes de propor evitou confirmar uma modelagem que misturava camadas.
- **Boxing no cache pesa mais que o transitório.** Medido: cache `object?` boxa 24 B/célula de **vida longa**
  → dispara **Gen1**; cache `Dictionary<string, ComputedValue>` → 0 B, zero coletas. O throughput no caminho
  cache-heavy sobe só 4–12% (lookups de Dictionary dominam), mas o ganho de GC é o argumento forte, não a
  velocidade bruta. Não vender um refactor de perf só pelo throughput sem medir o eixo de alocação/GC.

## MySheet — convenções e processo (2026-07-01)

- **Commits: inglês, sujeito curto + corpo descritivo, semantic/conventional commits.** Correção do usuário
  em 2026-07-01 — os commits em português eram a convenção antiga do repo; o padrão agora é
  `tipo(escopo): resumo em inglês` + parágrafo curto de contexto. Nunca incluir atribuição de IA.
- **Contagens em planos: recontar antes de publicar.** Escrevi "~85 exclusões" no roadmap de funções quando
  a lista explícita somava 35 — o usuário aprovou a LISTA, mas o número errado contaminou a meta de
  cobertura (~435 vs ~485 viáveis). Números derivados de listas devem ser contados por script/soma real,
  nunca estimados de memória (mesma família da lição dos golden values).
