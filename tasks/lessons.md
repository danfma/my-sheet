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
