# Documentação do MySheet

*Tradução do documento canônico em inglês ([README.md](../README.md)). Em caso de divergência, o inglês prevalece.*

O MySheet é uma engine de fórmulas de planilha rápida e em memória para .NET. Ele não substitui as
bibliotecas Excel completas (ClosedXML, EPPlus, NPOI, …) — é uma opção mais simples e mais rápida para um
trabalho específico: manter uma pasta de trabalho (workbook) do Excel em um servidor como a **fonte da
verdade de um cálculo**, carregá-la, reavaliar suas fórmulas em processo (sem Excel instalado, sem COM,
com alocação mínima) e expor ou gravar de volta os resultados.

## Guias

1. [Primeiros passos](getting-started.md) — instalar, montar um workbook, fazer o parse e avaliar fórmulas.
2. [Workbook, planilhas e expressões](workbook-and-expressions.md) — o modelo de objetos, as regras de
   parsing, os operadores, as referências e a reconversão de expressões em texto de fórmula.
3. [ComputedValue e erros](computed-value.md) — o tipo de resultado da avaliação e a struct `Error`.
4. [Funções personalizadas](custom-functions.md) — estenda a engine com suas próprias funções.
5. [Interop com Excel](excel-interop.md) — carregue arquivos `.xlsx`, exporte e mescle em templates.
6. [Serialização](serialization.md) — `Save`/`Load` com MemoryPack e o que é preservado no round-trip.
7. [Desempenho](performance.md) — memoização, `RunWithLargeStack` e o design sem alocações.
8. [Referência de funções](function-reference.md) — as 164 funções nativas, mais a tabela de cobertura do
   Excel.

## Pacotes

| Pacote | Conteúdo |
| --- | --- |
| `Danfma.MySheet` | Engine principal: parser, avaliador, funções nativas e personalizadas, memoização, serialização MemoryPack. |
| `Danfma.MySheet.Excel` | Interop com `.xlsx` (OpenXML SDK): `ExcelFile.Load`, `SaveAsExcel`, `MergeIntoExcel`. |

Ambos os pacotes têm como alvo o .NET 10 e são publicados em conjunto, sempre com a mesma versão.
