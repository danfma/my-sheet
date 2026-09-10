# Referência de funções

*Tradução do documento canônico em inglês ([function-reference.md](../function-reference.md)). Em caso de divergência, o inglês prevalece.*

O MySheet implementa **306 funções nativas (built-in)**. A lista registrada oficial é o mapa `ByName`
em [`Danfma.MySheet/Parsing/FunctionRegistry.cs`](../../Danfma.MySheet/Parsing/FunctionRegistry.cs) —
esta página é derivada dele. A quantidade de argumentos é validada **em tempo de parse**: chamar uma função nativa com um número
de argumentos não suportado lança uma `ParseException`, assim como o Excel rejeita a fórmula na
digitação.

As linhas abaixo descrevem o comportamento próprio de cada função e não mudam em contexto de array. Além
delas, uma função **puramente escalar** que recebe um intervalo em uma posição que consome arrays é aplicada
**elemento a elemento** — `SUM(LEN(A1:A3))` soma três comprimentos — enquanto uma função ciente de intervalos
continua consumindo o intervalo inteiro, como documentado na linha dela. 180 das 306 entradas podem ser
elevadas assim; veja
[argumentos implícitos de array](workbook-and-expressions.md#argumentos-implícitos-de-array) para saber quais
consumidores pedem um array, quais funções são elevadas e onde a elevação para.

Além dessas, você pode adicionar suas próprias funções com
[`workbook.RegisterFunction`](custom-functions.md); nomes desconhecidos são avaliados como `#NAME?`.

Convenções abaixo: `[argumento]` entre colchetes é opcional; `…` significa que a função é variádica.
"Aceita intervalos" (*range-aware*) significa que argumentos de intervalo (`A1:B10`, uniões e resultados
de referência como o de `OFFSET`) são expandidos célula a célula.

## Lógicas (12)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` se todo argumento for considerado verdadeiro; operandos de texto e em branco são ignorados (seja um literal ou alcançado por meio de uma referência); nenhum valor avaliável → `#VALUE!`. |
| `FALSE` | `FALSE()` | O valor lógico `FALSE` (forma de função do literal). |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Condicional; apenas o ramo escolhido é avaliado. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, ou o valor de contingência quando `value` é qualquer erro. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, ou o valor de contingência apenas quando `value` é `#N/A`. |
| `IFS` | `IFS(test1, value1, [test2, value2], …)` | Primeiro valor cujo teste é `TRUE` (avaliação preguiçosa, como `IF`); nenhum teste `TRUE` → `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Vincula nomes utilizáveis em `calculation` (ex.: `=LET(x, A1*2, x+x)`). Os nomes são locais à fórmula. |
| `NOT` | `NOT(logical)` | Negação lógica. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` se algum argumento for considerado verdadeiro; operandos de texto e em branco são ignorados (seja um literal ou alcançado por meio de uma referência); nenhum valor avaliável → `#VALUE!`. |
| `SWITCH` | `SWITCH(expression, value1, result1, …, [default])` | Primeiro resultado cujo valor é igual a `expression` (avaliação preguiçosa; semântica de igualdade do `=`); sem correspondência → o padrão ou `#N/A`. |
| `TRUE` | `TRUE()` | O valor lógico `TRUE` (forma de função do literal). |
| `XOR` | `XOR(logical1, [logical2], …)` | `TRUE` quando a quantidade de entradas `TRUE` é ímpar; operandos de texto e em branco são ignorados (seja um literal ou alcançado por meio de uma referência); nenhum valor avaliável → `#VALUE!`. |

## Matemática e trigonometria (75)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ABS` | `ABS(number)` | Valor absoluto. |
| `ACOS` | `ACOS(number)` | Arco cosseno; fora de `[-1, 1]` → `#NUM!`. |
| `ACOSH` | `ACOSH(number)` | Cosseno hiperbólico inverso; abaixo de 1 → `#NUM!`. |
| `ACOT` | `ACOT(number)` | Arco cotangente, em `(0, π)`. |
| `ACOTH` | `ACOTH(number)` | Cotangente hiperbólica inversa; `\|number\| <= 1` → `#NUM!`. |
| `AGGREGATE` | `AGGREGATE(function_num, options, ref1, [ref2], …)` / `AGGREGATE(function_num, options, array, k)` | As duas sintaxes documentadas do Excel sobre um único nome, distinguidas apenas pelo `function_num` — nada de sintático as separa. **1-13** é a forma-referência (todo argumento a partir do terceiro é outra referência): 1 `AVERAGE`, 2 `COUNT`, 3 `COUNTA`, 4 `MAX`, 5 `MIN`, 6 `PRODUCT`, 7 `STDEV.S`, 8 `STDEV.P`, 9 `SUM`, 10 `VAR.S`, 11 `VAR.P`, 12 `MEDIAN`, 13 `MODE.SNGL`. **14-19** é a forma-array, cujo quarto argumento é o `k` / `quart`: 14 `LARGE`, 15 `SMALL`, 16 `PERCENTILE.INC`, 17 `QUARTILE.INC`, 18 `PERCENTILE.EXC`, 19 `QUARTILE.EXC`. Por isso a mesma chamada de quatro argumentos significa duas coisas diferentes — com `A1:A3` = 5/0/9 e `B1:B3` = 1/2/3, `AGGREGATE(9,6,A1:A3,B1:B3)` = 20 (uma segunda referência), enquanto `AGGREGATE(15,6,A1:A3,2)` = 5 (o 2º menor). O `options` (0 ou omitido, até 7) é composto por três bits independentes: **+1** ignora linhas ocultas, **+2** ignora valores de erro, **+4** *mantém* as células com `SUBTOTAL`/`AGGREGATE` aninhados. Assim, 0-3 pulam uma célula referenciada cuja própria fórmula é um `SUBTOTAL` ou um `AGGREGATE`, e 4-7 a contam — com `C1:C3` = um `SUBTOTAL` aninhado que vale 3, um `AGGREGATE` aninhado que vale 3 e um 5 simples: `AGGREGATE(9,0,C1:C3)` = 5 contra `AGGREGATE(9,4,C1:C3)` = 11. Esse 5 é a tabela de options lida ao pé da letra e **não** é um valor medido: o Aspose.Cells 26.6.0 — o oráculo por trás dos números medidos no resto desta linha — responde 8, porque pula o `SUBTOTAL` aninhado mas conta o `AGGREGATE` aninhado (e 2, em vez de 1, para `AGGREGATE(3,0,C1:C3)`). É uma divergência conhecida, na qual a página documentada prevalece. As options 2/3/6/7 tiram as células de erro da população em vez de propagá-las: com `E1:E3` = 5 / `#DIV/0!` / 9, `AGGREGATE(9,6,E1:E3)` = 14 onde `AGGREGATE(9,4,E1:E3)` é `#DIV/0!`, e no `COUNTA` a própria contagem muda (`AGGREGATE(3,4,E1:E3)` = 3, `AGGREGATE(3,6,E1:E3)` = 2 — a página não diz o que o `COUNTA` passa a contar, mas os dois valores foram medidos no Aspose.Cells 26.6.0). Esse bit alcança apenas *células* de erro e *elementos* de array: um argumento que é ele próprio um erro propaga, independentemente das options — `AGGREGATE(9,6,1/0)` e `AGGREGATE(9,6,E1:E3,1/0)` dão os dois `#DIV/0!`, enquanto `AGGREGATE(9,6,E2)`, a mesma divisão morando em uma célula referenciada, dá 0 (tudo medido). O bit de linhas ocultas é um **no-op**: o MySheet não tem modelo de linhas ocultas, então 1/3/5/7 se comportam exatamente como 0/2/4/6 (limite documentado, o mesmo por trás dos códigos 101-111 do `SUBTOTAL`). O `k` da forma-array conta sobre a população que *sobrevive* aos descartes: nesse mesmo `E1:E3`, `AGGREGATE(15,6,E1:E3,2)` = 9, mas `AGGREGATE(15,6,E1:E3,3)` → `#NUM!`, ainda que o intervalo tenha três células. Só a **forma-array** aceita um [array computado](workbook-and-expressions.md#argumentos-implícitos-de-array) na sua posição `array`. Os argumentos `ref` da forma-referência são referências, e um array computado em um deles é `#VALUE!` — `AGGREGATE(9,4,ROW(A1:A3))`, `AGGREGATE(2,4,(A1:A3<>0)*1)` e até a constante `AGGREGATE(9,4,{1,2,3})` falham (medido no Aspose.Cells 26.6.0, em 2026-09-09, inclusive com entrada em CSE). Essa assimetria é justamente a razão de existirem duas sintaxes: com `options` = 6, é a forma-array que transforma o idioma de planilha `SMALL(posições/filtro, k)` de `#DIV/0!` na resposta pretendida — com `A1:A3` = 5/0/9, `AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0)),1)` = 1, e a mesma chamada com `k` = 2 dá 3. Erros: `function_num` fora de 1-19 ou `options` fora de 0-7 → `#VALUE!`; a forma-array com apenas três argumentos (sem o `k`) → `#VALUE!`, a resposta da própria página quando "um segundo argumento ref é necessário mas não é fornecido" (menos de três argumentos é erro de aridade, rejeitado em tempo de parse); um `k` fora da população sobrevivente → `#NUM!` — a resposta do próprio `LARGE`/`SMALL`; a página do AGGREGATE não a declara, mas ela foi medida no Aspose.Cells 26.6.0 (`AGGREGATE(15,6,E1:E3,3)` sobre uma população de dois sobreviventes, e uma população só de erros com qualquer `k`); uma referência a uma planilha inexistente → `#REF!` nas duas formas, antes de qualquer célula ser varrida. |
| `ARABIC` | `ARABIC(text)` | Numeral romano → número (não diferencia maiúsculas de minúsculas; `""` → 0; `-` inicial nega o valor). |
| `ASIN` | `ASIN(number)` | Arco seno; fora de `[-1, 1]` → `#NUM!`. |
| `ASINH` | `ASINH(number)` | Seno hiperbólico inverso. |
| `ATAN` | `ATAN(number)` | Arco tangente. |
| `ATAN2` | `ATAN2(x_num, y_num)` | Arco tangente a partir de coordenadas — na ordem `(x, y)` do Excel; `ATAN2(0,0)` → `#DIV/0!`. |
| `ATANH` | `ATANH(number)` | Tangente hiperbólica inversa; `\|number\| >= 1` → `#NUM!`. |
| `BASE` | `BASE(number, radix, [min_length])` | Número → texto na base `radix` (2-36), preenchido com zeros até `min_length`. |
| `CEILING` | `CEILING(number, significance)` | Teto legado com as regras de sinal do Excel (`CEILING(-2.5,-2)` = -4; número positivo com significância negativa → `#NUM!`). |
| `CEILING.MATH` | `CEILING.MATH(number, [significance], [mode])` | Arredonda para cima até um múltiplo; `mode` só afeta números negativos (para longe de zero quando diferente de zero). |
| `CEILING.PRECISE` | `CEILING.PRECISE(number, [significance])` | Arredonda em direção a +∞; o sinal da significância é ignorado. |
| `COMBIN` | `COMBIN(number, number_chosen)` | Combinações sem repetição. |
| `COMBINA` | `COMBINA(number, number_chosen)` | Combinações com repetição (`COMBIN(n+k-1, k)`). |
| `COS` | `COS(number)` | Cosseno (radianos). |
| `COSH` | `COSH(number)` | Cosseno hiperbólico. |
| `COT` | `COT(number)` | Cotangente; `COT(0)` → `#DIV/0!`. |
| `COTH` | `COTH(number)` | Cotangente hiperbólica; `COTH(0)` → `#DIV/0!`. |
| `CSC` | `CSC(number)` | Cossecante; `CSC(0)` → `#DIV/0!`. |
| `CSCH` | `CSCH(number)` | Cossecante hiperbólica; `CSCH(0)` → `#DIV/0!`. |
| `DECIMAL` | `DECIMAL(text, radix)` | Texto na base `radix` (2-36) → número; não diferencia maiúsculas de minúsculas. |
| `DEGREES` | `DEGREES(angle)` | Radianos → graus. |
| `EVEN` | `EVEN(number)` | Arredonda para longe de zero até o inteiro par mais próximo. |
| `EXP` | `EXP(number)` | e elevado a `number`. |
| `FACT` | `FACT(number)` | Fatorial (trunca; negativo → `#NUM!`). |
| `FACTDOUBLE` | `FACTDOUBLE(number)` | Fatorial duplo n!! (trunca; negativo → `#NUM!`). |
| `FLOOR` | `FLOOR(number, significance)` | Piso legado com as regras de sinal do Excel (`FLOOR(-2.5,-2)` = -2; significância 0 → `#DIV/0!`). |
| `FLOOR.MATH` | `FLOOR.MATH(number, [significance], [mode])` | Arredonda para baixo até um múltiplo; `mode` só afeta números negativos (em direção a zero quando diferente de zero). |
| `FLOOR.PRECISE` | `FLOOR.PRECISE(number, [significance])` | Arredonda em direção a -∞; o sinal da significância é ignorado. |
| `GCD` | `GCD(number1, …)` | Máximo divisor comum; aceita intervalos (trunca; negativo → `#NUM!`). |
| `INT` | `INT(number)` | Arredonda para baixo até o inteiro mais próximo. |
| `ISO.CEILING` | `ISO.CEILING(number, [significance])` | Comportamento de alias de `CEILING.PRECISE`. |
| `LCM` | `LCM(number1, …)` | Mínimo múltiplo comum; aceita intervalos (trunca; negativo → `#NUM!`). |
| `LN` | `LN(number)` | Logaritmo natural; não positivo → `#NUM!`. |
| `LOG` | `LOG(number, [base])` | Logaritmo (base padrão: 10); base 1 → `#DIV/0!`, base ≤ 0 → `#NUM!`. |
| `LOG10` | `LOG10(number)` | Logaritmo na base 10. |
| `MOD` | `MOD(number, divisor)` | Resto com o sinal do divisor (`MOD(-3,2)` = 1); divisor 0 → `#DIV/0!`. |
| `MROUND` | `MROUND(number, multiple)` | Arredonda para o múltiplo mais próximo; sinais opostos → `#NUM!`. |
| `MULTINOMIAL` | `MULTINOMIAL(number1, …)` | Coeficiente multinomial; aceita intervalos. |
| `ODD` | `ODD(number)` | Arredonda para longe de zero até o inteiro ímpar mais próximo. |
| `PI` | `PI()` | A constante π. |
| `POWER` | `POWER(number, power)` | Exponenciação; `0^0` → `#NUM!`, `0^negativo` → `#DIV/0!`. |
| `PRODUCT` | `PRODUCT(number1, …)` | Produto dos valores numéricos; aceita intervalos. |
| `QUOTIENT` | `QUOTIENT(numerator, denominator)` | Parte inteira de uma divisão (truncada). |
| `RADIANS` | `RADIANS(angle)` | Graus → radianos. |
| `RAND` | `RAND()` | Volátil: um número real aleatório em `[0, 1)`. Veja [Funções voláteis](workbook-and-expressions.md#funções-voláteis). |
| `RANDBETWEEN` | `RANDBETWEEN(bottom, top)` | Volátil: um número inteiro aleatório em `[bottom, top]` (inclusive); `bottom > top` → `#NUM!`; limites não inteiros truncam em direção a zero. |
| `ROMAN` | `ROMAN(number, [form])` | Número (0-3999) → numeral romano clássico; `ROMAN(0)` = `""`. As formas concisas 1-4/`FALSE` não são suportadas (→ `#VALUE!`). |
| `ROUND` | `ROUND(number, num_digits)` | Arredonda para a quantidade dada de dígitos. |
| `ROUNDDOWN` | `ROUNDDOWN(number, num_digits)` | Arredonda em direção a zero. |
| `ROUNDUP` | `ROUNDUP(number, num_digits)` | Arredonda para longe de zero. |
| `SEC` | `SEC(number)` | Secante. |
| `SECH` | `SECH(number)` | Secante hiperbólica. |
| `SERIESSUM` | `SERIESSUM(x, n, m, coefficients)` | Soma de série de potências; coeficientes via intervalo/valores. |
| `SIGN` | `SIGN(number)` | -1, 0 ou 1. |
| `SIN` | `SIN(number)` | Seno (radianos). |
| `SINH` | `SINH(number)` | Seno hiperbólico. |
| `SQRT` | `SQRT(number)` | Raiz quadrada; negativo → `#NUM!`. |
| `SQRTPI` | `SQRTPI(number)` | Raiz quadrada de `number × π`. |
| `SUBTOTAL` | `SUBTOTAL(function_num, ref1, [ref2], …)` | Agregação selecionada por `function_num` (1-11: AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/STDEV.S/STDEV.P/SUM/VAR.S/VAR.P); células referenciadas cuja própria fórmula é um `SUBTOTAL` são ignoradas (evita contagem em duplicidade). 101-111 se comportam como 1-11 — o MySheet não tem modelo de linhas ocultas (limite documentado). Código inválido → `#VALUE!`. Um argumento que é um **array computado**, e não uma referência, é `#VALUE!`, e não dobrado: a página define `ref1` como "the first named range or reference", e o Excel aplica isso ao pé da letra — `SUBTOTAL(9,ROW(A1:A3))`, `SUBTOTAL(9,(A1:A3<>0)*1)`, `SUBTOTAL(2,…)`, `SUBTOTAL(3,…)` e até a constante `SUBTOTAL(9,{1,2,3})` dão todos `#VALUE!` (medido no Aspose.Cells 26.6.0, em 2026-09-09, inclusive com entrada em CSE). O `SUM` dobra um [argumento implícito de array](workbook-and-expressions.md#argumentos-implícitos-de-array) e o `SUBTOTAL` não — uma versão anterior o dobrava aqui por analogia com o `SUM`, o que era uma inferência, e a medição a reverteu. A posição que de fato aceita um array computado é a da forma-array do `AGGREGATE` (`function_num` 14-19). Uma célula com `AGGREGATE` aninhado **não** é pulada, só uma com `SUBTOTAL` aninhado: sobre um intervalo com um `SUBTOTAL` aninhado (3), um `AGGREGATE` aninhado (3) e um 5 simples, `SUBTOTAL(9,…)` = 8 — medido no Aspose.Cells 26.6.0, de modo que a regra estreita é a do Excel, e não um palpite, ainda que a redação "nested SUBTOTAL and AGGREGATE" só exista na tabela de options do próprio AGGREGATE. O `AGGREGATE(9,0,…)` sobre esse mesmo intervalo dá 5 aqui, e esse 5 **não** é medido: ele segue a tabela de options do AGGREGATE ao pé da letra, enquanto o oráculo responde 8 — ele pula o `SUBTOTAL` aninhado e conta o `AGGREGATE` aninhado que a tabela manda pular. É uma divergência conhecida, na qual a página documentada prevalece. |
| `SUM` | `SUM([number1], …)` | Soma de todos os valores numéricos; aceita intervalos. Também dobra um [argumento implícito de array](workbook-and-expressions.md#argumentos-implícitos-de-array) — `SUM(IF(B2:B5="Show",1,0))` = 2. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Soma as células que atendem a um critério (ex.: `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Soma sob múltiplos pares de intervalo e critério. |
| `SUMPRODUCT` | `SUMPRODUCT(array1, [array2], …)` | Soma dos produtos posição a posição; entradas não numéricas contam como 0 — inclusive o `TRUE`/`FALSE` de uma comparação pura, e é por isso que o idioma do Excel multiplica por 1 — enquanto um elemento de erro se propaga. Aceita um [array computado](workbook-and-expressions.md#argumentos-implícitos-de-array) como argumento inteiro, misturado livremente com intervalos: `SUMPRODUCT((A1:A3<>0)*1)` = 2, `SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))` = 9. Os argumentos precisam ter as mesmas **dimensões**, e não apenas a mesma contagem de células: `SUMPRODUCT(A1:A3,A1:C1)` (3x1 contra 1x3) → `#VALUE!`. Um argumento *sem formato* — um nome definido, um intervalo aberto, uma união — não carrega um retângulo que a engine consiga enxergar e é julgado apenas pela contagem de células, então `SUMPRODUCT(MyName,A1:C1)` sobre um nome 3x1 calcula onde o Excel responde `#VALUE!` (uma divergência documentada); todo argumento que *carrega* um retângulo continua sendo comparado com os demais, então `SUMPRODUCT(MyName,A1:A3,A1:C1)` é `#VALUE!`. |
| `SUMSQ` | `SUMSQ(number1, …)` | Soma dos quadrados; aceita intervalos. |
| `SUMX2MY2` | `SUMX2MY2(array_x, array_y)` | Σ(x² − y²); pares com um lado não numérico são descartados; comprimentos diferentes → `#N/A`. |
| `SUMX2PY2` | `SUMX2PY2(array_x, array_y)` | Σ(x² + y²); mesmas regras de pareamento de `SUMX2MY2`. |
| `SUMXMY2` | `SUMXMY2(array_x, array_y)` | Σ(x − y)²; mesmas regras de pareamento de `SUMX2MY2`. |
| `TAN` | `TAN(number)` | Tangente (radianos). |
| `TANH` | `TANH(number)` | Tangente hiperbólica. |
| `TRUNC` | `TRUNC(number, [num_digits])` | Trunca em direção a zero (padrão: 0 dígitos). |

## Estatísticas (59)

Convenções desta família: os agregados simples ignoram texto/lógicos/células em branco referenciados
(como `SUM`); as variantes `*A` contam texto referenciado como 0 e lógicos como 1/0. As funções de duas
séries (`CORREL`, `SLOPE`, …) descartam um par inteiro quando QUALQUER um dos lados é não numérico,
retornam `#N/A` em caso de comprimentos diferentes e `#DIV/0!` quando a variância é zero. `GAUSS` fica
para a fase das distribuições estatísticas (precisa da CDF normal/erf); `PHI` — a densidade simples —
já está incluída.

`AVERAGE`, `COUNT`, `MIN`, `MAX`, `SMALL` e `LARGE` também aceitam um
[argumento implícito de array](workbook-and-expressions.md#argumentos-implícitos-de-array) — ex.:
`SMALL(IF(B2:B5="Show",ROW(B2:B5)),1)` — dobrando-o elemento a elemento com a mesma regra de ignorar
texto/lógicos (por isso o `FALSE` de um `IF` sem ramo é descartado).

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `AVEDEV` | `AVEDEV(number1, …)` | Média dos desvios absolutos em relação à média; nenhum valor → `#NUM!`. |
| `AVERAGE` | `AVERAGE([number1], …)` | Média aritmética dos valores numéricos; aceita intervalos. |
| `AVERAGEA` | `AVERAGEA(value1, …)` | `AVERAGE` com a regra `*A` (texto → 0, lógicos → 1/0). |
| `AVERAGEIF` | `AVERAGEIF(range, criteria, [average_range])` | Média das células que atendem a um critério; nenhuma correspondência numérica → `#DIV/0!`. |
| `AVERAGEIFS` | `AVERAGEIFS(average_range, criteria_range1, criteria1, …)` | Média sob múltiplos pares de intervalo e critério; nenhuma correspondência → `#DIV/0!`; incompatibilidade de formato → `#VALUE!`. |
| `CORREL` | `CORREL(array1, array2)` | Coeficiente de correlação produto-momento de Pearson. |
| `COUNT` | `COUNT([value1], …)` | Conta valores numéricos; aceita intervalos. |
| `COUNTA` | `COUNTA(value1, …)` | Conta valores não em branco; aceita intervalos. |
| `COUNTBLANK` | `COUNTBLANK(range, …)` | Conta células em branco. |
| `COUNTIF` | `COUNTIF(range, criteria)` | Conta as células que atendem a um critério. |
| `COUNTIFS` | `COUNTIFS(criteria_range1, criteria1, …)` | Contagem sob múltiplos pares de intervalo e critério. |
| `COVARIANCE.P` | `COVARIANCE.P(array1, array2)` | Covariância populacional Σ(x−x̄)(y−ȳ)/n. |
| `COVARIANCE.S` | `COVARIANCE.S(array1, array2)` | Covariância amostral (n−1); menos de 2 pares → `#DIV/0!`. |
| `DEVSQ` | `DEVSQ(number1, …)` | Soma dos quadrados dos desvios em relação à média; nenhum valor → `#NUM!`. |
| `FISHER` | `FISHER(x)` | Transformação de Fisher; `x` ≤ −1 ou ≥ 1 → `#NUM!`. |
| `FISHERINV` | `FISHERINV(y)` | Transformação inversa de Fisher. |
| `FORECAST.LINEAR` | `FORECAST.LINEAR(x, known_ys, known_xs)` | y previsto em `x` sobre a reta de mínimos quadrados — atenção à ordem dos argumentos (o novo x vem primeiro). |
| `GEOMEAN` | `GEOMEAN(number1, …)` | Média geométrica; qualquer valor ≤ 0 → `#NUM!`. |
| `HARMEAN` | `HARMEAN(number1, …)` | Média harmônica; qualquer valor ≤ 0 → `#NUM!`. |
| `INTERCEPT` | `INTERCEPT(known_ys, known_xs)` | Intercepto y da reta de mínimos quadrados. |
| `KURT` | `KURT(number1, …)` | Curtose em excesso (fórmula amostral do Excel); menos de 4 pontos ou s = 0 → `#DIV/0!`. |
| `LARGE` | `LARGE(array, k)` | k-ésimo maior valor; array vazio, k ≤ 0 ou k > n → `#NUM!`. |
| `MAX` | `MAX([number1], …)` | Maior valor numérico; aceita intervalos. |
| `MAXA` | `MAXA(value1, …)` | `MAX` com a regra `*A`. |
| `MAXIFS` | `MAXIFS(max_range, criteria_range1, criteria1, …)` | Maior valor correspondente; nenhuma correspondência → 0; incompatibilidade de formato → `#VALUE!`. |
| `MEDIAN` | `MEDIAN(number1, …)` | Valor central (média dos dois valores centrais quando a quantidade é par); nenhum valor → `#NUM!`. |
| `MIN` | `MIN([number1], …)` | Menor valor numérico; aceita intervalos. |
| `MINA` | `MINA(value1, …)` | `MIN` com a regra `*A`. |
| `MINIFS` | `MINIFS(min_range, criteria_range1, criteria1, …)` | Menor valor correspondente; nenhuma correspondência → 0; incompatibilidade de formato → `#VALUE!`. |
| `MODE.SNGL` | `MODE.SNGL(number1, …)` | Valor mais frequente; em caso de empate, prevalece o primeiro valor encontrado; sem duplicatas → `#N/A`. |
| `PEARSON` | `PEARSON(array1, array2)` | O mesmo coeficiente de `CORREL`. |
| `PERCENTILE.EXC` | `PERCENTILE.EXC(array, k)` | Percentil exclusivo: interpolação na posição `k·(n+1)`; `k` fora de `(0, 1)` ou uma posição inalcançável → `#NUM!`. |
| `PERCENTILE.INC` | `PERCENTILE.INC(array, k)` | Percentil inclusivo: interpolação na posição `k·(n−1)`; `k` fora de `[0, 1]` → `#NUM!`. |
| `PERCENTRANK.EXC` | `PERCENTRANK.EXC(array, x, [significance])` | Posição exclusiva de `x` como fração (`(abaixo+1)/(n+1)`, interpolada); TRUNCADA para `significance` dígitos (padrão 3); `x` fora do intervalo → `#N/A`. |
| `PERCENTRANK.INC` | `PERCENTRANK.INC(array, x, [significance])` | Posição inclusiva de `x` (`abaixo/(n−1)`, interpolada); mesma truncagem e erros de `.EXC`. |
| `PERMUT` | `PERMUT(number, number_chosen)` | Permutações sem repetição n!/(n−k)! (argumentos truncados); n ≤ 0, k < 0 ou n < k → `#NUM!`. |
| `PERMUTATIONA` | `PERMUTATIONA(number, number_chosen)` | Permutações com repetição n^k (argumentos truncados); negativo → `#NUM!`. |
| `PHI` | `PHI(x)` | Densidade da distribuição normal padrão. |
| `PROB` | `PROB(x_range, prob_range, lower_limit, [upper_limit])` | Soma das probabilidades de x em `[lower, upper]` (upper omitido → x = lower); probabilidades fora de `(0, 1]` ou que não somam 1 → `#NUM!`; comprimentos diferentes → `#N/A`. |
| `QUARTILE.EXC` | `QUARTILE.EXC(array, quart)` | Quartil exclusivo via `PERCENTILE.EXC(quart/4)`; quart (truncado) ≤ 0 ou ≥ 4 → `#NUM!`. |
| `QUARTILE.INC` | `QUARTILE.INC(array, quart)` | Quartil inclusivo via `PERCENTILE.INC(quart/4)`; quart (truncado) fora de 0-4 → `#NUM!`. |
| `RANK.AVG` | `RANK.AVG(number, ref, [order])` | Posição (rank) com empates calculados pela média; order 0/omitido → decrescente, senão crescente; valor ausente → `#N/A`. |
| `RANK.EQ` | `RANK.EQ(number, ref, [order])` | Posição (rank) com empates compartilhando a posição mais alta do grupo; mesma ordem/erros de `.AVG`. |
| `RSQ` | `RSQ(known_ys, known_xs)` | Quadrado do coeficiente de Pearson. |
| `SKEW` | `SKEW(number1, …)` | Assimetria amostral (fator n/((n−1)(n−2))); menos de 3 pontos ou s = 0 → `#DIV/0!`. |
| `SKEW.P` | `SKEW.P(number1, …)` | Assimetria populacional; menos de 3 pontos ou σ = 0 → `#DIV/0!`. |
| `SLOPE` | `SLOPE(known_ys, known_xs)` | Inclinação (slope) de mínimos quadrados; var(x) = 0 → `#DIV/0!`. |
| `SMALL` | `SMALL(array, k)` | k-ésimo menor valor; array vazio, k ≤ 0 ou k > n → `#NUM!`. |
| `STANDARDIZE` | `STANDARDIZE(x, mean, standard_dev)` | O z-score (x − mean)/sd; sd ≤ 0 → `#NUM!`. |
| `STDEV.P` | `STDEV.P(number1, …)` | Desvio padrão populacional ("n"); nenhum valor → `#DIV/0!`. |
| `STDEV.S` | `STDEV.S(number1, …)` | Desvio padrão amostral ("n−1"); menos de 2 valores → `#DIV/0!`. |
| `STDEVA` | `STDEVA(value1, …)` | `STDEV.S` com a regra `*A`. |
| `STDEVPA` | `STDEVPA(value1, …)` | `STDEV.P` com a regra `*A`. |
| `STEYX` | `STEYX(known_ys, known_xs)` | Erro padrão do y previsto; menos de 3 pares → `#DIV/0!`. |
| `TRIMMEAN` | `TRIMMEAN(array, percent)` | Média após cortar `INT(n·percent/2)` valores de CADA extremidade dos dados ordenados; `percent` fora de `[0, 1)` → `#NUM!`. |
| `VAR.P` | `VAR.P(number1, …)` | Variância populacional; nenhum valor → `#DIV/0!`. |
| `VAR.S` | `VAR.S(number1, …)` | Variância amostral; menos de 2 valores → `#DIV/0!`. |
| `VARA` | `VARA(value1, …)` | `VAR.S` com a regra `*A`. |
| `VARPA` | `VARPA(value1, …)` | `VAR.P` com a regra `*A`. |

## Texto (34)

As funções de texto seguem o contrato invariante de localidade da engine: comparações ordinais,
conversão de maiúsculas/minúsculas invariante, `.`/`,`/`$` em `FIXED`/`DOLLAR`. `CHAR`/`CODE` mapeiam
pontos de código Unicode (Latin-1 para 1-255), e não a página de código ANSI do Windows. As funções
`REGEX*` rodam sobre expressões regulares do .NET (o Excel especifica PCRE2; o subconjunto usual —
classes, quantificadores, âncoras, grupos, `$n` — se comporta de forma idêntica), com um tempo-limite
defensivo de correspondência de 1 segundo.

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `CHAR` | `CHAR(number)` | Caractere para um código 1-255 (fora do intervalo → `#VALUE!`). |
| `CLEAN` | `CLEAN(text)` | Remove os caracteres de controle 0-31 (127 etc. permanecem, como no Excel). |
| `CODE` | `CODE(text)` | Código do primeiro caractere (texto vazio → `#VALUE!`). |
| `CONCAT` | `CONCAT(text1, …)` | Concatena valores; aceita intervalos. |
| `CONCATENATE` | `CONCATENATE(text1, …)` | Alias legado de concatenação (argumentos escalares). |
| `DOLLAR` | `DOLLAR(number, [decimals])` | Número como TEXTO monetário — `$1,234.57`, negativos `($1,200)`; decimais com padrão 2, negativo arredonda à esquerda do ponto. |
| `EXACT` | `EXACT(text1, text2)` | Comparação que diferencia maiúsculas de minúsculas. |
| `FIND` | `FIND(find_text, within_text, [start_num])` | Posição (base 1) diferenciando maiúsculas de minúsculas; sem curingas; não encontrado → `#VALUE!`. |
| `FIXED` | `FIXED(number, [decimals], [no_commas])` | Número arredondado e renderizado como TEXTO — `1,234.6`; decimais com padrão 2 (máx. 127), negativo arredonda à esquerda do ponto. |
| `LEFT` | `LEFT(text, [num_chars])` | Caracteres iniciais (padrão: 1). |
| `LEN` | `LEN(text)` | Comprimento do texto. |
| `LOWER` | `LOWER(text)` | Converte o texto para minúsculas. |
| `MID` | `MID(text, start_num, num_chars)` | Subtexto por posição (base 1) e comprimento. |
| `NUMBERVALUE` | `NUMBERVALUE(text, [decimal_separator], [group_separator])` | Texto → número com localidade explícita (padrões `.` e `,`); espaços são ignorados; cada `%` ao final divide por 100. |
| `PROPER` | `PROPER(text)` | Coloca em maiúscula toda letra que segue um não-letra; converte o restante para minúsculas. |
| `REGEXEXTRACT` | `REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])` | Primeira correspondência do padrão (modo 0); modos de array 1/2 → `#VALUE!` até a fase de arrays; sem correspondência → `#N/A`. |
| `REGEXREPLACE` | `REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])` | Substitui as correspondências (referências de grupo `$n`); ocorrência 0 = todas, positiva = a n-ésima, negativa = a n-ésima a partir do fim. |
| `REGEXTEST` | `REGEXTEST(text, pattern, [case_sensitivity])` | `TRUE` quando o padrão corresponde em qualquer parte do texto. |
| `REPLACE` | `REPLACE(old_text, start_num, num_chars, new_text)` | Substitui por posição (base 1) e comprimento. |
| `REPT` | `REPT(text, number_times)` | Repete o texto (contagem truncada; negativa ou resultado acima de 32.767 caracteres → `#VALUE!`). |
| `RIGHT` | `RIGHT(text, [num_chars])` | Caracteres finais (padrão: 1). |
| `SEARCH` | `SEARCH(find_text, within_text, [start_num])` | Posição sem diferenciar maiúsculas de minúsculas, com curingas `?` `*` (`~` escapa); não encontrado → `#VALUE!`. |
| `SUBSTITUTE` | `SUBSTITUTE(text, old_text, new_text, [instance_num])` | Substituição que diferencia maiúsculas de minúsculas — toda ocorrência, ou apenas a de índice `instance_num` (base 1). |
| `T` | `T(value)` | O valor, se for texto; caso contrário, `""`. |
| `TEXT` | `TEXT(value, format_text)` | Formata um valor (formatos de número e data, ex.: `"0.00"`, `"dd/mm/yyyy"`). Um formato de data lê o calendário de 1900 do próprio Excel — veja [Data e hora](#data-e-hora-25). |
| `TEXTAFTER` | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Texto após o n-ésimo delimitador (negativo conta a partir do fim); sem correspondência → `if_not_found` ou `#N/A`. |
| `TEXTBEFORE` | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Texto antes do n-ésimo delimitador (negativo conta a partir do fim); sem correspondência → `if_not_found` ou `#N/A`. |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Junta valores com um delimitador; aceita intervalos. |
| `TRIM` | `TRIM(text)` | Remove espaços em excesso. |
| `UNICHAR` | `UNICHAR(number)` | Caractere para um ponto de código Unicode completo (0/fora do intervalo → `#VALUE!`; pontos de código substitutos (*surrogates*) → `#N/A`). |
| `UNICODE` | `UNICODE(text)` | Ponto de código do primeiro caractere (pares substitutos lidos como um só). |
| `UPPER` | `UPPER(text)` | Converte o texto para maiúsculas. |
| `VALUE` | `VALUE(text)` | Converte texto em número. |
| `VALUETOTEXT` | `VALUETOTEXT(value, [format])` | Valor como texto — formato 0 conciso (padrão), 1 estrito (texto entre aspas); erros viram seu texto de exibição. |

## Pesquisa e referência (17)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ADDRESS` | `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])` | O endereço da célula como TEXTO (`abs_num` 1-4 → `$C$2`/`C$2`/`$C2`/`C2`); `a1=FALSE` renderiza apenas a forma R1C1 absoluta (`R2C3` — R1C1 relativo → `#VALUE!`); `sheet_text` entra como prefixo, entre aspas quando necessário. |
| `AREAS` | `AREAS(reference)` | Número de áreas (intervalos contíguos ou células individuais) na referência — uma verificação sintática, como `ISREF`; não referência → `#VALUE!`, e um argumento que não consegue ser resolvido informa o próprio erro (`AREAS(NoSuchName)` → `#NAME?`). |
| `CHOOSE` | `CHOOSE(index_num, value1, [value2], …)` | O valor na posição `index_num` (truncado); avaliação preguiçosa — apenas o argumento escolhido é avaliado; um intervalo escolhido permanece *range-aware* (`SUM(CHOOSE(…))`); fora do intervalo → `#VALUE!`. |
| `COLUMN` | `COLUMN([reference])` | Número da coluna da referência (a coluna mais à esquerda para um intervalo) — ou da célula atual, quando chamada sem argumento. Aceita QUALQUER expressão que produza uma referência, não apenas uma referência literal: um nome definido, `INDEX`/`OFFSET`/`INDIRECT`/`CHOOSE`, um intervalo `:` com extremidades que retornam referências (`COLUMN(INDEX(A1:C1,1,2))` = 2). Uma referência de coluna/linha inteira usa o limite DECLARADO (`COLUMN(A:A)` = 1, `COLUMN(1:1)` = 1) — enquanto o `COLUMNS` da linha ao lado usa a extensão POPULADA em um eixo aberto; um argumento que não é referência (ou uma união) → `#VALUE!`, e um argumento que não consegue ser resolvido informa o próprio erro (`#NAME?`, `#REF!`). Em [posição de array](workbook-and-expressions.md#argumentos-implícitos-de-array) produz o vetor inteiro de números de coluna, sobre um intervalo literal *e* sobre um nome ou um intervalo `:` que denote um (`SUM(COLUMN(A1:C3))` = 18, `SUM(COLUMN(MyName))` = 3 para um nome de coluna única sobre três linhas); ali um intervalo aberto é recusado e um argumento que é uma *função* retornando referência permanece escalar. |
| `COLUMNS` | `COLUMNS(range)` | Número de colunas do intervalo. Sobre uma [referência de coluna/linha inteira](workbook-and-expressions.md#referências-de-coluna-e-linha-inteira), um eixo de coluna limitado é exato (`COLUMNS(A:C)` = 3), um aberto usa a extensão populada. Um argumento que não consegue ser resolvido para uma referência informa o próprio erro (`#NAME?`, `#REF!`); um valor escalar simples conta como 1 (um array 1x1). |
| `FORMULATEXT` | `FORMULATEXT(reference)` | A fórmula da célula referenciada como TEXTO, com o `=` incluído (reescrita — *unparse* — no contexto de planilha da célula referenciada); uma célula literal ou vazia → `#N/A`. |
| `HLOOKUP` | `HLOOKUP(lookup_value, table_range, row_index_num, [range_lookup])` | Pesquisa horizontal na primeira linha de uma tabela; exata ou aproximada; `row_index_num` < 1 → `#VALUE!`, além da tabela → `#REF!`. |
| `INDEX` | `INDEX(range, row_num, [column_num])` | O valor em uma posição (base 1) dentro de um intervalo. Aceita um [primeiro argumento implícito de array](workbook-and-expressions.md#argumentos-implícitos-de-array) (`INDEX(ROW(B2:B5),1)` = 2), incluindo a identidade `INDEX(ROW($A:$A), n)` que retorna `n` sem materializar a coluna; fora do intervalo → `#REF!`. |
| `INDIRECT` | `INDIRECT(ref_text, [a1])` | A referência nomeada por `ref_text`, resolvida em tempo de avaliação: o texto é interpretado como um corpo de fórmula no contexto da **planilha atual**, de modo que um `"A1"` sem qualificação significa o `A1` da planilha chamadora, enquanto `"Data!B2"` atravessa planilhas. Uma única célula é desreferenciada para o seu **valor** (`INDIRECT("A1")`); um resultado multicélula permanece uma **referência** para consumidores que entendem intervalos (`SUM(INDIRECT("A1:A3"))`, `ROWS(INDIRECT("A1:A3"))` = 3, e como extremidade de `:` — `SUM(INDIRECT("A1"):A3)`), de modo que, sozinha em uma célula, ela recebe [interseção implícita na fronteira da célula](workbook-and-expressions.md#interseção-implícita-na-fronteira-da-célula) (`=INDIRECT("A1:A3")` escrita em `B2` mostra `A2`; em `B5`, `#VALUE!`). [Nomes definidos](workbook-and-expressions.md#intervalos-nomeados) também são resolvidos, inclusive um montado em tempo de execução (`SUM(INDIRECT("R"&"ng"))`, `INDIRECT("Data!A"&2)`). **Volátil**: a referência só é conhecida em tempo de avaliação, então a célula é sempre recalculada — veja [Funções voláteis](workbook-and-expressions.md#funções-voláteis). Tudo o que falha vira `#REF!`, nunca `#NAME?`: `a1` = `FALSE`/`0` (o estilo R1C1 não é suportado — apenas A1) ou um `a1` que não seja um número; um `ref_text` que não seja texto (um número, um lógico, uma célula numérica); um texto que não seja interpretável; e um nome ou uma planilha desconhecidos. |
| `LOOKUP` | `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Forma vetorial (sempre aproximada: o maior valor ≤ pesquisado); a forma matricial de 2 argumentos busca na primeira linha e retorna da última linha quando o intervalo é mais largo que alto; caso contrário, primeira/última coluna. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | Posição (base 1) de um valor em um intervalo (`match_type`: 1 aproximado crescente — padrão, 0 exato, -1 aproximado decrescente). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | Uma referência deslocada (e opcionalmente redimensionada) a partir de uma referência inicial; pode retornar uma referência multicélula para consumidores que aceitam intervalos. |
| `ROW` | `ROW([reference])` | Número da linha da referência (a linha superior para um intervalo) — ou da célula atual, quando chamada sem argumento. Aceita QUALQUER expressão que produza uma referência, não apenas uma referência literal: um nome definido, `INDEX`/`OFFSET`/`INDIRECT`/`CHOOSE`, um intervalo `:` com extremidades que retornam referências (`ROW(INDEX(A1:A3,2,1))` = 2). Uma referência de coluna/linha inteira usa o limite DECLARADO (`ROW(A:A)` = 1, `ROW(A2:A)` = 2) — enquanto o `ROWS` da linha ao lado usa a extensão POPULADA em um eixo aberto; um argumento que não é referência (ou uma união) → `#VALUE!`, e um argumento que não consegue ser resolvido informa o próprio erro (`#NAME?`, `#REF!`). Em [posição de array](workbook-and-expressions.md#argumentos-implícitos-de-array) produz o vetor inteiro de números de linha, sobre um intervalo literal *e* sobre um nome ou um intervalo `:` que denote um (`SUM(ROW(MyName))` = 6 para um nome sobre três linhas); ali um intervalo aberto é recusado e um argumento que é uma *função* retornando referência permanece escalar. |
| `ROWS` | `ROWS(range)` | Número de linhas do intervalo. Sobre uma [referência de coluna/linha inteira](workbook-and-expressions.md#referências-de-coluna-e-linha-inteira), um eixo de linha aberto usa a extensão populada (`ROWS(A:A)` = linha populada máxima − linha populada mínima + 1, 0 se vazia — uma divergência documentada em relação à grade fixa do Excel), um limitado é exato (`ROWS(1:5)` = 5). Um argumento que não consegue ser resolvido para uma referência informa o próprio erro (`ROWS(NoSuchName)` → `#NAME?`, `ROWS(INDIRECT("zz"))` → `#REF!`); um valor escalar simples conta como 1 (um array 1x1). |
| `VLOOKUP` | `VLOOKUP(lookup_value, table_range, col_index_num, [range_lookup])` | Pesquisa vertical na primeira coluna de uma tabela; exata ou aproximada. |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])` | Pesquisa moderna, com contingência para "não encontrado" e modos de correspondência/busca. |
| `XMATCH` | `XMATCH(lookup_value, lookup_range, [match_mode], [search_mode])` | Posição (base 1) com os modos do `XLOOKUP` (0 exato — padrão, -1 exato-ou-menor, 1 exato-ou-maior, 2 curinga; busca 1/-1). |

## Informações (18)

As funções `IS*` inspecionam o valor avaliado sem coerção (`ISNUMBER("19")` é `FALSE`) e nunca propagam
erros — elas os relatam.

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ERROR.TYPE` | `ERROR.TYPE(error_val)` | `#NULL!`=1, `#DIV/0!`=2, `#VALUE!`=3, `#REF!`=4, `#NAME?`=5, `#NUM!`=6, `#N/A`=7; não erro → `#N/A`. |
| `ISBLANK` | `ISBLANK(value)` | `TRUE` para um valor em branco. |
| `ISERR` | `ISERR(value)` | `TRUE` para qualquer erro, exceto `#N/A`. |
| `ISERROR` | `ISERROR(value)` | `TRUE` para qualquer valor de erro. |
| `ISEVEN` | `ISEVEN(number)` | `TRUE` para um número par (truncado antes); não numérico → `#VALUE!`. |
| `ISFORMULA` | `ISFORMULA(reference)` | `TRUE` quando a célula referenciada contém uma fórmula (e não um literal simples); não referência → `#VALUE!`. |
| `ISLOGICAL` | `ISLOGICAL(value)` | `TRUE` para um valor lógico. |
| `ISNA` | `ISNA(value)` | `TRUE` apenas para `#N/A`. |
| `ISNONTEXT` | `ISNONTEXT(value)` | `TRUE` para qualquer coisa que não seja texto (incluindo valores em branco). |
| `ISNUMBER` | `ISNUMBER(value)` | `TRUE` para um valor numérico. |
| `ISODD` | `ISODD(number)` | `TRUE` para um número ímpar (truncado antes); não numérico → `#VALUE!`. |
| `ISREF` | `ISREF(value)` | `TRUE` quando o argumento é uma referência (célula/intervalo/união) — uma verificação sintática, independentemente do valor. |
| `ISTEXT` | `ISTEXT(value)` | `TRUE` para texto. |
| `N` | `N(value)` | Número → ele mesmo; `TRUE`→1/`FALSE`→0; erro → o próprio erro; qualquer outra coisa → 0. |
| `NA` | `NA()` | O valor de erro `#N/A`. |
| `SHEET` | `SHEET([value])` | Posição (base 1, na ordem das abas) da planilha de uma referência ou de um nome de planilha — ou da planilha atual, sem argumento. |
| `SHEETS` | `SHEETS()` | Número de planilhas no workbook (a forma com referência 3-D não se aplica: toda referência abrange uma única planilha). |
| `TYPE` | `TYPE(value)` | 1 número (incluindo em branco), 2 texto, 4 lógico, 16 erro (inspecionado, não propagado), 64 referência multicélula. |

## Financeiras (55)

Semântica padrão de valor do dinheiro no tempo: `rate` por período, `nper` é o total de períodos, `type`
0 = fim do período (padrão) / 1 = início. As funções de títulos, cupom e fluxo de caixa datado recebem
**seriais de data** (construa-os com `DATE`, exatamente como nas funções de data) e uma `basis` de
contagem de dias: 0 = US (NASD) 30/360 (padrão), 1 = real/real, 2 = real/360, 3 = real/365, 4 = europeu
30/360. A `frequency` do cupom é 1 (anual), 2 (semestral) ou 4 (trimestral). As datas de cupom são
construídas andando **para trás a partir do vencimento**. Resultados iterativos (`RATE`, `IRR`, `XIRR`,
`YIELD`, `ODDFYIELD`) usam o mesmo solver robusto de bracketing + bisseção que `RATE`/`IRR` (validado
contra um caso rígido de 30 anos). Os valores de referência de toda a família são conferidos contra o
oráculo `ExcelFinancialFunctions`. Violações de domínio mapeiam para `#NUM!` (liquidação ≥ vencimento,
frequency ∉ {1,2,4}, basis ∉ 0..4, etc.).

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `FV` | `FV(rate, nper, pmt, [pv], [type])` | Valor futuro de um investimento. |
| `IPMT` | `IPMT(rate, per, nper, pv, [fv], [type])` | Parcela de juros de um dado período de pagamento. |
| `IRR` | `IRR(values, [guess])` | Taxa interna de retorno de um intervalo de fluxos de caixa. |
| `NPER` | `NPER(rate, pmt, pv, [fv], [type])` | Número de períodos de pagamento. |
| `NPV` | `NPV(rate, value1, …)` | Valor presente líquido de fluxos de caixa futuros; aceita intervalos. |
| `PMT` | `PMT(rate, nper, pv, [fv], [type])` | Pagamento periódico constante de um empréstimo/anuidade. |
| `PPMT` | `PPMT(rate, per, nper, pv, [fv], [type])` | Parcela de principal de um dado período de pagamento. |
| `PV` | `PV(rate, nper, pmt, [fv], [type])` | Valor presente de um investimento. |
| `RATE` | `RATE(nper, pmt, pv, [fv], [type], [guess])` | Taxa de juros por período (iterativa). |
| `SLN` | `SLN(cost, salvage, life)` | Depreciação linear por período. |
| `SYD` | `SYD(cost, salvage, life, per)` | Depreciação pela soma dos dígitos dos anos. |
| `DB` | `DB(cost, salvage, life, period, [month])` | Depreciação por saldo decrescente fixo. |
| `DDB` | `DDB(cost, salvage, life, period, [factor])` | Depreciação por saldo decrescente em dobro. |
| `VDB` | `VDB(cost, salvage, life, start, end, [factor], [no_switch])` | Depreciação por saldo decrescente variável. |
| `AMORLINC` | `AMORLINC(cost, purchased, first_period, salvage, period, rate, [basis])` | Depreciação linear francesa (proporcional). |
| `AMORDEGRC` | `AMORDEGRC(cost, purchased, first_period, salvage, period, rate, [basis])` | Depreciação decrescente francesa com coeficiente baseado na vida útil. |
| `EFFECT` | `EFFECT(nominal_rate, npery)` | Taxa de juros anual efetiva. |
| `NOMINAL` | `NOMINAL(effect_rate, npery)` | Taxa de juros anual nominal. |
| `MIRR` | `MIRR(values, finance_rate, reinvest_rate)` | Taxa interna de retorno modificada. |
| `RRI` | `RRI(nper, pv, fv)` | Taxa de juros equivalente para o crescimento de um investimento. |
| `PDURATION` | `PDURATION(rate, pv, fv)` | Períodos para um investimento atingir um valor. |
| `ISPMT` | `ISPMT(rate, per, nper, pv)` | Juros pagos durante um período de empréstimo linear. |
| `CUMIPMT` | `CUMIPMT(rate, nper, pv, start, end, type)` | Juros acumulados em um intervalo de períodos. |
| `CUMPRINC` | `CUMPRINC(rate, nper, pv, start, end, type)` | Principal acumulado em um intervalo de períodos. |
| `FVSCHEDULE` | `FVSCHEDULE(principal, schedule)` | Valor futuro após uma série de taxas compostas. |
| `DOLLARDE` | `DOLLARDE(fractional_dollar, fraction)` | Preço em notação fracionária → decimal. |
| `DOLLARFR` | `DOLLARFR(decimal_dollar, fraction)` | Preço decimal → notação fracionária. |
| `XNPV` | `XNPV(rate, values, dates)` | Valor presente líquido de fluxos de caixa datados (real/365). |
| `XIRR` | `XIRR(values, dates, [guess])` | Taxa interna de retorno de fluxos de caixa datados. |
| `ACCRINT` | `ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])` | Juros acumulados de um título com juros periódicos. |
| `ACCRINTM` | `ACCRINTM(issue, settlement, rate, par, [basis])` | Juros acumulados de um título que paga no vencimento. |
| `DISC` | `DISC(settlement, maturity, pr, redemption, [basis])` | Taxa de desconto de um título. |
| `INTRATE` | `INTRATE(settlement, maturity, investment, redemption, [basis])` | Taxa de juros de um título totalmente investido. |
| `RECEIVED` | `RECEIVED(settlement, maturity, investment, discount, [basis])` | Valor recebido no vencimento. |
| `PRICEDISC` | `PRICEDISC(settlement, maturity, discount, redemption, [basis])` | Preço por $100 de um título descontado. |
| `PRICEMAT` | `PRICEMAT(settlement, maturity, issue, rate, yld, [basis])` | Preço por $100 de um título com juros pagos no vencimento. |
| `YIELDDISC` | `YIELDDISC(settlement, maturity, pr, redemption, [basis])` | Rendimento anual de um título descontado. |
| `YIELDMAT` | `YIELDMAT(settlement, maturity, issue, rate, pr, [basis])` | Rendimento anual de um título com juros pagos no vencimento. |
| `TBILLEQ` | `TBILLEQ(settlement, maturity, discount)` | Rendimento equivalente a título de uma letra do Tesouro (T-bill). |
| `TBILLPRICE` | `TBILLPRICE(settlement, maturity, discount)` | Preço por $100 de uma letra do Tesouro (T-bill). |
| `TBILLYIELD` | `TBILLYIELD(settlement, maturity, pr)` | Rendimento de uma letra do Tesouro (T-bill). |
| `COUPPCD` | `COUPPCD(settlement, maturity, frequency, [basis])` | Data do cupom anterior à liquidação. |
| `COUPNCD` | `COUPNCD(settlement, maturity, frequency, [basis])` | Data do próximo cupom após a liquidação. |
| `COUPNUM` | `COUPNUM(settlement, maturity, frequency, [basis])` | Número de cupons entre a liquidação e o vencimento. |
| `COUPDAYS` | `COUPDAYS(settlement, maturity, frequency, [basis])` | Dias no período do cupom que contém a liquidação. |
| `COUPDAYBS` | `COUPDAYBS(settlement, maturity, frequency, [basis])` | Dias do início do período até a liquidação. |
| `COUPDAYSNC` | `COUPDAYSNC(settlement, maturity, frequency, [basis])` | Dias da liquidação até o próximo cupom. |
| `PRICE` | `PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])` | Preço por $100 de um título com cupons periódicos. |
| `YIELD` | `YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])` | Rendimento até o vencimento (iterativo). |
| `DURATION` | `DURATION(settlement, maturity, coupon, yld, frequency, [basis])` | Duração de Macaulay em anos. |
| `MDURATION` | `MDURATION(settlement, maturity, coupon, yld, frequency, [basis])` | Duração de Macaulay modificada. |
| `ODDFPRICE` | `ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])` | Preço de um título com primeiro período irregular. |
| `ODDFYIELD` | `ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])` | Rendimento de um título com primeiro período irregular. |
| `ODDLPRICE` | `ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])` | Preço de um título com último período irregular. |
| `ODDLYIELD` | `ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])` | Rendimento de um título com último período irregular. |

## Data e hora (25)

Datas são **números seriais** (`double`), exatamente como no Excel: a parte inteira conta os dias e a fração
é a hora do dia. As funções de data recebem seriais numéricos (construa-os com `DATE`/`TIME`, ou texto
numérico que o `CoerceToNumber` já aceita); elas **não** fazem o parse implícito de *strings* de data (use
`DATEVALUE`/`TIMEVALUE` para isso). Um serial negativo está fora do intervalo → `#NUM!`. `TODAY`/`NOW` leem o
relógio e são **voláteis** — veja [Funções voláteis](workbook-and-expressions.md#funções-voláteis).

**Onde a contagem começa.** O serial **1** é **1900-01-01**, o serial 2 é 1900-01-02, e assim por diante —
`YEAR(1)`/`MONTH(1)`/`DAY(1)` dão 1900/1/1. O serial **0** é o *dia zero* do Excel: não é uma data real, mas um
marcador que o Excel escreve como `1900-01-00`, então `DAY(0)` = 0 e `DATE(1900,1,0)` = 0. Nada existe abaixo
dele — `DATE(1900,1,-1)` é `#NUM!` e `DATEVALUE("1899-12-31")` é `#VALUE!`.

**O serial 60 é um dia que nunca existiu.** O Lotus 1-2-3 tratava 1900 como ano bissexto; o Excel copiou o bug
para continuar compatível com os arquivos dele e carrega um **1900-02-29** fantasma no serial 60 desde então.
O MySheet carrega o mesmo fantasma pelo mesmo motivo — para que um serial signifique a mesma data nos dois
programas. Mas 1900 não foi bissexto, então as funções não concordam sobre o que o serial 60 *é*:

- **O `DATE` não consegue construí-lo.** `DATE(1900,2,29)` avança para março e dá **61**, e tanto
  `DATE(1900,2,28)` quanto `DATE(1900,3,0)` dão 59. Só duas coisas chegam ao 60: aritmética sobre seriais
  (59 + 1) e `DATEVALUE("1900-02-29")` = **60** (também `"2/29/1900"` e `"29-Feb-1900"`).
- **As funções de calendário o leem como 1900-02-28.** `DAY(60)` = 28, `EDATE(60,0)` = `EOMONTH(60,0)` = 59 e
  `WEEKDAY(60)` = 3 — o serial 60 colapsa sobre o serial 59 e informa o dia da semana dele.
- **O `TEXT` e as contagens 30/360 veem um 29 de fevereiro de verdade**, porque para eles fevereiro de 1900
  tem 29 dias: `TEXT(60,"yyyy-mm-dd")` = `1900-02-29`, `DAYS360(59,61)` = 3, `YEARFRAC(60,61)` = 2/360.
- **Tudo o que conta dias conta seriais**, então um intervalo que contém o serial 60 é um dia mais longo do
  que o calendário gregoriano entre as mesmas duas datas: `DATEDIF(1,61,"D")` = 60, `NETWORKDAYS(1,61)` = 45.

A partir do serial **61** (1900-03-01) — onde vive toda data que uma planilha real guarda — o dia fantasma
ficou para trás e nada disso se aplica.

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `DATE` | `DATE(year, month, day)` | Serial a partir das partes; overflow do Excel (mês 13 → janeiro seguinte, dia 0 → fim do mês anterior); ano 0–1899 soma 1900. |
| `DATEDIF` | `DATEDIF(start, end, unit)` | Diferença em `"Y"`/`"M"`/`"D"`/`"MD"`/`"YM"`/`"YD"`; `start > end` → `#NUM!`. `"MD"` é oficialmente não confiável. `"MD"`/`"YD"` contam seriais a partir de `start` empurrado para frente por cada mês (ano) **inteiro** do intervalo, com o deslocamento limitado ao último dia do mês de destino (`DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")` = 1). |
| `DATEVALUE` | `DATEVALUE(date_text)` | Converte uma string de data (invariant `yyyy-MM-dd`, `M/d/yyyy`, `d-MMM-yyyy`, …) em um serial de dia inteiro; não interpretável → `#VALUE!`, e o mesmo vale para qualquer data **anterior a 1900-01-01** (`DATEVALUE("1899-12-31")`). O `"1900-02-29"` fantasma do Excel é interpretado, como serial 60. |
| `DAY` | `DAY(serial)` | Dia do mês (1–31). |
| `DAYS` | `DAYS(end, start)` | Dias inteiros entre duas datas (pode ser negativo). |
| `DAYS360` | `DAYS360(start, end, [method])` | Contagem de dias 30/360; padrão US (NASD), `TRUE` = europeu. O método US **não tem ajuste de fim de fevereiro** e **não avança um `end` que caia no fim do mês para o dia 1º do mês seguinte** — ele acompanha o Excel conforme medido, não a página da Microsoft (veja as notas abaixo da tabela). |
| `EDATE` | `EDATE(start, months)` | O mesmo dia do mês, `months` à frente/atrás, limitado ao fim do mês; um resultado anterior ao serial 0 → `#NUM!`. |
| `EOMONTH` | `EOMONTH(start, months)` | Último dia do mês `months` à frente/atrás de `start`; um resultado anterior ao serial 0 → `#NUM!` (`EOMONTH(1,-1)` = 0 continua válido). |
| `HOUR` | `HOUR(serial)` | Hora (0–23) da fração de tempo. |
| `ISOWEEKNUM` | `ISOWEEKNUM(serial)` | Número da semana ISO 8601 (semanas começam na segunda-feira; a semana 1 contém a primeira quinta-feira). |
| `MINUTE` | `MINUTE(serial)` | Minuto (0–59) da fração de tempo. |
| `MONTH` | `MONTH(serial)` | Mês (1–12). |
| `NETWORKDAYS` | `NETWORKDAYS(start, end, [holidays])` | Dias úteis em `[start, end]` (inclusivo); sábado/domingo e `holidays` são excluídos. |
| `NETWORKDAYS.INTL` | `NETWORKDAYS.INTL(start, end, [weekend], [holidays])` | `NETWORKDAYS` com um fim de semana personalizável (número 1–7/11–17 ou uma máscara de 7 caracteres `"0000011"`). Um número de fim de semana fora da tabela → `#NUM!`; uma máscara `"1111111"` (todos os dias de descanso) é um **0** legítimo aqui, não um erro — ao contrário do `WORKDAY.INTL`. |
| `NOW` | `NOW()` | Volátil: a data **e** hora local atuais como um serial. Veja [Funções voláteis](workbook-and-expressions.md#funções-voláteis). |
| `SECOND` | `SECOND(serial)` | Segundo (0–59), arredondado ao segundo mais próximo. |
| `TIME` | `TIME(hour, minute, second)` | Fração de hora do dia; componentes 0–32767 rolam, aplicados módulo 24h; negativo → `#NUM!`. |
| `TIMEVALUE` | `TIMEVALUE(time_text)` | Converte uma string de hora (`HH:mm[:ss]`, `h:mm[:ss] AM/PM`) em uma fração `[0,1)`; não interpretável → `#VALUE!`. |
| `TODAY` | `TODAY()` | Volátil: a data local atual como um serial de dia inteiro (o piso de `NOW()`). Veja [Funções voláteis](workbook-and-expressions.md#funções-voláteis). |
| `WEEKDAY` | `WEEKDAY(serial, [return_type])` | Dia da semana; `return_type` 1/2/3 e 11–17 (veja a tabela do WEEKDAY). O dia da semana é o do Excel, herdado do Lotus e lido direto do serial: `WEEKDAY(1)` = 1 (domingo) embora 1900-01-01 tenha sido de fato uma segunda-feira — igual ao Excel. |
| `WEEKNUM` | `WEEKNUM(serial, [return_type])` | Semana do ano; Sistema 1 para 1/2/11–17, ISO 8601 (Sistema 2) para 21. |
| `WORKDAY` | `WORKDAY(start, days, [holidays])` | Data `days` dias úteis a partir de `start` (start excluído); negativo anda para trás. |
| `WORKDAY.INTL` | `WORKDAY.INTL(start, days, [weekend], [holidays])` | `WORKDAY` com um fim de semana personalizável. São três respostas distintas, todas medidas no Aspose.Cells 26.6.0 (2026-09-09): um **número** de fim de semana fora de 1–7/11–17 → `#NUM!`; uma máscara de **fim de semana total** `"1111111"` → `#VALUE!`, não o `#NUM!` que a página da Microsoft sugere; e `days` = 0 nunca sai do lugar, então responde o serial de `start` mesmo sob uma máscara de fim de semana total (`WORKDAY.INTL(45366,0,"1111111")` = 45366). |
| `YEAR` | `YEAR(serial)` | Ano civil (1900–9999). |
| `YEARFRAC` | `YEARFRAC(start, end, [basis])` | Fração do ano na base 0 (US 30/360), 1 (real/real), 2 (real/360), 3 (real/365), 4 (europeu 30/360). A base 0 **não tem regra de fim de fevereiro** e só puxa um `start` no fim de fevereiro para o dia 30 **depois** de testar um `end` no dia 31, então ela discorda deliberadamente do `DAYS360` em alguns pares de fim de fevereiro (veja as notas abaixo da tabela). |

**O `TEXT` é o único lugar em que um 1900-02-29 é impresso** (`TEXT(60,"yyyy-mm-dd")` = `1900-02-29`) **e o
único em que um dia zero aparece** (`TEXT(0,"yyyy-mm-dd")` = `1900-01-00`). A única exceção dentro de um
formato é `ddd`/`dddd`: eles nomeiam o dia da semana do Lotus, e pedir um deles também puxa o *número* do dia
impresso de volta para 28 de fevereiro — `TEXT(60,"yyyy-mm-dd dddd")` = `1900-02-28 Tuesday`, enquanto
`TEXT(60,"yyyy-mm-dd")` mantém o 29. Um serial negativo é `#VALUE!` no `TEXT`, não o `#NUM!` que as funções de
data respondem.

**Os dias úteis em janeiro–fevereiro de 1900 caminham pelo calendário real, não pelo do Excel.** `WORKDAY`,
`WORKDAY.INTL`, `NETWORKDAYS` e `NETWORKDAYS.INTL` batem com o Excel exatamente a partir do serial 61
(1900-03-01) — toda data que uma planilha real guarda. Abaixo disso, as próprias respostas do Excel não são
coerentes entre si, ou seja, não há regra alguma a reproduzir, e a caminhada simplesmente continua seguindo o
dia da semana gregoriano real que ela segue em todo o resto — acima do serial 61 os dois concordam, abaixo dele
podem divergir. Medido no Aspose.Cells 26.6.0 (2026-09-09), o Excel:

1. dá o mesmo dia para o 4º e para o 5º dia útil a partir do mesmo início — `WORKDAY(6,4)` = `WORKDAY(6,5)` =
   12 — ou seja, o resultado dele não é uma função de `days`;
2. faz o mesmo com um fim de semana de um único dia: `WORKDAY.INTL(1,5,"1000000")` =
   `WORKDAY.INTL(1,6,"1000000")` = 7;
3. responde `WORKDAY(58,4)` = 64, um serial que ele *mesmo* trata como fim de semana (`WEEKDAY(64)` = 1,
   `NETWORKDAYS(64,64)` = 0, `TEXT(64,"dddd")` = Sunday);
4. e não é aditivo quando o intervalo é dividido: `NETWORKDAYS(58,62)` = 4, enquanto `NETWORKDAYS(58,58)` +
   `NETWORKDAYS(59,61)` + `NETWORKDAYS(62,62)` = 1 + 3 + 1 = 5 — um total que nenhum critério de dia útil
   avaliado dia a dia consegue produzir.

**A divergência abaixo do serial 61 não é uma regra só, nem um punhado fixo de linhas.** Há um dia
visivelmente envolvido — 1900-01-05 é uma **sexta-feira** no calendário real e uma quinta-feira no dia da
semana do Lotus usado pelo Excel — mas isso é uma observação sobre um dia, não a explicação da janela:
enquanto isso era decidido, catorze regras candidatas foram ajustadas a uma varredura de 580 linhas de
`WORKDAY` e a melhor delas ainda errava **17** dessas linhas. Então as linhas abaixo são **exemplos** da
divergência, não um conjunto exaustivo — MySheet primeiro, Excel depois: `WORKDAY(5,1)` **8** / 6,
`WORKDAY(6,1)` **8** / 9, `WORKDAY(6,4)` **11** / 12, `WORKDAY(13,1)` **15** / 16. Outras formas abaixo do
serial 61 também divergem, entre elas as quatro autocontradições acima: `WORKDAY(58,4)` **62** / 64,
`NETWORKDAYS(58,62)` **5** / 4, `WORKDAY.INTL(1,5,"1000000")` **6** / 7 e `NETWORKDAYS(58,62,H)` com `H`
guardando o serial 59 **4** / 3.

O que se sustenta sem ressalva é apenas o próprio limite: a partir do serial 61, toda resposta corresponde.
Abaixo dele algumas linhas concordam, entre elas `WORKDAY(59,1)` = 60, `WORKDAY(60,-1)` = 59,
`NETWORKDAYS(59,61)` = 3 e `NETWORKDAYS(1,61)` = 45 — mas a concordância ali é um fato linha por linha, não uma
regra que se possa estender. Dois feriados em torno do dia fantasma ilustram isso: `NETWORKDAYS(58,62,H)` com
`H` guardando os seriais 59 e 60 dá **3** aqui contra 2, mais uma vez a mesma autocontradição, porque o Aspose
volta a responder menos que a soma das próprias partes. Todo RESULTADO DE FÓRMULA dos parágrafos acima é medido
(Aspose.Cells 26.6.0, 2026-09-09, entrada simples na célula) e fixado por um teste, inclusive o lado do MySheet;
as contagens de regras candidatas e de linhas varridas vêm da varredura do próprio plano da fase e não são
asserções linha a linha.

**Três mudanças da 3.17.0 que não têm nada a ver com 1900** — elas movem resultados em datas modernas também:

- **A base 0 do `YEARFRAC` perdeu a regra de fim de fevereiro, e seus dois passos 30/360 trocaram de ordem.**
  Um `end` no último dia de fevereiro não é mais promovido a um dia 30 nominal, então
  `YEARFRAC(DATE(2024,2,29),DATE(2025,2,28),0)` é 358/360 e `YEARFRAC(DATE(2023,2,28),DATE(2024,2,29),0)` é
  359/360 — ambos davam exatamente 1 antes da 3.17.0. A puxada de um `start` no fim de fevereiro para o dia 30
  continua existindo, mas agora roda *depois* do teste do `end` no dia 31 e por isso não arrasta mais um `end`
  no dia 31 junto com ela: `YEARFRAC(DATE(2023,2,28),DATE(2023,3,31),0)` = 31/360 (era 30/360).
- **O `DAYS360` (US) perdeu por completo o avanço para o dia 1º do mês seguinte** e ordena sua puxada de
  fevereiro ao *contrário* da base 0 do `YEARFRAC`. Um `end` no fim do mês não avança mais para o dia 1º do mês
  seguinte: `DAYS360(DATE(2011,1,1),DATE(2011,4,30))` = 119 (era 120),
  `DAYS360(DATE(2011,1,15),DATE(2011,9,30))` = 255 (era 256), `DAYS360(DATE(2024,1,31),DATE(2024,2,29))` = 29
  (era 30), `DAYS360(DATE(2024,1,16),DATE(2024,2,29))` = 43 (era 45),
  `DAYS360(DATE(2024,2,28),DATE(2024,2,29))` = 1 (era 3). Um `end` no dia 31 continua caindo para 30 quando o
  `start` ajustado já chegou a 30, então `DAYS360(DATE(2011,1,1),DATE(2011,12,31))` = 360 não muda. E, como
  aqui a puxada de fevereiro roda *antes*, ela arrasta um `end` no dia 31 junto:
  `DAYS360(DATE(2023,2,28),DATE(2023,3,31))` = 30, contra os 31/360 do `YEARFRAC` acima — as duas funções agora
  **discordam deliberadamente** em pares de fim de fevereiro e de dia 31, exatamente como o Excel faz.
- **O `"MD"` e o `"YD"` do `DATEDIF` se ancoram de outra forma.** Os dois agora contam seriais a partir de
  `start` empurrado para frente por cada mês (ano) inteiro do intervalo, com esse deslocamento limitado ao
  último dia do mês de destino, em vez de emprestar o tamanho do mês anterior:
  `DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")` = 1, onde a fórmula antiga respondia −1, e
  `DATEDIF(DATE(2024,2,29),DATE(2025,3,1),"YD")` = 1.

## Compatibilidade — aliases legados (11)

Os nomes anteriores a 2010 das funções estatísticas modernas. Cada alias é um **nó de AST
distinto**, não apenas uma grafia alternativa do registro moderno: ele é avaliado exatamente como o
seu equivalente moderno, mas `FORMULATEXT`, a serialização e a exportação para xlsx preservam a
grafia que você escreveu — `STDEV(…)` nunca vira `STDEV.S(…)`. (`CONCATENATE` e o `FLOOR` legado,
também presentes na categoria Compatibilidade da Microsoft, estão documentados em suas seções de
Texto/Matemática.)

| Função | Argumentos | Equivalente moderno |
| --- | --- | --- |
| `COVAR` | `COVAR(array1, array2)` | `COVARIANCE.P` |
| `FORECAST` | `FORECAST(x, known_ys, known_xs)` | `FORECAST.LINEAR` |
| `MODE` | `MODE(number1, …)` | `MODE.SNGL` |
| `PERCENTILE` | `PERCENTILE(array, k)` | `PERCENTILE.INC` |
| `PERCENTRANK` | `PERCENTRANK(array, x, [significance])` | `PERCENTRANK.INC` |
| `QUARTILE` | `QUARTILE(array, quart)` | `QUARTILE.INC` |
| `RANK` | `RANK(number, ref, [order])` | `RANK.EQ` |
| `STDEV` | `STDEV(number1, …)` | `STDEV.S` |
| `STDEVP` | `STDEVP(number1, …)` | `STDEV.P` |
| `VAR` | `VAR(number1, …)` | `VAR.S` |
| `VARP` | `VARP(number1, …)` | `VAR.P` |

## Cobertura de funções do Excel

O MySheet implementa 306 das ~520 funções do [catálogo oficial de funções do Excel da
Microsoft](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
agrupadas abaixo pelas próprias categorias da Microsoft (✅ implementada, ⬜ ainda não, ✖ fora de escopo
por design). **35 funções estão permanentemente fora de escopo** — elas dependem de serviços externos, do
ambiente de interface do aplicativo ou de recursos que a engine deliberadamente não modela (veja
[Fora de escopo](#fora-de-escopo-por-design) abaixo) — restando um catálogo viável de ~485 funções, sobre
o qual o roadmap é acompanhado. Alguns nomes são listados pela Microsoft em mais de uma categoria —
`CONCATENATE` (Texto e Compatibilidade), `FLOOR` (Matemática e Compatibilidade) e `FORECAST`
(Estatísticas e Compatibilidade) —, então as contagens por categoria não somam um total único, e uma
categoria abaixo pode ser **maior** que a tabela de mesmo nome acima, que documenta cada função uma
única vez. Esses três nomes são toda a diferença: Estatísticas aparece como 60 abaixo, contra uma tabela
de 59, porque `FORECAST` está documentada junto aos aliases de compatibilidade; e Compatibilidade
aparece como 13, contra uma tabela de 11, porque `CONCATENATE` e `FLOOR` estão documentadas em Texto e
em Matemática. Veja o
[`FunctionRegistry.cs`](../../Danfma.MySheet/Parsing/FunctionRegistry.cs) para a lista registrada
oficial.

<details open>
<summary><strong>Financeiras</strong> — 55/55</summary>

✅ `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FV` `FVSCHEDULE` `INTRATE` `IPMT` `IRR` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `NPER` `NPV` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PMT` `PPMT` `PRICE` `PRICEDISC` `PRICEMAT` `PV` `RATE` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

</details>

<details open>
<summary><strong>Lógicas</strong> — 12/19</summary>

✅ `AND` `FALSE` `IF` `IFERROR` `IFNA` `IFS` `LET` `NOT` `OR` `SWITCH` `TRUE` `XOR`

⬜ `BYCOL` `BYROW` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN`

</details>

<details open>
<summary><strong>Pesquisa e referência</strong> — 17/40</summary>

✅ `ADDRESS` `AREAS` `CHOOSE` `COLUMN` `COLUMNS` `FORMULATEXT` `HLOOKUP` `INDEX` `INDIRECT` `LOOKUP` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP` `XMATCH`

⬜ `CHOOSECOLS` `CHOOSEROWS` `DROP` `EXPAND` `FILTER` `HSTACK` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS`

✖ `GETPIVOTDATA` `GROUPBY` `HYPERLINK` `IMAGE` `PIVOTBY` `RTD`

</details>

<details open>
<summary><strong>Matemática e trigonometria</strong> — 75/82</summary>

✅ `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `AGGREGATE` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MOD` `MROUND` `MULTINOMIAL` `ODD` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `RAND` `RANDBETWEEN` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

⬜ `MDETERM` `MINVERSE` `MMULT` `MUNIT` `PERCENTOF` `RANDARRAY` `SEQUENCE`

</details>

<details open>
<summary><strong>Estatísticas</strong> — 60/111</summary>

✅ `AVEDEV` `AVERAGE` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `CORREL` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `COVARIANCE.P` `COVARIANCE.S` `DEVSQ` `FISHER` `FISHERINV` `FORECAST` `FORECAST.LINEAR` `GEOMEAN` `HARMEAN` `INTERCEPT` `KURT` `LARGE` `MAX` `MAXA` `MAXIFS` `MEDIAN` `MIN` `MINA` `MINIFS` `MODE.SNGL` `PEARSON` `PERCENTILE.EXC` `PERCENTILE.INC` `PERCENTRANK.EXC` `PERCENTRANK.INC` `PERMUT` `PERMUTATIONA` `PHI` `PROB` `QUARTILE.EXC` `QUARTILE.INC` `RANK.AVG` `RANK.EQ` `RSQ` `SKEW` `SKEW.P` `SLOPE` `SMALL` `STANDARDIZE` `STDEV.P` `STDEV.S` `STDEVA` `STDEVPA` `STEYX` `TRIMMEAN` `VAR.P` `VAR.S` `VARA` `VARPA`

⬜ `BETA.DIST` `BETA.INV` `BINOM.DIST` `BINOM.DIST.RANGE` `BINOM.INV` `CHISQ.DIST` `CHISQ.DIST.RT` `CHISQ.INV` `CHISQ.INV.RT` `CHISQ.TEST` `CONFIDENCE.NORM` `CONFIDENCE.T` `EXPON.DIST` `F.DIST` `F.DIST.RT` `F.INV` `F.INV.RT` `F.TEST` `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` `FREQUENCY` `GAMMA` `GAMMA.DIST` `GAMMA.INV` `GAMMALN` `GAMMALN.PRECISE` `GAUSS` `GROWTH` `HYPGEOM.DIST` `LINEST` `LOGEST` `LOGNORM.DIST` `LOGNORM.INV` `MODE.MULT` `NEGBINOM.DIST` `NORM.DIST` `NORM.INV` `NORM.S.DIST` `NORM.S.INV` `POISSON.DIST` `T.DIST` `T.DIST.2T` `T.DIST.RT` `T.INV` `T.INV.2T` `T.TEST` `TREND` `WEIBULL.DIST` `Z.TEST`

Os nomes ⬜ restantes são quase todos distribuições estatísticas — elas dependem de funções
especiais validadas (gama/beta incompleta regularizada, erf, inversas numéricas) e serão entregues
juntas em uma fase posterior. `GAUSS` aguarda com elas: é a CDF normal menos ½, que precisa de erf
(`PHI`, a densidade simples, já está incluída).

</details>

<details open>
<summary><strong>Texto</strong> — 34/49</summary>

✅ `CHAR` `CLEAN` `CODE` `CONCAT` `CONCATENATE` `DOLLAR` `EXACT` `FIND` `FIXED` `LEFT` `LEN` `LOWER` `MID` `NUMBERVALUE` `PROPER` `REGEXEXTRACT` `REGEXREPLACE` `REGEXTEST` `REPLACE` `REPT` `RIGHT` `SEARCH` `SUBSTITUTE` `T` `TEXT` `TEXTAFTER` `TEXTBEFORE` `TEXTJOIN` `TRIM` `UNICHAR` `UNICODE` `UPPER` `VALUE` `VALUETOTEXT`

⬜ `ARRAYTOTEXT` `TEXTSPLIT`

✖ `ASC` `BAHTTEXT` `DBCS` `DETECTLANGUAGE` `FINDB` `LEFTB` `LENB` `MIDB` `PHONETIC` `REPLACEB` `RIGHTB` `SEARCHB` `TRANSLATE`

</details>

<details open>
<summary><strong>Informações</strong> — 18/22</summary>

✅ `ERROR.TYPE` `ISBLANK` `ISERR` `ISERROR` `ISEVEN` `ISFORMULA` `ISLOGICAL` `ISNA` `ISNONTEXT` `ISNUMBER` `ISODD` `ISREF` `ISTEXT` `N` `NA` `SHEET` `SHEETS` `TYPE`

⬜ `ISOMITTED`

✖ `CELL` `INFO` `STOCKHISTORY`

</details>

<details open>
<summary><strong>Data e hora</strong> — 25/25</summary>

✅ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

`NOW` e `TODAY` são **voláteis** — elas leem o relógio injetável do workbook e são atualizadas por
`Recalculate()`. A categoria está agora completa (25/25). Veja [Funções
voláteis](workbook-and-expressions.md#funções-voláteis).

</details>

<details>
<summary><strong>Compatibilidade (aliases legados)</strong> — 13/41</summary>

✅ `CONCATENATE` `COVAR` `FLOOR` `FORECAST` `MODE` `PERCENTILE` `PERCENTRANK` `QUARTILE` `RANK` `STDEV` `STDEVP` `VAR` `VARP`

⬜ `BETADIST` `BETAINV` `BINOMDIST` `CHIDIST` `CHIINV` `CHITEST` `CONFIDENCE` `CRITBINOM` `EXPONDIST` `FDIST` `FINV` `FTEST` `GAMMADIST` `GAMMAINV` `HYPGEOMDIST` `LOGINV` `LOGNORMDIST` `NEGBINOMDIST` `NORMDIST` `NORMINV` `NORMSDIST` `NORMSINV` `POISSON` `TDIST` `TINV` `TTEST` `WEIBULL` `ZTEST`

Os aliases ⬜ restantes são os nomes legados das distribuições estatísticas e as acompanham
(incluindo `CONFIDENCE`/`CRITBINOM`).

</details>

<details>
<summary><strong>Engenharia</strong> — 0/54</summary>

⬜ `BESSELI` `BESSELJ` `BESSELK` `BESSELY` `BIN2DEC` `BIN2HEX` `BIN2OCT` `BITAND` `BITLSHIFT` `BITOR` `BITRSHIFT` `BITXOR` `COMPLEX` `CONVERT` `DEC2BIN` `DEC2HEX` `DEC2OCT` `DELTA` `ERF` `ERF.PRECISE` `ERFC` `ERFC.PRECISE` `GESTEP` `HEX2BIN` `HEX2DEC` `HEX2OCT` `IMABS` `IMAGINARY` `IMARGUMENT` `IMCONJUGATE` `IMCOS` `IMCOSH` `IMCOT` `IMCSC` `IMCSCH` `IMDIV` `IMEXP` `IMLN` `IMLOG10` `IMLOG2` `IMPOWER` `IMPRODUCT` `IMREAL` `IMSEC` `IMSECH` `IMSIN` `IMSINH` `IMSQRT` `IMSUB` `IMSUM` `IMTAN` `OCT2BIN` `OCT2DEC` `OCT2HEX`

</details>

<details>
<summary><strong>Banco de dados</strong> — 0/12</summary>

⬜ `DAVERAGE` `DCOUNT` `DCOUNTA` `DGET` `DMAX` `DMIN` `DPRODUCT` `DSTDEV` `DSTDEVP` `DSUM` `DVAR` `DVARP`

</details>

<details>
<summary><strong>Cubos</strong> — 0/7</summary>

✖ `CUBEKPIMEMBER` `CUBEMEMBER` `CUBEMEMBERPROPERTY` `CUBERANKEDMEMBER` `CUBESET` `CUBESETCOUNT` `CUBEVALUE`

</details>

<details>
<summary><strong>Web</strong> — 0/3</summary>

✖ `ENCODEURL` `FILTERXML` `WEBSERVICE`

</details>

<details>
<summary><strong>Definidas pelo usuário</strong> — 0/3</summary>

✖ `CALL` `EUROCONVERT` `REGISTER.ID`

</details>

## Veja também

- [Funções personalizadas](custom-functions.md) — preenchendo você mesmo as lacunas da cobertura.
- [Workbook, planilhas e expressões](workbook-and-expressions.md) — operadores e sintaxe de referências.

## Fora de escopo (por design)

O MySheet é uma engine de cálculo do lado do servidor, então 35 funções do catálogo são
**permanentemente excluídas**, e não apenas "ainda não implementadas":

- **Serviços externos**: a família Cube/OLAP (`CUBE*`), a família Web (`WEBSERVICE`, `FILTERXML`,
  `ENCODEURL`), `RTD`, `STOCKHISTORY`, `IMAGE` e os serviços de tradução (`DETECTLANGUAGE`, `TRANSLATE`)
  — uma engine determinística não faz chamadas de rede.
- **Ambiente do aplicativo de planilha**: `CELL`, `INFO` e `HYPERLINK` descrevem a interface/host do
  Excel, que não existe aqui.
- **Modelo de tabela dinâmica**: `GETPIVOTDATA`, `PIVOTBY`, `GROUPBY` — o MySheet não tem modelo de
  tabela dinâmica.
- **Registro legado/XLM**: `CALL`, `REGISTER.ID`, `EUROCONVERT`.
- **Semântica de texto de byte duplo / localidades CJK**: `ASC`, `DBCS`, `BAHTTEXT`, `PHONETIC` e as
  variantes de byte `*B` (`LENB`, `FINDB`, `LEFTB`, `MIDB`, `RIGHTB`, `SEARCHB`, `REPLACEB`) — a engine é
  invariante de localidade por design.

Se o seu workbook depende de uma dessas, as [funções personalizadas](custom-functions.md) permitem que o
host forneça o comportamento (inclusive com chamadas de rede) sob o mesmo nome.
