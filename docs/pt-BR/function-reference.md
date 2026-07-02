# Referência de funções

*Tradução do documento canônico em inglês ([function-reference.md](../function-reference.md)). Em caso de divergência, o inglês prevalece.*

O MySheet implementa **231 funções nativas (built-in)**. A lista registrada oficial é o mapa `Functions`
em [`Danfma.MySheet/Parsing/Parser.cs`](../../Danfma.MySheet/Parsing/Parser.cs) — esta página é derivada
dele. A quantidade de argumentos é validada **em tempo de parse**: chamar uma função nativa com um número
de argumentos não suportado lança uma `ParseException`, assim como o Excel rejeita a fórmula na
digitação.

Além dessas, você pode adicionar suas próprias funções com
[`workbook.RegisterFunction`](custom-functions.md); nomes desconhecidos são avaliados como `#NAME?`.

Convenções abaixo: `[argumento]` entre colchetes é opcional; `…` significa que a função é variádica.
"Aceita intervalos" (*range-aware*) significa que argumentos de intervalo (`A1:B10`, uniões e resultados
de referência como o de `OFFSET`) são expandidos célula a célula.

## Lógicas (12)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` se todo argumento for considerado verdadeiro. |
| `FALSE` | `FALSE()` | O valor lógico `FALSE` (forma de função do literal). |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Condicional; apenas o ramo escolhido é avaliado. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, ou o valor de contingência quando `value` é qualquer erro. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, ou o valor de contingência apenas quando `value` é `#N/A`. |
| `IFS` | `IFS(test1, value1, [test2, value2], …)` | Primeiro valor cujo teste é `TRUE` (avaliação preguiçosa, como `IF`); nenhum teste `TRUE` → `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Vincula nomes utilizáveis em `calculation` (ex.: `=LET(x, A1*2, x+x)`). Os nomes são locais à fórmula. |
| `NOT` | `NOT(logical)` | Negação lógica. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` se algum argumento for considerado verdadeiro. |
| `SWITCH` | `SWITCH(expression, value1, result1, …, [default])` | Primeiro resultado cujo valor é igual a `expression` (avaliação preguiçosa; semântica de igualdade do `=`); sem correspondência → o padrão ou `#N/A`. |
| `TRUE` | `TRUE()` | O valor lógico `TRUE` (forma de função do literal). |
| `XOR` | `XOR(logical1, [logical2], …)` | `TRUE` quando a quantidade de entradas `TRUE` é ímpar; células de texto/em branco em intervalos são ignoradas. |

## Matemática e trigonometria (72)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ABS` | `ABS(number)` | Valor absoluto. |
| `ACOS` | `ACOS(number)` | Arco cosseno; fora de `[-1, 1]` → `#NUM!`. |
| `ACOSH` | `ACOSH(number)` | Cosseno hiperbólico inverso; abaixo de 1 → `#NUM!`. |
| `ACOT` | `ACOT(number)` | Arco cotangente, em `(0, π)`. |
| `ACOTH` | `ACOTH(number)` | Cotangente hiperbólica inversa; `\|number\| <= 1` → `#NUM!`. |
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
| `SUBTOTAL` | `SUBTOTAL(function_num, ref1, [ref2], …)` | Agregação selecionada por `function_num` (1-11: AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/STDEV.S/STDEV.P/SUM/VAR.S/VAR.P); células referenciadas cuja própria fórmula é um `SUBTOTAL` são ignoradas (evita contagem em duplicidade). 101-111 se comportam como 1-11 — o MySheet não tem modelo de linhas ocultas (limite documentado). Código inválido → `#VALUE!`. |
| `SUM` | `SUM([number1], …)` | Soma de todos os valores numéricos; aceita intervalos. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Soma as células que atendem a um critério (ex.: `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Soma sob múltiplos pares de intervalo e critério. |
| `SUMPRODUCT` | `SUMPRODUCT(array1, [array2], …)` | Soma dos produtos posição a posição; entradas não numéricas contam como 0; formatos diferentes → `#VALUE!`. |
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
| `TEXT` | `TEXT(value, format_text)` | Formata um valor (formatos de número e data, ex.: `"0.00"`, `"dd/mm/yyyy"`). |
| `TEXTAFTER` | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Texto após o n-ésimo delimitador (negativo conta a partir do fim); sem correspondência → `if_not_found` ou `#N/A`. |
| `TEXTBEFORE` | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Texto antes do n-ésimo delimitador (negativo conta a partir do fim); sem correspondência → `if_not_found` ou `#N/A`. |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Junta valores com um delimitador; aceita intervalos. |
| `TRIM` | `TRIM(text)` | Remove espaços em excesso. |
| `UNICHAR` | `UNICHAR(number)` | Caractere para um ponto de código Unicode completo (0/fora do intervalo → `#VALUE!`; pontos de código substitutos (*surrogates*) → `#N/A`). |
| `UNICODE` | `UNICODE(text)` | Ponto de código do primeiro caractere (pares substitutos lidos como um só). |
| `UPPER` | `UPPER(text)` | Converte o texto para maiúsculas. |
| `VALUE` | `VALUE(text)` | Converte texto em número. |
| `VALUETOTEXT` | `VALUETOTEXT(value, [format])` | Valor como texto — formato 0 conciso (padrão), 1 estrito (texto entre aspas); erros viram seu texto de exibição. |

## Pesquisa e referência (16)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ADDRESS` | `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])` | O endereço da célula como TEXTO (`abs_num` 1-4 → `$C$2`/`C$2`/`$C2`/`C2`); `a1=FALSE` renderiza apenas a forma R1C1 absoluta (`R2C3` — R1C1 relativo → `#VALUE!`); `sheet_text` entra como prefixo, entre aspas quando necessário. |
| `AREAS` | `AREAS(reference)` | Número de áreas (intervalos contíguos ou células individuais) na referência — uma verificação sintática, como `ISREF`; não referência → `#VALUE!`. |
| `CHOOSE` | `CHOOSE(index_num, value1, [value2], …)` | O valor na posição `index_num` (truncado); avaliação preguiçosa — apenas o argumento escolhido é avaliado; um intervalo escolhido permanece *range-aware* (`SUM(CHOOSE(…))`); fora do intervalo → `#VALUE!`. |
| `COLUMN` | `COLUMN([reference])` | Número da coluna da referência (a coluna mais à esquerda para um intervalo) — ou da célula atual, quando chamada sem argumento. |
| `COLUMNS` | `COLUMNS(range)` | Número de colunas do intervalo. |
| `FORMULATEXT` | `FORMULATEXT(reference)` | A fórmula da célula referenciada como TEXTO, com o `=` incluído (reescrita — *unparse* — no contexto de planilha da célula referenciada); uma célula literal ou vazia → `#N/A`. |
| `HLOOKUP` | `HLOOKUP(lookup_value, table_range, row_index_num, [range_lookup])` | Pesquisa horizontal na primeira linha de uma tabela; exata ou aproximada; `row_index_num` < 1 → `#VALUE!`, além da tabela → `#REF!`. |
| `INDEX` | `INDEX(range, row_num, [column_num])` | O valor em uma posição (base 1) dentro de um intervalo. |
| `LOOKUP` | `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Forma vetorial (sempre aproximada: o maior valor ≤ pesquisado); a forma matricial de 2 argumentos busca na primeira linha e retorna da última linha quando o intervalo é mais largo que alto; caso contrário, primeira/última coluna. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | Posição (base 1) de um valor em um intervalo (`match_type`: 1 aproximado crescente — padrão, 0 exato, -1 aproximado decrescente). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | Uma referência deslocada (e opcionalmente redimensionada) a partir de uma referência inicial; pode retornar uma referência multicélula para consumidores que aceitam intervalos. |
| `ROW` | `ROW([reference])` | Número da linha da referência — ou da célula atual, quando chamada sem argumento. |
| `ROWS` | `ROWS(range)` | Número de linhas do intervalo. |
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

## Financeiras (9)

Semântica padrão de valor do dinheiro no tempo: `rate` por período, `nper` é o total de períodos, `type`
0 = fim do período (padrão) / 1 = início.

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

O MySheet implementa 231 das ~520 funções do [catálogo oficial de funções do Excel da
Microsoft](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
agrupadas abaixo pelas próprias categorias da Microsoft (✅ implementada, ⬜ ainda não, ✖ fora de escopo
por design). **35 funções estão permanentemente fora de escopo** — elas dependem de serviços externos, do
ambiente de interface do aplicativo ou de recursos que a engine deliberadamente não modela (veja
[Fora de escopo](#fora-de-escopo-por-design) abaixo) — restando um catálogo viável de ~485 funções, sobre
o qual o roadmap é acompanhado. Alguns nomes são listados pela Microsoft em mais de uma categoria (ex.:
`CONCATENATE` em Texto e em Compatibilidade, `LET` em Lógicas e em Matemática), então as contagens por
categoria não somam um total único — veja o `Parser.cs` para a lista registrada oficial.

<details open>
<summary><strong>Financeiras</strong> — 9/55</summary>

✅ `FV` `IPMT` `IRR` `NPER` `NPV` `PMT` `PPMT` `PV` `RATE`

⬜ `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FVSCHEDULE` `INTRATE` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PRICE` `PRICEDISC` `PRICEMAT` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

</details>

<details open>
<summary><strong>Lógicas</strong> — 12/19</summary>

✅ `AND` `FALSE` `IF` `IFERROR` `IFNA` `IFS` `LET` `NOT` `OR` `SWITCH` `TRUE` `XOR`

⬜ `BYCOL` `BYROW` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN`

</details>

<details open>
<summary><strong>Pesquisa e referência</strong> — 16/40</summary>

✅ `ADDRESS` `AREAS` `CHOOSE` `COLUMN` `COLUMNS` `FORMULATEXT` `HLOOKUP` `INDEX` `LOOKUP` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP` `XMATCH`

⬜ `CHOOSECOLS` `CHOOSEROWS` `DROP` `EXPAND` `FILTER` `HSTACK` `INDIRECT` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS`

✖ `GETPIVOTDATA` `GROUPBY` `HYPERLINK` `IMAGE` `PIVOTBY` `RTD`

</details>

<details open>
<summary><strong>Matemática e trigonometria</strong> — 72/82</summary>

✅ `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MOD` `MROUND` `MULTINOMIAL` `ODD` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

⬜ `AGGREGATE` `MDETERM` `MINVERSE` `MMULT` `MUNIT` `PERCENTOF` `RAND` `RANDARRAY` `RANDBETWEEN` `SEQUENCE`

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

<details>
<summary><strong>Data e hora</strong> — 0/25</summary>

⬜ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

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
