# Referência de funções

*Tradução do documento canônico em inglês ([function-reference.md](../function-reference.md)). Em caso de divergência, o inglês prevalece.*

O MySheet implementa **112 funções nativas (built-in)**. A lista registrada oficial é o mapa `Functions`
em [`Danfma.MySheet/Parsing/Parser.cs`](../../Danfma.MySheet/Parsing/Parser.cs) — esta página é derivada
dele. A quantidade de argumentos é validada **em tempo de parse**: chamar uma função nativa com um número
de argumentos não suportado lança uma `ParseException`, assim como o Excel rejeita a fórmula na
digitação.

Além dessas, você pode adicionar suas próprias funções com
[`workbook.RegisterFunction`](custom-functions.md); nomes desconhecidos são avaliados como `#NAME?`.

Convenções abaixo: `[argumento]` entre colchetes é opcional; `…` significa que a função é variádica.
"Aceita intervalos" (*range-aware*) significa que argumentos de intervalo (`A1:B10`, uniões e resultados
de referência como o de `OFFSET`) são expandidos célula a célula.

## Lógicas (7)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` se todo argumento for considerado verdadeiro. |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Condicional; apenas o ramo escolhido é avaliado. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, ou o valor de contingência quando `value` é qualquer erro. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, ou o valor de contingência apenas quando `value` é `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Vincula nomes utilizáveis em `calculation` (ex.: `=LET(x, A1*2, x+x)`). Os nomes são locais à fórmula. |
| `NOT` | `NOT(logical)` | Negação lógica. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` se algum argumento for considerado verdadeiro. |

## Matemática e trigonometria (67)

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
| `SUM` | `SUM([number1], …)` | Soma de todos os valores numéricos; aceita intervalos. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Soma as células que atendem a um critério (ex.: `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Soma sob múltiplos pares de intervalo e critério. |
| `SUMSQ` | `SUMSQ(number1, …)` | Soma dos quadrados; aceita intervalos. |
| `TAN` | `TAN(number)` | Tangente (radianos). |
| `TANH` | `TANH(number)` | Tangente hiperbólica. |
| `TRUNC` | `TRUNC(number, [num_digits])` | Trunca em direção a zero (padrão: 0 dígitos). |

## Estatísticas (8)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `AVERAGE` | `AVERAGE([number1], …)` | Média aritmética dos valores numéricos; aceita intervalos. |
| `COUNT` | `COUNT([value1], …)` | Conta valores numéricos; aceita intervalos. |
| `COUNTA` | `COUNTA(value1, …)` | Conta valores não em branco; aceita intervalos. |
| `COUNTBLANK` | `COUNTBLANK(range, …)` | Conta células em branco. |
| `COUNTIF` | `COUNTIF(range, criteria)` | Conta as células que atendem a um critério. |
| `COUNTIFS` | `COUNTIFS(criteria_range1, criteria1, …)` | Contagem sob múltiplos pares de intervalo e critério. |
| `MAX` | `MAX([number1], …)` | Maior valor numérico; aceita intervalos. |
| `MIN` | `MIN([number1], …)` | Menor valor numérico; aceita intervalos. |

## Texto (11)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `CONCAT` | `CONCAT(text1, …)` | Concatena valores; aceita intervalos. |
| `CONCATENATE` | `CONCATENATE(text1, …)` | Alias legado de concatenação (argumentos escalares). |
| `LEFT` | `LEFT(text, [num_chars])` | Caracteres iniciais (padrão: 1). |
| `LEN` | `LEN(text)` | Comprimento do texto. |
| `LOWER` | `LOWER(text)` | Converte o texto para minúsculas. |
| `MID` | `MID(text, start_num, num_chars)` | Subtexto por posição (base 1) e comprimento. |
| `TEXT` | `TEXT(value, format_text)` | Formata um valor (formatos de número e data, ex.: `"0.00"`, `"dd/mm/yyyy"`). |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Junta valores com um delimitador; aceita intervalos. |
| `TRIM` | `TRIM(text)` | Remove espaços em excesso. |
| `UPPER` | `UPPER(text)` | Converte o texto para maiúsculas. |
| `VALUE` | `VALUE(text)` | Converte texto em número. |

## Pesquisa e referência (7)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `INDEX` | `INDEX(range, row_num, [column_num])` | O valor em uma posição (base 1) dentro de um intervalo. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | Posição (base 1) de um valor em um intervalo (`match_type`: 1 aproximado crescente — padrão, 0 exato, -1 aproximado decrescente). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | Uma referência deslocada (e opcionalmente redimensionada) a partir de uma referência inicial; pode retornar uma referência multicélula para consumidores que aceitam intervalos. |
| `ROW` | `ROW([reference])` | Número da linha da referência — ou da célula atual, quando chamada sem argumento. |
| `ROWS` | `ROWS(range)` | Número de linhas do intervalo. |
| `VLOOKUP` | `VLOOKUP(lookup_value, table_range, col_index_num, [range_lookup])` | Pesquisa vertical na primeira coluna de uma tabela; exata ou aproximada. |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])` | Pesquisa moderna, com contingência para "não encontrado" e modos de correspondência/busca. |

## Informações (3)

| Função | Argumentos | Descrição |
| --- | --- | --- |
| `ISBLANK` | `ISBLANK(value)` | `TRUE` para um valor em branco. |
| `ISNUMBER` | `ISNUMBER(value)` | `TRUE` para um valor numérico. |
| `SHEET` | `SHEET([value])` | Posição (base 1, na ordem das abas) da planilha de uma referência ou de um nome de planilha — ou da planilha atual, sem argumento. |

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

## Cobertura de funções do Excel

O MySheet implementa 112 das ~520 funções do [catálogo oficial de funções do Excel da
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
<summary><strong>Lógicas</strong> — 7/19</summary>

✅ `AND` `IF` `IFERROR` `IFNA` `LET` `NOT` `OR`

⬜ `BYCOL` `BYROW` `FALSE` `IFS` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN` `SWITCH` `TRUE` `XOR`

</details>

<details open>
<summary><strong>Pesquisa e referência</strong> — 7/40</summary>

✅ `INDEX` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP`

⬜ `ADDRESS` `AREAS` `CHOOSE` `CHOOSECOLS` `CHOOSEROWS` `COLUMN` `COLUMNS` `DROP` `EXPAND` `FILTER` `FORMULATEXT` `HLOOKUP` `HSTACK` `INDIRECT` `LOOKUP` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS` `XMATCH`

✖ `GETPIVOTDATA` `GROUPBY` `HYPERLINK` `IMAGE` `PIVOTBY` `RTD`

</details>

<details open>
<summary><strong>Matemática e trigonometria</strong> — 67/82</summary>

✅ `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MOD` `MROUND` `MULTINOMIAL` `ODD` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUM` `SUMIF` `SUMIFS` `SUMSQ` `TAN` `TANH` `TRUNC`

⬜ `AGGREGATE` `MDETERM` `MINVERSE` `MMULT` `MUNIT` `PERCENTOF` `RAND` `RANDARRAY` `RANDBETWEEN` `SEQUENCE` `SUBTOTAL` `SUMPRODUCT` `SUMX2MY2` `SUMX2PY2` `SUMXMY2`

</details>

<details open>
<summary><strong>Estatísticas</strong> — 8/111</summary>

✅ `AVERAGE` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `MAX` `MIN`

⬜ `AVEDEV` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `BETA.DIST` `BETA.INV` `BINOM.DIST` `BINOM.DIST.RANGE` `BINOM.INV` `CHISQ.DIST` `CHISQ.DIST.RT` `CHISQ.INV` `CHISQ.INV.RT` `CHISQ.TEST` `CONFIDENCE.NORM` `CONFIDENCE.T` `CORREL` `COVARIANCE.P` `COVARIANCE.S` `DEVSQ` `EXPON.DIST` `F.DIST` `F.DIST.RT` `F.INV` `F.INV.RT` `F.TEST` `FISHER` `FISHERINV` `FORECAST` `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` `FORECAST.LINEAR` `FREQUENCY` `GAMMA` `GAMMA.DIST` `GAMMA.INV` `GAMMALN` `GAMMALN.PRECISE` `GAUSS` `GEOMEAN` `GROWTH` `HARMEAN` `HYPGEOM.DIST` `INTERCEPT` `KURT` `LARGE` `LINEST` `LOGEST` `LOGNORM.DIST` `LOGNORM.INV` `MAXA` `MAXIFS` `MEDIAN` `MINA` `MINIFS` `MODE.MULT` `MODE.SNGL` `NEGBINOM.DIST` `NORM.DIST` `NORM.INV` `NORM.S.DIST` `NORM.S.INV` `PEARSON` `PERCENTILE.EXC` `PERCENTILE.INC` `PERCENTRANK.EXC` `PERCENTRANK.INC` `PERMUT` `PERMUTATIONA` `PHI` `POISSON.DIST` `PROB` `QUARTILE.EXC` `QUARTILE.INC` `RANK.AVG` `RANK.EQ` `RSQ` `SKEW` `SKEW.P` `SLOPE` `SMALL` `STANDARDIZE` `STDEV.P` `STDEV.S` `STDEVA` `STDEVPA` `STEYX` `T.DIST` `T.DIST.2T` `T.DIST.RT` `T.INV` `T.INV.2T` `T.TEST` `TREND` `TRIMMEAN` `VAR.P` `VAR.S` `VARA` `VARPA` `WEIBULL.DIST` `Z.TEST`

</details>

<details open>
<summary><strong>Texto</strong> — 11/49</summary>

✅ `CONCAT` `CONCATENATE` `LEFT` `LEN` `LOWER` `MID` `TEXT` `TEXTJOIN` `TRIM` `UPPER` `VALUE`

⬜ `ARRAYTOTEXT` `CHAR` `CLEAN` `CODE` `DOLLAR` `EXACT` `FIND` `FIXED` `NUMBERVALUE` `PROPER` `REGEXEXTRACT` `REGEXREPLACE` `REGEXTEST` `REPLACE` `REPT` `RIGHT` `SEARCH` `SUBSTITUTE` `T` `TEXTAFTER` `TEXTBEFORE` `TEXTSPLIT` `UNICHAR` `UNICODE` `VALUETOTEXT`

✖ `ASC` `BAHTTEXT` `DBCS` `DETECTLANGUAGE` `FINDB` `LEFTB` `LENB` `MIDB` `PHONETIC` `REPLACEB` `RIGHTB` `SEARCHB` `TRANSLATE`

</details>

<details open>
<summary><strong>Informações</strong> — 3/22</summary>

✅ `ISBLANK` `ISNUMBER` `SHEET`

⬜ `ERROR.TYPE` `ISERR` `ISERROR` `ISEVEN` `ISFORMULA` `ISLOGICAL` `ISNA` `ISNONTEXT` `ISODD` `ISOMITTED` `ISREF` `ISTEXT` `N` `NA` `SHEETS` `TYPE`

✖ `CELL` `INFO` `STOCKHISTORY`

</details>

<details>
<summary><strong>Data e hora</strong> — 0/25</summary>

⬜ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

</details>

<details>
<summary><strong>Compatibilidade (aliases legados)</strong> — 1/41</summary>

✅ `CONCATENATE`

⬜ `BETADIST` `BETAINV` `BINOMDIST` `CHIDIST` `CHIINV` `CHITEST` `CONFIDENCE` `COVAR` `CRITBINOM` `EXPONDIST` `FDIST` `FINV` `FLOOR` `FORECAST` `FTEST` `GAMMADIST` `GAMMAINV` `HYPGEOMDIST` `LOGINV` `LOGNORMDIST` `MODE` `NEGBINOMDIST` `NORMDIST` `NORMINV` `NORMSDIST` `NORMSINV` `PERCENTILE` `PERCENTRANK` `POISSON` `QUARTILE` `RANK` `STDEV` `STDEVP` `TDIST` `TINV` `TTEST` `VAR` `VARP` `WEIBULL` `ZTEST`

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
