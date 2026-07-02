# Function reference

MySheet implements **254 built-in functions**. The authoritative registered list is the `Functions` map
in [`Danfma.MySheet/Parsing/Parser.cs`](../Danfma.MySheet/Parsing/Parser.cs) — this page is derived from
it. Argument counts are validated **at parse time**: calling a built-in with an unsupported number of
arguments throws a `ParseException`, just as Excel rejects the formula at entry.

Beyond these, you can add your own functions with
[`workbook.RegisterFunction`](custom-functions.md); unknown names evaluate to `#NAME?`.

Conventions below: `[argument]` is optional; `…` means the function is variadic. "Range-aware" means
range arguments (`A1:B10`, unions, and reference results such as `OFFSET`'s) are expanded cell by cell.

## Logical (12)

| Function | Arguments | Description |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` if every argument is truthy. |
| `FALSE` | `FALSE()` | The logical value `FALSE` (function form of the literal). |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Conditional; only the taken branch is evaluated. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, or the fallback when `value` is any error. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, or the fallback only when `value` is `#N/A`. |
| `IFS` | `IFS(test1, value1, [test2, value2], …)` | First value whose test is `TRUE` (lazy, like `IF`); no `TRUE` test → `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Binds names usable in `calculation` (e.g. `=LET(x, A1*2, x+x)`). Names are local to the formula. |
| `NOT` | `NOT(logical)` | Logical negation. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` if any argument is truthy. |
| `SWITCH` | `SWITCH(expression, value1, result1, …, [default])` | First result whose value equals `expression` (lazy; `=` equality semantics); no match → default or `#N/A`. |
| `TRUE` | `TRUE()` | The logical value `TRUE` (function form of the literal). |
| `XOR` | `XOR(logical1, [logical2], …)` | `TRUE` when the number of `TRUE` inputs is odd; text/blank cells in ranges are ignored. |

## Math and trigonometry (72)

| Function | Arguments | Description |
| --- | --- | --- |
| `ABS` | `ABS(number)` | Absolute value. |
| `ACOS` | `ACOS(number)` | Arccosine; outside `[-1, 1]` → `#NUM!`. |
| `ACOSH` | `ACOSH(number)` | Inverse hyperbolic cosine; below 1 → `#NUM!`. |
| `ACOT` | `ACOT(number)` | Arccotangent, in `(0, π)`. |
| `ACOTH` | `ACOTH(number)` | Inverse hyperbolic cotangent; `\|number\| <= 1` → `#NUM!`. |
| `ARABIC` | `ARABIC(text)` | Roman numeral → number (case-insensitive; `""` → 0; leading `-` negates). |
| `ASIN` | `ASIN(number)` | Arcsine; outside `[-1, 1]` → `#NUM!`. |
| `ASINH` | `ASINH(number)` | Inverse hyperbolic sine. |
| `ATAN` | `ATAN(number)` | Arctangent. |
| `ATAN2` | `ATAN2(x_num, y_num)` | Arctangent from coordinates — Excel's `(x, y)` order; `ATAN2(0,0)` → `#DIV/0!`. |
| `ATANH` | `ATANH(number)` | Inverse hyperbolic tangent; `\|number\| >= 1` → `#NUM!`. |
| `BASE` | `BASE(number, radix, [min_length])` | Number → text in base `radix` (2-36), zero-padded to `min_length`. |
| `CEILING` | `CEILING(number, significance)` | Legacy ceiling with Excel's sign rules (`CEILING(-2.5,-2)` = -4; positive number with negative significance → `#NUM!`). |
| `CEILING.MATH` | `CEILING.MATH(number, [significance], [mode])` | Rounds up to a multiple; `mode` only affects negative numbers (away from zero when non-zero). |
| `CEILING.PRECISE` | `CEILING.PRECISE(number, [significance])` | Rounds toward +∞; significance sign ignored. |
| `COMBIN` | `COMBIN(number, number_chosen)` | Combinations without repetition. |
| `COMBINA` | `COMBINA(number, number_chosen)` | Combinations with repetition (`COMBIN(n+k-1, k)`). |
| `COS` | `COS(number)` | Cosine (radians). |
| `COSH` | `COSH(number)` | Hyperbolic cosine. |
| `COT` | `COT(number)` | Cotangent; `COT(0)` → `#DIV/0!`. |
| `COTH` | `COTH(number)` | Hyperbolic cotangent; `COTH(0)` → `#DIV/0!`. |
| `CSC` | `CSC(number)` | Cosecant; `CSC(0)` → `#DIV/0!`. |
| `CSCH` | `CSCH(number)` | Hyperbolic cosecant; `CSCH(0)` → `#DIV/0!`. |
| `DECIMAL` | `DECIMAL(text, radix)` | Text in base `radix` (2-36) → number; case-insensitive. |
| `DEGREES` | `DEGREES(angle)` | Radians → degrees. |
| `EVEN` | `EVEN(number)` | Rounds away from zero to the nearest even integer. |
| `EXP` | `EXP(number)` | e raised to `number`. |
| `FACT` | `FACT(number)` | Factorial (truncates; negative → `#NUM!`). |
| `FACTDOUBLE` | `FACTDOUBLE(number)` | Double factorial n!! (truncates; negative → `#NUM!`). |
| `FLOOR` | `FLOOR(number, significance)` | Legacy floor with Excel's sign rules (`FLOOR(-2.5,-2)` = -2; significance 0 → `#DIV/0!`). |
| `FLOOR.MATH` | `FLOOR.MATH(number, [significance], [mode])` | Rounds down to a multiple; `mode` only affects negative numbers (toward zero when non-zero). |
| `FLOOR.PRECISE` | `FLOOR.PRECISE(number, [significance])` | Rounds toward -∞; significance sign ignored. |
| `GCD` | `GCD(number1, …)` | Greatest common divisor; range-aware (truncates; negative → `#NUM!`). |
| `INT` | `INT(number)` | Rounds down to the nearest integer. |
| `ISO.CEILING` | `ISO.CEILING(number, [significance])` | Alias behaviour of `CEILING.PRECISE`. |
| `LCM` | `LCM(number1, …)` | Least common multiple; range-aware (truncates; negative → `#NUM!`). |
| `LN` | `LN(number)` | Natural logarithm; non-positive → `#NUM!`. |
| `LOG` | `LOG(number, [base])` | Logarithm (default base 10); base 1 → `#DIV/0!`, base ≤ 0 → `#NUM!`. |
| `LOG10` | `LOG10(number)` | Base-10 logarithm. |
| `MOD` | `MOD(number, divisor)` | Remainder with the sign of the divisor (`MOD(-3,2)` = 1); divisor 0 → `#DIV/0!`. |
| `MROUND` | `MROUND(number, multiple)` | Rounds to the nearest multiple; opposite signs → `#NUM!`. |
| `MULTINOMIAL` | `MULTINOMIAL(number1, …)` | Multinomial coefficient; range-aware. |
| `ODD` | `ODD(number)` | Rounds away from zero to the nearest odd integer. |
| `PI` | `PI()` | The constant π. |
| `POWER` | `POWER(number, power)` | Exponentiation; `0^0` → `#NUM!`, `0^negative` → `#DIV/0!`. |
| `PRODUCT` | `PRODUCT(number1, …)` | Product of the numeric values; range-aware. |
| `QUOTIENT` | `QUOTIENT(numerator, denominator)` | Integer portion of a division (truncated). |
| `RADIANS` | `RADIANS(angle)` | Degrees → radians. |
| `ROMAN` | `ROMAN(number, [form])` | Number (0-3999) → classic Roman numeral; `ROMAN(0)` = `""`. Concise forms 1-4/`FALSE` are not supported (→ `#VALUE!`). |
| `ROUND` | `ROUND(number, num_digits)` | Rounds to the given number of digits. |
| `ROUNDDOWN` | `ROUNDDOWN(number, num_digits)` | Rounds toward zero. |
| `ROUNDUP` | `ROUNDUP(number, num_digits)` | Rounds away from zero. |
| `SEC` | `SEC(number)` | Secant. |
| `SECH` | `SECH(number)` | Hyperbolic secant. |
| `SERIESSUM` | `SERIESSUM(x, n, m, coefficients)` | Power series sum; coefficients via range/values. |
| `SIGN` | `SIGN(number)` | -1, 0 or 1. |
| `SIN` | `SIN(number)` | Sine (radians). |
| `SINH` | `SINH(number)` | Hyperbolic sine. |
| `SQRT` | `SQRT(number)` | Square root; negative → `#NUM!`. |
| `SQRTPI` | `SQRTPI(number)` | Square root of `number × π`. |
| `SUBTOTAL` | `SUBTOTAL(function_num, ref1, [ref2], …)` | Aggregate selected by `function_num` (1-11: AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/STDEV.S/STDEV.P/SUM/VAR.S/VAR.P); referenced cells whose own formula is a `SUBTOTAL` are skipped (no double counting). 101-111 behave like 1-11 — MySheet has no hidden-row model (documented limit). Invalid code → `#VALUE!`. |
| `SUM` | `SUM([number1], …)` | Sum of all numeric values; range-aware. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Sums the cells matching a criteria (e.g. `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Sum under multiple criteria-range pairs. |
| `SUMPRODUCT` | `SUMPRODUCT(array1, [array2], …)` | Sum of the position-wise products; non-numeric entries count as 0; different shapes → `#VALUE!`. |
| `SUMSQ` | `SUMSQ(number1, …)` | Sum of squares; range-aware. |
| `SUMX2MY2` | `SUMX2MY2(array_x, array_y)` | Σ(x² − y²); pairs with a non-numeric side are dropped; different lengths → `#N/A`. |
| `SUMX2PY2` | `SUMX2PY2(array_x, array_y)` | Σ(x² + y²); same pairing rules as `SUMX2MY2`. |
| `SUMXMY2` | `SUMXMY2(array_x, array_y)` | Σ(x − y)²; same pairing rules as `SUMX2MY2`. |
| `TAN` | `TAN(number)` | Tangent (radians). |
| `TANH` | `TANH(number)` | Hyperbolic tangent. |
| `TRUNC` | `TRUNC(number, [num_digits])` | Truncates toward zero (default 0 digits). |

## Statistical (59)

Conventions of this family: the plain aggregates ignore referenced text/logicals/blanks (like
`SUM`); the `*A` variants count referenced text as 0 and logicals as 1/0. The two-range functions
(`CORREL`, `SLOPE`, …) drop a pair entirely when EITHER side is non-numeric, return `#N/A` on a
length mismatch and `#DIV/0!` on zero variance. `GAUSS` is deferred with the statistical
distributions (it needs the normal CDF/erf); `PHI` — the plain density — is included.

| Function | Arguments | Description |
| --- | --- | --- |
| `AVEDEV` | `AVEDEV(number1, …)` | Mean of the absolute deviations from the mean; no values → `#NUM!`. |
| `AVERAGE` | `AVERAGE([number1], …)` | Arithmetic mean of the numeric values; range-aware. |
| `AVERAGEA` | `AVERAGEA(value1, …)` | `AVERAGE` with the `*A` rule (text → 0, logicals → 1/0). |
| `AVERAGEIF` | `AVERAGEIF(range, criteria, [average_range])` | Average of the cells matching a criteria; no numeric match → `#DIV/0!`. |
| `AVERAGEIFS` | `AVERAGEIFS(average_range, criteria_range1, criteria1, …)` | Average under multiple criteria-range pairs; no match → `#DIV/0!`; shape mismatch → `#VALUE!`. |
| `CORREL` | `CORREL(array1, array2)` | Pearson product-moment correlation coefficient. |
| `COUNT` | `COUNT([value1], …)` | Counts numeric values; range-aware. |
| `COUNTA` | `COUNTA(value1, …)` | Counts non-blank values; range-aware. |
| `COUNTBLANK` | `COUNTBLANK(range, …)` | Counts blank cells. |
| `COUNTIF` | `COUNTIF(range, criteria)` | Counts the cells matching a criteria. |
| `COUNTIFS` | `COUNTIFS(criteria_range1, criteria1, …)` | Count under multiple criteria-range pairs. |
| `COVARIANCE.P` | `COVARIANCE.P(array1, array2)` | Population covariance Σ(x−x̄)(y−ȳ)/n. |
| `COVARIANCE.S` | `COVARIANCE.S(array1, array2)` | Sample covariance (n−1); fewer than 2 pairs → `#DIV/0!`. |
| `DEVSQ` | `DEVSQ(number1, …)` | Sum of squared deviations from the mean; no values → `#NUM!`. |
| `FISHER` | `FISHER(x)` | Fisher transformation; `x` ≤ −1 or ≥ 1 → `#NUM!`. |
| `FISHERINV` | `FISHERINV(y)` | Inverse Fisher transformation. |
| `FORECAST.LINEAR` | `FORECAST.LINEAR(x, known_ys, known_xs)` | Predicted y at `x` on the least-squares line — note the argument order (new x first). |
| `GEOMEAN` | `GEOMEAN(number1, …)` | Geometric mean; any value ≤ 0 → `#NUM!`. |
| `HARMEAN` | `HARMEAN(number1, …)` | Harmonic mean; any value ≤ 0 → `#NUM!`. |
| `INTERCEPT` | `INTERCEPT(known_ys, known_xs)` | y-intercept of the least-squares line. |
| `KURT` | `KURT(number1, …)` | Excess kurtosis (Excel's sample formula); fewer than 4 points or s = 0 → `#DIV/0!`. |
| `LARGE` | `LARGE(array, k)` | k-th largest value; empty array, k ≤ 0 or k > n → `#NUM!`. |
| `MAX` | `MAX([number1], …)` | Largest numeric value; range-aware. |
| `MAXA` | `MAXA(value1, …)` | `MAX` with the `*A` rule. |
| `MAXIFS` | `MAXIFS(max_range, criteria_range1, criteria1, …)` | Largest matching value; no match → 0; shape mismatch → `#VALUE!`. |
| `MEDIAN` | `MEDIAN(number1, …)` | Middle value (mean of the two middle values for an even count); no values → `#NUM!`. |
| `MIN` | `MIN([number1], …)` | Smallest numeric value; range-aware. |
| `MINA` | `MINA(value1, …)` | `MIN` with the `*A` rule. |
| `MINIFS` | `MINIFS(min_range, criteria_range1, criteria1, …)` | Smallest matching value; no match → 0; shape mismatch → `#VALUE!`. |
| `MODE.SNGL` | `MODE.SNGL(number1, …)` | Most frequent value; a tie resolves to the first value encountered; no duplicates → `#N/A`. |
| `PEARSON` | `PEARSON(array1, array2)` | The same coefficient as `CORREL`. |
| `PERCENTILE.EXC` | `PERCENTILE.EXC(array, k)` | Exclusive percentile: interpolation at rank `k·(n+1)`; `k` outside `(0, 1)` or an unreachable rank → `#NUM!`. |
| `PERCENTILE.INC` | `PERCENTILE.INC(array, k)` | Inclusive percentile: interpolation at `k·(n−1)`; `k` outside `[0, 1]` → `#NUM!`. |
| `PERCENTRANK.EXC` | `PERCENTRANK.EXC(array, x, [significance])` | Exclusive rank of `x` as a fraction (`(below+1)/(n+1)`, interpolated); TRUNCATED to `significance` digits (default 3); `x` out of range → `#N/A`. |
| `PERCENTRANK.INC` | `PERCENTRANK.INC(array, x, [significance])` | Inclusive rank of `x` (`below/(n−1)`, interpolated); same truncation and errors as `.EXC`. |
| `PERMUT` | `PERMUT(number, number_chosen)` | Permutations without repetition n!/(n−k)! (arguments truncated); n ≤ 0, k < 0 or n < k → `#NUM!`. |
| `PERMUTATIONA` | `PERMUTATIONA(number, number_chosen)` | Permutations with repetition n^k (arguments truncated); negative → `#NUM!`. |
| `PHI` | `PHI(x)` | Density of the standard normal distribution. |
| `PROB` | `PROB(x_range, prob_range, lower_limit, [upper_limit])` | Sum of the probabilities of x in `[lower, upper]` (upper omitted → x = lower); probabilities outside `(0, 1]` or not summing to 1 → `#NUM!`; length mismatch → `#N/A`. |
| `QUARTILE.EXC` | `QUARTILE.EXC(array, quart)` | Exclusive quartile via `PERCENTILE.EXC(quart/4)`; quart (truncated) ≤ 0 or ≥ 4 → `#NUM!`. |
| `QUARTILE.INC` | `QUARTILE.INC(array, quart)` | Inclusive quartile via `PERCENTILE.INC(quart/4)`; quart (truncated) outside 0-4 → `#NUM!`. |
| `RANK.AVG` | `RANK.AVG(number, ref, [order])` | Rank with ties averaged; order 0/omitted → descending, otherwise ascending; value absent → `#N/A`. |
| `RANK.EQ` | `RANK.EQ(number, ref, [order])` | Rank with ties sharing the top rank of the group; same order/errors as `.AVG`. |
| `RSQ` | `RSQ(known_ys, known_xs)` | Square of the Pearson coefficient. |
| `SKEW` | `SKEW(number1, …)` | Sample skewness (factor n/((n−1)(n−2))); fewer than 3 points or s = 0 → `#DIV/0!`. |
| `SKEW.P` | `SKEW.P(number1, …)` | Population skewness; fewer than 3 points or σ = 0 → `#DIV/0!`. |
| `SLOPE` | `SLOPE(known_ys, known_xs)` | Least-squares slope; var(x) = 0 → `#DIV/0!`. |
| `SMALL` | `SMALL(array, k)` | k-th smallest value; empty array, k ≤ 0 or k > n → `#NUM!`. |
| `STANDARDIZE` | `STANDARDIZE(x, mean, standard_dev)` | The z-score (x − mean)/sd; sd ≤ 0 → `#NUM!`. |
| `STDEV.P` | `STDEV.P(number1, …)` | Population standard deviation ("n"); no values → `#DIV/0!`. |
| `STDEV.S` | `STDEV.S(number1, …)` | Sample standard deviation ("n−1"); fewer than 2 values → `#DIV/0!`. |
| `STDEVA` | `STDEVA(value1, …)` | `STDEV.S` with the `*A` rule. |
| `STDEVPA` | `STDEVPA(value1, …)` | `STDEV.P` with the `*A` rule. |
| `STEYX` | `STEYX(known_ys, known_xs)` | Standard error of the predicted y; fewer than 3 pairs → `#DIV/0!`. |
| `TRIMMEAN` | `TRIMMEAN(array, percent)` | Mean after cutting `INT(n·percent/2)` values from EACH end of the sorted data; `percent` outside `[0, 1)` → `#NUM!`. |
| `VAR.P` | `VAR.P(number1, …)` | Population variance; no values → `#DIV/0!`. |
| `VAR.S` | `VAR.S(number1, …)` | Sample variance; fewer than 2 values → `#DIV/0!`. |
| `VARA` | `VARA(value1, …)` | `VAR.S` with the `*A` rule. |
| `VARPA` | `VARPA(value1, …)` | `VAR.P` with the `*A` rule. |

## Text (34)

Text functions follow the engine's locale-invariant contract: ordinal comparisons, invariant casing,
`.`/`,`/`$` in `FIXED`/`DOLLAR`. `CHAR`/`CODE` map Unicode code points (Latin-1 for 1-255), not the
Windows ANSI code page. The `REGEX*` functions run on .NET regular expressions (Excel specifies
PCRE2; the usual subset — classes, quantifiers, anchors, groups, `$n` — behaves identically) with a
defensive 1-second match timeout.

| Function | Arguments | Description |
| --- | --- | --- |
| `CHAR` | `CHAR(number)` | Character for a code 1-255 (out of range → `#VALUE!`). |
| `CLEAN` | `CLEAN(text)` | Removes control characters 0-31 (127 etc. stay, like Excel). |
| `CODE` | `CODE(text)` | Code of the first character (empty text → `#VALUE!`). |
| `CONCAT` | `CONCAT(text1, …)` | Concatenates values; range-aware. |
| `CONCATENATE` | `CONCATENATE(text1, …)` | Legacy alias of concatenation (scalar arguments). |
| `DOLLAR` | `DOLLAR(number, [decimals])` | Number as currency TEXT — `$1,234.57`, negatives `($1,200)`; decimals default 2, negative rounds left of the point. |
| `EXACT` | `EXACT(text1, text2)` | Case-sensitive comparison. |
| `FIND` | `FIND(find_text, within_text, [start_num])` | Case-sensitive position (1-based); no wildcards; not found → `#VALUE!`. |
| `FIXED` | `FIXED(number, [decimals], [no_commas])` | Number rounded and rendered as TEXT — `1,234.6`; decimals default 2 (max 127), negative rounds left of the point. |
| `LEFT` | `LEFT(text, [num_chars])` | Leading characters (default 1). |
| `LEN` | `LEN(text)` | Text length. |
| `LOWER` | `LOWER(text)` | Lower-cases the text. |
| `MID` | `MID(text, start_num, num_chars)` | Substring by 1-based position and length. |
| `NUMBERVALUE` | `NUMBERVALUE(text, [decimal_separator], [group_separator])` | Locale-explicit text → number (defaults `.` and `,`); spaces ignored; trailing `%` divide by 100 each. |
| `PROPER` | `PROPER(text)` | Capitalizes every letter that follows a non-letter; lower-cases the rest. |
| `REGEXEXTRACT` | `REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])` | First match of the pattern (mode 0); array modes 1/2 → `#VALUE!` until the arrays phase; no match → `#N/A`. |
| `REGEXREPLACE` | `REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])` | Replaces matches (`$n` group references); occurrence 0 = all, positive = nth, negative = nth from the end. |
| `REGEXTEST` | `REGEXTEST(text, pattern, [case_sensitivity])` | `TRUE` when the pattern matches anywhere in the text. |
| `REPLACE` | `REPLACE(old_text, start_num, num_chars, new_text)` | Replaces by 1-based position and length. |
| `REPT` | `REPT(text, number_times)` | Repeats text (count truncated; negative or a result over 32,767 chars → `#VALUE!`). |
| `RIGHT` | `RIGHT(text, [num_chars])` | Trailing characters (default 1). |
| `SEARCH` | `SEARCH(find_text, within_text, [start_num])` | Case-insensitive position with `?` `*` wildcards (`~` escapes); not found → `#VALUE!`. |
| `SUBSTITUTE` | `SUBSTITUTE(text, old_text, new_text, [instance_num])` | Case-sensitive replacement — every occurrence, or only the 1-based `instance_num`. |
| `T` | `T(value)` | The value if it is text, otherwise `""`. |
| `TEXT` | `TEXT(value, format_text)` | Formats a value (number and date formats, e.g. `"0.00"`, `"dd/mm/yyyy"`). |
| `TEXTAFTER` | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text after the nth delimiter (negative counts from the end); miss → `if_not_found` or `#N/A`. |
| `TEXTBEFORE` | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text before the nth delimiter (negative counts from the end); miss → `if_not_found` or `#N/A`. |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Joins values with a delimiter; range-aware. |
| `TRIM` | `TRIM(text)` | Removes excess whitespace. |
| `UNICHAR` | `UNICHAR(number)` | Character for a full Unicode code point (0/out of range → `#VALUE!`; surrogate code points → `#N/A`). |
| `UNICODE` | `UNICODE(text)` | Code point of the first character (surrogate pairs read as one). |
| `UPPER` | `UPPER(text)` | Upper-cases the text. |
| `VALUE` | `VALUE(text)` | Converts text to a number. |
| `VALUETOTEXT` | `VALUETOTEXT(value, [format])` | Value as text — format 0 concise (default), 1 strict (text quoted); errors become their display text. |

## Lookup and reference (16)

| Function | Arguments | Description |
| --- | --- | --- |
| `ADDRESS` | `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])` | The cell address as TEXT (`abs_num` 1-4 → `$C$2`/`C$2`/`$C2`/`C2`); `a1=FALSE` renders only the absolute R1C1 form (`R2C3` — relative R1C1 → `#VALUE!`); `sheet_text` is prefixed, quoted when needed. |
| `AREAS` | `AREAS(reference)` | Number of areas (contiguous ranges or single cells) in the reference — a syntactic check, like `ISREF`; non-reference → `#VALUE!`. |
| `CHOOSE` | `CHOOSE(index_num, value1, [value2], …)` | The value at `index_num` (truncated); lazy — only the chosen argument is evaluated; a chosen range stays range-aware (`SUM(CHOOSE(…))`); out of range → `#VALUE!`. |
| `COLUMN` | `COLUMN([reference])` | Column number of the reference (leftmost column for a range) — or of the current cell when called with no argument. |
| `COLUMNS` | `COLUMNS(range)` | Number of columns in the range. |
| `FORMULATEXT` | `FORMULATEXT(reference)` | The referenced cell's formula as TEXT, `=` included (un-parsed in the referenced cell's sheet context); a literal or empty cell → `#N/A`. |
| `HLOOKUP` | `HLOOKUP(lookup_value, table_range, row_index_num, [range_lookup])` | Horizontal lookup in the first row of a table; exact or approximate; `row_index_num` < 1 → `#VALUE!`, beyond the table → `#REF!`. |
| `INDEX` | `INDEX(range, row_num, [column_num])` | The value at a 1-based position inside a range. |
| `LOOKUP` | `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Vector form (always approximate: largest value ≤ lookup); the 2-argument array form searches the first row and returns from the last row when the range is wider than tall, otherwise first/last column. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | 1-based position of a value in a range (`match_type`: 1 approximate ascending — default, 0 exact, -1 approximate descending). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | A reference displaced (and optionally resized) from a starting reference; may return a multi-cell reference for range-aware consumers. |
| `ROW` | `ROW([reference])` | Row number of the reference — or of the current cell when called with no argument. |
| `ROWS` | `ROWS(range)` | Number of rows in the range. |
| `VLOOKUP` | `VLOOKUP(lookup_value, table_range, col_index_num, [range_lookup])` | Vertical lookup in the first column of a table; exact or approximate. |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])` | Modern lookup with not-found fallback and match/search modes. |
| `XMATCH` | `XMATCH(lookup_value, lookup_range, [match_mode], [search_mode])` | 1-based position with `XLOOKUP`'s modes (0 exact — default, -1 exact-or-smaller, 1 exact-or-larger, 2 wildcard; search 1/-1). |

## Information (18)

The `IS*` functions inspect the evaluated value without coercion (`ISNUMBER("19")` is `FALSE`) and
never propagate errors — they report on them.

| Function | Arguments | Description |
| --- | --- | --- |
| `ERROR.TYPE` | `ERROR.TYPE(error_val)` | `#NULL!`=1, `#DIV/0!`=2, `#VALUE!`=3, `#REF!`=4, `#NAME?`=5, `#NUM!`=6, `#N/A`=7; non-error → `#N/A`. |
| `ISBLANK` | `ISBLANK(value)` | `TRUE` for a blank value. |
| `ISERR` | `ISERR(value)` | `TRUE` for any error except `#N/A`. |
| `ISERROR` | `ISERROR(value)` | `TRUE` for any error value. |
| `ISEVEN` | `ISEVEN(number)` | `TRUE` for an even number (truncated first); nonnumeric → `#VALUE!`. |
| `ISFORMULA` | `ISFORMULA(reference)` | `TRUE` when the referenced cell contains a formula (not a plain literal); non-reference → `#VALUE!`. |
| `ISLOGICAL` | `ISLOGICAL(value)` | `TRUE` for a logical value. |
| `ISNA` | `ISNA(value)` | `TRUE` only for `#N/A`. |
| `ISNONTEXT` | `ISNONTEXT(value)` | `TRUE` for anything that is not text (blanks included). |
| `ISNUMBER` | `ISNUMBER(value)` | `TRUE` for a numeric value. |
| `ISODD` | `ISODD(number)` | `TRUE` for an odd number (truncated first); nonnumeric → `#VALUE!`. |
| `ISREF` | `ISREF(value)` | `TRUE` when the argument is a reference (cell/range/union) — a syntactic check, regardless of the value. |
| `ISTEXT` | `ISTEXT(value)` | `TRUE` for text. |
| `N` | `N(value)` | Number → itself; `TRUE`→1/`FALSE`→0; error → the error; anything else → 0. |
| `NA` | `NA()` | The `#N/A` error value. |
| `SHEET` | `SHEET([value])` | 1-based sheet position (tab order) of a reference or sheet name — or of the current sheet with no argument. |
| `SHEETS` | `SHEETS()` | Number of sheets in the workbook (the 3-D reference form does not apply: every reference spans one sheet). |
| `TYPE` | `TYPE(value)` | 1 number (blanks included), 2 text, 4 logical, 16 error (inspected, not propagated), 64 multi-cell reference. |

## Financial (9)

Standard time-value-of-money semantics: `rate` per period, `nper` total periods, `type` 0 = end of
period (default) / 1 = beginning.

| Function | Arguments | Description |
| --- | --- | --- |
| `FV` | `FV(rate, nper, pmt, [pv], [type])` | Future value of an investment. |
| `IPMT` | `IPMT(rate, per, nper, pv, [fv], [type])` | Interest portion of a given payment period. |
| `IRR` | `IRR(values, [guess])` | Internal rate of return of a cash-flow range. |
| `NPER` | `NPER(rate, pmt, pv, [fv], [type])` | Number of payment periods. |
| `NPV` | `NPV(rate, value1, …)` | Net present value of future cash flows; range-aware. |
| `PMT` | `PMT(rate, nper, pv, [fv], [type])` | Constant periodic payment of a loan/annuity. |
| `PPMT` | `PPMT(rate, per, nper, pv, [fv], [type])` | Principal portion of a given payment period. |
| `PV` | `PV(rate, nper, pmt, [fv], [type])` | Present value of an investment. |
| `RATE` | `RATE(nper, pmt, pv, [fv], [type], [guess])` | Interest rate per period (iterative). |

## Date and time (23)

Dates are **serial numbers** (`double`), exactly like Excel: the integer part counts days from the
1899-12-30 epoch and the fraction is the time of day. Date functions take numeric serials (build them with
`DATE`/`TIME`, or numeric text `CoerceToNumber` accepts); they do **not** implicitly parse date *strings*
(use `DATEVALUE`/`TIMEVALUE` for that). A negative serial is out of range → `#NUM!`. `TODAY`/`NOW` are
deferred (they are volatile — a future phase). Documented limitation: serials 1..59 (Jan–Feb 1900) render
one day behind Excel and serial 60 (Excel's fictitious 1900-02-29) is not representable; real dates
(serial ≥ 61, 1900-03-01) are exact.

| Function | Arguments | Description |
| --- | --- | --- |
| `DATE` | `DATE(year, month, day)` | Serial from parts; Excel overflow (month 13 → next Jan, day 0 → prior month-end); year 0–1899 adds 1900. |
| `DATEDIF` | `DATEDIF(start, end, unit)` | Difference in `"Y"`/`"M"`/`"D"`/`"MD"`/`"YM"`/`"YD"`; `start > end` → `#NUM!`. `"MD"` is officially unreliable. |
| `DATEVALUE` | `DATEVALUE(date_text)` | Parses a date string (invariant `yyyy-MM-dd`, `M/d/yyyy`, `d-MMM-yyyy`, …) to a whole-day serial; unparseable → `#VALUE!`. |
| `DAY` | `DAY(serial)` | Day of the month (1–31). |
| `DAYS` | `DAYS(end, start)` | Whole days between two dates (may be negative). |
| `DAYS360` | `DAYS360(start, end, [method])` | 30/360 day count; US (NASD) default, `TRUE` = European. |
| `EDATE` | `EDATE(start, months)` | The same day-of-month `months` away, clamped to the month end. |
| `EOMONTH` | `EOMONTH(start, months)` | Last day of the month `months` away from `start`. |
| `HOUR` | `HOUR(serial)` | Hour (0–23) of the time fraction. |
| `ISOWEEKNUM` | `ISOWEEKNUM(serial)` | ISO 8601 week number (weeks start Monday; week 1 holds the first Thursday). |
| `MINUTE` | `MINUTE(serial)` | Minute (0–59) of the time fraction. |
| `MONTH` | `MONTH(serial)` | Month (1–12). |
| `NETWORKDAYS` | `NETWORKDAYS(start, end, [holidays])` | Working days in `[start, end]` (inclusive); Sat/Sun and `holidays` excluded. |
| `NETWORKDAYS.INTL` | `NETWORKDAYS.INTL(start, end, [weekend], [holidays])` | `NETWORKDAYS` with a custom weekend (number 1–7/11–17 or a 7-char `"0000011"` mask). |
| `SECOND` | `SECOND(serial)` | Second (0–59), rounded to the nearest second. |
| `TIME` | `TIME(hour, minute, second)` | Time-of-day fraction; components 0–32767 roll over, taken mod 24h; negative → `#NUM!`. |
| `TIMEVALUE` | `TIMEVALUE(time_text)` | Parses a time string (`HH:mm[:ss]`, `h:mm[:ss] AM/PM`) to a `[0,1)` fraction; unparseable → `#VALUE!`. |
| `WEEKDAY` | `WEEKDAY(serial, [return_type])` | Day of week; `return_type` 1/2/3 and 11–17 (see the WEEKDAY table). |
| `WEEKNUM` | `WEEKNUM(serial, [return_type])` | Week of year; System 1 for 1/2/11–17, ISO 8601 (System 2) for 21. |
| `WORKDAY` | `WORKDAY(start, days, [holidays])` | Date `days` working days from `start` (start excluded); negative walks backward. |
| `WORKDAY.INTL` | `WORKDAY.INTL(start, days, [weekend], [holidays])` | `WORKDAY` with a custom weekend; invalid/all-weekend → `#NUM!`. |
| `YEAR` | `YEAR(serial)` | Calendar year (1900–9999). |
| `YEARFRAC` | `YEARFRAC(start, end, [basis])` | Year fraction on basis 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), 4 (European 30/360). |

## Compatibility — legacy aliases (11)

The pre-2010 names of the modern statistical functions. Each alias is a **distinct AST node**, not
a re-spelling of the modern record: it evaluates exactly like its modern equivalent, but
`FORMULATEXT`, serialization and xlsx export preserve the spelling you wrote — `STDEV(…)` never
becomes `STDEV.S(…)`. (`CONCATENATE` and legacy `FLOOR`, also in Microsoft's Compatibility
category, are documented in their Text/Math sections.)

| Function | Arguments | Modern equivalent |
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

## Excel function coverage

MySheet implements 254 of the ~520 functions in [Microsoft's official Excel function
catalog](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
grouped below by Microsoft's own categories (✅ implemented, ⬜ not yet, ✖ out of scope by design).
**35 functions are permanently out of scope** — they depend on external services, UI environment, or
features the engine deliberately does not model (see [Out of scope](#out-of-scope-by-design) below) —
leaving a viable catalog of ~485 that the roadmap tracks against. A few names are cross-listed by
Microsoft in more than one category (e.g. `CONCATENATE` in both Text and Compatibility, `LET` in both
Logical and Math), so per-category counts don't sum to a single unique total — see `Parser.cs` for the
authoritative registered list.

<details open>
<summary><strong>Financial</strong> — 9/55</summary>

✅ `FV` `IPMT` `IRR` `NPER` `NPV` `PMT` `PPMT` `PV` `RATE`

⬜ `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FVSCHEDULE` `INTRATE` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PRICE` `PRICEDISC` `PRICEMAT` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

</details>

<details open>
<summary><strong>Logical</strong> — 12/19</summary>

✅ `AND` `FALSE` `IF` `IFERROR` `IFNA` `IFS` `LET` `NOT` `OR` `SWITCH` `TRUE` `XOR`

⬜ `BYCOL` `BYROW` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN`

</details>

<details open>
<summary><strong>Lookup and Reference</strong> — 16/40</summary>

✅ `ADDRESS` `AREAS` `CHOOSE` `COLUMN` `COLUMNS` `FORMULATEXT` `HLOOKUP` `INDEX` `LOOKUP` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP` `XMATCH`

⬜ `CHOOSECOLS` `CHOOSEROWS` `DROP` `EXPAND` `FILTER` `HSTACK` `INDIRECT` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS`

✖ `GETPIVOTDATA` `GROUPBY` `HYPERLINK` `IMAGE` `PIVOTBY` `RTD`

</details>

<details open>
<summary><strong>Math and Trigonometry</strong> — 72/82</summary>

✅ `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MOD` `MROUND` `MULTINOMIAL` `ODD` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

⬜ `AGGREGATE` `MDETERM` `MINVERSE` `MMULT` `MUNIT` `PERCENTOF` `RAND` `RANDARRAY` `RANDBETWEEN` `SEQUENCE`

</details>

<details open>
<summary><strong>Statistical</strong> — 60/111</summary>

✅ `AVEDEV` `AVERAGE` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `CORREL` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `COVARIANCE.P` `COVARIANCE.S` `DEVSQ` `FISHER` `FISHERINV` `FORECAST` `FORECAST.LINEAR` `GEOMEAN` `HARMEAN` `INTERCEPT` `KURT` `LARGE` `MAX` `MAXA` `MAXIFS` `MEDIAN` `MIN` `MINA` `MINIFS` `MODE.SNGL` `PEARSON` `PERCENTILE.EXC` `PERCENTILE.INC` `PERCENTRANK.EXC` `PERCENTRANK.INC` `PERMUT` `PERMUTATIONA` `PHI` `PROB` `QUARTILE.EXC` `QUARTILE.INC` `RANK.AVG` `RANK.EQ` `RSQ` `SKEW` `SKEW.P` `SLOPE` `SMALL` `STANDARDIZE` `STDEV.P` `STDEV.S` `STDEVA` `STDEVPA` `STEYX` `TRIMMEAN` `VAR.P` `VAR.S` `VARA` `VARPA`

⬜ `BETA.DIST` `BETA.INV` `BINOM.DIST` `BINOM.DIST.RANGE` `BINOM.INV` `CHISQ.DIST` `CHISQ.DIST.RT` `CHISQ.INV` `CHISQ.INV.RT` `CHISQ.TEST` `CONFIDENCE.NORM` `CONFIDENCE.T` `EXPON.DIST` `F.DIST` `F.DIST.RT` `F.INV` `F.INV.RT` `F.TEST` `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` `FREQUENCY` `GAMMA` `GAMMA.DIST` `GAMMA.INV` `GAMMALN` `GAMMALN.PRECISE` `GAUSS` `GROWTH` `HYPGEOM.DIST` `LINEST` `LOGEST` `LOGNORM.DIST` `LOGNORM.INV` `MODE.MULT` `NEGBINOM.DIST` `NORM.DIST` `NORM.INV` `NORM.S.DIST` `NORM.S.INV` `POISSON.DIST` `T.DIST` `T.DIST.2T` `T.DIST.RT` `T.INV` `T.INV.2T` `T.TEST` `TREND` `WEIBULL.DIST` `Z.TEST`

The remaining ⬜ names are almost all statistical distributions — they need validated special
functions (regularized incomplete gamma/beta, erf, numeric inverses) and ship together in a later
phase. `GAUSS` waits with them: it is the normal CDF minus ½, which needs erf (`PHI`, the plain
density, is already in).

</details>

<details open>
<summary><strong>Text</strong> — 34/49</summary>

✅ `CHAR` `CLEAN` `CODE` `CONCAT` `CONCATENATE` `DOLLAR` `EXACT` `FIND` `FIXED` `LEFT` `LEN` `LOWER` `MID` `NUMBERVALUE` `PROPER` `REGEXEXTRACT` `REGEXREPLACE` `REGEXTEST` `REPLACE` `REPT` `RIGHT` `SEARCH` `SUBSTITUTE` `T` `TEXT` `TEXTAFTER` `TEXTBEFORE` `TEXTJOIN` `TRIM` `UNICHAR` `UNICODE` `UPPER` `VALUE` `VALUETOTEXT`

⬜ `ARRAYTOTEXT` `TEXTSPLIT`

✖ `ASC` `BAHTTEXT` `DBCS` `DETECTLANGUAGE` `FINDB` `LEFTB` `LENB` `MIDB` `PHONETIC` `REPLACEB` `RIGHTB` `SEARCHB` `TRANSLATE`

</details>

<details open>
<summary><strong>Information</strong> — 18/22</summary>

✅ `ERROR.TYPE` `ISBLANK` `ISERR` `ISERROR` `ISEVEN` `ISFORMULA` `ISLOGICAL` `ISNA` `ISNONTEXT` `ISNUMBER` `ISODD` `ISREF` `ISTEXT` `N` `NA` `SHEET` `SHEETS` `TYPE`

⬜ `ISOMITTED`

✖ `CELL` `INFO` `STOCKHISTORY`

</details>

<details open>
<summary><strong>Date and Time</strong> — 23/25</summary>

✅ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `SECOND` `TIME` `TIMEVALUE` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

⬜ `NOW` `TODAY`

`NOW` and `TODAY` are **deferred: volatile** — they read the wall clock, which needs the volatility
infrastructure (a `TimeProvider` plus no-cache propagation) planned for a later phase, not the date math
shipped here.

</details>

<details>
<summary><strong>Compatibility (legacy aliases)</strong> — 13/41</summary>

✅ `CONCATENATE` `COVAR` `FLOOR` `FORECAST` `MODE` `PERCENTILE` `PERCENTRANK` `QUARTILE` `RANK` `STDEV` `STDEVP` `VAR` `VARP`

⬜ `BETADIST` `BETAINV` `BINOMDIST` `CHIDIST` `CHIINV` `CHITEST` `CONFIDENCE` `CRITBINOM` `EXPONDIST` `FDIST` `FINV` `FTEST` `GAMMADIST` `GAMMAINV` `HYPGEOMDIST` `LOGINV` `LOGNORMDIST` `NEGBINOMDIST` `NORMDIST` `NORMINV` `NORMSDIST` `NORMSINV` `POISSON` `TDIST` `TINV` `TTEST` `WEIBULL` `ZTEST`

The remaining ⬜ aliases are the legacy names of the statistical distributions and follow them
(`CONFIDENCE`/`CRITBINOM` included).

</details>

<details>
<summary><strong>Engineering</strong> — 0/54</summary>

⬜ `BESSELI` `BESSELJ` `BESSELK` `BESSELY` `BIN2DEC` `BIN2HEX` `BIN2OCT` `BITAND` `BITLSHIFT` `BITOR` `BITRSHIFT` `BITXOR` `COMPLEX` `CONVERT` `DEC2BIN` `DEC2HEX` `DEC2OCT` `DELTA` `ERF` `ERF.PRECISE` `ERFC` `ERFC.PRECISE` `GESTEP` `HEX2BIN` `HEX2DEC` `HEX2OCT` `IMABS` `IMAGINARY` `IMARGUMENT` `IMCONJUGATE` `IMCOS` `IMCOSH` `IMCOT` `IMCSC` `IMCSCH` `IMDIV` `IMEXP` `IMLN` `IMLOG10` `IMLOG2` `IMPOWER` `IMPRODUCT` `IMREAL` `IMSEC` `IMSECH` `IMSIN` `IMSINH` `IMSQRT` `IMSUB` `IMSUM` `IMTAN` `OCT2BIN` `OCT2DEC` `OCT2HEX`

</details>

<details>
<summary><strong>Database</strong> — 0/12</summary>

⬜ `DAVERAGE` `DCOUNT` `DCOUNTA` `DGET` `DMAX` `DMIN` `DPRODUCT` `DSTDEV` `DSTDEVP` `DSUM` `DVAR` `DVARP`

</details>

<details>
<summary><strong>Cubes</strong> — 0/7</summary>

✖ `CUBEKPIMEMBER` `CUBEMEMBER` `CUBEMEMBERPROPERTY` `CUBERANKEDMEMBER` `CUBESET` `CUBESETCOUNT` `CUBEVALUE`

</details>

<details>
<summary><strong>Web</strong> — 0/3</summary>

✖ `ENCODEURL` `FILTERXML` `WEBSERVICE`

</details>

<details>
<summary><strong>User Defined</strong> — 0/3</summary>

✖ `CALL` `EUROCONVERT` `REGISTER.ID`

</details>

## See also

- [Custom functions](custom-functions.md) — filling gaps in the coverage yourself.
- [Workbook, sheets and expressions](workbook-and-expressions.md) — operators and reference syntax.

## Out of scope (by design)

MySheet is a server-side calculation engine, so 35 catalog functions are **permanently excluded** rather
than "not yet implemented":

- **External services**: the Cube/OLAP family (`CUBE*`), the Web family (`WEBSERVICE`, `FILTERXML`,
  `ENCODEURL`), `RTD`, `STOCKHISTORY`, `IMAGE`, and the translation services (`DETECTLANGUAGE`,
  `TRANSLATE`) — a deterministic engine does not call out to the network.
- **Spreadsheet-application environment**: `CELL`, `INFO` and `HYPERLINK` describe the Excel UI/host, which
  does not exist here.
- **Pivot-table model**: `GETPIVOTDATA`, `PIVOTBY`, `GROUPBY` — MySheet has no pivot model.
- **Legacy registration/XLM**: `CALL`, `REGISTER.ID`, `EUROCONVERT`.
- **Double-byte / CJK-locale text semantics**: `ASC`, `DBCS`, `BAHTTEXT`, `PHONETIC` and the `*B` byte
  variants (`LENB`, `FINDB`, `LEFTB`, `MIDB`, `RIGHTB`, `SEARCHB`, `REPLACEB`) — the engine is
  locale-invariant by design.

If your workbook depends on one of these, [custom functions](custom-functions.md) let the host supply the
behavior (including network calls) under the same name.
