# Function reference

MySheet implements **306 built-in functions**. The authoritative registered list is the `ByName` map
in [`Danfma.MySheet/Parsing/FunctionRegistry.cs`](../Danfma.MySheet/Parsing/FunctionRegistry.cs) — this
page is derived from it. Argument counts are validated **at parse time**: calling a built-in with an unsupported number of
arguments throws a `ParseException`, just as Excel rejects the formula at entry.

The rows below describe each function's own behaviour and are unchanged by array context. On top of them, a
**pure-scalar** function handed a range in an array-consuming position is applied **element by element** —
`SUM(LEN(A1:A3))` sums three lengths — while a range-aware one keeps consuming the whole range as documented
in its row. 180 of the 306 entries can be lifted that way; see
[implicit array arguments](workbook-and-expressions.md#implicit-array-arguments) for which consumers ask for
an array, which functions are lifted, and where the lift stops.

Beyond these, you can add your own functions with
[`workbook.RegisterFunction`](custom-functions.md); unknown names evaluate to `#NAME?`.

Conventions below: `[argument]` is optional; `…` means the function is variadic. "Range-aware" means
range arguments (`A1:B10`, unions, and reference results such as `OFFSET`'s) are expanded cell by cell.

## Logical (12)

| Function | Arguments | Description |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` if every argument is truthy; text and blank operands are ignored (whether a literal or reached through a reference); no evaluable value → `#VALUE!`. |
| `FALSE` | `FALSE()` | The logical value `FALSE` (function form of the literal). |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Conditional; only the taken branch is evaluated. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, or the fallback when `value` is any error. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, or the fallback only when `value` is `#N/A`. |
| `IFS` | `IFS(test1, value1, [test2, value2], …)` | First value whose test is `TRUE` (lazy, like `IF`); no `TRUE` test → `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Binds names usable in `calculation` (e.g. `=LET(x, A1*2, x+x)`). Names are local to the formula. A range bound to a name stays a range, so `LET(hdr, Data!$1:$1, MATCH(x, hdr, 0))` and `LET(r, A1:C9, SUM(r))` work; a single cell is bound by value. |
| `NOT` | `NOT(logical)` | Logical negation. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` if any argument is truthy; text and blank operands are ignored (whether a literal or reached through a reference); no evaluable value → `#VALUE!`. |
| `SWITCH` | `SWITCH(expression, value1, result1, …, [default])` | First result whose value equals `expression` (lazy; `=` equality semantics); no match → default or `#N/A`. |
| `TRUE` | `TRUE()` | The logical value `TRUE` (function form of the literal). |
| `XOR` | `XOR(logical1, [logical2], …)` | `TRUE` when the number of `TRUE` inputs is odd; text and blank operands are ignored (whether a literal or reached through a reference); no evaluable value → `#VALUE!`. |

## Math and trigonometry (75)

| Function | Arguments | Description |
| --- | --- | --- |
| `ABS` | `ABS(number)` | Absolute value. |
| `ACOS` | `ACOS(number)` | Arccosine; outside `[-1, 1]` → `#NUM!`. |
| `ACOSH` | `ACOSH(number)` | Inverse hyperbolic cosine; below 1 → `#NUM!`. |
| `ACOT` | `ACOT(number)` | Arccotangent, in `(0, π)`. |
| `ACOTH` | `ACOTH(number)` | Inverse hyperbolic cotangent; `\|number\| <= 1` → `#NUM!`. |
| `AGGREGATE` | `AGGREGATE(function_num, options, ref1, [ref2], …)` / `AGGREGATE(function_num, options, array, k)` | Excel's two documented syntaxes over one name, told apart by `function_num` alone — nothing syntactic separates them. **1-13** is the reference form (every argument from the third on is another ref): 1 `AVERAGE`, 2 `COUNT`, 3 `COUNTA`, 4 `MAX`, 5 `MIN`, 6 `PRODUCT`, 7 `STDEV.S`, 8 `STDEV.P`, 9 `SUM`, 10 `VAR.S`, 11 `VAR.P`, 12 `MEDIAN`, 13 `MODE.SNGL`. **14-19** is the array form, whose fourth argument is the `k` / `quart`: 14 `LARGE`, 15 `SMALL`, 16 `PERCENTILE.INC`, 17 `QUARTILE.INC`, 18 `PERCENTILE.EXC`, 19 `QUARTILE.EXC`. The same four-argument shape therefore means two different things — over `A1:A3` = 5/0/9 and `B1:B3` = 1/2/3, `AGGREGATE(9,6,A1:A3,B1:B3)` = 20 (a second ref) while `AGGREGATE(15,6,A1:A3,2)` = 5 (the 2nd smallest). `options` (0 or omitted, through 7) is three independent bits: **+1** ignore hidden rows, **+2** ignore error values, **+4** *keep* nested `SUBTOTAL`/`AGGREGATE` cells. So 0-3 skip a referenced cell whose own formula is a `SUBTOTAL` or an `AGGREGATE` and 4-7 count it — over `C1:C3` = a nested `SUBTOTAL` worth 3, a nested `AGGREGATE` worth 3 and a plain 5, `AGGREGATE(9,0,C1:C3)` = 5 against `AGGREGATE(9,4,C1:C3)` = 11. That 5 is the options table read literally and is **not** a measured value: Aspose.Cells 26.6.0 — the oracle behind the measured numbers elsewhere in this row — answers 8, because it skips the nested `SUBTOTAL` but counts the nested `AGGREGATE` (and 2 rather than 1 for `AGGREGATE(3,0,C1:C3)`). A known divergence in which the documented page wins. Options 2/3/6/7 drop error cells from the population instead of propagating them: over `E1:E3` = 5 / `#DIV/0!` / 9, `AGGREGATE(9,6,E1:E3)` = 14 where `AGGREGATE(9,4,E1:E3)` is `#DIV/0!`, and for `COUNTA` the count itself changes (`AGGREGATE(3,4,E1:E3)` = 3, `AGGREGATE(3,6,E1:E3)` = 2 — the page does not say what `COUNTA` then counts, but both are measured on Aspose.Cells 26.6.0). That bit reaches error *cells* and array *elements* only: an argument that is itself an error propagates whatever the options say, so `AGGREGATE(9,6,1/0)` and `AGGREGATE(9,6,E1:E3,1/0)` are both `#DIV/0!`, while `AGGREGATE(9,6,E2)` — the same division living in a referenced cell — is 0 (all measured). The hidden-row bit is a **no-op** — MySheet has no hidden-row model, so 1/3/5/7 behave exactly like 0/2/4/6 (documented limit, the same one behind `SUBTOTAL`'s 101-111). The array form's `k` counts over the population that *survives* the skips: on that same `E1:E3`, `AGGREGATE(15,6,E1:E3,2)` = 9, but `AGGREGATE(15,6,E1:E3,3)` → `#NUM!` even though the range holds three cells. Only the **array form** takes a [computed array](workbook-and-expressions.md#implicit-array-arguments) in its `array` position. The reference form's `ref` arguments are references, and a computed array in one of them is `#VALUE!` — `AGGREGATE(9,4,ROW(A1:A3))`, `AGGREGATE(2,4,(A1:A3<>0)*1)` and even the constant `AGGREGATE(9,4,{1,2,3})` all fail (measured on Aspose.Cells 26.6.0, 2026-09-09, under CSE entry too). That asymmetry is what the two syntaxes are FOR: with option 6, the array form is what turns the worksheet idiom `SMALL(positions/filter, k)` from `#DIV/0!` into the answer it means — over `A1:A3` = 5/0/9, `AGGREGATE(15,6,(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>"")*(A1:A3<>0)),1)` = 1 and the same call with `k` = 2 is 3. Errors: `function_num` outside 1-19 or `options` outside 0-7 → `#VALUE!`; the array form with only three arguments (no `k`) → `#VALUE!`, the page's own answer when "a second ref argument is necessary but not provided" (fewer than three arguments is an arity error, rejected at parse time); a `k` outside the surviving population → `#NUM!` — `LARGE`/`SMALL`'s answer; AGGREGATE's page does not state it, but it is measured on Aspose.Cells 26.6.0 (`AGGREGATE(15,6,E1:E3,3)` over a two-survivor population, and an all-error population at any `k`); a ref to a missing sheet → `#REF!` in both forms, before any cell is scanned. |
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
| `RAND` | `RAND()` | Volatile: a random real in `[0, 1)`. See [Volatile functions](workbook-and-expressions.md#volatile-functions). |
| `RANDBETWEEN` | `RANDBETWEEN(bottom, top)` | Volatile: a random integer in `[bottom, top]` (inclusive); `bottom > top` → `#NUM!`; non-integer bounds truncate toward zero. |
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
| `SUBTOTAL` | `SUBTOTAL(function_num, ref1, [ref2], …)` | Aggregate selected by `function_num` (1-11: AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/STDEV.S/STDEV.P/SUM/VAR.S/VAR.P); referenced cells whose own formula is a `SUBTOTAL` are skipped (no double counting). 101-111 behave like 1-11 — MySheet has no hidden-row model (documented limit). Invalid code → `#VALUE!`. An argument that is a **computed array** rather than a reference is `#VALUE!`, NOT folded: the page defines `ref1` as "the first named range or reference" and Excel enforces that literally — `SUBTOTAL(9,ROW(A1:A3))`, `SUBTOTAL(9,(A1:A3<>0)*1)`, `SUBTOTAL(2,…)`, `SUBTOTAL(3,…)` and even the constant `SUBTOTAL(9,{1,2,3})` all give `#VALUE!` (measured on Aspose.Cells 26.6.0, 2026-09-09, under CSE entry too). `SUM` does fold an [implicit array argument](workbook-and-expressions.md#implicit-array-arguments) and `SUBTOTAL` does not — an earlier release folded it here by analogy with `SUM`, which was an inference, and the measurement reversed it. The array position that *does* take a computed array is `AGGREGATE`'s array form (`function_num` 14-19). A nested `AGGREGATE` cell is **not** skipped, only a nested `SUBTOTAL` is: over a range holding a nested `SUBTOTAL` (3), a nested `AGGREGATE` (3) and a plain 5, `SUBTOTAL(9,…)` = 8 — measured on Aspose.Cells 26.6.0, so the narrow rule is Excel's, not a guess, even though the "nested SUBTOTAL and AGGREGATE" wording exists only in AGGREGATE's own options table. `AGGREGATE(9,0,…)` over that same range is 5 here, which is **not** measured: it follows AGGREGATE's options table as written, while the oracle answers 8 — it skips the nested `SUBTOTAL` and counts the nested `AGGREGATE` the table says to skip. A known divergence in which the documented page wins. |
| `SUM` | `SUM([number1], …)` | Sum of all numeric values; range-aware. Also folds an [implicit array argument](workbook-and-expressions.md#implicit-array-arguments) — `SUM(IF(B2:B5="Show",1,0))` = 2. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Sums the cells matching a criteria (e.g. `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Sum under multiple criteria-range pairs. |
| `SUMPRODUCT` | `SUMPRODUCT(array1, [array2], …)` | Sum of the position-wise products; non-numeric entries count as 0 — including the `TRUE`/`FALSE` of a bare comparison, which is why the Excel idiom multiplies by 1 — while an error element propagates. Accepts a [computed array](workbook-and-expressions.md#implicit-array-arguments) as a whole argument, mixed freely with ranges: `SUMPRODUCT((A1:A3<>0)*1)` = 2, `SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))` = 9. The arguments must share the same **dimensions**, not merely the same cell count: `SUMPRODUCT(A1:A3,A1:C1)` (3x1 against 1x3) → `#VALUE!`. A *shapeless* argument — a defined name, an open range, a union — carries no rectangle the engine can see and is judged by cell count alone, so `SUMPRODUCT(MyName,A1:C1)` over a 3x1 name computes where Excel answers `#VALUE!` (a documented deviation); every argument that does carry a rectangle is still compared with the others, so `SUMPRODUCT(MyName,A1:A3,A1:C1)` is `#VALUE!`. |
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

`AVERAGE`, `COUNT`, `MIN`, `MAX`, `SMALL`, and `LARGE` also accept an
[implicit array argument](workbook-and-expressions.md#implicit-array-arguments) — e.g.
`SMALL(IF(B2:B5="Show",ROW(B2:B5)),1)` — folding it element-by-element with the same ignore-text/logicals
rule (so a branch-less `IF`'s `FALSE` drops out).

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
| `TEXT` | `TEXT(value, format_text)` | Formats a value (number and date formats, e.g. `"0.00"`, `"dd/mm/yyyy"`). A date format reads Excel's own 1900 calendar — see [Date and time](#date-and-time-25). |
| `TEXTAFTER` | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text after the nth delimiter (negative counts from the end); miss → `if_not_found` or `#N/A`. |
| `TEXTBEFORE` | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text before the nth delimiter (negative counts from the end); miss → `if_not_found` or `#N/A`. |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Joins values with a delimiter; range-aware. |
| `TRIM` | `TRIM(text)` | Removes excess whitespace. |
| `UNICHAR` | `UNICHAR(number)` | Character for a full Unicode code point (0/out of range → `#VALUE!`; surrogate code points → `#N/A`). |
| `UNICODE` | `UNICODE(text)` | Code point of the first character (surrogate pairs read as one). |
| `UPPER` | `UPPER(text)` | Upper-cases the text. |
| `VALUE` | `VALUE(text)` | Converts text to a number. |
| `VALUETOTEXT` | `VALUETOTEXT(value, [format])` | Value as text — format 0 concise (default), 1 strict (text quoted); errors become their display text. |

## Lookup and reference (17)

| Function | Arguments | Description |
| --- | --- | --- |
| `ADDRESS` | `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])` | The cell address as TEXT (`abs_num` 1-4 → `$C$2`/`C$2`/`$C2`/`C2`); `a1=FALSE` renders only the absolute R1C1 form (`R2C3` — relative R1C1 → `#VALUE!`); `sheet_text` is prefixed, quoted when needed. |
| `AREAS` | `AREAS(reference)` | Number of areas (contiguous ranges or single cells) in the reference — a syntactic check, like `ISREF`; non-reference → `#VALUE!`, and an argument that fails to resolve reports its own error (`AREAS(NoSuchName)` → `#NAME?`). |
| `CHOOSE` | `CHOOSE(index_num, value1, [value2], …)` | The value at `index_num` (truncated); lazy — only the chosen argument is evaluated; a chosen range stays range-aware (`SUM(CHOOSE(…))`); out of range → `#VALUE!`. |
| `COLUMN` | `COLUMN([reference])` | Column number of the reference (leftmost column for a range) — or of the current cell when called with no argument. Accepts ANY reference-producing expression, not only a literal reference: a defined name, `INDEX`/`OFFSET`/`INDIRECT`/`CHOOSE`, a `:` range with reference-returning endpoints (`COLUMN(INDEX(A1:C1,1,2))` = 2). A whole-column/row reference uses its DECLARED bound (`COLUMN(A:A)` = 1, `COLUMN(1:1)` = 1) — where the adjacent `COLUMNS` uses the POPULATED extent on an open axis; a non-reference (or a union) → `#VALUE!`, and an argument that fails to resolve reports its own error (`#NAME?`, `#REF!`). In an [array position](workbook-and-expressions.md#implicit-array-arguments) it yields the whole column-number vector, over a literal range *and* over a name or a `:` range that denotes one (`SUM(COLUMN(A1:C3))` = 18, `SUM(COLUMN(MyName))` = 3 for a single-column name over three rows); an open range is refused there and a reference-returning *function* argument stays scalar. |
| `COLUMNS` | `COLUMNS(range)` | Number of columns in the range. Over a [whole-column/row reference](workbook-and-expressions.md#whole-column-and-whole-row-references) a bounded column axis is exact (`COLUMNS(A:C)` = 3), an open one uses the populated extent. An argument that fails to resolve to a reference reports its own error (`#NAME?`, `#REF!`); a plain scalar value counts as 1 (a 1x1 array). |
| `FORMULATEXT` | `FORMULATEXT(reference)` | The referenced cell's formula as TEXT, `=` included (un-parsed in the referenced cell's sheet context); a literal or empty cell → `#N/A`. |
| `HLOOKUP` | `HLOOKUP(lookup_value, table_range, row_index_num, [range_lookup])` | Horizontal lookup in the first row of a table; exact or approximate; `row_index_num` < 1 → `#VALUE!`, beyond the table → `#REF!`. |
| `INDEX` | `INDEX(range, row_num, [column_num])` | The value at a 1-based position inside a range. Accepts an [implicit array first argument](workbook-and-expressions.md#implicit-array-arguments) (`INDEX(ROW(B2:B5),1)` = 2), including the `INDEX(ROW($A:$A), n)` identity that returns `n` without materializing the column; out of range → `#REF!`. |
| `INDIRECT` | `INDIRECT(ref_text, [a1])` | The reference named by `ref_text`, resolved at evaluation time: the text is parsed as a formula body in the **current sheet's** context, so an unqualified `"A1"` means the calling sheet's `A1` while `"Data!B2"` crosses sheets. A single cell is dereferenced to its **value** (`INDIRECT("A1")`); a multi-cell result stays a **reference** for range-aware consumers (`SUM(INDIRECT("A1:A3"))`, `ROWS(INDIRECT("A1:A3"))` = 3, and as a `:` endpoint — `SUM(INDIRECT("A1"):A3)`), so bare in a cell it gets [implicit intersection at the cell boundary](workbook-and-expressions.md#implicit-intersection-at-the-cell-boundary) (`=INDIRECT("A1:A3")` written in `B2` shows `A2`, in `B5` `#VALUE!`). [Defined names](workbook-and-expressions.md#named-ranges) resolve too, including one assembled at run time (`SUM(INDIRECT("R"&"ng"))`, `INDIRECT("Data!A"&2)`). **Volatile**: the reference is known only at evaluation time, so the cell is always recomputed — see [Volatile functions](workbook-and-expressions.md#volatile-functions). Everything that fails is `#REF!`, never `#NAME?`: `a1` = `FALSE`/`0` (R1C1 style is not supported — A1 only) or an `a1` that is not a number; a `ref_text` that is not text (a number, a logical, a numeric cell); text that does not parse; and an unknown name or sheet. |
| `LOOKUP` | `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Vector form (always approximate: largest value ≤ lookup); the 2-argument array form searches the first row and returns from the last row when the range is wider than tall, otherwise first/last column. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | 1-based position of a value in a range (`match_type`: 1 approximate ascending — default, 0 exact, -1 approximate descending). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | A reference displaced (and optionally resized) from a starting reference; may return a multi-cell reference for range-aware consumers. |
| `ROW` | `ROW([reference])` | Row number of the reference (top row for a range) — or of the current cell when called with no argument. Accepts ANY reference-producing expression, not only a literal reference: a defined name, `INDEX`/`OFFSET`/`INDIRECT`/`CHOOSE`, a `:` range with reference-returning endpoints (`ROW(INDEX(A1:A3,2,1))` = 2). A whole-column/row reference uses its DECLARED bound (`ROW(A:A)` = 1, `ROW(A2:A)` = 2) — where the adjacent `ROWS` uses the POPULATED extent on an open axis; a non-reference (or a union) → `#VALUE!`, and an argument that fails to resolve reports its own error (`#NAME?`, `#REF!`). In an [array position](workbook-and-expressions.md#implicit-array-arguments) it yields the whole row-number vector, over a literal range *and* over a name or a `:` range that denotes one (`SUM(ROW(MyName))` = 6 for a name over three rows); an open range is refused there and a reference-returning *function* argument stays scalar. |
| `ROWS` | `ROWS(range)` | Number of rows in the range. Over a [whole-column/row reference](workbook-and-expressions.md#whole-column-and-whole-row-references) an open row axis uses the populated extent (`ROWS(A:A)` = max − min populated row + 1, 0 if empty — a documented divergence from Excel's fixed grid), a bounded one is exact (`ROWS(1:5)` = 5). An argument that fails to resolve to a reference reports its own error (`ROWS(NoSuchName)` → `#NAME?`, `ROWS(INDIRECT("zz"))` → `#REF!`); a plain scalar value counts as 1 (a 1x1 array). |
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

## Financial (55)

Standard time-value-of-money semantics: `rate` per period, `nper` total periods, `type` 0 = end of
period (default) / 1 = beginning. The bond, coupon and dated-cash-flow functions take **date serials**
(build them with `DATE`, exactly like the date functions) and a day-count `basis`: 0 = US (NASD) 30/360
(default), 1 = actual/actual, 2 = actual/360, 3 = actual/365, 4 = European 30/360. Coupon `frequency` is
1 (annual), 2 (semi-annual) or 4 (quarterly). Coupon dates are built by stepping **backward from
maturity**. Iterative results (`RATE`, `IRR`, `XIRR`, `YIELD`, `ODDFYIELD`) use the same robust
bracketing + bisection solver as `RATE`/`IRR` (validated against a stiff 30-year case). Golden values for
the whole family are cross-checked against the `ExcelFinancialFunctions` oracle. Domain violations map to
`#NUM!` (settlement ≥ maturity, frequency ∉ {1,2,4}, basis ∉ 0..4, etc.).

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
| `SLN` | `SLN(cost, salvage, life)` | Straight-line depreciation per period. |
| `SYD` | `SYD(cost, salvage, life, per)` | Sum-of-years'-digits depreciation. |
| `DB` | `DB(cost, salvage, life, period, [month])` | Fixed-declining-balance depreciation. |
| `DDB` | `DDB(cost, salvage, life, period, [factor])` | Double-declining-balance depreciation. |
| `VDB` | `VDB(cost, salvage, life, start, end, [factor], [no_switch])` | Variable declining-balance depreciation. |
| `AMORLINC` | `AMORLINC(cost, purchased, first_period, salvage, period, rate, [basis])` | French linear depreciation (prorated). |
| `AMORDEGRC` | `AMORDEGRC(cost, purchased, first_period, salvage, period, rate, [basis])` | French declining depreciation with a life-based coefficient. |
| `EFFECT` | `EFFECT(nominal_rate, npery)` | Effective annual interest rate. |
| `NOMINAL` | `NOMINAL(effect_rate, npery)` | Nominal annual interest rate. |
| `MIRR` | `MIRR(values, finance_rate, reinvest_rate)` | Modified internal rate of return. |
| `RRI` | `RRI(nper, pv, fv)` | Equivalent interest rate for an investment's growth. |
| `PDURATION` | `PDURATION(rate, pv, fv)` | Periods for an investment to reach a value. |
| `ISPMT` | `ISPMT(rate, per, nper, pv)` | Interest paid during a straight-loan period. |
| `CUMIPMT` | `CUMIPMT(rate, nper, pv, start, end, type)` | Cumulative interest over a period range. |
| `CUMPRINC` | `CUMPRINC(rate, nper, pv, start, end, type)` | Cumulative principal over a period range. |
| `FVSCHEDULE` | `FVSCHEDULE(principal, schedule)` | Future value after a series of compound rates. |
| `DOLLARDE` | `DOLLARDE(fractional_dollar, fraction)` | Fractional-notation price → decimal. |
| `DOLLARFR` | `DOLLARFR(decimal_dollar, fraction)` | Decimal price → fractional notation. |
| `XNPV` | `XNPV(rate, values, dates)` | Net present value of dated cash flows (actual/365). |
| `XIRR` | `XIRR(values, dates, [guess])` | Internal rate of return of dated cash flows. |
| `ACCRINT` | `ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])` | Accrued interest for a periodic-interest security. |
| `ACCRINTM` | `ACCRINTM(issue, settlement, rate, par, [basis])` | Accrued interest for a maturity-paying security. |
| `DISC` | `DISC(settlement, maturity, pr, redemption, [basis])` | Discount rate of a security. |
| `INTRATE` | `INTRATE(settlement, maturity, investment, redemption, [basis])` | Interest rate of a fully-invested security. |
| `RECEIVED` | `RECEIVED(settlement, maturity, investment, discount, [basis])` | Amount received at maturity. |
| `PRICEDISC` | `PRICEDISC(settlement, maturity, discount, redemption, [basis])` | Price per $100 of a discounted security. |
| `PRICEMAT` | `PRICEMAT(settlement, maturity, issue, rate, yld, [basis])` | Price per $100 of an interest-at-maturity security. |
| `YIELDDISC` | `YIELDDISC(settlement, maturity, pr, redemption, [basis])` | Annual yield of a discounted security. |
| `YIELDMAT` | `YIELDMAT(settlement, maturity, issue, rate, pr, [basis])` | Annual yield of an interest-at-maturity security. |
| `TBILLEQ` | `TBILLEQ(settlement, maturity, discount)` | Bond-equivalent yield of a Treasury bill. |
| `TBILLPRICE` | `TBILLPRICE(settlement, maturity, discount)` | Price per $100 of a Treasury bill. |
| `TBILLYIELD` | `TBILLYIELD(settlement, maturity, pr)` | Yield of a Treasury bill. |
| `COUPPCD` | `COUPPCD(settlement, maturity, frequency, [basis])` | Previous coupon date before settlement. |
| `COUPNCD` | `COUPNCD(settlement, maturity, frequency, [basis])` | Next coupon date after settlement. |
| `COUPNUM` | `COUPNUM(settlement, maturity, frequency, [basis])` | Number of coupons between settlement and maturity. |
| `COUPDAYS` | `COUPDAYS(settlement, maturity, frequency, [basis])` | Days in the coupon period containing settlement. |
| `COUPDAYBS` | `COUPDAYBS(settlement, maturity, frequency, [basis])` | Days from the period start to settlement. |
| `COUPDAYSNC` | `COUPDAYSNC(settlement, maturity, frequency, [basis])` | Days from settlement to the next coupon. |
| `PRICE` | `PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])` | Price per $100 of a periodic-coupon bond. |
| `YIELD` | `YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])` | Yield to maturity (iterative). |
| `DURATION` | `DURATION(settlement, maturity, coupon, yld, frequency, [basis])` | Macaulay duration in years. |
| `MDURATION` | `MDURATION(settlement, maturity, coupon, yld, frequency, [basis])` | Modified Macaulay duration. |
| `ODDFPRICE` | `ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])` | Price of a bond with an odd first period. |
| `ODDFYIELD` | `ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])` | Yield of a bond with an odd first period. |
| `ODDLPRICE` | `ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])` | Price of a bond with an odd last period. |
| `ODDLYIELD` | `ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])` | Yield of a bond with an odd last period. |

## Date and time (25)

Dates are **serial numbers** (`double`), exactly like Excel: the integer part counts days and the fraction is
the time of day. Date functions take numeric serials (build them with `DATE`/`TIME`, or numeric text
`CoerceToNumber` accepts); they do **not** implicitly parse date *strings* (use `DATEVALUE`/`TIMEVALUE` for
that). A negative serial is out of range → `#NUM!`. `TODAY`/`NOW` read the clock and are **volatile** — see
[Volatile functions](workbook-and-expressions.md#volatile-functions).

**Where the count starts.** Serial **1** is **1900-01-01**, serial 2 is 1900-01-02, and so on —
`YEAR(1)`/`MONTH(1)`/`DAY(1)` are 1900/1/1. Serial **0** is Excel's *day zero*: not a real date, but a
placeholder Excel writes as `1900-01-00`, so `DAY(0)` = 0 and `DATE(1900,1,0)` = 0. Nothing exists below it —
`DATE(1900,1,-1)` is `#NUM!` and `DATEVALUE("1899-12-31")` is `#VALUE!`.

**Serial 60 is a day that never happened.** Lotus 1-2-3 treated 1900 as a leap year; Excel copied the bug to
stay file-compatible with it, and has carried a phantom **1900-02-29** at serial 60 ever since. MySheet
carries it for the same reason — so that a serial means the same date in both. But 1900 was not a leap year,
so the functions do not agree on what serial 60 *is*:

- **`DATE` cannot build it.** `DATE(1900,2,29)` rolls into March and gives **61**, and both `DATE(1900,2,28)`
  and `DATE(1900,3,0)` give 59. Only two things reach 60: arithmetic on serials (59 + 1), and
  `DATEVALUE("1900-02-29")` = **60** (also `"2/29/1900"` and `"29-Feb-1900"`).
- **The calendar functions read it as 1900-02-28.** `DAY(60)` = 28, `EDATE(60,0)` = `EOMONTH(60,0)` = 59, and
  `WEEKDAY(60)` = 3 — serial 60 collapses onto serial 59 and reports its weekday.
- **`TEXT` and the 30/360 day counts see a real February 29**, because for them February 1900 has 29 days:
  `TEXT(60,"yyyy-mm-dd")` = `1900-02-29`, `DAYS360(59,61)` = 3, `YEARFRAC(60,61)` = 2/360.
- **Everything that counts days counts serials**, so a span containing serial 60 is one day longer than the
  Gregorian calendar between the same two dates: `DATEDIF(1,61,"D")` = 60, `NETWORKDAYS(1,61)` = 45.

From serial **61** (1900-03-01) upward — where every date a real workbook holds lives — the phantom day is
behind you and none of this applies.

| Function | Arguments | Description |
| --- | --- | --- |
| `DATE` | `DATE(year, month, day)` | Serial from parts; Excel overflow (month 13 → next Jan, day 0 → prior month-end); year 0–1899 adds 1900. |
| `DATEDIF` | `DATEDIF(start, end, unit)` | Difference in `"Y"`/`"M"`/`"D"`/`"MD"`/`"YM"`/`"YD"`; `start > end` → `#NUM!`. `"MD"` is officially unreliable. `"MD"`/`"YD"` count serials from `start` pushed forward by every **whole** month (year) in the span, clamped to the target month's last day (`DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")` = 1). |
| `DATEVALUE` | `DATEVALUE(date_text)` | Parses a date string (invariant `yyyy-MM-dd`, `M/d/yyyy`, `d-MMM-yyyy`, …) to a whole-day serial; unparseable → `#VALUE!`, and so is any date **before 1900-01-01** (`DATEVALUE("1899-12-31")`). Excel's phantom `"1900-02-29"` parses, to serial 60. |
| `DAY` | `DAY(serial)` | Day of the month (1–31). |
| `DAYS` | `DAYS(end, start)` | Whole days between two dates (may be negative). |
| `DAYS360` | `DAYS360(start, end, [method])` | 30/360 day count; US (NASD) default, `TRUE` = European. The US method has **no end-of-February adjustment** and **no roll of a month-end `end` to the 1st of the next month** — it matches Excel as measured, not the Microsoft page (see the notes below the table). |
| `EDATE` | `EDATE(start, months)` | The same day-of-month `months` away, clamped to the month end; a result before serial 0 → `#NUM!`. |
| `EOMONTH` | `EOMONTH(start, months)` | Last day of the month `months` away from `start`; a result before serial 0 → `#NUM!` (`EOMONTH(1,-1)` = 0 is still valid). |
| `HOUR` | `HOUR(serial)` | Hour (0–23) of the time fraction. |
| `ISOWEEKNUM` | `ISOWEEKNUM(serial)` | ISO 8601 week number (weeks start Monday; week 1 holds the first Thursday). |
| `MINUTE` | `MINUTE(serial)` | Minute (0–59) of the time fraction. |
| `MONTH` | `MONTH(serial)` | Month (1–12). |
| `NETWORKDAYS` | `NETWORKDAYS(start, end, [holidays])` | Working days in `[start, end]` (inclusive); Sat/Sun and `holidays` excluded. |
| `NETWORKDAYS.INTL` | `NETWORKDAYS.INTL(start, end, [weekend], [holidays])` | `NETWORKDAYS` with a custom weekend (number 1–7/11–17 or a 7-char `"0000011"` mask). A weekend number outside the table → `#NUM!`; an all-weekend `"1111111"` mask is a legitimate **0** here, not an error — unlike `WORKDAY.INTL`. |
| `NOW` | `NOW()` | Volatile: the current local date **and** time as a serial. See [Volatile functions](workbook-and-expressions.md#volatile-functions). |
| `SECOND` | `SECOND(serial)` | Second (0–59), rounded to the nearest second. |
| `TIME` | `TIME(hour, minute, second)` | Time-of-day fraction; components 0–32767 roll over, taken mod 24h; negative → `#NUM!`. |
| `TIMEVALUE` | `TIMEVALUE(time_text)` | Parses a time string (`HH:mm[:ss]`, `h:mm[:ss] AM/PM`) to a `[0,1)` fraction; unparseable → `#VALUE!`. |
| `TODAY` | `TODAY()` | Volatile: the current local date as a whole-day serial (the floor of `NOW()`). See [Volatile functions](workbook-and-expressions.md#volatile-functions). |
| `WEEKDAY` | `WEEKDAY(serial, [return_type])` | Day of week; `return_type` 1/2/3 and 11–17 (see the WEEKDAY table). The weekday is Excel's Lotus-inherited one, read straight off the serial: `WEEKDAY(1)` = 1 (Sunday) although 1900-01-01 was really a Monday — matching Excel. |
| `WEEKNUM` | `WEEKNUM(serial, [return_type])` | Week of year; System 1 for 1/2/11–17, ISO 8601 (System 2) for 21. |
| `WORKDAY` | `WORKDAY(start, days, [holidays])` | Date `days` working days from `start` (start excluded); negative walks backward. |
| `WORKDAY.INTL` | `WORKDAY.INTL(start, days, [weekend], [holidays])` | `WORKDAY` with a custom weekend. Three separate answers, all measured on Aspose.Cells 26.6.0 (2026-09-09): a weekend **number** outside 1–7/11–17 → `#NUM!`; an **all-weekend** `"1111111"` mask → `#VALUE!`, not the `#NUM!` the Microsoft page implies; and `days` = 0 never moves, so it answers the `start` serial even under an all-weekend mask (`WORKDAY.INTL(45366,0,"1111111")` = 45366). |
| `YEAR` | `YEAR(serial)` | Calendar year (1900–9999). |
| `YEARFRAC` | `YEARFRAC(start, end, [basis])` | Year fraction on basis 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), 4 (European 30/360). Basis 0 has **no end-of-February rule** and pulls a February-end `start` to day 30 only **after** testing a day-31 `end`, so it deliberately disagrees with `DAYS360` on some February-end pairs (see the notes below the table). |

**`TEXT` is the only place a 1900-02-29 is printed** (`TEXT(60,"yyyy-mm-dd")` = `1900-02-29`) **and the only
place a day zero is** (`TEXT(0,"yyyy-mm-dd")` = `1900-01-00`). The one exception inside a format is
`ddd`/`dddd`: they name the Lotus weekday, and asking for one also pulls the printed day *number* back onto
February 28 — `TEXT(60,"yyyy-mm-dd dddd")` = `1900-02-28 Tuesday`, where `TEXT(60,"yyyy-mm-dd")` keeps the 29.
A negative serial is `#VALUE!` in `TEXT`, not the `#NUM!` the date functions answer.

**Working days in January–February 1900 walk the real calendar, not Excel's.** `WORKDAY`, `WORKDAY.INTL`,
`NETWORKDAYS` and `NETWORKDAYS.INTL` match Excel exactly from serial 61 (1900-03-01) on — every date a real
workbook holds. Below that, Excel's own answers are not self-consistent, so there is no rule to reproduce and
the walk simply keeps following the real Gregorian weekday it follows everywhere else — above serial 61 the two
agree, below it they can differ. Measured on Aspose.Cells 26.6.0 (2026-09-09), Excel:

1. gives the same day for the 4th and the 5th working day from one start — `WORKDAY(6,4)` = `WORKDAY(6,5)` =
   12 — so its result is not a function of `days`;
2. does the same under a one-day weekend: `WORKDAY.INTL(1,5,"1000000")` = `WORKDAY.INTL(1,6,"1000000")` = 7;
3. answers `WORKDAY(58,4)` = 64, a serial it *itself* treats as a weekend (`WEEKDAY(64)` = 1,
   `NETWORKDAYS(64,64)` = 0, `TEXT(64,"dddd")` = Sunday);
4. and is not additive over a split range: `NETWORKDAYS(58,62)` = 4, while `NETWORKDAYS(58,58)` +
   `NETWORKDAYS(59,61)` + `NETWORKDAYS(62,62)` = 1 + 3 + 1 = 5 — a total no per-day working/non-working
   verdict can produce.

**The divergence below serial 61 is not one rule, and not a fixed handful of rows.** One day is visibly
involved — 1900-01-05 is a **Friday** on the real calendar and a Thursday on Excel's Lotus weekday — but that
is an observation about one day, not an explanation of the window: while this was being decided, fourteen
candidate rules were fitted to a 580-row `WORKDAY` sweep and the best of them was still wrong on **17** of
those rows. So the rows below are **examples** of the divergence, not an exhaustive set — MySheet first, Excel
second: `WORKDAY(5,1)` **8** / 6, `WORKDAY(6,1)` **8** / 9, `WORKDAY(6,4)` **11** / 12, `WORKDAY(13,1)` **15**
/ 16. Other shapes below serial 61 differ too, the four self-contradictions above among them: `WORKDAY(58,4)`
**62** / 64, `NETWORKDAYS(58,62)` **5** / 4, `WORKDAY.INTL(1,5,"1000000")` **6** / 7, and
`NETWORKDAYS(58,62,H)` with `H` holding serial 59 **4** / 3.

What *does* hold is the boundary at serial 61, and the rows that depend on the phantom day being a working day
of the walk — these agree on both engines: `WORKDAY(59,1)` = 60, `WORKDAY(60,-1)` = 59, `NETWORKDAYS(59,61)` =
3, `NETWORKDAYS(1,61)` = 45. Every number in the two paragraphs above is measured (Aspose.Cells 26.6.0,
2026-09-09, plain cell entry) and pinned by a test, MySheet's side included.

**Three changes in 3.17.0 that are not about 1900** — they move results on modern dates too:

- **`YEARFRAC` basis 0 lost its end-of-February rule, and its two 30/360 steps swapped order.** An `end` on
  the last day of February is no longer promoted to a nominal day 30, so
  `YEARFRAC(DATE(2024,2,29),DATE(2025,2,28),0)` is 358/360 and `YEARFRAC(DATE(2023,2,28),DATE(2024,2,29),0)` is
  359/360 — both were exactly 1 before 3.17.0. The pull of a February-end `start` to day 30 survives, but it
  now runs *after* the day-31 `end` test and so no longer drags a day-31 end down with it:
  `YEARFRAC(DATE(2023,2,28),DATE(2023,3,31),0)` = 31/360 (was 30/360).
- **`DAYS360` (US) lost the first-of-next-month roll entirely**, and orders its February pull the *other* way
  round from `YEARFRAC` basis 0. A month-end `end` is no longer rolled to the 1st of the next month:
  `DAYS360(DATE(2011,1,1),DATE(2011,4,30))` = 119 (was 120), `DAYS360(DATE(2011,1,15),DATE(2011,9,30))` = 255
  (was 256), `DAYS360(DATE(2024,1,31),DATE(2024,2,29))` = 29 (was 30),
  `DAYS360(DATE(2024,1,16),DATE(2024,2,29))` = 43 (was 45), `DAYS360(DATE(2024,2,28),DATE(2024,2,29))` = 1
  (was 3). A day-31 `end` still drops to 30 once the adjusted `start` reached 30, so
  `DAYS360(DATE(2011,1,1),DATE(2011,12,31))` = 360 is unchanged. And because the February pull runs *first*
  here, it does drag a day-31 end down: `DAYS360(DATE(2023,2,28),DATE(2023,3,31))` = 30, against `YEARFRAC`'s
  31/360 above — the two functions now **deliberately disagree** on February-end and day-31 pairs, exactly as
  Excel does.
- **`DATEDIF`'s `"MD"` and `"YD"` anchor differently.** Both now count serials from `start` pushed forward by
  every whole month (year) in the span, with that shift clamped to the target month's last day, instead of
  borrowing the previous month's length: `DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")` = 1, where the old
  formula answered −1, and `DATEDIF(DATE(2024,2,29),DATE(2025,3,1),"YD")` = 1.

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

MySheet implements 306 of the ~520 functions in [Microsoft's official Excel function
catalog](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
grouped below by Microsoft's own categories (✅ implemented, ⬜ not yet, ✖ out of scope by design).
**35 functions are permanently out of scope** — they depend on external services, UI environment, or
features the engine deliberately does not model (see [Out of scope](#out-of-scope-by-design) below) —
leaving a viable catalog of ~485 that the roadmap tracks against. A few names are cross-listed by
Microsoft in more than one category — `CONCATENATE` (Text and Compatibility), `FLOOR` (Math and
Compatibility) and `FORECAST` (Statistical and Compatibility) — so per-category counts don't sum to a
single unique total, and a category below can be **larger** than the same-named table above, which
documents each function exactly once. Those three names are the whole of the difference: Statistical is
60 below against a table of 59 because `FORECAST` is documented with the compatibility aliases, and
Compatibility is 13 below against a table of 11 because `CONCATENATE` and `FLOOR` are documented under
Text and Math. See [`FunctionRegistry.cs`](../Danfma.MySheet/Parsing/FunctionRegistry.cs) for the
authoritative registered list.

<details open>
<summary><strong>Financial</strong> — 55/55</summary>

✅ `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FV` `FVSCHEDULE` `INTRATE` `IPMT` `IRR` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `NPER` `NPV` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PMT` `PPMT` `PRICE` `PRICEDISC` `PRICEMAT` `PV` `RATE` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

</details>

<details open>
<summary><strong>Logical</strong> — 12/19</summary>

✅ `AND` `FALSE` `IF` `IFERROR` `IFNA` `IFS` `LET` `NOT` `OR` `SWITCH` `TRUE` `XOR`

⬜ `BYCOL` `BYROW` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN`

</details>

<details open>
<summary><strong>Lookup and Reference</strong> — 17/40</summary>

✅ `ADDRESS` `AREAS` `CHOOSE` `COLUMN` `COLUMNS` `FORMULATEXT` `HLOOKUP` `INDEX` `INDIRECT` `LOOKUP` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP` `XMATCH`

⬜ `CHOOSECOLS` `CHOOSEROWS` `DROP` `EXPAND` `FILTER` `HSTACK` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS`

✖ `GETPIVOTDATA` `GROUPBY` `HYPERLINK` `IMAGE` `PIVOTBY` `RTD`

</details>

<details open>
<summary><strong>Math and Trigonometry</strong> — 75/82</summary>

✅ `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `AGGREGATE` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MOD` `MROUND` `MULTINOMIAL` `ODD` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `RAND` `RANDBETWEEN` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

⬜ `MDETERM` `MINVERSE` `MMULT` `MUNIT` `PERCENTOF` `RANDARRAY` `SEQUENCE`

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
<summary><strong>Date and Time</strong> — 25/25</summary>

✅ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

`NOW` and `TODAY` are **volatile** — they read the workbook's injectable clock and refresh on
`Recalculate()`. The category is now complete (25/25). See
[Volatile functions](workbook-and-expressions.md#volatile-functions).

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
