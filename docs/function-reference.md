# Function reference

MySheet implements **52 built-in functions**. The authoritative registered list is the `Functions` map
in [`Danfma.MySheet/Parsing/Parser.cs`](../Danfma.MySheet/Parsing/Parser.cs) — this page is derived from
it. Argument counts are validated **at parse time**: calling a built-in with an unsupported number of
arguments throws a `ParseException`, just as Excel rejects the formula at entry.

Beyond these, you can add your own functions with
[`workbook.RegisterFunction`](custom-functions.md); unknown names evaluate to `#NAME?`.

Conventions below: `[argument]` is optional; `…` means the function is variadic. "Range-aware" means
range arguments (`A1:B10`, unions, and reference results such as `OFFSET`'s) are expanded cell by cell.

## Logical (7)

| Function | Arguments | Description |
| --- | --- | --- |
| `AND` | `AND(logical1, [logical2], …)` | `TRUE` if every argument is truthy. |
| `IF` | `IF(condition, value_if_true, [value_if_false])` | Conditional; only the taken branch is evaluated. |
| `IFERROR` | `IFERROR(value, value_if_error)` | `value`, or the fallback when `value` is any error. |
| `IFNA` | `IFNA(value, value_if_na)` | `value`, or the fallback only when `value` is `#N/A`. |
| `LET` | `LET(name1, value1, [name2, value2, …], calculation)` | Binds names usable in `calculation` (e.g. `=LET(x, A1*2, x+x)`). Names are local to the formula. |
| `NOT` | `NOT(logical)` | Logical negation. |
| `OR` | `OR(logical1, [logical2], …)` | `TRUE` if any argument is truthy. |

## Math and trigonometry (7)

| Function | Arguments | Description |
| --- | --- | --- |
| `ABS` | `ABS(number)` | Absolute value. |
| `INT` | `INT(number)` | Rounds down to the nearest integer. |
| `ROUND` | `ROUND(number, num_digits)` | Rounds to the given number of digits. |
| `ROUNDUP` | `ROUNDUP(number, num_digits)` | Rounds away from zero. |
| `SUM` | `SUM([number1], …)` | Sum of all numeric values; range-aware. |
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | Sums the cells matching a criteria (e.g. `">10"`). |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | Sum under multiple criteria-range pairs. |

## Statistical (8)

| Function | Arguments | Description |
| --- | --- | --- |
| `AVERAGE` | `AVERAGE([number1], …)` | Arithmetic mean of the numeric values; range-aware. |
| `COUNT` | `COUNT([value1], …)` | Counts numeric values; range-aware. |
| `COUNTA` | `COUNTA(value1, …)` | Counts non-blank values; range-aware. |
| `COUNTBLANK` | `COUNTBLANK(range, …)` | Counts blank cells. |
| `COUNTIF` | `COUNTIF(range, criteria)` | Counts the cells matching a criteria. |
| `COUNTIFS` | `COUNTIFS(criteria_range1, criteria1, …)` | Count under multiple criteria-range pairs. |
| `MAX` | `MAX([number1], …)` | Largest numeric value; range-aware. |
| `MIN` | `MIN([number1], …)` | Smallest numeric value; range-aware. |

## Text (11)

| Function | Arguments | Description |
| --- | --- | --- |
| `CONCAT` | `CONCAT(text1, …)` | Concatenates values; range-aware. |
| `CONCATENATE` | `CONCATENATE(text1, …)` | Legacy alias of concatenation (scalar arguments). |
| `LEFT` | `LEFT(text, [num_chars])` | Leading characters (default 1). |
| `LEN` | `LEN(text)` | Text length. |
| `LOWER` | `LOWER(text)` | Lower-cases the text. |
| `MID` | `MID(text, start_num, num_chars)` | Substring by 1-based position and length. |
| `TEXT` | `TEXT(value, format_text)` | Formats a value (number and date formats, e.g. `"0.00"`, `"dd/mm/yyyy"`). |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, …)` | Joins values with a delimiter; range-aware. |
| `TRIM` | `TRIM(text)` | Removes excess whitespace. |
| `UPPER` | `UPPER(text)` | Upper-cases the text. |
| `VALUE` | `VALUE(text)` | Converts text to a number. |

## Lookup and reference (7)

| Function | Arguments | Description |
| --- | --- | --- |
| `INDEX` | `INDEX(range, row_num, [column_num])` | The value at a 1-based position inside a range. |
| `MATCH` | `MATCH(lookup_value, lookup_range, [match_type])` | 1-based position of a value in a range (`match_type`: 1 approximate ascending — default, 0 exact, -1 approximate descending). |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | A reference displaced (and optionally resized) from a starting reference; may return a multi-cell reference for range-aware consumers. |
| `ROW` | `ROW([reference])` | Row number of the reference — or of the current cell when called with no argument. |
| `ROWS` | `ROWS(range)` | Number of rows in the range. |
| `VLOOKUP` | `VLOOKUP(lookup_value, table_range, col_index_num, [range_lookup])` | Vertical lookup in the first column of a table; exact or approximate. |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])` | Modern lookup with not-found fallback and match/search modes. |

## Information (3)

| Function | Arguments | Description |
| --- | --- | --- |
| `ISBLANK` | `ISBLANK(value)` | `TRUE` for a blank value. |
| `ISNUMBER` | `ISNUMBER(value)` | `TRUE` for a numeric value. |
| `SHEET` | `SHEET([value])` | 1-based sheet position (tab order) of a reference or sheet name — or of the current sheet with no argument. |

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

## Excel function coverage

MySheet implements 52 of the ~520 functions in [Microsoft's official Excel function
catalog](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
grouped below by Microsoft's own categories (✅ implemented, ⬜ not yet). A few names are cross-listed by
Microsoft in more than one category (e.g. `CONCATENATE` in both Text and Compatibility, `LET` in both
Logical and Math), so per-category counts don't sum to a single unique total — see `Parser.cs` for the
authoritative registered list.

<details open>
<summary><strong>Financial</strong> — 9/55</summary>

✅ `FV` `IPMT` `IRR` `NPER` `NPV` `PMT` `PPMT` `PV` `RATE`

⬜ `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FVSCHEDULE` `INTRATE` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PRICE` `PRICEDISC` `PRICEMAT` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

</details>

<details open>
<summary><strong>Logical</strong> — 7/19</summary>

✅ `AND` `IF` `IFERROR` `IFNA` `LET` `NOT` `OR`

⬜ `BYCOL` `BYROW` `FALSE` `IFS` `LAMBDA` `MAKEARRAY` `MAP` `REDUCE` `SCAN` `SWITCH` `TRUE` `XOR`

</details>

<details open>
<summary><strong>Lookup and Reference</strong> — 7/40</summary>

✅ `INDEX` `MATCH` `OFFSET` `ROW` `ROWS` `VLOOKUP` `XLOOKUP`

⬜ `ADDRESS` `AREAS` `CHOOSE` `CHOOSECOLS` `CHOOSEROWS` `COLUMN` `COLUMNS` `DROP` `EXPAND` `FILTER` `FORMULATEXT` `GETPIVOTDATA` `GROUPBY` `HLOOKUP` `HSTACK` `HYPERLINK` `IMAGE` `INDIRECT` `LOOKUP` `PIVOTBY` `RTD` `SORT` `SORTBY` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VSTACK` `WRAPCOLS` `WRAPROWS` `XMATCH`

</details>

<details open>
<summary><strong>Math and Trigonometry</strong> — 7/82</summary>

✅ `ABS` `INT` `ROUND` `ROUNDUP` `SUM` `SUMIF` `SUMIFS`

⬜ `ACOS` `ACOSH` `ACOT` `ACOTH` `AGGREGATE` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MDETERM` `MINVERSE` `MMULT` `MOD` `MROUND` `MULTINOMIAL` `MUNIT` `ODD` `PERCENTOF` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `RAND` `RANDARRAY` `RANDBETWEEN` `ROMAN` `ROUNDDOWN` `SEC` `SECH` `SERIESSUM` `SEQUENCE` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

</details>

<details open>
<summary><strong>Statistical</strong> — 8/111</summary>

✅ `AVERAGE` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `MAX` `MIN`

⬜ `AVEDEV` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `BETA.DIST` `BETA.INV` `BINOM.DIST` `BINOM.DIST.RANGE` `BINOM.INV` `CHISQ.DIST` `CHISQ.DIST.RT` `CHISQ.INV` `CHISQ.INV.RT` `CHISQ.TEST` `CONFIDENCE.NORM` `CONFIDENCE.T` `CORREL` `COVARIANCE.P` `COVARIANCE.S` `DEVSQ` `EXPON.DIST` `F.DIST` `F.DIST.RT` `F.INV` `F.INV.RT` `F.TEST` `FISHER` `FISHERINV` `FORECAST` `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` `FORECAST.LINEAR` `FREQUENCY` `GAMMA` `GAMMA.DIST` `GAMMA.INV` `GAMMALN` `GAMMALN.PRECISE` `GAUSS` `GEOMEAN` `GROWTH` `HARMEAN` `HYPGEOM.DIST` `INTERCEPT` `KURT` `LARGE` `LINEST` `LOGEST` `LOGNORM.DIST` `LOGNORM.INV` `MAXA` `MAXIFS` `MEDIAN` `MINA` `MINIFS` `MODE.MULT` `MODE.SNGL` `NEGBINOM.DIST` `NORM.DIST` `NORM.INV` `NORM.S.DIST` `NORM.S.INV` `PEARSON` `PERCENTILE.EXC` `PERCENTILE.INC` `PERCENTRANK.EXC` `PERCENTRANK.INC` `PERMUT` `PERMUTATIONA` `PHI` `POISSON.DIST` `PROB` `QUARTILE.EXC` `QUARTILE.INC` `RANK.AVG` `RANK.EQ` `RSQ` `SKEW` `SKEW.P` `SLOPE` `SMALL` `STANDARDIZE` `STDEV.P` `STDEV.S` `STDEVA` `STDEVPA` `STEYX` `T.DIST` `T.DIST.2T` `T.DIST.RT` `T.INV` `T.INV.2T` `T.TEST` `TREND` `TRIMMEAN` `VAR.P` `VAR.S` `VARA` `VARPA` `WEIBULL.DIST` `Z.TEST`

</details>

<details open>
<summary><strong>Text</strong> — 11/49</summary>

✅ `CONCAT` `CONCATENATE` `LEFT` `LEN` `LOWER` `MID` `TEXT` `TEXTJOIN` `TRIM` `UPPER` `VALUE`

⬜ `ASC` `ARRAYTOTEXT` `BAHTTEXT` `CHAR` `CLEAN` `CODE` `DBCS` `DETECTLANGUAGE` `DOLLAR` `EXACT` `FIND` `FINDB` `FIXED` `LEFTB` `LENB` `MIDB` `NUMBERVALUE` `PHONETIC` `PROPER` `REGEXEXTRACT` `REGEXREPLACE` `REGEXTEST` `REPLACE` `REPLACEB` `REPT` `RIGHT` `RIGHTB` `SEARCH` `SEARCHB` `SUBSTITUTE` `T` `TEXTAFTER` `TEXTBEFORE` `TEXTSPLIT` `TRANSLATE` `UNICHAR` `UNICODE` `VALUETOTEXT`

</details>

<details open>
<summary><strong>Information</strong> — 3/22</summary>

✅ `ISBLANK` `ISNUMBER` `SHEET`

⬜ `CELL` `ERROR.TYPE` `INFO` `ISERR` `ISERROR` `ISEVEN` `ISFORMULA` `ISLOGICAL` `ISNA` `ISNONTEXT` `ISODD` `ISOMITTED` `ISREF` `ISTEXT` `N` `NA` `SHEETS` `STOCKHISTORY` `TYPE`

</details>

<details>
<summary><strong>Date and Time</strong> — 0/25</summary>

⬜ `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

</details>

<details>
<summary><strong>Compatibility (legacy aliases)</strong> — 1/41</summary>

✅ `CONCATENATE`

⬜ `BETADIST` `BETAINV` `BINOMDIST` `CHIDIST` `CHIINV` `CHITEST` `CONFIDENCE` `COVAR` `CRITBINOM` `EXPONDIST` `FDIST` `FINV` `FLOOR` `FORECAST` `FTEST` `GAMMADIST` `GAMMAINV` `HYPGEOMDIST` `LOGINV` `LOGNORMDIST` `MODE` `NEGBINOMDIST` `NORMDIST` `NORMINV` `NORMSDIST` `NORMSINV` `PERCENTILE` `PERCENTRANK` `POISSON` `QUARTILE` `RANK` `STDEV` `STDEVP` `TDIST` `TINV` `TTEST` `VAR` `VARP` `WEIBULL` `ZTEST`

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

⬜ `CUBEKPIMEMBER` `CUBEMEMBER` `CUBEMEMBERPROPERTY` `CUBERANKEDMEMBER` `CUBESET` `CUBESETCOUNT` `CUBEVALUE`

</details>

<details>
<summary><strong>Web</strong> — 0/3</summary>

⬜ `ENCODEURL` `FILTERXML` `WEBSERVICE`

</details>

<details>
<summary><strong>User Defined</strong> — 0/3</summary>

⬜ `CALL` `EUROCONVERT` `REGISTER.ID`

</details>

## See also

- [Custom functions](custom-functions.md) — filling gaps in the coverage yourself.
- [Workbook, sheets and expressions](workbook-and-expressions.md) — operators and reference syntax.
