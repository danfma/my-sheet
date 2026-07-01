# Danfma.MySheet

[![NuGet](https://img.shields.io/nuget/v/Danfma.MySheet.svg)](https://www.nuget.org/packages/Danfma.MySheet/)

A fast, in-memory spreadsheet formula engine for .NET — parse Excel-style formulas, evaluate them, and
extract data, without a full spreadsheet application.

## Features

- **Formula parser** (Pratt parser) for the Excel operator set: `+ - * / ^ %`, `&` (text), comparisons
  `= <> < > <= >=` (Excel cross-type ordering), and references `: ! ,` plus grouping `( )`.
- **~50 built-in functions**: `IF/AND/OR/NOT/IFERROR/IFNA`, `SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/COUNTBLANK`,
  `COUNTIF(S)/SUMIF(S)`, text (`UPPER/LOWER/TRIM/LEN/LEFT/MID/VALUE/CONCAT/CONCATENATE/TEXTJOIN/TEXT`),
  math (`INT/ROUND/ROUNDUP/ABS`), info (`ISNUMBER/ISBLANK`), lookup (`ROW/ROWS/SHEET/MATCH/INDEX/VLOOKUP/
  XLOOKUP/OFFSET`), financial (`PMT/PV/FV/NPER/IPMT/PPMT/NPV/RATE/IRR`), and `LET`.
- **References**: sheet-qualified (`Sheet2!A1`, `'My Sheet'!A1:B2`), absolute markers (`$A$1`), and
  case-insensitive sheet names.
- **Custom functions**: register host functions by name (`Workbook.RegisterFunction`) — they parse and
  serialize with the workbook.
- **Allocation-free evaluation**: `expression.Evaluate(workbook)` returns a `ComputedValue` — an opaque
  value-type union (number / boolean / text / blank / error / reference) that does **not** box numbers.
  Extract strictly with `TryGetNumber`/`AsDouble`/`ToDouble` (and `TryGetError(out Error)`), or get the
  boxed `object?` via `AsObject()` for interop.
- **Memoization**: per-cell cache (storing `ComputedValue` inline — no long-lived per-cell box) with
  explicit invalidation; circular references become `#REF!` instead of a stack overflow.
- **MemoryPack serialization** of the workbook.
- **Excel (.xlsx) interop** via the companion `Danfma.MySheet.Excel` package (OpenXML SDK, cross-platform,
  no Excel installation): load `.xlsx` files into a `Workbook` (formulas become real expression trees,
  re-evaluated by this engine), export workbooks (`ValuesOnly` snapshot or `Formulas` with cached values),
  and merge computed values into an existing template preserving its formatting.

## Excel interop (`Danfma.MySheet.Excel`)

```csharp
using Danfma.MySheet.Excel;

// Load: formulas come back as real expression trees, re-evaluated by MySheet.
Workbook workbook = ExcelFile.Load("input.xlsx");
double total = workbook.GetCellValue("Data", "A4").ToDouble();

// Export: a flattened snapshot (default), or keep formulas + cached values.
workbook.SaveAsExcel("snapshot.xlsx");
workbook.SaveAsExcel("live.xlsx", new ExcelExportOptions { FormulaMode = FormulaMode.Formulas });

// Merge: inject computed values into an existing file, in place (formatting preserved,
// target-cell formulas replaced by the literal values). For the template→report flow,
// copy the pristine template first and merge into the copy:
File.Copy("template.xlsx", "report.xlsx");
workbook.MergeIntoExcel("report.xlsx");
```

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

## Quick start

```
dotnet add package Danfma.MySheet
```

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Parsing;
using Danfma.MySheet.Expressions;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

sheet["A1"] = new NumberValue(1);
sheet["A2"] = new NumberValue(2);
sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);

// Typed, allocation-free result:
ComputedValue result = sheet["A3"].Evaluate(workbook);
double total = result.ToDouble();            // 3.0 (throws if not a number)
if (result.TryGetError(out Error error))     // e.g. error.Display == "#DIV/0!"
{
    // handle the error
}

object? boxed = result.AsObject();           // 3.0 (double) — for object?-based interop
```

For deep dependency chains, wrap the evaluation in a large-stack thread:

```csharp
var value = Workbook.RunWithLargeStack(() => sheet["A3"].Evaluate(workbook));
```

## License

MIT — see [LICENSE](LICENSE).
