# Custom functions

MySheet's function set is extensible: register a .NET delegate under a name, and formulas can call it
like any built-in — `=DOUBLE(A1)`, `=RISK_SCORE(B1:B10, 0.95)`, and so on. This is the main extension
point of the engine, and it is designed around two ideas: **arguments arrive unevaluated** (so your
function controls evaluation, enabling laziness and short-circuiting), and **returning scalars is
frictionless** (implicit conversions do the wrapping).

## Registering

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

workbook.RegisterFunction("DOUBLE", (arguments, wb) =>
{
    var value = arguments[0].Evaluate(wb);

    return value.TryGetNumber(out var number)
        ? number * 2
        : ComputedValue.Error(Error.Value);
});

sheet["A1"] = ExpressionParser.Parse("=DOUBLE(21)", sheet);
double result = workbook.GetCellValue("Sheet1", "A1").ToDouble();   // 42.0
```

The delegate signature is:

```csharp
public delegate ComputedValue CustomFunction(Expression[] arguments, Workbook workbook);
```

- `arguments` are the **raw, unevaluated** expression nodes from the call site.
- `workbook` is the evaluation context — pass it to `arguments[i].Evaluate(workbook)` when you want a
  value.
- The return type is [`ComputedValue`](computed-value.md).

Function names:

- Names are **case-insensitive** (`=double(21)` works) and may contain underscores (`=RISK_SCORE(...)`).
- Excel stores newer functions with an `_xlfn.` prefix; the parser normalizes it away, so a custom
  function named `MYFN` also matches `=_xlfn.MYFN()` in a loaded `.xlsx` file.
- **Built-ins cannot be overridden.** The names in the parser's function table always parse to their
  built-in nodes; the custom registry is only consulted for other names.
- Calling a name that was never registered is not an exception — it evaluates to `#NAME?`, like Excel.

## Lazy arguments and short-circuiting

Because arguments arrive unevaluated, your function decides *what* gets evaluated and *when*. Expensive
or side-effecting branches are simply never touched unless you evaluate them:

```csharp
// Returns the first argument that evaluates to a non-blank value.
// Later arguments are never evaluated once a match is found.
workbook.RegisterFunction("FIRSTNONBLANK", (arguments, wb) =>
{
    foreach (var argument in arguments)
    {
        var value = argument.Evaluate(wb);

        if (value.Kind != ComputedValueKind.Blank)
        {
            return value;
        }
    }

    return ComputedValue.Blank;
});
```

This is the same mechanism that lets `IF` evaluate only the taken branch.

## Returning values

Scalars convert implicitly, so the common cases need no ceremony:

```csharp
workbook.RegisterFunction("PI2", (_, _) => 6.283185307179586);   // double → Number
workbook.RegisterFunction("YES", (_, _) => true);                // bool → Boolean
workbook.RegisterFunction("GREET", (_, _) => "hello");           // string → Text (null → Blank)
```

For everything else, use the factories:

```csharp
workbook.RegisterFunction("SAFEDIV", (arguments, wb) =>
{
    var left = arguments[0].Evaluate(wb);
    var right = arguments[1].Evaluate(wb);

    if (!left.TryGetNumber(out var numerator) || !right.TryGetNumber(out var denominator))
    {
        return ComputedValue.Error(Error.Value);   // wrong argument types → #VALUE!
    }

    if (denominator == 0)
    {
        return ComputedValue.Blank;                // this function's choice: blank instead of #DIV/0!
    }

    return numerator / denominator;
});
```

Guidelines:

- Signal user-facing failures by **returning** `ComputedValue.Error(...)` — not by throwing. Errors
  propagate through formulas the way Excel errors do (and `IFERROR`/`IFNA` can catch them).
- Reserve exceptions for genuine bugs; they will propagate out of `Evaluate` as exceptions.

## Accepting ranges and references

A range argument (`=MYFN(A1:A10)`) arrives as a `RangeReference` node — evaluating it directly yields
`#VALUE!` (a bare range has no scalar value). To consume the cells, wrap the reference in a
`ComputedValue` and enumerate its values through the memoized cache:

```csharp
workbook.RegisterFunction("PRODUCT", (arguments, wb) =>
{
    var product = 1.0;

    foreach (var argument in arguments)
    {
        IEnumerable<ComputedValue> values = argument is Reference reference
            ? ComputedValue.Reference(reference).EnumerateValues(wb)
            : [argument.Evaluate(wb)];

        foreach (var value in values)
        {
            if (value.TryGetError(out var error))
            {
                return error;                       // propagate errors, Excel-style
            }

            if (value.TryGetNumber(out var number))
            {
                product *= number;
            }
        }
    }

    return product;
});

sheet["B1"] = ExpressionParser.Parse("=PRODUCT(A1:A3, 2)", sheet);
```

`Reference` covers single cells (`CellReference`), rectangles (`RangeReference`) and unions
(`UnionReference`), so the same branch handles `=PRODUCT(A1)`, `=PRODUCT(A1:A10)` and
`=PRODUCT((A1:A3, C1:C3))`.

## Interaction with memoization

Custom functions participate in the per-cell cache like everything else: once a cell's value is
computed, `GetCellValue` serves it from the cache and the delegate is **not** invoked again until you
call `workbook.InvalidateCache()`. Do not rely on a custom function being re-executed per read — if it
reads external state (time, database, …), the value observed is the one from the first evaluation until
the cache is invalidated. See [Performance](performance.md).

## Serialization and Excel files

A custom-function *call* is a regular AST node (`FunctionCall`) holding the name and the argument
expressions. That node:

- **round-trips through MemoryPack** (`Workbook.Save`/`Load`),
- **parses from `.xlsx` files** loaded with `ExcelFile.Load`, and
- **un-parses** via `ToFormula` (so `FormulaMode.Formulas` exports keep the call text).

The **delegate itself is never serialized** — behavior is code, not data. After loading, re-register
the implementation before evaluating, or the call evaluates to `#NAME?`:

```csharp
workbook.Save("model.mysheet");

var restored = Workbook.Load("model.mysheet");

// Without this, any cell calling CUSTOM(...) evaluates to #NAME?.
restored.RegisterFunction("CUSTOM", (arguments, wb) =>
{
    var a = arguments[0].Evaluate(wb).AsDouble() ?? 0;
    var b = arguments[1].Evaluate(wb).AsDouble() ?? 0;

    return a + b;
});
```

A good pattern for applications is a single `RegisterAll(Workbook workbook)` method that installs every
function your formulas use, called both after `new Workbook()` and after every `Workbook.Load` /
`ExcelFile.Load`.

## See also

- [ComputedValue and errors](computed-value.md) — the values your function receives and returns.
- [Serialization](serialization.md) — what persists and what must be re-registered.
