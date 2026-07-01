# ComputedValue and errors

Evaluating an expression returns a `ComputedValue` — MySheet's single result type. It is a
`readonly struct` that emulates a discriminated union: one `double` field (numbers, booleans, error
codes), one `object?` field (text, references), and a one-byte tag. Producing a number, boolean, blank
or error **allocates nothing**; text and references only carry a reference that already existed.

```csharp
using Danfma.MySheet.Expressions;

ComputedValue value = workbook.GetCellValue("Sheet1", "A3");
```

## Kinds

`value.Kind` is a `ComputedValueKind`:

| Kind | Meaning | Allocation |
| --- | --- | --- |
| `Blank` | An empty cell / omitted argument. | none |
| `Number` | A `double`. Dates are numbers too (Excel serial dates). | none |
| `Boolean` | `TRUE`/`FALSE`. | none |
| `Text` | A `string`. | carries the existing string |
| `Error` | An Excel error (`#DIV/0!`, `#N/A`, …) as an [`Error`](#the-error-struct) struct. | none |
| `Reference` | A reference produced by a function such as `OFFSET` (see [References](#references-and-enumeratevalues)). | carries the existing reference |

## Extracting values

Extraction is **explicit and strict by exact kind** — there is no Excel-style coercion at this surface
(coercion is internal to the engine). In particular, `Number` and `Boolean` do not cross:
`TryGetNumber` on a boolean returns `false`.

Three styles, one per ergonomic need:

### `TryGet*` — safe, out + bool

```csharp
if (value.TryGetNumber(out double number)) { /* ... */ }
if (value.TryGetBoolean(out bool flag)) { /* ... */ }
if (value.TryGetText(out string? text)) { /* ... */ }
if (value.TryGetError(out Error error)) { /* e.g. error == Error.DivZero */ }
if (value.TryGetReference(out Reference? reference)) { /* ... */ }
```

### `As*` — nullable sugar

```csharp
double? number = value.AsDouble();    // null unless Kind == Number
bool? flag = value.AsBoolean();       // null unless Kind == Boolean
string? text = value.AsString();      // null unless Kind == Text
```

### `To*` — strict asserts (throw on mismatch)

```csharp
double number = value.ToDouble();     // throws InvalidOperationException unless Kind == Number
bool flag = value.ToBoolean();        // throws unless Kind == Boolean
string text = value.ToText();         // throws unless Kind == Text
```

Use `To*` when the workbook contract guarantees the type (for example, a numeric result cell in a model
you control) and a mismatch is a genuine bug.

### `AsObject()` — the `object?` bridge

For interop with loosely-typed code, `AsObject()` boxes the value:

| Kind | `AsObject()` returns |
| --- | --- |
| `Blank` | `null` |
| `Number` | boxed `double` |
| `Boolean` | boxed `bool` |
| `Text` | the `string` |
| `Error` | the corresponding `ErrorValue` AST node (e.g. `ErrorValue.DivByZero`) |
| `Reference` | the `Reference` expression |

This is a permanent escape hatch, not the main path — it re-introduces the boxing the struct exists to
avoid, so keep it out of hot loops.

## Constructing values

You mostly construct `ComputedValue`s when writing [custom functions](custom-functions.md). Factories
plus implicit conversions **on input only** (extraction is never implicit):

```csharp
ComputedValue a = ComputedValue.Number(42);
ComputedValue b = ComputedValue.Boolean(true);
ComputedValue c = ComputedValue.Text("hello");
ComputedValue d = ComputedValue.Blank;
ComputedValue e = ComputedValue.Error(Error.Num);

// Implicit conversions from double / bool / string / Error:
ComputedValue f = 42.0;
ComputedValue g = true;
ComputedValue h = "hello";       // a null string converts to Blank
ComputedValue i = Error.Value;
```

## The `Error` struct

`Error` is an allocation-free smart-enum struct wrapping an `int` code — it fits entirely inside a
`ComputedValue`. The well-known Excel errors are named static instances:

| Instance | `Display` |
| --- | --- |
| `Error.Null` | `#NULL!` |
| `Error.DivZero` | `#DIV/0!` |
| `Error.Value` | `#VALUE!` |
| `Error.Ref` | `#REF!` |
| `Error.Name` | `#NAME?` |
| `Error.Num` | `#NUM!` |
| `Error.NA` | `#N/A` |

`ToString()` prints the display text, and equality is by code:

```csharp
if (value.TryGetError(out var error))
{
    Console.WriteLine(error);                    // "#DIV/0!"
    Console.WriteLine(error.Display);            // "#DIV/0!"

    if (error == Error.NA)
    {
        // a lookup found nothing — often fine to treat as "no data"
    }
}
```

Notes:

- A **circular reference** is detected by the memoization layer and surfaces as `Error.Ref` (`#REF!`)
  instead of a stack overflow.
- A **call to an unregistered function** evaluates to `Error.Name` (`#NAME?`).
- `Error` is distinct from `ErrorValue`, which is the serializable AST *node* for a literal error stored
  in a cell (e.g. loaded from an `.xlsx` file). `AsObject()` maps an `Error` back to the corresponding
  `ErrorValue` singleton; evaluating an `ErrorValue` produces the corresponding `Error`.
- Registering custom error codes is deliberately out of scope for now.

## References and `EnumerateValues`

A few functions (currently `OFFSET`) evaluate to a *reference* rather than a scalar —
`Kind == ComputedValueKind.Reference`. `EnumerateValues` walks the referenced cells and yields their
**computed values** (through the memoization cache):

```csharp
var offset = ExpressionParser.Parse("=OFFSET(A1, 0, 0, 3, 1)", sheet);
ComputedValue reference = offset.Evaluate(workbook);

foreach (ComputedValue cell in reference.EnumerateValues(workbook))
{
    if (cell.TryGetNumber(out var number))
    {
        Console.WriteLine(number);
    }
}
```

On any non-`Reference` value, `EnumerateValues` yields nothing. Note that a *bare range* in a formula
(`=A1:B2` used as a scalar) does not produce a `Reference` value — it evaluates to `#VALUE!`, as in
Excel; ranges are consumed by the functions that accept them. To enumerate a range from a custom
function, see [Custom functions — range arguments](custom-functions.md#accepting-ranges-and-references).

## See also

- [Custom functions](custom-functions.md) — producing `ComputedValue`s of your own.
- [Performance](performance.md) — why the union design exists, with measured numbers.
