using System.Diagnostics.CodeAnalysis;

namespace Danfma.MySheet.Expressions;

/// <summary>Discriminador do conteúdo de um <see cref="ComputedValue"/>.</summary>
public enum ComputedValueKind : byte
{
    Blank,
    Number,
    Boolean,
    Text,
    Error,
    Reference,
}

/// <summary>
/// Resultado de avaliar uma <see cref="Expression"/>, como um value type opaco que emula uma union sem
/// <c>Unsafe</c>: um campo <c>double</c> (Number/Boolean/Error-code) mais um <c>object?</c> (Text/Reference)
/// e uma tag de 1 byte. Number/Boolean/Blank/Error não alocam — só Text/Reference guardam uma referência que
/// já existe. A extração é explícita e por tipo exato (<c>TryGet*</c>/<c>As*</c>/<c>To*</c>); a coerção
/// estilo-Excel é interna ao engine, não faz parte desta superfície.
/// </summary>
public readonly struct ComputedValue
{
    private readonly double _num;    // Number | Boolean(0/1) | Error(código)
    private readonly object? _ref;   // Text(string) | Reference(Reference) — null nos escalares
    private readonly ComputedValueKind _kind;

    private ComputedValue(double num, object? reference, ComputedValueKind kind)
    {
        _num = num;
        _ref = reference;
        _kind = kind;
    }

    public static readonly ComputedValue Blank = new(0d, null, ComputedValueKind.Blank);

    public static ComputedValue Number(double value) => new(value, null, ComputedValueKind.Number);

    public static ComputedValue Boolean(bool value) => new(value ? 1d : 0d, null, ComputedValueKind.Boolean);

    public static ComputedValue Text(string? value) =>
        value is null ? Blank : new(0d, value, ComputedValueKind.Text);

    public static ComputedValue Error(Error error) => new(error.Code, null, ComputedValueKind.Error);

    public static ComputedValue Reference(Reference reference) =>
        new(0d, reference, ComputedValueKind.Reference);

    public ComputedValueKind Kind => _kind;

    // --- Conversões implícitas: SÓ de entrada (construir), nunca de saída (extrair é sempre explícito) ---

    public static implicit operator ComputedValue(double value) => Number(value);

    public static implicit operator ComputedValue(bool value) => Boolean(value);

    public static implicit operator ComputedValue(string? value) => Text(value);

    public static implicit operator ComputedValue(Error error) => Error(error);

    // --- TryGet*: seguro, out+bool, tipo exato (Number e Boolean NÃO se cruzam) ---

    public bool TryGetNumber(out double value)
    {
        if (_kind == ComputedValueKind.Number)
        {
            value = _num;
            return true;
        }

        value = 0d;
        return false;
    }

    public bool TryGetBoolean(out bool value)
    {
        if (_kind == ComputedValueKind.Boolean)
        {
            value = _num != 0d;
            return true;
        }

        value = false;
        return false;
    }

    public bool TryGetText([NotNullWhen(true)] out string? value)
    {
        if (_kind == ComputedValueKind.Text)
        {
            value = (string)_ref!;
            return true;
        }

        value = null;
        return false;
    }

    public bool TryGetError(out Error error)
    {
        if (_kind == ComputedValueKind.Error)
        {
            error = Expressions.Error.FromCode((int)_num);
            return true;
        }

        error = default;
        return false;
    }

    public bool TryGetReference([NotNullWhen(true)] out Reference? reference)
    {
        if (_kind == ComputedValueKind.Reference)
        {
            reference = (Reference)_ref!;
            return true;
        }

        reference = null;
        return false;
    }

    // --- As*: açúcar nullable sobre o TryGet ---

    public double? AsDouble() => _kind == ComputedValueKind.Number ? _num : null;

    public bool? AsBoolean() => _kind == ComputedValueKind.Boolean ? _num != 0d : null;

    public string? AsString() => _kind == ComputedValueKind.Text ? (string)_ref! : null;

    // --- To*: assert estrito — LANÇA em não-correspondência (sem coerção) ---

    public double ToDouble() =>
        _kind == ComputedValueKind.Number ? _num : throw NotOfKind(ComputedValueKind.Number);

    public bool ToBoolean() =>
        _kind == ComputedValueKind.Boolean ? _num != 0d : throw NotOfKind(ComputedValueKind.Boolean);

    public string ToText() =>
        _kind == ComputedValueKind.Text ? (string)_ref! : throw NotOfKind(ComputedValueKind.Text);

    private InvalidOperationException NotOfKind(ComputedValueKind expected) =>
        new($"ComputedValue é {_kind}, esperado {expected}.");

    /// <summary>
    /// Enumera os VALORES de uma referência (via cache), não expressões. Vazio quando não for
    /// <see cref="ComputedValueKind.Reference"/>.
    /// </summary>
    public IEnumerable<ComputedValue> EnumerateValues(EvaluationContext context)
    {
        if (_kind != ComputedValueKind.Reference)
        {
            yield break;
        }

        switch (_ref)
        {
            case RangeReference range:
                foreach (var value in range.ExpandValues(context))
                {
                    yield return From(value);
                }

                break;

            case Reference reference:
                yield return From(reference.Compute(context));
                break;
        }
    }

    /// <summary>Overload de conveniência (espelha <see cref="Expression.Compute(Workbook)"/>).</summary>
    public IEnumerable<ComputedValue> EnumerateValues(Workbook workbook) =>
        EnumerateValues(new EvaluationContext(workbook));

    /// <summary>Ponte para o mundo <c>object?</c> (interop / call sites legados). Boxa escalares numéricos.</summary>
    public object? AsObject() => _kind switch
    {
        ComputedValueKind.Number => _num,
        ComputedValueKind.Boolean => _num != 0d,
        ComputedValueKind.Text => _ref,
        ComputedValueKind.Error => Expressions.Error.FromCode((int)_num).ToErrorValue(),
        ComputedValueKind.Reference => _ref,
        _ => null,
    };

    /// <summary>Ponte inversa: envolve um valor cru do avaliador legado num <see cref="ComputedValue"/>.</summary>
    public static ComputedValue From(object? value) => value switch
    {
        null => Blank,
        double d => Number(d),
        bool b => Boolean(b),
        string s => Text(s),
        ErrorValue e => Error(Expressions.Error.FromDisplay(e.ErrorCode)),
        Reference r => Reference(r),
        ComputedValue cv => cv,
        _ => Error(Expressions.Error.Value),
    };
}
