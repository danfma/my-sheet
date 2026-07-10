using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// Conversão de bases e numerais romanos: BASE, DECIMAL, ROMAN e ARABIC. ROMAN implementa a forma
// clássica (form 0/TRUE/omitida); as formas concisas 1-4/FALSE são uma limitação documentada e
// devolvem #VALUE! em vez de um numeral potencialmente errado.

[MemoryPackable]
public sealed partial record Base(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var radix) is { } radixError)
        {
            return ComputedValue.Error(radixError);
        }

        var minLength = 0.0;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out minLength) is { } minLengthError
        )
        {
            return ComputedValue.Error(minLengthError);
        }

        var value = Math.Truncate(number);
        var numberBase = Math.Truncate(radix);
        var padding = Math.Truncate(minLength);

        // Docs: número em [0, 2^53), radix em [2, 36], min_length em [0, 255]; fora disso → #NUM!.
        if (
            value < 0
            || value >= ExcelMath.MaxSafeInteger
            || numberBase < 2
            || numberBase > 36
            || padding < 0
            || padding > 255
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        Span<char> buffer = stackalloc char[64];
        var position = buffer.Length;
        var remaining = (long)value;
        var divisor = (long)numberBase;

        if (remaining == 0)
        {
            buffer[--position] = '0';
        }

        while (remaining > 0)
        {
            var digit = (int)(remaining % divisor);

            buffer[--position] = (char)(digit < 10 ? '0' + digit : 'A' + digit - 10);
            remaining /= divisor;
        }

        var text = new string(buffer[position..]);

        return ComputedValue.Text(
            text.Length < (int)padding ? text.PadLeft((int)padding, '0') : text
        );
    }
}

[MemoryPackable]
public sealed partial record DecimalNumber(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var radix) is { } radixError)
        {
            return ComputedValue.Error(radixError);
        }

        var numberBase = Math.Truncate(radix);

        if (numberBase < 2 || numberBase > 36)
        {
            return ComputedValue.Error(Error.Num);
        }

        // Docs: o texto tem no máximo 255 caracteres.
        if (text.Length > 255)
        {
            return ComputedValue.Error(Error.Value);
        }

        var result = 0d;

        foreach (var character in text)
        {
            var digit = character switch
            {
                >= '0' and <= '9' => character - '0',
                >= 'A' and <= 'Z' => character - 'A' + 10,
                >= 'a' and <= 'z' => character - 'a' + 10,
                _ => -1,
            };

            // Dígito inválido para o radix → #NUM!.
            if (digit < 0 || digit >= numberBase)
            {
                return ComputedValue.Error(Error.Num);
            }

            result = result * numberBase + digit;
        }

        return ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record Roman(Expression[] Arguments) : Function
{
    private static readonly (int Value, string Numeral)[] ClassicNumerals =
    [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments.Length == 2)
        {
            var form = Arguments[1].Evaluate(context);

            if (form.TryGetError(out var formError))
            {
                return ComputedValue.Error(formError);
            }

            // TRUE/0/omitido = forma clássica. As formas concisas (1-4 e FALSE) não são
            // implementadas — devolvemos #VALUE! em vez de um numeral não-clássico errado.
            // O booleano é testado antes da coerção porque TRUE coagiria para o número 1 (forma 1).
            if (form.TryGetBoolean(out var boolean))
            {
                if (!boolean)
                {
                    return ComputedValue.Error(Error.Value);
                }
            }
            else if (form.CoerceToNumber(out var formNumber) is { } formNumberError)
            {
                return ComputedValue.Error(formNumberError);
            }
            else if (formNumber != 0)
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        // Docs: número negativo ou maior que 3999 → #VALUE!. ROMAN(0) é a string vazia.
        if (number < 0 || number > 3999)
        {
            return ComputedValue.Error(Error.Value);
        }

        var value = (int)Math.Truncate(number);
        var builder = new StringBuilder();

        foreach (var (numeralValue, numeral) in ClassicNumerals)
        {
            while (value >= numeralValue)
            {
                builder.Append(numeral);
                value -= numeralValue;
            }
        }

        return ComputedValue.Text(builder.ToString());
    }
}

[MemoryPackable]
public sealed partial record Arabic(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = Arguments[0].Evaluate(context);

        if (value.TryGetError(out var error))
        {
            return ComputedValue.Error(error);
        }

        // Docs: números, datas e texto que não é numeral romano → #VALUE! (portanto sem coerção
        // numérica aqui); string vazia → 0; espaços nas pontas ignorados; caixa ignorada.
        string text;

        if (value.Kind == ComputedValueKind.Blank)
        {
            text = string.Empty;
        }
        else if (!value.TryGetText(out text!))
        {
            return ComputedValue.Error(Error.Value);
        }

        text = text.Trim();

        if (text.Length == 0)
        {
            return ComputedValue.Number(0);
        }

        if (text.Length > 255)
        {
            return ComputedValue.Error(Error.Value);
        }

        var negative = text[0] == '-';

        if (negative)
        {
            text = text[1..];

            if (text.Length == 0)
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        var total = 0d;

        for (var i = 0; i < text.Length; i++)
        {
            var current = RomanDigit(text[i]);

            if (current < 0)
            {
                return ComputedValue.Error(Error.Value);
            }

            var next = i + 1 < text.Length ? RomanDigit(text[i + 1]) : 0;

            // Notação subtrativa: um dígito menor antes de um maior subtrai (ex.: CM = 900).
            total += current < next ? -current : current;
        }

        return ComputedValue.Number(negative ? -total : total);
    }

    private static int RomanDigit(char character) =>
        char.ToUpperInvariant(character) switch
        {
            'I' => 1,
            'V' => 5,
            'X' => 10,
            'L' => 50,
            'C' => 100,
            'D' => 500,
            'M' => 1000,
            _ => -1,
        };
}
