using System.Globalization;
using System.Xml;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Number-to-XML formatting shared by the write paths. Small non-negative integers — the dominant
/// numeric shape in business sheets (counts, codes, shared-string indexes) — come from a cached
/// table; everything else formats with invariant shortest-round-trip semantics, byte-identical to
/// <c>ToString(CultureInfo.InvariantCulture)</c>. The <see cref="XmlWriter"/> overloads write
/// through a reused per-thread buffer (<c>TryFormat</c> + <c>WriteChars</c>), so a non-cached
/// number costs no string at all.
/// </summary>
internal static class XlsxNumbers
{
    private static readonly string[] SmallIntegers = CreateSmallIntegers();

    [ThreadStatic]
    private static char[]? _buffer;

    private static string[] CreateSmallIntegers()
    {
        var values = new string[1024];

        for (var i = 0; i < values.Length; i++)
        {
            values[i] = i.ToString(CultureInfo.InvariantCulture);
        }

        return values;
    }

    public static string Format(double value)
    {
        if (value >= 0 && value < SmallIntegers.Length)
        {
            var integer = (int)value;

            if (integer == value)
            {
                return SmallIntegers[integer];
            }
        }

        return value.ToString(CultureInfo.InvariantCulture);
    }

    public static string Format(int value) =>
        value >= 0 && value < SmallIntegers.Length
            ? SmallIntegers[value]
            : value.ToString(CultureInfo.InvariantCulture);

    /// <summary>Writes the number as element/attribute content without allocating a string.</summary>
    public static void Write(XmlWriter writer, double value)
    {
        if (value >= 0 && value < SmallIntegers.Length)
        {
            var integer = (int)value;

            if (integer == value)
            {
                writer.WriteString(SmallIntegers[integer]);

                return;
            }
        }

        var buffer = _buffer ??= new char[32];

        // Digits, sign, '.', 'E' — nothing XML needs escaping, so WriteChars is safe and exact.
        value.TryFormat(buffer, out var written, default, CultureInfo.InvariantCulture);
        writer.WriteChars(buffer, 0, written);
    }

    /// <summary>Writes the integer as element/attribute content without allocating a string.</summary>
    public static void Write(XmlWriter writer, int value)
    {
        if (value >= 0 && value < SmallIntegers.Length)
        {
            writer.WriteString(SmallIntegers[value]);

            return;
        }

        var buffer = _buffer ??= new char[32];

        value.TryFormat(buffer, out var written, default, CultureInfo.InvariantCulture);
        writer.WriteChars(buffer, 0, written);
    }
}
