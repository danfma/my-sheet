using System.Globalization;
using System.Text;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Renders the common Excel date/time format codes (<c>TEXT</c>'s second argument), including the
/// month-vs-minute disambiguation of <c>m</c>/<c>mm</c>. Covers y, m, d, h, s, AM/PM and literals; the
/// full Excel format spec (sections, colours, fractions, locale tags) is out of scope.
///
/// The fields are substituted one token at a time from <see cref="DateSerial.TryGetCalendar"/>,
/// <see cref="DateSerial.LotusDayOfWeek"/> and <see cref="DateSerial.TimeOfDaySeconds"/> rather than by
/// translating the format to a .NET custom format string and calling <see cref="DateTime.ToString(string)"/>:
/// no <see cref="DateTime"/> can hold Excel's day zero (1900-01-<b>00</b>) or its phantom 1900-02-<b>29</b>,
/// which are exactly the two dates this formatter has to print.
/// </summary>
internal static class ExcelDateFormat
{
    public static bool IsDateOrTime(string format)
    {
        var inQuote = false;

        foreach (var c in format)
        {
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && c is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M')
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Renders <paramref name="serial"/> through <paramref name="format"/>, or returns <c>false</c> when no
    /// calendar date represents the serial (negative, or past 9999-12-31). The range VERDICT is returned
    /// instead of an <see cref="Error"/> because the callers disagree on which error it is:
    /// <c>TEXT</c> answers <c>#VALUE!</c> where the date functions answer <c>#NUM!</c>.
    /// </summary>
    public static bool TryRender(string format, double serial, out string text)
    {
        // A weekday NAME in the format pulls the day NUMBER off the collapsed DateTime map — measured on
        // Aspose.Cells 26.6.0: TEXT(60,"yyyy-mm-dd") is 1900-02-29 but TEXT(60,"yyyy-mm-dd dddd") is
        // "1900-02-28 Tuesday". Day zero is NOT affected (TEXT(0,"yyyy-mm-dd dddd") = "1900-01-00 Saturday"),
        // which is why the exception rides on phantomFeb29 alone and not on TryGetCalendar as a whole.
        var phantomFeb29 = !HasWeekdayName(format);

        if (
            DateSerial.TryGetCalendar(
                serial,
                phantomFeb29,
                out var year,
                out var month,
                out var day
            )
            is not null
        )
        {
            text = string.Empty;
            return false;
        }

        text = Render(
            format,
            new DateTimeFields(
                year,
                month,
                day,
                DateSerial.LotusDayOfWeek(serial),
                DateSerial.TimeOfDaySeconds(serial)
            )
        );

        return true;
    }

    /// <summary>Whether the format asks for a weekday name (<c>ddd</c> or longer) outside a literal.</summary>
    private static bool HasWeekdayName(string format)
    {
        var i = 0;

        while (i < format.Length)
        {
            var c = format[i];

            if (c == '"')
            {
                i++;
                while (i < format.Length && format[i] != '"')
                {
                    i++;
                }

                i++; // closing quote
                continue;
            }

            if (c == '\\')
            {
                i += 2;
                continue;
            }

            if (c is 'd' or 'D')
            {
                var run = 1;
                while (i + run < format.Length && format[i + run] is 'd' or 'D')
                {
                    run++;
                }

                if (run >= 3)
                {
                    return true;
                }

                i += run;
                continue;
            }

            i++;
        }

        return false;
    }

    private static string Render(string format, DateTimeFields fields)
    {
        var twelveHour = ContainsIgnoreCase(format, "AM/PM") || ContainsIgnoreCase(format, "A/P");
        var result = new StringBuilder();
        var i = 0;

        while (i < format.Length)
        {
            if (StartsWithIgnoreCase(format, i, "AM/PM"))
            {
                result.Append(fields.Meridiem);
                i += 5;
                continue;
            }

            if (StartsWithIgnoreCase(format, i, "A/P"))
            {
                result.Append(fields.Meridiem);
                i += 3;
                continue;
            }

            var c = format[i];

            if (c == '"')
            {
                i++;
                while (i < format.Length && format[i] != '"')
                {
                    result.Append(format[i]);
                    i++;
                }

                i++; // closing quote
                continue;
            }

            if (c == '\\' && i + 1 < format.Length)
            {
                result.Append(format[i + 1]);
                i += 2;
                continue;
            }

            if (char.IsLetter(c))
            {
                var lower = char.ToLowerInvariant(c);
                var run = 1;
                while (i + run < format.Length && char.ToLowerInvariant(format[i + run]) == lower)
                {
                    run++;
                }

                AppendField(result, new Token(format, i, lower, run), twelveHour, fields);
                i += run;
                continue;
            }

            // Separators (/ : - space) and anything else non-alphabetic pass through untouched.
            result.Append(c);
            i++;
        }

        return result.ToString();
    }

    private static void AppendField(
        StringBuilder result,
        Token token,
        bool twelveHour,
        DateTimeFields fields
    )
    {
        var run = token.Run;

        switch (token.Lower)
        {
            case 'y':
                result.Append(
                    run >= 3
                        ? fields.Year.ToString("0000", CultureInfo.InvariantCulture)
                        : (fields.Year % 100).ToString("00", CultureInfo.InvariantCulture)
                );
                break;

            case 'd' when run >= 4:
                result.Append(Invariant.GetDayName(fields.Weekday));
                break;

            case 'd' when run == 3:
                result.Append(Invariant.GetAbbreviatedDayName(fields.Weekday));
                break;

            case 'd':
                AppendNumber(result, fields.Day, run);
                break;

            case 'h':
                AppendNumber(result, twelveHour ? fields.Hour12 : fields.Hour24, run);
                break;

            case 's':
                AppendNumber(result, fields.Second, run);
                break;

            case 'm' when IsMinute(token):
                AppendNumber(result, fields.Minute, run);
                break;

            case 'm' when run >= 4:
                result.Append(Invariant.GetMonthName(fields.Month));
                break;

            case 'm' when run == 3:
                result.Append(Invariant.GetAbbreviatedMonthName(fields.Month));
                break;

            case 'm':
                AppendNumber(result, fields.Month, run);
                break;

            default:
                result.Append(token.Format, token.Index, run);
                break;
        }
    }

    private static void AppendNumber(StringBuilder result, int value, int run) =>
        result.Append(value.ToString(run >= 2 ? "00" : "0", CultureInfo.InvariantCulture));

    // 'm' means minutes when, ignoring separators, it follows an hour token or precedes a seconds token.
    private static bool IsMinute(Token token)
    {
        var format = token.Format;
        var before = token.Index - 1;

        while (before >= 0 && !char.IsLetter(format[before]))
        {
            before--;
        }

        if (before >= 0 && format[before] is 'h' or 'H')
        {
            return true;
        }

        var after = token.Index + token.Run;

        while (after < format.Length && !char.IsLetter(format[after]))
        {
            after++;
        }

        return after < format.Length && format[after] is 's' or 'S';
    }

    private static DateTimeFormatInfo Invariant => CultureInfo.InvariantCulture.DateTimeFormat;

    private static bool ContainsIgnoreCase(string text, string value) =>
        text.Contains(value, StringComparison.OrdinalIgnoreCase);

    private static bool StartsWithIgnoreCase(string text, int index, string value) =>
        index + value.Length <= text.Length
        && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase)
            == 0;

    /// <summary>One run of like letters in the format string, e.g. the <c>mmm</c> of <c>d-mmm-yy</c>.</summary>
    private readonly record struct Token(string Format, int Index, char Lower, int Run);

    /// <summary>
    /// The fields one <c>TEXT</c> call prints: Excel's calendar date (day zero and the phantom Feb 29
    /// included, which is why they are plain <c>int</c>s and not a <see cref="DateTime"/>), the Lotus weekday
    /// and the time of day as a whole second count.
    /// </summary>
    private readonly record struct DateTimeFields(
        int Year,
        int Month,
        int Day,
        DayOfWeek Weekday,
        int SecondOfDay
    )
    {
        public int Hour24 => SecondOfDay / 3600;

        public int Hour12 => Hour24 % 12 == 0 ? 12 : Hour24 % 12;

        public int Minute => SecondOfDay / 60 % 60;

        public int Second => SecondOfDay % 60;

        // Both AM/PM and A/P print the two-letter form, as the .NET "tt" this used to translate to did.
        // MEASURED divergence, recorded and NOT fixed here because it is unrelated to the epoch and no
        // Phase 9 item covers it: Aspose.Cells 26.6.0 prints A/P as one letter (TEXT(0.5,"h:mm A/P") =
        // "12:00 P" against "12:00 PM" here).
        public string Meridiem => SecondOfDay < 12 * 3600 ? "AM" : "PM";
    }
}
