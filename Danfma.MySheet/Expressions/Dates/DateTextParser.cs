using System.Globalization;

namespace Danfma.MySheet.Expressions.Dates;

/// <summary>
/// Invariant-culture text parsing for <c>DATEVALUE</c> and <c>TIMEVALUE</c>. The engine is locale-invariant
/// by design (§A7): a fixed set of documented formats is accepted, no current-locale short-date parsing and
/// no year-less dates (which would depend on the clock — a volatile concern deferred to F1). Anything that
/// matches none of the formats is a <c>#VALUE!</c> at the call site.
/// </summary>
internal static class DateTextParser
{
    private const DateTimeStyles Styles =
        DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;

    // Date formats (ISO with '-' or '/', US M/d/yyyy, and the d-MMM-yyyy long form), plus a few date+time
    // combinations so a string carrying a time still yields its date. Month names parse case-insensitively.
    private static readonly string[] DateFormats =
    [
        "yyyy-MM-dd",
        "yyyy/MM/dd",
        "M/d/yyyy",
        "MM/dd/yyyy",
        "d-MMM-yyyy",
        "d-MMM-yy",
        "MMM d, yyyy",
        "MMMM d, yyyy",
        "M/d/yyyy h:mm tt",
        "M/d/yyyy H:mm",
        "yyyy-MM-dd H:mm",
        "yyyy-MM-dd h:mm tt",
        "d-MMM-yyyy h:mm tt",
        "d-MMM-yyyy H:mm",
    ];

    // Time formats: 24-hour HH:mm(:ss) and 12-hour h:mm(:ss) AM/PM, standalone or preceded by a date (whose
    // day part is then discarded, matching Excel: TIMEVALUE keeps only the time fraction).
    private static readonly string[] TimeFormats =
    [
        "H:mm",
        "H:mm:ss",
        "h:mm tt",
        "h:mm:ss tt",
        "M/d/yyyy h:mm tt",
        "M/d/yyyy h:mm:ss tt",
        "M/d/yyyy H:mm",
        "M/d/yyyy H:mm:ss",
        "yyyy-MM-dd H:mm",
        "yyyy-MM-dd H:mm:ss",
        "yyyy-MM-dd h:mm tt",
        "d-MMM-yyyy h:mm tt",
        "d-MMM-yyyy h:mm:ss tt",
        "d-MMM-yyyy H:mm",
        "d-MMM-yyyy H:mm:ss",
    ];

    // A REAL leap year, four digits wide like 1900, so substituting it leaves every format's field widths
    // untouched.
    private const int PhantomSubstituteYear = 2000;

    public static bool TryParseDate(string text, out DateTime date) =>
        DateTime.TryParseExact(
            text.Trim(),
            DateFormats,
            CultureInfo.InvariantCulture,
            Styles,
            out date
        );

    /// <summary>
    /// Whether the text spells Excel's phantom 1900-02-29 (the caller answers
    /// <see cref="DateSerial.PhantomFeb29Serial"/>). <c>DATEVALUE</c> is the only path that reaches that
    /// serial, and it accepts every spelling the ordinary formats do — measured on Aspose.Cells 26.6.0:
    /// <c>1900-02-29</c>, <c>1900/02/29</c>, <c>2/29/1900</c>, <c>29-Feb-1900</c>, <c>Feb 29, 1900</c>,
    /// <c>February 29, 1900</c> and any of them with a trailing time all give 60.
    /// </summary>
    public static bool TryParsePhantomFeb29(string text) =>
        TrySwapPhantomYear(text.Trim(), out var substituted)
        && TryParseDate(substituted, out var date)
        && date is { Year: PhantomSubstituteYear, Month: 2, Day: 29 };

    public static bool TryParseTime(string text, out double fraction)
    {
        var trimmed = text.Trim();

        if (TryParseTimeExact(trimmed, out fraction))
        {
            return true;
        }

        // The phantom day carries a time of day like any other date: TIMEVALUE("1900-02-29 12:00") is 0.5
        // (measured), even though no proleptic-Gregorian parse accepts that calendar day.
        return TrySwapPhantomYear(trimmed, out var substituted)
            && TryParseTimeExact(substituted, out fraction);
    }

    private static bool TryParseTimeExact(string text, out double fraction)
    {
        if (
            DateTime.TryParseExact(
                text,
                TimeFormats,
                CultureInfo.InvariantCulture,
                Styles,
                out var parsed
            )
        )
        {
            fraction = parsed.TimeOfDay.TotalDays;
            return true;
        }

        fraction = 0d;
        return false;
    }

    // February 1900 has 29 days in Excel and 28 in every calendar .NET knows, so TryParseExact rejects the
    // phantom day whatever the spelling. Swapping the year for a real leap year lets the SAME format tables
    // recognise it, which is why no second table is needed; the caller then checks the parsed date really is
    // the 29th of February, so an ordinary 1900 date (or a literal 2000 one) is not mistaken for the phantom.
    private static bool TrySwapPhantomYear(string trimmed, out string substituted)
    {
        const string phantomYear = "1900";

        if (!trimmed.Contains(phantomYear, StringComparison.Ordinal))
        {
            substituted = trimmed;
            return false;
        }

        substituted = trimmed.Replace(
            phantomYear,
            PhantomSubstituteYear.ToString(CultureInfo.InvariantCulture),
            StringComparison.Ordinal
        );

        return true;
    }
}
