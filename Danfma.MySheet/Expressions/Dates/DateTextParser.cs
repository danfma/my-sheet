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

    public static bool TryParseDate(string text, out DateTime date) =>
        DateTime.TryParseExact(
            text.Trim(),
            DateFormats,
            CultureInfo.InvariantCulture,
            Styles,
            out date
        );

    public static bool TryParseTime(string text, out double fraction)
    {
        if (
            DateTime.TryParseExact(
                text.Trim(),
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
}
