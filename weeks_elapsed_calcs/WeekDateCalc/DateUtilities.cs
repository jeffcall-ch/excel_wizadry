using System.Globalization;

namespace WeekDateCalc;

public static class DateUtilities
{
    private static readonly string[] SupportedFormats = ["dd/MM/yyyy", "dd.MM.yyyy"];

    public static bool TryParseDate(string? value, out DateTime date)
    {
        date = default;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return DateTime.TryParseExact(
            value.Trim(),
            SupportedFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out date);
    }

    public static double WeeksBetween(DateTime a, DateTime b)
    {
        var days = Math.Abs((b.Date - a.Date).TotalDays);
        return Math.Round(days / 7.0, 1, MidpointRounding.AwayFromZero);
    }

    public static DateTime AddWeeksWithWorkdayRule(DateTime startDate, int weeks)
    {
        var result = startDate.Date.AddDays(weeks * 7);

        if (result.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday))
        {
            return result;
        }

        if (weeks < 0)
        {
            return result.DayOfWeek == DayOfWeek.Saturday
                ? result.AddDays(-1)
                : result.AddDays(-2);
        }

        return result.DayOfWeek == DayOfWeek.Saturday
            ? result.AddDays(2)
            : result.AddDays(1);
    }
}
