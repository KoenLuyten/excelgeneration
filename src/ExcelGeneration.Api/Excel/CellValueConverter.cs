using System.Globalization;
using ClosedXML.Excel;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Api.Excel;

/// <summary>
/// Converts a deserialized cell value (<see cref="string"/>, <see cref="long"/>, <see cref="decimal"/>,
/// <see cref="double"/>, <see cref="bool"/> or <c>null</c>) to an Excel value according to the column type.
/// </summary>
internal static class CellValueConverter
{
    private const DateTimeStyles DateStyles = DateTimeStyles.AllowWhiteSpaces;

    /// <summary>Returns <c>false</c> when the value cannot be represented as <paramref name="type"/>.</summary>
    public static bool TryConvert(object? value, ColumnType type, out XLCellValue result)
    {
        if (value is null || value is string { Length: 0 } && type != ColumnType.String)
        {
            result = Blank.Value;
            return true;
        }

        switch (type)
        {
            case ColumnType.String:
            case ColumnType.Hyperlink:
            case ColumnType.Formula:
                result = ToInvariantString(value);
                return true;

            case ColumnType.Integer:
                if (TryGetInteger(value, out var integer))
                {
                    result = integer;
                    return true;
                }
                break;

            case ColumnType.Decimal:
                if (TryGetNumber(value, out var number))
                {
                    result = number;
                    return true;
                }
                break;

            case ColumnType.Boolean:
                if (TryGetBoolean(value, out var boolean))
                {
                    result = boolean;
                    return true;
                }
                break;

            case ColumnType.DateTime:
            case ColumnType.Date:
                if (TryGetDateTime(value, out var dateTime))
                {
                    result = type == ColumnType.Date ? dateTime.Date : dateTime;
                    return true;
                }
                break;

            case ColumnType.Time:
                if (TryGetTimeSpan(value, out var timeSpan))
                {
                    result = timeSpan;
                    return true;
                }
                break;
        }

        result = Blank.Value;
        return false;
    }

    private static string ToInvariantString(object value) => value switch
    {
        string s => s,
        bool b => b ? "TRUE" : "FALSE",
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };

    private static bool TryGetInteger(object value, out double result)
    {
        switch (value)
        {
            case long l:
                result = l;
                return true;
            case decimal m when decimal.Truncate(m) == m:
                result = (double)m;
                return true;
            case double d when Math.Truncate(d) == d && double.IsFinite(d):
                result = d;
                return true;
            case string s when long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed):
                result = parsed;
                return true;
            default:
                result = 0;
                return false;
        }
    }

    private static bool TryGetNumber(object value, out double result)
    {
        switch (value)
        {
            case long l:
                result = l;
                return true;
            case decimal m:
                result = (double)m;
                return true;
            case double d when double.IsFinite(d):
                result = d;
                return true;
            case string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && double.IsFinite(parsed):
                result = parsed;
                return true;
            default:
                result = 0;
                return false;
        }
    }

    private static bool TryGetBoolean(object value, out bool result)
    {
        switch (value)
        {
            case bool b:
                result = b;
                return true;
            case long l and (0 or 1):
                result = l == 1;
                return true;
            case string s when bool.TryParse(s, out var parsed):
                result = parsed;
                return true;
            case string s when s is "0" or "1":
                result = s == "1";
                return true;
            default:
                result = false;
                return false;
        }
    }

    private static bool TryGetDateTime(object value, out DateTime result)
    {
        switch (value)
        {
            case string s:
                // Keep the clock time as written; Excel has no notion of time zones.
                if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateStyles, out var dto))
                {
                    result = dto.DateTime;
                    return true;
                }
                break;
            case long or decimal or double:
                // OLE automation date (Excel serial number).
                var oa = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (oa is > -657435.0 and < 2958466.0)
                {
                    result = DateTime.FromOADate(oa);
                    return true;
                }
                break;
        }

        result = default;
        return false;
    }

    private static bool TryGetTimeSpan(object value, out TimeSpan result)
    {
        switch (value)
        {
            case string s:
                if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out result))
                    return true;
                if (DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateStyles, out var dto))
                {
                    result = dto.TimeOfDay;
                    return true;
                }
                break;
            case long or decimal or double:
                // Fraction of a day, as Excel stores times.
                var days = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsFinite(days) && Math.Abs(days) < TimeSpan.MaxValue.TotalDays)
                {
                    result = TimeSpan.FromDays(days);
                    return true;
                }
                break;
        }

        result = default;
        return false;
    }
}
