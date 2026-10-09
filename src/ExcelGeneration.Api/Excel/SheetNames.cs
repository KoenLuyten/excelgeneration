namespace ExcelGeneration.Api.Excel;

internal static class SheetNames
{
    public const int MaxLength = 31;
    private static readonly char[] InvalidChars = ['[', ']', ':', '*', '?', '/', '\\'];

    /// <summary>Replaces characters Excel does not allow and truncates to 31 characters.</summary>
    public static string Sanitize(string name)
    {
        var chars = name.Trim().ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            if (Array.IndexOf(InvalidChars, chars[i]) >= 0 || char.IsControl(chars[i]))
                chars[i] = '_';
        }

        var sanitized = new string(chars).Trim('\'');
        return sanitized.Length > MaxLength ? sanitized[..MaxLength] : sanitized;
    }
}
