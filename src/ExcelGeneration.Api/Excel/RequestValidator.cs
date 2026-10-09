using ExcelGeneration.Contracts;

namespace ExcelGeneration.Api.Excel;

/// <summary>Structural validation of the request. Cell values are validated while writing.</summary>
internal static class RequestValidator
{
    public const int MaxColumns = 16_384;

    public static void Validate(GenerateExcelRequest request)
    {
        var errors = new List<PayloadError>();

        if (request.Sheets is null || request.Sheets.Count == 0)
        {
            throw new PayloadValidationException(new PayloadError("At least one sheet is required."));
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in request.Sheets)
        {
            if (sheet is null)
            {
                errors.Add(new PayloadError("Sheet definitions cannot be null."));
                continue;
            }

            var name = SheetNames.Sanitize(sheet.Name ?? string.Empty);
            if (name.Length == 0)
                errors.Add(new PayloadError("Sheet name is required.", sheet.Name));
            else if (!names.Add(name))
                errors.Add(new PayloadError($"Duplicate sheet name '{name}'.", sheet.Name));

            if (sheet.Columns is null || sheet.Columns.Count == 0)
                errors.Add(new PayloadError("At least one column is required.", sheet.Name));
            else if (sheet.Columns.Count > MaxColumns)
                errors.Add(new PayloadError($"A sheet can have at most {MaxColumns} columns.", sheet.Name));
            else
            {
                foreach (var column in sheet.Columns)
                {
                    if (string.IsNullOrWhiteSpace(column?.Name))
                        errors.Add(new PayloadError("Column name is required.", sheet.Name));
                    else if (!Enum.IsDefined(column.Type))
                        errors.Add(new PayloadError($"Unknown column type '{column.Type}'.", sheet.Name, Column: column.Name));
                    else if (column.Width is <= 0 or > 255)
                        errors.Add(new PayloadError("Column width must be between 0 and 255.", sheet.Name, Column: column.Name));
                }
            }

            if (sheet.AutoFilter && !sheet.ShowHeader)
                errors.Add(new PayloadError("AutoFilter requires ShowHeader.", sheet.Name));

            if (sheet.Header is { } header)
            {
                ValidateColor(header.FontColor, nameof(header.FontColor), sheet.Name, errors);
                ValidateColor(header.BackgroundColor, nameof(header.BackgroundColor), sheet.Name, errors);
            }
        }

        if (errors.Count > 0)
            throw new PayloadValidationException(errors);
    }

    private static void ValidateColor(string? color, string property, string? sheet, List<PayloadError> errors)
    {
        if (color is null)
            return;

        var hex = color.TrimStart('#');
        if (hex.Length is not (6 or 8) || !hex.All(char.IsAsciiHexDigit))
            errors.Add(new PayloadError($"{property} '{color}' must be an HTML hex color like #1F4E78.", sheet));
    }
}
