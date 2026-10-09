namespace ExcelGeneration.Api.Excel;

internal static class FileNames
{
    private const string Default = "export.xlsx";

    public static string Normalize(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return Default;

        var name = Path.GetFileName(fileName.Trim());
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');

        if (string.IsNullOrWhiteSpace(Path.GetFileNameWithoutExtension(name)))
            return Default;

        return name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? name : name + ".xlsx";
    }
}
