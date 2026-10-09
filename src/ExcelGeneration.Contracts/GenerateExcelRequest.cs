namespace ExcelGeneration.Contracts;

/// <summary>
/// Parameters of the <c>Excel.Generate</c> JSON-RPC method. Describes a workbook with one or more sheets.
/// </summary>
public record GenerateExcelRequest
{
    /// <summary>File name of the generated workbook. Defaults to <c>export.xlsx</c>.</summary>
    public string? FileName { get; init; }

    /// <summary>The sheets (tabs) of the workbook, in order.</summary>
    public required IReadOnlyList<SheetDefinition> Sheets { get; init; }
}
