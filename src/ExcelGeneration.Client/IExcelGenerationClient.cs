using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client;

/// <summary>Client for the Excel generation JSON-RPC service.</summary>
public interface IExcelGenerationClient
{
    /// <summary>Generates a workbook and returns it in memory.</summary>
    /// <exception cref="ExcelGenerationException">The service returned a JSON-RPC error.</exception>
    Task<ExcelFile> GenerateAsync(GenerateExcelRequest request, CancellationToken cancellationToken = default);

    /// <summary>Generates a workbook described by a <see cref="WorkbookBuilder"/>.</summary>
    Task<ExcelFile> GenerateAsync(WorkbookBuilder workbook, CancellationToken cancellationToken = default);

    /// <summary>Generates a workbook and copies it to <paramref name="destination"/> without buffering it in memory.</summary>
    /// <returns>The file name suggested by the service.</returns>
    Task<string> GenerateToStreamAsync(GenerateExcelRequest request, Stream destination, CancellationToken cancellationToken = default);
}
