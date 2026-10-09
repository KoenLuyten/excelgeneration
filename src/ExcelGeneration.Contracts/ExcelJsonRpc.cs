namespace ExcelGeneration.Contracts;

/// <summary>Well-known JSON-RPC constants of the Excel generation service.</summary>
public static class ExcelJsonRpc
{
    /// <summary>Relative route of the JSON-RPC endpoint.</summary>
    public const string Route = "jsonrpc";

    /// <summary>JSON-RPC service name.</summary>
    public const string ServiceName = "Excel";

    /// <summary>JSON-RPC method that generates a workbook.</summary>
    public const string GenerateMethodName = "Generate";

    /// <summary>Fully qualified JSON-RPC method name: <c>Excel.Generate</c>.</summary>
    public const string GenerateMethod = ServiceName + "." + GenerateMethodName;

    /// <summary>Content type of the generated workbook.</summary>
    public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
}
