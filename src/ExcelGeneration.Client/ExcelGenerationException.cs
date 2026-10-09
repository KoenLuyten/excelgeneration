using System.Text.Json;

namespace ExcelGeneration.Client;

/// <summary>The service returned a JSON-RPC error.</summary>
public sealed class ExcelGenerationException(int code, string message, JsonElement? data = null) : Exception(message)
{
    public const int InvalidRequest = -32600;
    public const int MethodNotFound = -32601;
    public const int InvalidParams = -32602;
    public const int InternalError = -32603;
    public const int ParseError = -32700;

    /// <summary>JSON-RPC error code.</summary>
    public int Code { get; } = code;

    /// <summary>Additional error information. For <see cref="InvalidParams"/> this is a list of
    /// <c>{ message, sheet, row, column }</c> objects.</summary>
    public JsonElement? ErrorData { get; } = data;
}
