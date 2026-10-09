using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelGeneration.Client.Http;

internal sealed record JsonRpcRequest<TParams>(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("method")] string Method,
    [property: JsonPropertyName("params")] TParams Params)
{
    [JsonPropertyName("jsonrpc")]
    [JsonPropertyOrder(-1)]
    public string JsonRpc => "2.0";
}

internal sealed record JsonRpcErrorEnvelope(
    [property: JsonPropertyName("id")] string? Id,
    [property: JsonPropertyName("error")] JsonRpcError? Error);

internal sealed record JsonRpcError(
    [property: JsonPropertyName("code")] int Code,
    [property: JsonPropertyName("message")] string? Message,
    [property: JsonPropertyName("data")] JsonElement? Data);
