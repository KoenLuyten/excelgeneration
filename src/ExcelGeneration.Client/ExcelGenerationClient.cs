using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelGeneration.Client.Http;
using ExcelGeneration.Contracts;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace ExcelGeneration.Client;

/// <summary>
/// Calls <c>Excel.Generate</c> over JSON-RPC. The request body is gzip-compressed; the workbook is returned as a
/// binary file, while errors are returned as JSON-RPC error envelopes.
/// </summary>
public sealed class ExcelGenerationClient : IExcelGenerationClient
{
    internal static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly Uri Route = new(ExcelJsonRpc.Route, UriKind.Relative);

    private readonly HttpClient _httpClient;
    private readonly ExcelGenerationClientOptions _options;

    /// <param name="httpClient">An HttpClient whose <see cref="HttpClient.BaseAddress"/> points to the service.</param>
    public ExcelGenerationClient(HttpClient httpClient) : this(httpClient, new ExcelGenerationClientOptions())
    {
    }

    [ActivatorUtilitiesConstructor]
    public ExcelGenerationClient(HttpClient httpClient, IOptions<ExcelGenerationClientOptions> options) : this(httpClient, options.Value)
    {
    }

    public ExcelGenerationClient(HttpClient httpClient, ExcelGenerationClientOptions options)
    {
        _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    public Task<ExcelFile> GenerateAsync(WorkbookBuilder workbook, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        return GenerateAsync(workbook.Build(), cancellationToken);
    }

    public async Task<ExcelFile> GenerateAsync(GenerateExcelRequest request, CancellationToken cancellationToken = default)
    {
        using var buffer = new MemoryStream();
        var fileName = await GenerateToStreamAsync(request, buffer, cancellationToken).ConfigureAwait(false);
        return new ExcelFile(fileName, buffer.ToArray());
    }

    public async Task<string> GenerateToStreamAsync(GenerateExcelRequest request, Stream destination, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(destination);

        var envelope = new JsonRpcRequest<GenerateExcelRequest>(Guid.NewGuid().ToString(), ExcelJsonRpc.GenerateMethod, request);
        using var message = new HttpRequestMessage(HttpMethod.Post, Route)
        {
            Content = new GzipJsonContent<JsonRpcRequest<GenerateExcelRequest>>(envelope, SerializerOptions, _options.CompressionLevel)
        };
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(ExcelJsonRpc.XlsxContentType));
        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var response = await _httpClient
            .SendAsync(message, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
            .ConfigureAwait(false);

        var mediaType = response.Content.Headers.ContentType?.MediaType;
        if (string.Equals(mediaType, "application/json", StringComparison.OrdinalIgnoreCase))
        {
            throw await ReadErrorAsync(response, cancellationToken).ConfigureAwait(false);
        }

        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            throw new HttpRequestException($"Excel generation failed with HTTP {(int)response.StatusCode}: {body}", null, response.StatusCode);
        }

        await response.Content.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
        return GetFileName(response) ?? request.FileName ?? "export.xlsx";
    }

    private static async Task<Exception> ReadErrorAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        await using (stream.ConfigureAwait(false))
        {
            JsonRpcErrorEnvelope? envelope;
            try
            {
                envelope = await JsonSerializer.DeserializeAsync<JsonRpcErrorEnvelope>(stream, SerializerOptions, cancellationToken).ConfigureAwait(false);
            }
            catch (JsonException exc)
            {
                return new HttpRequestException($"The service returned an unreadable JSON response (HTTP {(int)response.StatusCode}).", exc, response.StatusCode);
            }

            if (envelope?.Error is not { } error)
            {
                return new HttpRequestException($"The service returned JSON instead of a workbook (HTTP {(int)response.StatusCode}).", null, response.StatusCode);
            }

            return new ExcelGenerationException(error.Code, error.Message ?? "Unknown error", error.Data);
        }
    }

    private static string? GetFileName(HttpResponseMessage response)
    {
        var disposition = response.Content.Headers.ContentDisposition;
        var name = disposition?.FileNameStar ?? disposition?.FileName;
        return string.IsNullOrWhiteSpace(name) ? null : name.Trim('"');
    }
}
