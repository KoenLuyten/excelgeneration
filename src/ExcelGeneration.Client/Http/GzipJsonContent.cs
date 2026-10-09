using System.IO.Compression;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace ExcelGeneration.Client.Http;

/// <summary>
/// HTTP content that serializes a value as JSON straight into a gzip stream. Nothing is buffered as a whole:
/// rows are pulled from their <see cref="IEnumerable{T}"/> while the request body is written (chunked transfer).
/// </summary>
internal sealed class GzipJsonContent<T> : HttpContent
{
    private readonly T _value;
    private readonly JsonSerializerOptions _options;
    private readonly CompressionLevel _level;

    public GzipJsonContent(T value, JsonSerializerOptions options, CompressionLevel level = CompressionLevel.Fastest)
    {
        _value = value;
        _options = options;
        _level = level;
        Headers.ContentType = new MediaTypeHeaderValue("application/json") { CharSet = "utf-8" };
        Headers.ContentEncoding.Add("gzip");
    }

    protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context)
        => SerializeToStreamAsync(stream, context, CancellationToken.None);

    protected override async Task SerializeToStreamAsync(Stream stream, TransportContext? context, CancellationToken cancellationToken)
    {
        await using var gzip = new GZipStream(stream, _level, leaveOpen: true);
        await JsonSerializer.SerializeAsync(gzip, _value, _options, cancellationToken).ConfigureAwait(false);
    }

    protected override bool TryComputeLength(out long length)
    {
        length = -1;
        return false;
    }
}
