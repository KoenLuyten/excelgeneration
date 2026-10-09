using System.IO.Compression;

namespace ExcelGeneration.Client;

public sealed class ExcelGenerationClientOptions
{
    /// <summary>Base address of the service, e.g. <c>https://excel.example.com/</c>.</summary>
    public Uri? BaseAddress { get; set; }

    /// <summary>HTTP timeout. Generating large workbooks can take a while. Defaults to 5 minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);

    /// <summary>Gzip level used for the request body. Defaults to <see cref="CompressionLevel.Fastest"/>.</summary>
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Fastest;
}
