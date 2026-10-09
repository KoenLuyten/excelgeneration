using System.Net;
using ExcelGeneration.Client;
using Microsoft.Extensions.Options;

namespace Microsoft.Extensions.DependencyInjection;

public static class ExcelGenerationServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IExcelGenerationClient"/> as a typed HttpClient. Responses are decompressed
    /// automatically (gzip/brotli); requests are gzip-compressed by the client.
    /// </summary>
    public static IHttpClientBuilder AddExcelGenerationClient(this IServiceCollection services, Action<ExcelGenerationClientOptions> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        services.AddOptions<ExcelGenerationClientOptions>()
            .Configure(configure)
            .Validate(o => o.BaseAddress is { IsAbsoluteUri: true }, "ExcelGenerationClientOptions.BaseAddress must be an absolute URI.");

        return services
            .AddHttpClient<IExcelGenerationClient, ExcelGenerationClient>((sp, http) =>
            {
                var options = sp.GetRequiredService<IOptions<ExcelGenerationClientOptions>>().Value;
                http.BaseAddress = EnsureTrailingSlash(options.BaseAddress!);
                http.Timeout = options.Timeout;
            })
            .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Brotli
            });
    }

    private static Uri EnsureTrailingSlash(Uri uri)
        => uri.AbsoluteUri.EndsWith('/') ? uri : new Uri(uri.AbsoluteUri + "/");
}
