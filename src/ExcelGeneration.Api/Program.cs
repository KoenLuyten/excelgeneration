using System.IO.Compression;
using ExcelGeneration.Api.Endpoints;
using ExcelGeneration.Api.Excel;
using MediatorEndpoint;
using Microsoft.AspNetCore.ResponseCompression;

var builder = WebApplication.CreateBuilder(args);

// Large payloads: Kestrel's limit is also enforced by the request decompression middleware on the
// *decompressed* body, which protects against decompression bombs.
var maxRequestBodySize = builder.Configuration.GetValue<long?>("ExcelGeneration:MaxRequestBodySizeBytes") ?? 200L * 1024 * 1024;
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = maxRequestBodySize);

// Compression: accept gzip/brotli/deflate request bodies (Content-Encoding) and compress JSON responses.
builder.Services.AddRequestDecompression();
builder.Services.AddResponseCompression(o =>
{
    o.EnableForHttps = true;
    o.Providers.Add<BrotliCompressionProvider>();
    o.Providers.Add<GzipCompressionProvider>();
});
builder.Services.Configure<BrotliCompressionProviderOptions>(o => o.Level = CompressionLevel.Fastest);
builder.Services.Configure<GzipCompressionProviderOptions>(o => o.Level = CompressionLevel.Fastest);

// MediatorEndpoint uses the minimal API JSON options (web defaults: camelCase, case-insensitive).
builder.Services.ConfigureHttpJsonOptions(_ => { });

var assemblies = new[] { typeof(Program).Assembly };
builder.Services.AddMediatR(cfg => cfg.RegisterServicesFromAssemblies(assemblies));
builder.Services.AddMediatorEndpoint(cfg =>
{
    cfg.RegisterServicesFromAssemblies(assemblies);
    cfg.RequestName = type => RequestNameAttribute.Get(type) ?? new RequestName(null, null, type.Name);
});

builder.Services.AddSingleton<WorkbookWriter>();

var app = builder.Build();

app.UseRequestDecompression();
app.UseResponseCompression();

app.MapJsonRpc();
app.MapGet("/health", () => Results.Ok("Healthy"));

app.Run();

public partial class Program;
