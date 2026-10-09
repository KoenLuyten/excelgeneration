using ExcelGeneration.Client;
using Microsoft.AspNetCore.Mvc.Testing;

namespace ExcelGeneration.Tests;

public sealed class ApiFactory : WebApplicationFactory<Program>
{
    public ExcelGenerationClient CreateExcelClient(params DelegatingHandler[] handlers)
        => new(CreateDefaultClient(new Uri("http://localhost/"), handlers));
}

/// <summary>Records the outgoing request headers.</summary>
public sealed class CapturingHandler : DelegatingHandler
{
    public List<HttpRequestMessage> Requests { get; } = [];

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        Requests.Add(request);
        return base.SendAsync(request, cancellationToken);
    }
}
