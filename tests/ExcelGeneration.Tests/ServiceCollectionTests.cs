using ExcelGeneration.Client;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelGeneration.Tests;

public sealed class ServiceCollectionTests
{
    [Fact]
    public void Registers_typed_client()
    {
        var services = new ServiceCollection();
        services.AddExcelGenerationClient(o => o.BaseAddress = new Uri("https://excel.example.com/api"));

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<IExcelGenerationClient>();

        Assert.IsType<ExcelGenerationClient>(client);
    }
}
