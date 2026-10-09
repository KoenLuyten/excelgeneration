using System.Data;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using ExcelGeneration.Client;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Tests;

public sealed class EndToEndTests(ApiFactory factory) : IClassFixture<ApiFactory>
{
    public enum Status { Open, Closed }

    public sealed record Order(
        int Id,
        string Customer,
        decimal Amount,
        bool Paid,
        DateTime CreatedAt,
        DateOnly DueDate,
        TimeOnly Slot,
        Status Status,
        Uri? Link,
        double? Discount);

    private static readonly Order[] Orders =
    [
        new(1, "Contoso", 1234.5m, true, new DateTime(2026, 10, 1, 13, 45, 30), new DateOnly(2026, 11, 1), new TimeOnly(9, 30), Status.Open, new Uri("https://example.com/1"), 0.1),
        new(2, "Fabrikam", 99.99m, false, new DateTime(2026, 10, 2, 8, 0, 0), new DateOnly(2026, 11, 2), new TimeOnly(14, 15, 5), Status.Closed, null, null),
    ];

    [Fact]
    public async Task Generates_workbook_with_multiple_sheets_types_header_and_autofilter()
    {
        var table = new DataTable("Customers");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Revenue", typeof(decimal)).Caption = "Total revenue";
        table.Rows.Add("Contoso", 1000m);
        table.Rows.Add("Fabrikam", DBNull.Value);

        var workbook = new WorkbookBuilder("orders")
            .AddSheet("Orders", Orders, s =>
            {
                s.AutoFilter = true;
                s.Header = new HeaderStyle { BackgroundColor = "#1F4E78", FontColor = "#FFFFFF", BottomBorder = true };
                s.Column(o => o.Amount, c => c.Format = "#,##0.00");
                s.Column(o => o.Customer, c => { c.Name = "Customer name"; c.Width = 30; });
            })
            .AddSheet(table, s => s.AutoFilter = false);

        var file = await factory.CreateExcelClient().GenerateAsync(workbook);

        Assert.Equal("orders.xlsx", file.FileName);
        using var xl = new XLWorkbook(file.OpenRead());
        Assert.Equal(["Orders", "Customers"], xl.Worksheets.Select(w => w.Name));

        var orders = xl.Worksheet("Orders");
        Assert.Equal("Id", orders.Cell("A1").GetText());
        Assert.Equal("Customer name", orders.Cell("B1").GetText());
        Assert.True(orders.Cell("A1").Style.Font.Bold);
        Assert.Equal(XLColor.FromHtml("#1F4E78"), orders.Cell("A1").Style.Fill.BackgroundColor);
        Assert.Equal(1, orders.SheetView.SplitRow);
        Assert.Equal("A1:J3", orders.AutoFilter.Range.RangeAddress.ToString());
        Assert.Equal(30, orders.Column(2).Width);

        Assert.Equal(1, orders.Cell("A2").GetDouble());
        Assert.Equal("Contoso", orders.Cell("B2").GetText());
        Assert.Equal(1234.5, orders.Cell("C2").GetDouble());
        Assert.Equal("#,##0.00", orders.Cell("C2").Style.NumberFormat.Format);
        Assert.True(orders.Cell("D2").GetBoolean());
        Assert.Equal(new DateTime(2026, 10, 1, 13, 45, 30), orders.Cell("E2").GetDateTime());
        Assert.Equal(new DateTime(2026, 11, 1), orders.Cell("F2").GetDateTime());
        Assert.Equal("yyyy-mm-dd", orders.Cell("F2").Style.NumberFormat.Format);
        Assert.Equal(new TimeSpan(9, 30, 0), orders.Cell("G2").GetTimeSpan());
        Assert.Equal("Open", orders.Cell("H2").GetText());
        Assert.Equal("https://example.com/1", orders.Cell("I2").GetText());
        Assert.True(orders.Cell("I2").HasHyperlink);
        Assert.Equal(0.1, orders.Cell("J2").GetDouble());
        Assert.True(orders.Cell("I3").IsEmpty());
        Assert.True(orders.Cell("J3").IsEmpty());

        var customers = xl.Worksheet("Customers");
        Assert.Equal("Total revenue", customers.Cell("B1").GetText());
        Assert.Equal(1000, customers.Cell("B2").GetDouble());
        Assert.True(customers.Cell("B3").IsEmpty());
        Assert.False(customers.AutoFilter.IsEnabled);
    }

    [Fact]
    public async Task Request_body_is_gzip_compressed()
    {
        var capture = new CapturingHandler();
        var client = factory.CreateExcelClient(capture);

        var file = await client.GenerateAsync(new WorkbookBuilder().AddSheet("Orders", Orders));

        var request = Assert.Single(capture.Requests);
        Assert.Contains("gzip", request.Content!.Headers.ContentEncoding);
        Assert.Equal("export.xlsx", file.FileName);
        using var xl = new XLWorkbook(file.OpenRead());
        Assert.Equal(3, xl.Worksheet(1).LastRowUsed()!.RowNumber());
    }

    [Fact]
    public async Task Raw_json_payload_without_compression_is_accepted()
    {
        const string payload = """
            { "id":"1", "jsonrpc":"2.0", "method":"Excel.Generate", "params": {
              "fileName":"report.xlsx",
              "sheets":[{
                "name":"Orders",
                "columns":[{"name":"Id","type":"Integer"},{"name":"Date","type":"Date","format":"dd/mm/yyyy"},
                           {"name":"Amount","type":"Decimal","format":"#,##0.00","width":14},
                           {"name":"Total","type":"Formula"}],
                "header":{"bold":true,"fontColor":"#FFFFFF","backgroundColor":"#1F4E78","freeze":true},
                "autoFilter":true,
                "rows":[[1,"2026-10-01",12.5,"=C2*2"],[2,"2026-10-02",null,"C3*2"]]
            }]}}
            """;

        using var http = factory.CreateClient();
        using var response = await http.PostAsync("/jsonrpc", new StringContent(payload, Encoding.UTF8, "application/json"));

        Assert.Equal(ExcelJsonRpc.XlsxContentType, response.Content.Headers.ContentType?.MediaType);
        Assert.Equal("report.xlsx", response.Content.Headers.ContentDisposition?.FileNameStar ?? response.Content.Headers.ContentDisposition?.FileName);

        using var xl = new XLWorkbook(await response.Content.ReadAsStreamAsync());
        var sheet = xl.Worksheet("Orders");
        Assert.Equal(new DateTime(2026, 10, 1), sheet.Cell("B2").GetDateTime());
        Assert.Equal("dd/mm/yyyy", sheet.Cell("B2").Style.NumberFormat.Format);
        Assert.Equal("C2*2", sheet.Cell("D2").FormulaA1);
        Assert.Equal(25, sheet.Cell("D2").Value.GetNumber());
        Assert.Equal(14, sheet.Column(3).Width);
    }

    [Fact]
    public async Task Invalid_cell_value_returns_invalid_params_with_location()
    {
        var request = new GenerateExcelRequest
        {
            Sheets =
            [
                new SheetDefinition
                {
                    Name = "Data",
                    Columns = [new ColumnDefinition("Amount", ColumnType.Decimal), new ColumnDefinition("When", ColumnType.Date)],
                    Rows = [new ExcelRow(1.5, "2026-01-01"), new ExcelRow("abc", "not a date")]
                }
            ]
        };

        var exc = await Assert.ThrowsAsync<ExcelGenerationException>(() => factory.CreateExcelClient().GenerateAsync(request));

        Assert.Equal(ExcelGenerationException.InvalidParams, exc.Code);
        var errors = exc.ErrorData!.Value.EnumerateArray().ToArray();
        Assert.Equal(2, errors.Length);
        Assert.Equal("Data", errors[0].GetProperty("sheet").GetString());
        Assert.Equal(2, errors[0].GetProperty("row").GetInt32());
        Assert.Equal("Amount", errors[0].GetProperty("column").GetString());
        Assert.Equal("When", errors[1].GetProperty("column").GetString());
    }

    [Fact]
    public async Task Structural_errors_are_reported()
    {
        var request = new GenerateExcelRequest
        {
            Sheets =
            [
                new SheetDefinition { Name = "Same", Columns = [new ColumnDefinition("A")], Rows = [] },
                new SheetDefinition { Name = "same", Columns = [new ColumnDefinition("A")], Rows = [], ShowHeader = false, AutoFilter = true }
            ]
        };

        var exc = await Assert.ThrowsAsync<ExcelGenerationException>(() => factory.CreateExcelClient().GenerateAsync(request));

        Assert.Equal(ExcelGenerationException.InvalidParams, exc.Code);
        Assert.Equal(2, exc.ErrorData!.Value.GetArrayLength());
    }

    [Fact]
    public async Task Unknown_method_returns_method_not_found()
    {
        using var http = factory.CreateClient();
        using var response = await http.PostAsJsonAsync("/jsonrpc", new { id = "1", jsonrpc = "2.0", method = "Excel.Nope", @params = new { } });

        var json = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(-32601, json.GetProperty("error").GetProperty("code").GetInt32());
    }

    [Fact]
    public async Task Unknown_column_type_returns_error()
    {
        const string payload = """
            {"id":"1","jsonrpc":"2.0","method":"Excel.Generate","params":{"sheets":[{"name":"A","columns":[{"name":"X","type":"Money"}],"rows":[]}]}}
            """;

        using var http = factory.CreateClient();
        using var response = await http.PostAsync("/jsonrpc", new StringContent(payload, Encoding.UTF8, "application/json"));

        var json = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal(-32602, json.GetProperty("error").GetProperty("code").GetInt32());
    }

    public sealed record Measurement(int Index, string Label, double Value, DateTime At, bool Flag);

    [Fact]
    public async Task Large_payload_is_streamed_and_generated()
    {
        const int rows = 200_000;
        var start = new DateTime(2026, 1, 1);
        var data = Enumerable.Range(1, rows).Select(i => new Measurement(i, $"Label {i}", i * 1.5, start.AddMinutes(i), i % 2 == 0));

        var file = await factory.CreateExcelClient().GenerateAsync(
            new WorkbookBuilder("large.xlsx").AddSheet("Data", data, s => s.AutoFilter = true));

        using var xl = new XLWorkbook(file.OpenRead());
        var sheet = xl.Worksheet("Data");
        Assert.Equal(rows + 1, sheet.LastRowUsed()!.RowNumber());
        Assert.Equal(rows, sheet.Cell(rows + 1, 1).GetDouble());
    }
}
