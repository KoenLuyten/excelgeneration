# ExcelGeneration.Client

Client for the Excel generation JSON-RPC service (`Excel.Generate`). It handles these for you:

- the JSON-RPC envelope;
- **gzip compression** of the request body. The body is streamed, so rows are serialized while they are sent;
- decompression of responses;
- turning `IEnumerable<T>` and `DataTable`/`DataSet` into the positional row payload;
- turning JSON-RPC errors into `ExcelGenerationException`.

Targets net8.0 and net10.0, and brings in `ExcelGeneration.Contracts`.

```bash
dotnet add package ExcelGeneration.Client
```

## Registration

```csharp
using System.IO.Compression;
using ExcelGeneration.Client;

builder.Services
    .AddExcelGenerationClient(o =>
    {
        o.BaseAddress = new Uri("https://excel.example.com/");   // must be absolute
        o.Timeout = TimeSpan.FromMinutes(10);                     // default 5 minutes
        o.CompressionLevel = CompressionLevel.Optimal;            // default Fastest
    });
// AddExcelGenerationClient returns an IHttpClientBuilder, so you can chain
// handlers onto it, e.g. .AddHttpMessageHandler<AuthHandler>().
```

Then inject `IExcelGenerationClient`. Without DI, pass an `HttpClient` that has `BaseAddress` set:

```csharp
var client = new ExcelGenerationClient(new HttpClient { BaseAddress = new Uri("https://excel.example.com/") });
```

## Usage

### From a collection

Each public property becomes a column, in declaration order. Use attributes to control the columns:

```csharp
public record Order(
    [property: ExcelColumn("Order #", Order = 1)] int Id,
    DateOnly Date,
    string Customer,
    [property: ExcelColumn(Format = "#,##0.00", Width = 14)] decimal Amount,
    [property: ExcelIgnore] string InternalNote);

var workbook = new WorkbookBuilder("orders.xlsx")
    .AddSheet("Orders", orders, s =>
    {
        s.AutoFilter = true;
        s.Header = new HeaderStyle { BackgroundColor = "#1F4E78", FontColor = "#FFFFFF" };
        s.Column(o => o.Date, c => c.Format = "dd/mm/yyyy");   // fluent options override attributes
        s.Column(o => o.Customer, c => c.Name = "Customer name");
    });

ExcelFile file = await client.GenerateAsync(workbook, cancellationToken);
await file.SaveAsAsync("orders.xlsx", cancellationToken);
```

### From a `DataTable` or `DataSet`

Columns are configured by name:

```csharp
var workbook = new WorkbookBuilder("report.xlsx")
    .AddSheet(customersTable)                                   // sheet name = TableName
    .AddSheet("Invoices", invoicesTable, s =>
    {
        s.Column("Amount", c => c.Format = "#,##0.00");
        s.Ignore("RowVersion");
    });

// One sheet per table:
var all = new WorkbookBuilder("export.xlsx")
    .AddSheets(dataSet, (table, s) => s.AutoFilter = true);
```

### Hand-written sheets (formulas, hyperlinks)

When mapping doesn't fit, add a sheet built directly against the contract. You can mix it with mapped sheets:

```csharp
using ExcelGeneration.Contracts;

var summary = new SheetDefinition
{
    Name = "Summary",
    Columns =
    [
        new ColumnDefinition("Label"),
        new ColumnDefinition("Value", ColumnType.Formula, "#,##0.00"),
        new ColumnDefinition("Link", ColumnType.Hyperlink),
    ],
    Rows =
    [
        new ExcelRow("Total", "SUM(Orders!D:D)", "https://example.com/orders"),
        new ExcelRow("Count", "COUNTA(Orders!A:A)-1", null),
    ],
};

var workbook = new WorkbookBuilder("orders.xlsx")
    .AddSheet("Orders", orders)
    .AddSheet(summary);
```

### Streaming to an HTTP response

`GenerateAsync` holds the whole workbook in memory. For large files in an ASP.NET Core endpoint, use `GenerateToStreamAsync` to copy the service response straight to the caller:

```csharp
app.MapGet("/orders/export", async (IExcelGenerationClient excel, OrderRepository repo, HttpContext ctx, CancellationToken ct) =>
{
    var request = new WorkbookBuilder("orders.xlsx")
        .AddSheet("Orders", repo.StreamAll())   // the IEnumerable<T> is enumerated while the request is sent
        .Build();

    ctx.Response.ContentType = ExcelJsonRpc.XlsxContentType;
    ctx.Response.Headers.ContentDisposition = "attachment; filename=orders.xlsx";
    await excel.GenerateToStreamAsync(request, ctx.Response.Body, ct);
});
```

For smaller files, the in-memory result plugs straight into `Results.File`:

```csharp
var file = await excel.GenerateAsync(workbook, ct);
return Results.File(file.Content, file.ContentType, file.FileName);
```

## Column mapping

- **Properties:** public readable properties, in declaration order.
- **Attributes:** use `[ExcelColumn(Name, Order, Format, Width, Type)]`, `[ExcelIgnore]` or `[Display(Name, Order)]`.
- **Fluent options:** `SheetOptions.Column(...)` and `Ignore(...)` override the attributes.
- **DataTable:** the column name is the `DataColumn.Caption`, and `DBNull` becomes an empty cell.

Column types are inferred from the CLR type:

| CLR type | Column type |
|---|---|
| integer types | `Integer` |
| `float`, `double`, `decimal` | `Decimal` |
| `bool` | `Boolean` |
| `DateTime`, `DateTimeOffset` | `DateTime` |
| `DateOnly` | `Date` |
| `TimeOnly`, `TimeSpan` | `Time` |
| `Uri` | `Hyperlink` |
| enums and everything else | `String` |

## Errors

`ExcelGenerationException.Code` holds the JSON-RPC error code. For `-32602` (invalid params), `ErrorData` is an array of `{ message, sheet, row, column }`. Transport failures and timeouts still surface as `HttpRequestException` and `TaskCanceledException`.

```csharp
try
{
    var file = await client.GenerateAsync(workbook, ct);
}
catch (ExcelGenerationException ex) when (ex.Code == ExcelGenerationException.InvalidParams && ex.ErrorData is { } errors)
{
    foreach (var error in errors.EnumerateArray())
    {
        // each error: { message, sheet, row, column }
        logger.LogWarning("Invalid cell: {Error}", error.GetRawText());
    }
}
catch (ExcelGenerationException ex)
{
    logger.LogError(ex, "Excel generation failed with code {Code}", ex.Code);
}
```
