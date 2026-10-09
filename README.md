# Excel generation service

A .NET 10 Web API that builds `.xlsx` workbooks with [ClosedXML](https://github.com/ClosedXML/ClosedXML). Callers reach it over JSON-RPC 2.0 through [MediatorEndpoint](https://github.com/christophdebaene/MediatorEndpoint) + MediatR 12.5 (the last Apache-2.0 release).

| Project | Purpose |
|---|---|
| `src/ExcelGeneration.Contracts` | Payload contract, shared by server and client. NuGet package. |
| `src/ExcelGeneration.Api` | JSON-RPC endpoint `POST /jsonrpc`, method `Excel.Generate`. |
| `src/ExcelGeneration.Client` | Client NuGet package: gzip, `IEnumerable<T>`/`DataTable` mapping, typed HttpClient. See [Using the client package](#using-the-client-package). |
| `tests/ExcelGeneration.Tests` | End-to-end tests (WebApplicationFactory) and mapping tests. |

## Explicit contract vs. `JObject`

The service uses an **explicit, typed contract with dynamic cell values**.

- **The library binds with System.Text.Json.** MediatorEndpoint deserializes `params` with System.Text.Json into the request type. `JObject` is a Newtonsoft.Json type, so using it would add a second serializer and bypass the library's binding.
- **The structure is fixed.** Sheets, columns, types, header layout and autofilter are known up front, so a typed contract gives:
  - validation, with error messages that point to the sheet, row and column;
  - IntelliSense for callers;
  - room for OpenAPI generation later (`MediatorEndpoint.JsonRpc.OpenApi`).
- **Only cell values are dynamic.** Rows are positional arrays (`[1,"abc",true,null]`) whose meaning comes from the column `type`:
  - this is much smaller than an array of objects, because property names are not repeated on every row;
  - a custom converter (`ExcelRowJsonConverter`) reads cells as plain primitives instead of a `JsonElement` per cell, which keeps server memory down on big payloads.

## Payload

```json
{ "id": "<string>", "jsonrpc": "2.0", "method": "Excel.Generate", "params": {
  "fileName": "report.xlsx",
  "sheets": [{
    "name": "Orders",
    "columns": [
      { "name": "Id", "type": "Integer" },
      { "name": "Date", "type": "Date", "format": "dd/mm/yyyy" },
      { "name": "Amount", "type": "Decimal", "format": "#,##0.00", "width": 14 }
    ],
    "showHeader": true,
    "header": { "bold": true, "fontColor": "#FFFFFF", "backgroundColor": "#1F4E78", "freeze": true,
                "italic": false, "fontSize": 11, "wrapText": false, "horizontalAlignment": "Center", "bottomBorder": true },
    "autoFilter": true,
    "autoFitColumns": true,
    "rows": [[1, "2026-10-01", 12.5], [2, "2026-10-02", null]]
  }]
}}
```

### Column types

| Type | Accepted values |
|---|---|
| `String` | any primitive |
| `Integer` | number or numeric string |
| `Decimal` | number or numeric string |
| `Boolean` | `true` / `false`, `0` / `1` |
| `DateTime`, `Date` | ISO 8601 string, or an OLE automation date number |
| `Time` | `HH:mm:ss` or `d.HH:mm:ss` |
| `Hyperlink` | absolute URL |
| `Formula` | A1 formula, with or without `=` |

`null` becomes an empty cell.

### Response

- **Success:** the raw `.xlsx` (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`), with the file name in `Content-Disposition`.
- **Failure:** a JSON-RPC error envelope (`application/json`, HTTP 200). The codes are:

| Code | Meaning |
|---|---|
| `-32600` | invalid request |
| `-32601` | unknown method |
| `-32602` | invalid params; `error.data` is `[{message, sheet, row, column}]` |
| `-32603` | internal error |
| `-32700` | parse error |

Note on MediatorEndpoint: `id` must be a **string**, and batch requests are not supported.

## Compression and large payloads

- **Requests:** bodies can be sent with `Content-Encoding: gzip`, `br` or `deflate`. The `ExcelGeneration.Client` package always gzips.
- **Responses:** JSON responses are compressed (gzip or brotli) when the caller sends `Accept-Encoding`. The xlsx itself is already a zip, so it is not compressed again.
- **Size limit:** `ExcelGeneration:MaxRequestBodySizeBytes` sets the limit (default 200 MB). It is checked against the **decompressed** body size.
- **Memory:** the server holds the decompressed JSON and the ClosedXML workbook in memory, so size its memory to match. The decompressed JSON is held once as a string (from MediatorEndpoint) and once as the parsed document.
- **Auto-fit:** column auto-fit only looks at the first 1,000 rows.

## Using the client package

`ExcelGeneration.Client` targets net8.0 and net10.0 and brings in `ExcelGeneration.Contracts`. It takes care of the JSON-RPC envelope, gzip and row mapping, so callers work with collections and `DataTable`s instead of raw payloads.

```bash
dotnet add package ExcelGeneration.Client
```

### Registration

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

Column types are inferred from the CLR type. See [Column mapping](src/ExcelGeneration.Client/README.md#column-mapping) for the full table.

### From a `DataTable` or `DataSet`

Columns are configured by name. The header text is the `DataColumn.Caption`, and `DBNull` becomes an empty cell.

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

### Error handling

When the service returns a JSON-RPC error, the client throws `ExcelGenerationException`. Transport failures and timeouts still surface as `HttpRequestException` and `TaskCanceledException`.

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

## Run

```bash
dotnet run --project src/ExcelGeneration.Api      # http://localhost:5088, see ExcelGeneration.Api.http
dotnet test
dotnet pack -c Release -o artifacts                # ExcelGeneration.Contracts + ExcelGeneration.Client nupkgs
```
