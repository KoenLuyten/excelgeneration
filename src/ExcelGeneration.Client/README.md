# ExcelGeneration.Client

Client for the Excel generation JSON-RPC service (`Excel.Generate`). It handles these for you:

- the JSON-RPC envelope;
- **gzip compression** of the request body. The body is streamed, so rows are serialized while they are sent;
- decompression of responses;
- turning `IEnumerable<T>` and `DataTable`/`DataSet` into the positional row payload;
- turning JSON-RPC errors into `ExcelGenerationException`.

## Registration

```csharp
services.AddExcelGenerationClient(o =>
{
    o.BaseAddress = new Uri("https://excel.example.com/");
    o.Timeout = TimeSpan.FromMinutes(10);
});
```

Or use it without DI: `new ExcelGenerationClient(httpClient)`. The `httpClient` must have its `BaseAddress` set.

## Usage

```csharp
var workbook = new WorkbookBuilder("orders.xlsx")
    .AddSheet("Orders", orders, s =>
    {
        s.AutoFilter = true;
        s.Header = new HeaderStyle { BackgroundColor = "#1F4E78", FontColor = "#FFFFFF" };
        s.Column(o => o.Amount, c => c.Format = "#,##0.00");
        s.Ignore(o => o.InternalNote);
    })
    .AddSheet(customersDataTable);           // sheet name = TableName

ExcelFile file = await client.GenerateAsync(workbook);
await file.SaveAsAsync("orders.xlsx");

// For very large files, write straight to a stream:
await client.GenerateToStreamAsync(workbook.Build(), responseBody);
```

### Column mapping

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

`ExcelGenerationException.Code` holds the JSON-RPC error code. For `-32602` (invalid params), `ErrorData` is an array of `{ message, sheet, row, column }`.
