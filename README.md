# Excel generation service

A .NET 10 Web API that builds `.xlsx` workbooks with [ClosedXML](https://github.com/ClosedXML/ClosedXML). Callers reach it over JSON-RPC 2.0 through [MediatorEndpoint](https://github.com/christophdebaene/MediatorEndpoint) + MediatR 12.5 (the last Apache-2.0 release).

| Project | Purpose |
|---|---|
| `src/ExcelGeneration.Contracts` | Payload contract, shared by server and client. NuGet package. |
| `src/ExcelGeneration.Api` | JSON-RPC endpoint `POST /jsonrpc`, method `Excel.Generate`. |
| `src/ExcelGeneration.Client` | Client NuGet package: gzip, `IEnumerable<T>`/`DataTable` mapping, typed HttpClient. |
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

## Run

```bash
dotnet run --project src/ExcelGeneration.Api      # http://localhost:5088, see ExcelGeneration.Api.http
dotnet test
dotnet pack -c Release -o artifacts                # ExcelGeneration.Contracts + ExcelGeneration.Client nupkgs
```
