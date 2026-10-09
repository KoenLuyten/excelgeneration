using ClosedXML.Excel;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Api.Excel;

/// <summary>Builds an .xlsx workbook from a <see cref="GenerateExcelRequest"/> using ClosedXML.</summary>
public sealed class WorkbookWriter
{
    public const int MaxRows = 1_048_576;

    /// <summary>Column auto-fit only measures this many rows; measuring every row is very slow on large sheets.</summary>
    private const int AutoFitSampleRows = 1_000;

    /// <summary>Stop collecting cell errors after this many to keep the error response small.</summary>
    private const int MaxReportedErrors = 50;

    public byte[] Write(GenerateExcelRequest request, CancellationToken cancellationToken = default)
    {
        using var workbook = new XLWorkbook();

        foreach (var sheet in request.Sheets)
        {
            WriteSheet(workbook, sheet, cancellationToken);
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static void WriteSheet(XLWorkbook workbook, SheetDefinition sheet, CancellationToken cancellationToken)
    {
        var worksheet = workbook.Worksheets.Add(SheetNames.Sanitize(sheet.Name));
        var columns = sheet.Columns;
        var errors = new List<PayloadError>();

        var firstDataRow = sheet.ShowHeader ? 2 : 1;
        if (sheet.ShowHeader)
        {
            for (var c = 0; c < columns.Count; c++)
            {
                worksheet.Cell(1, c + 1).Value = columns[c].Name;
            }
        }

        var rowNumber = firstDataRow - 1;
        long dataRowIndex = 0;
        foreach (var row in sheet.Rows)
        {
            dataRowIndex++;
            rowNumber++;
            if (rowNumber > MaxRows)
            {
                throw new PayloadValidationException(new PayloadError($"A sheet can contain at most {MaxRows} rows (including the header).", sheet.Name));
            }

            if ((dataRowIndex & 0x3FF) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var values = row.Values;
            if (values.Count > columns.Count)
            {
                AddError(errors, new PayloadError($"Row has {values.Count} values but only {columns.Count} columns are defined.", sheet.Name, dataRowIndex));
                continue;
            }

            for (var c = 0; c < values.Count; c++)
            {
                var value = values[c];
                if (value is null)
                    continue;

                var column = columns[c];
                if (!CellValueConverter.TryConvert(value, column.Type, out var cellValue))
                {
                    AddError(errors, new PayloadError($"Value '{value}' is not a valid {column.Type}.", sheet.Name, dataRowIndex, column.Name));
                    continue;
                }

                var cell = worksheet.Cell(rowNumber, c + 1);
                try
                {
                    SetCell(cell, column.Type, cellValue);
                }
                catch (Exception exc) when (exc is ArgumentException or FormatException or UriFormatException)
                {
                    AddError(errors, new PayloadError($"Value '{value}' is not a valid {column.Type}: {exc.Message}", sheet.Name, dataRowIndex, column.Name));
                }
            }
        }

        if (errors.Count > 0)
        {
            throw new PayloadValidationException(errors);
        }

        var lastRow = Math.Max(rowNumber, 1);
        ApplyColumnLayout(worksheet, sheet, firstDataRow, rowNumber);

        if (sheet.ShowHeader)
        {
            ApplyHeaderStyle(worksheet, sheet.Header, columns.Count);
            if (sheet.AutoFilter)
            {
                worksheet.Range(1, 1, lastRow, columns.Count).SetAutoFilter();
            }
        }
    }

    private static void SetCell(IXLCell cell, ColumnType type, XLCellValue value)
    {
        switch (type)
        {
            case ColumnType.Formula:
                cell.FormulaA1 = value.GetText().TrimStart('=');
                break;
            case ColumnType.Hyperlink:
                var url = value.GetText();
                cell.Value = url;
                cell.SetHyperlink(new XLHyperlink(new Uri(url, UriKind.Absolute)));
                break;
            default:
                cell.Value = value;
                break;
        }
    }

    private static void ApplyColumnLayout(IXLWorksheet worksheet, SheetDefinition sheet, int firstDataRow, int lastDataRow)
    {
        var columns = sheet.Columns;
        var hasData = lastDataRow >= firstDataRow;

        for (var c = 0; c < columns.Count; c++)
        {
            var definition = columns[c];
            var column = worksheet.Column(c + 1);

            var format = definition.Format ?? DefaultFormat(definition.Type);
            if (format is not null && hasData)
            {
                worksheet.Range(firstDataRow, c + 1, lastDataRow, c + 1).Style.NumberFormat.Format = format;
            }

            if (definition.Width is { } width)
            {
                column.Width = width;
            }
            else if (sheet.AutoFitColumns)
            {
                var lastSampleRow = hasData ? Math.Min(lastDataRow, firstDataRow + AutoFitSampleRows - 1) : 1;
                column.AdjustToContents(1, lastSampleRow);
            }
        }
    }

    private static string? DefaultFormat(ColumnType type) => type switch
    {
        ColumnType.Date => "yyyy-mm-dd",
        ColumnType.DateTime => "yyyy-mm-dd hh:mm:ss",
        ColumnType.Time => "[h]:mm:ss",
        _ => null
    };

    private static void ApplyHeaderStyle(IXLWorksheet worksheet, HeaderStyle? header, int columnCount)
    {
        if (header is null)
            return;

        var style = worksheet.Range(1, 1, 1, columnCount).Style;
        style.Font.Bold = header.Bold;
        style.Font.Italic = header.Italic;
        style.Alignment.WrapText = header.WrapText;

        if (header.FontSize is { } size)
            style.Font.FontSize = size;
        if (header.FontColor is { } fontColor)
            style.Font.FontColor = XLColor.FromHtml(NormalizeColor(fontColor));
        if (header.BackgroundColor is { } background)
            style.Fill.BackgroundColor = XLColor.FromHtml(NormalizeColor(background));
        if (header.BottomBorder)
            style.Border.BottomBorder = XLBorderStyleValues.Thin;
        if (header.HorizontalAlignment is { } alignment)
        {
            style.Alignment.Horizontal = alignment switch
            {
                HorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
                HorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
                _ => XLAlignmentHorizontalValues.Left
            };
        }

        if (header.Freeze)
            worksheet.SheetView.FreezeRows(1);
    }

    private static string NormalizeColor(string color) => color.StartsWith('#') ? color : "#" + color;

    private static void AddError(List<PayloadError> errors, PayloadError error)
    {
        if (errors.Count < MaxReportedErrors)
            errors.Add(error);
        else if (errors.Count == MaxReportedErrors)
            errors.Add(new PayloadError($"More than {MaxReportedErrors} errors; remaining errors are not reported."));
    }
}
