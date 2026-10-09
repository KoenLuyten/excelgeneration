using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Text.Json;
using ExcelGeneration.Client;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Tests;

public sealed class MappingTests
{
    public sealed class Product
    {
        [ExcelColumn("Product name", Order = 1)]
        public string Name { get; init; } = "";

        [ExcelColumn(Order = 0, Format = "0.00", Width = 12)]
        public decimal Price { get; init; }

        [ExcelIgnore]
        public string Secret { get; init; } = "";

        [Display(Name = "In stock")]
        public bool InStock { get; init; }

        public int? Quantity { get; init; }

        [ExcelColumn(Type = ColumnType.Hyperlink)]
        public string Url { get; init; } = "";
    }

    [Fact]
    public void Maps_properties_with_attributes()
    {
        var request = new WorkbookBuilder()
            .AddSheet("Products", [new Product { Name = "Pen", Price = 1.5m, Secret = "x", InStock = true, Url = "https://x" }])
            .Build();

        var sheet = Assert.Single(request.Sheets);
        Assert.Equal(["Price", "Product name", "In stock", "Quantity", "Url"], sheet.Columns.Select(c => c.Name));
        Assert.Equal([ColumnType.Decimal, ColumnType.String, ColumnType.Boolean, ColumnType.Integer, ColumnType.Hyperlink], sheet.Columns.Select(c => c.Type));
        Assert.Equal("0.00", sheet.Columns[0].Format);
        Assert.Equal(12, sheet.Columns[0].Width);
        Assert.Equal([1.5m, "Pen", true, null, "https://x"], Assert.Single(sheet.Rows).Values);
    }

    [Fact]
    public void Options_override_attributes_and_can_ignore()
    {
        var request = new WorkbookBuilder()
            .AddSheet("Products", Array.Empty<Product>(), s => s
                .Column(p => p.Name, c => { c.Name = "Name"; c.Order = -1; })
                .Ignore(p => p.Url)
                .Ignore(p => p.Quantity))
            .Build();

        Assert.Equal(["Name", "Price", "In stock"], request.Sheets[0].Columns.Select(c => c.Name));
    }

    [Fact]
    public void Rows_are_enumerated_lazily()
    {
        var enumerated = 0;
        IEnumerable<Product> Source()
        {
            enumerated++;
            yield return new Product();
        }

        var request = new WorkbookBuilder().AddSheet("P", Source()).Build();
        Assert.Equal(0, enumerated);

        _ = request.Sheets[0].Rows.ToList();
        Assert.Equal(1, enumerated);
    }

    [Fact]
    public void Maps_datatable_with_dbnull_and_deleted_rows()
    {
        var table = new DataTable("T");
        table.Columns.Add("Id", typeof(long));
        table.Columns.Add("When", typeof(DateTime));
        table.Columns.Add("Hidden", typeof(string));
        table.Rows.Add(1L, DBNull.Value, "h");
        table.Rows.Add(2L, new DateTime(2026, 1, 1), "h");
        table.AcceptChanges();
        table.Rows[1].Delete();

        var sheet = new WorkbookBuilder().AddSheet(table, s => s.Ignore("Hidden")).Build().Sheets[0];

        Assert.Equal("T", sheet.Name);
        Assert.Equal([ColumnType.Integer, ColumnType.DateTime], sheet.Columns.Select(c => c.Type));
        Assert.Equal([1L, null], Assert.Single(sheet.Rows).Values);
    }

    [Fact]
    public void Rows_serialize_as_positional_arrays_and_roundtrip_to_primitives()
    {
        var row = new ExcelRow(1, "a", true, null, 1.5m, new DateOnly(2026, 1, 2), new TimeOnly(10, 30), DayOfWeek.Monday, double.NaN, new DateTime(2026, 1, 2, 3, 4, 5));

        var json = JsonSerializer.Serialize(row);
        Assert.Equal("""[1,"a",true,null,1.5,"2026-01-02","10:30:00","Monday",null,"2026-01-02T03:04:05"]""", json);

        var back = JsonSerializer.Deserialize<ExcelRow>(json);
        Assert.Equal([1L, "a", true, null, 1.5m, "2026-01-02", "10:30:00", "Monday", null, "2026-01-02T03:04:05"], back.Values);
    }
}
