using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client;

/// <summary>Configures how a property is written as an Excel column.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute : Attribute
{
    private ColumnType? _type;
    private int? _order;
    private double? _width;

    public ExcelColumnAttribute()
    {
    }

    public ExcelColumnAttribute(string name) => Name = name;

    /// <summary>Column header. Defaults to the property name.</summary>
    public string? Name { get; set; }

    /// <summary>Overrides the column type inferred from the property type.</summary>
    public ColumnType Type { get => _type ?? ColumnType.String; set => _type = value; }

    /// <summary>Column position; columns without an order keep their declaration order after ordered ones.</summary>
    public int Order { get => _order ?? 0; set => _order = value; }

    /// <summary>Excel number format, e.g. <c>#,##0.00</c>.</summary>
    public string? Format { get; set; }

    /// <summary>Fixed column width.</summary>
    public double Width { get => _width ?? 0; set => _width = value; }

    internal ColumnType? TypeOrNull => _type;
    internal int? OrderOrNull => _order;
    internal double? WidthOrNull => _width;
}

/// <summary>Excludes a property from the generated sheet.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelIgnoreAttribute : Attribute;
