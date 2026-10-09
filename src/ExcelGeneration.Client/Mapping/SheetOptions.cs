using System.Linq.Expressions;
using System.Reflection;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client;

/// <summary>Sheet level options used when mapping a <see cref="System.Data.DataTable"/> or a collection.</summary>
public class SheetOptions
{
    internal Dictionary<string, ColumnOptions> Columns { get; } = new(StringComparer.Ordinal);
    internal HashSet<string> Ignored { get; } = new(StringComparer.Ordinal);

    /// <summary>Write the column names as first row. Defaults to <c>true</c>.</summary>
    public bool ShowHeader { get; set; } = true;

    /// <summary>Optional layout of the header row.</summary>
    public HeaderStyle? Header { get; set; }

    /// <summary>Add an autofilter on the header row.</summary>
    public bool AutoFilter { get; set; }

    /// <summary>Adjust column widths to content. Defaults to <c>true</c>.</summary>
    public bool AutoFitColumns { get; set; } = true;

    /// <summary>Configures the column mapped from the property or <see cref="System.Data.DataColumn"/> named <paramref name="sourceName"/>.</summary>
    public SheetOptions Column(string sourceName, Action<ColumnOptions> configure)
    {
        ArgumentNullException.ThrowIfNull(sourceName);
        ArgumentNullException.ThrowIfNull(configure);
        if (!Columns.TryGetValue(sourceName, out var options))
        {
            Columns[sourceName] = options = new ColumnOptions();
        }
        configure(options);
        return this;
    }

    /// <summary>Excludes the property or <see cref="System.Data.DataColumn"/> named <paramref name="sourceName"/>.</summary>
    public SheetOptions Ignore(string sourceName)
    {
        Ignored.Add(sourceName);
        return this;
    }

    internal ColumnOptions? GetColumn(string sourceName) => Columns.GetValueOrDefault(sourceName);
}

/// <summary>Options for mapping a collection of <typeparamref name="T"/>.</summary>
public sealed class SheetOptions<T> : SheetOptions
{
    /// <summary>Configures the column mapped from <paramref name="property"/>.</summary>
    public SheetOptions<T> Column<TProperty>(Expression<Func<T, TProperty>> property, Action<ColumnOptions> configure)
    {
        Column(GetMemberName(property), configure);
        return this;
    }

    /// <summary>Excludes <paramref name="property"/> from the sheet.</summary>
    public SheetOptions<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property)
    {
        Ignore(GetMemberName(property));
        return this;
    }

    private static string GetMemberName<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        var body = expression.Body is UnaryExpression { NodeType: ExpressionType.Convert } convert ? convert.Operand : expression.Body;
        return body is MemberExpression { Member: PropertyInfo property } && property.DeclaringType!.IsAssignableFrom(typeof(T))
            ? property.Name
            : throw new ArgumentException("Expression must select a property, e.g. x => x.Name.", nameof(expression));
    }
}

/// <summary>Column level options. Unset values fall back to attributes and inferred defaults.</summary>
public sealed class ColumnOptions
{
    public string? Name { get; set; }
    public ColumnType? Type { get; set; }
    public string? Format { get; set; }
    public double? Width { get; set; }
    public int? Order { get; set; }
}
