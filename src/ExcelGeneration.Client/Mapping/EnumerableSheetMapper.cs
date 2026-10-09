using System.Collections.Concurrent;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;
using System.Reflection;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client.Mapping;

/// <summary>Maps a collection of objects to a <see cref="SheetDefinition"/> using its public properties.</summary>
internal static class EnumerableSheetMapper
{
    public static SheetDefinition Map<T>(string name, IEnumerable<T> items, SheetOptions options)
    {
        ArgumentNullException.ThrowIfNull(items);

        var columns = PropertyAccessors<T>.All
            .Where(p => !options.Ignored.Contains(p.Property.Name))
            .Select((p, index) => (Accessor: p, Index: index, Definition: Describe(p.Property, options.GetColumn(p.Property.Name))))
            .OrderBy(x => x.Definition.Order ?? int.MaxValue)
            .ThenBy(x => x.Index)
            .ToArray();

        if (columns.Length == 0)
            throw new InvalidOperationException($"Type {typeof(T)} has no public readable properties to map to columns.");

        var getters = columns.Select(c => c.Accessor.Getter).ToArray();

        return new SheetDefinition
        {
            Name = name,
            Columns = columns.Select(c => c.Definition.Column).ToArray(),
            // Lazy: items are only enumerated while the request body is being written.
            Rows = items.Select(item => ToRow(item, getters)),
            ShowHeader = options.ShowHeader,
            Header = options.Header,
            AutoFilter = options.AutoFilter,
            AutoFitColumns = options.AutoFitColumns
        };
    }

    private static ExcelRow ToRow<T>(T item, Func<T, object?>[] getters)
    {
        var values = new object?[getters.Length];
        if (item is not null)
        {
            for (var i = 0; i < getters.Length; i++)
                values[i] = getters[i](item);
        }
        return new ExcelRow(values);
    }

    private static (ColumnDefinition Column, int? Order) Describe(PropertyInfo property, ColumnOptions? options)
    {
        var excel = property.GetCustomAttribute<ExcelColumnAttribute>();
        var display = property.GetCustomAttribute<DisplayAttribute>();

        var column = new ColumnDefinition
        {
            Name = options?.Name ?? excel?.Name ?? display?.GetName() ?? property.Name,
            Type = options?.Type ?? excel?.TypeOrNull ?? ClrTypeMap.Map(property.PropertyType),
            Format = options?.Format ?? excel?.Format,
            Width = options?.Width ?? excel?.WidthOrNull
        };

        return (column, options?.Order ?? excel?.OrderOrNull ?? display?.GetOrder());
    }

    private sealed record PropertyAccessor<T>(PropertyInfo Property, Func<T, object?> Getter);

    private static class PropertyAccessors<T>
    {
        public static readonly PropertyAccessor<T>[] All = typeof(T)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead && p.GetMethod!.IsPublic && p.GetIndexParameters().Length == 0)
            .Where(p => p.GetCustomAttribute<ExcelIgnoreAttribute>() is null)
            .OrderBy(p => p.MetadataToken)
            .Select(p => new PropertyAccessor<T>(p, Compile(p)))
            .ToArray();

        private static Func<T, object?> Compile(PropertyInfo property)
        {
            var instance = Expression.Parameter(typeof(T), "x");
            var body = Expression.Convert(Expression.Property(instance, property), typeof(object));
            return Expression.Lambda<Func<T, object?>>(body, instance).Compile();
        }
    }
}
