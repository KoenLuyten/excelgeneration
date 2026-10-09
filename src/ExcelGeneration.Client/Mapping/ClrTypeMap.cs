using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client.Mapping;

/// <summary>Infers the <see cref="ColumnType"/> from a CLR type.</summary>
internal static class ClrTypeMap
{
    public static ColumnType Map(Type type)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;

        if (type.IsEnum)
            return ColumnType.String;

        return Type.GetTypeCode(type) switch
        {
            TypeCode.Byte or TypeCode.SByte or TypeCode.Int16 or TypeCode.UInt16 or
            TypeCode.Int32 or TypeCode.UInt32 or TypeCode.Int64 or TypeCode.UInt64 => ColumnType.Integer,
            TypeCode.Single or TypeCode.Double or TypeCode.Decimal => ColumnType.Decimal,
            TypeCode.Boolean => ColumnType.Boolean,
            TypeCode.DateTime => ColumnType.DateTime,
            _ when type == typeof(DateTimeOffset) => ColumnType.DateTime,
            _ when type == typeof(DateOnly) => ColumnType.Date,
            _ when type == typeof(TimeOnly) || type == typeof(TimeSpan) => ColumnType.Time,
            _ when type == typeof(Uri) => ColumnType.Hyperlink,
            _ => ColumnType.String
        };
    }
}
