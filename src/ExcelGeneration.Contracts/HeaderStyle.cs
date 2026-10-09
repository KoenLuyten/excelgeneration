using System.Text.Json.Serialization;

namespace ExcelGeneration.Contracts;

/// <summary>Optional layout of the header row.</summary>
public sealed record HeaderStyle
{
    public bool Bold { get; init; } = true;
    public bool Italic { get; init; }

    /// <summary>Font color as HTML hex, e.g. <c>#FFFFFF</c>.</summary>
    public string? FontColor { get; init; }

    /// <summary>Fill color as HTML hex, e.g. <c>#1F4E78</c>.</summary>
    public string? BackgroundColor { get; init; }

    public double? FontSize { get; init; }
    public bool WrapText { get; init; }
    public HorizontalAlignment? HorizontalAlignment { get; init; }

    /// <summary>Draw a thin border below the header row.</summary>
    public bool BottomBorder { get; init; }

    /// <summary>Freeze the header row so it stays visible while scrolling. Defaults to <c>true</c>.</summary>
    public bool Freeze { get; init; } = true;
}

[JsonConverter(typeof(JsonStringEnumConverter<HorizontalAlignment>))]
public enum HorizontalAlignment
{
    Left,
    Center,
    Right
}
