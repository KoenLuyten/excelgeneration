namespace ExcelGeneration.Api.Excel;

/// <summary>The payload is well-formed JSON but cannot be turned into a workbook. Mapped to JSON-RPC -32602.</summary>
public sealed class PayloadValidationException(IReadOnlyList<PayloadError> errors)
    : Exception(errors.Count == 1 ? errors[0].Message : $"The request contains {errors.Count} errors. See error data.")
{
    public IReadOnlyList<PayloadError> Errors { get; } = errors;

    public PayloadValidationException(PayloadError error) : this([error])
    {
    }
}

/// <param name="Sheet">Sheet name, if applicable.</param>
/// <param name="Row">1-based data row index (excluding the header), if applicable.</param>
/// <param name="Column">Column name, if applicable.</param>
public sealed record PayloadError(string Message, string? Sheet = null, long? Row = null, string? Column = null);
