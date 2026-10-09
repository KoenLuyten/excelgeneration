namespace ExcelGeneration.Client;

/// <summary>A generated workbook.</summary>
public sealed record ExcelFile(string FileName, byte[] Content)
{
    public string ContentType => Contracts.ExcelJsonRpc.XlsxContentType;

    public Stream OpenRead() => new MemoryStream(Content, writable: false);

    public Task SaveAsAsync(string path, CancellationToken cancellationToken = default)
        => File.WriteAllBytesAsync(path, Content, cancellationToken);
}
