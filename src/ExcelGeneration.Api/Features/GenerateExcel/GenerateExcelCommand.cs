using ExcelGeneration.Api.Excel;
using ExcelGeneration.Contracts;
using MediatorEndpoint;
using MediatorEndpoint.Responses;
using MediatR;

namespace ExcelGeneration.Api.Features.GenerateExcel;

/// <summary>Server-side request of <c>Excel.Generate</c>; same JSON shape as the public contract.</summary>
[Command]
[RequestName(ExcelJsonRpc.ServiceName, ExcelJsonRpc.GenerateMethodName)]
public sealed record GenerateExcelCommand : GenerateExcelRequest, IRequest<FileResponse>;

public sealed class GenerateExcelHandler(WorkbookWriter writer) : IRequestHandler<GenerateExcelCommand, FileResponse>
{
    public Task<FileResponse> Handle(GenerateExcelCommand request, CancellationToken cancellationToken)
    {
        RequestValidator.Validate(request);

        var content = writer.Write(request, cancellationToken);
        var fileName = FileNames.Normalize(request.FileName);

        return Task.FromResult(new FileResponse(fileName, ExcelJsonRpc.XlsxContentType, content));
    }
}
