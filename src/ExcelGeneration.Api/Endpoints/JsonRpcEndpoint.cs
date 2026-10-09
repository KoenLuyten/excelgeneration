using System.Text.Json;
using ExcelGeneration.Api.Excel;
using ExcelGeneration.Contracts;
using MediatorEndpoint.JsonRpc;
using MediatR;

namespace ExcelGeneration.Api.Endpoints;

public static class JsonRpcEndpoint
{
    public static IEndpointRouteBuilder MapJsonRpc(this IEndpointRouteBuilder app)
    {
        app.MapPost("/" + ExcelJsonRpc.Route, async (HttpContext context, ISender sender, JsonRpc? jsonRpc, ILogger<JsonRpc> logger, CancellationToken cancellationToken) =>
        {
            var id = jsonRpc!.Request.Id;
            try
            {
                var message = await jsonRpc.CreateMessageAsync(context);
                var response = await sender.Send(message!, cancellationToken);
                return JsonRpcResults.Response(id, response);
            }
            catch (JsonException exc)
            {
                return JsonRpcResults.Response(JsonRpcErrorResponse.Create(id, JsonRpcErrorCode.InvalidParams, exc.Message, new { exc.Path }));
            }
            catch (PayloadValidationException exc)
            {
                return JsonRpcResults.Response(JsonRpcErrorResponse.Create(id, JsonRpcErrorCode.InvalidParams, exc.Message, exc.Errors));
            }
            catch (Exception exc) when (exc is not OperationCanceledException)
            {
                logger.LogError(exc, "JSON-RPC method {Method} ({Id}) failed", jsonRpc.Request.Method, id);
                return JsonRpcResults.Response(JsonRpcErrorResponse.Create(id, JsonRpcErrorCode.InternalError, "An unexpected error occurred."));
            }
        })
        .AddEndpointFilter<JsonRpcValidationFilter>()
        .DisableAntiforgery();

        return app;
    }
}
