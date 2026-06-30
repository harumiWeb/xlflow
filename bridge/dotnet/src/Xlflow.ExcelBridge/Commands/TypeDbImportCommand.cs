using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class TypeDbImportCommand : ICommandHandler
{
    private readonly TypeLibImporterService _service;

    public TypeDbImportCommand(TypeLibImporterService? service = null)
    {
        _service = service ?? new TypeLibImporterService();
    }

    public string CommandName => "type-db-import";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new TypeDbImportArguments(
            OutputDir: BridgePayload.GetString(request.Payload, "OutputDir") ?? "",
            GeneratorVersion: BridgePayload.GetString(request.Payload, "GeneratorVersion") ?? "dev",
            Libraries: BridgePayload.GetString(request.Payload, "Libraries") ?? "excel");
        return _service.Execute(request, args, cancellationToken);
    }
}
