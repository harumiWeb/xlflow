using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class InspectCommand : ICommandHandler
{
    private readonly IInspectService _service;

    public InspectCommand(IInspectService? service = null)
    {
        _service = service ?? new ExcelInspectionService();
    }

    public string CommandName => "inspect";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new InspectCommandArguments(
            Target: BridgePayload.GetString(request.Payload, "Target") ?? "",
            Sheet: BridgePayload.GetString(request.Payload, "Sheet") ?? "",
            Address: BridgePayload.GetString(request.Payload, "Address") ?? "",
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            IncludeStyle: BridgePayload.GetBool(request.Payload, "IncludeStyle"),
            MaxRows: BridgePayload.GetInt(request.Payload, "MaxRows"),
            MaxCols: BridgePayload.GetInt(request.Payload, "MaxCols"));

        if (string.IsNullOrWhiteSpace(args.Target))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "inspect_args_invalid",
                Message: "Target is required",
                Phase: "inspect",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
