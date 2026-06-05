using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class AttachCommand : ICommandHandler
{
    private readonly IAttachService _service;

    public AttachCommand(IAttachService? service = null)
    {
        _service = service ?? new ExcelAttachService();
    }

    public string CommandName => "attach";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var workbookPath = BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "";
        if (string.IsNullOrWhiteSpace(workbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "attach_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "attach",
                Source: "xlflow-excel-bridge"));
        }

        var args = new AttachCommandArguments(
            WorkbookPath: workbookPath,
            Active: BridgePayload.GetBool(request.Payload, "Active"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");
        return _service.Execute(request, args, cancellationToken);
    }
}
