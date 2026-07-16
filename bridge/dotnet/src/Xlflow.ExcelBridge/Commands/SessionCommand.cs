using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class SessionCommand : ICommandHandler
{
    private readonly ISessionService _service;

    public SessionCommand(ISessionService? service = null)
    {
        _service = service ?? new ExcelSessionService();
    }

    public string CommandName => "session";

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
                Code: "session_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "session",
                Source: "xlflow-excel-bridge"));
        }

        var args = new SessionCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: workbookPath,
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            Discard: BridgePayload.GetBool(request.Payload, "Discard"));
        return _service.Execute(request, args, cancellationToken);
    }
}
