using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class ListCommand : ICommandHandler
{
    private readonly IListService _service;

    public ListCommand(IListService? service = null)
    {
        _service = service ?? new ExcelListService();
    }

    public string CommandName => "list";

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
                Code: "list_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "list",
                Source: "xlflow-excel-bridge"));
        }

        var args = new ListCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: workbookPath,
            FormsDir: BridgePayload.GetString(request.Payload, "FormsDir") ?? "",
            ProjectRoot: BridgePayload.GetString(request.Payload, "ProjectRoot") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");
        return _service.Execute(request, args, cancellationToken);
    }
}
