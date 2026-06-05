using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class UICommand : ICommandHandler
{
    private readonly IUIService _service;

    public UICommand(IUIService? service = null)
    {
        _service = service ?? new ExcelUIService();
    }

    public string CommandName => "ui";

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
                Code: "ui_button_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "ui",
                Source: "xlflow-excel-bridge"));
        }

        var args = new UICommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: workbookPath,
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            Sheet: BridgePayload.GetString(request.Payload, "Sheet") ?? "",
            Cell: BridgePayload.GetString(request.Payload, "Cell") ?? "",
            Text: BridgePayload.GetString(request.Payload, "Text") ?? "",
            Macro: BridgePayload.GetString(request.Payload, "Macro") ?? "",
            Id: BridgePayload.GetString(request.Payload, "Id") ?? "",
            Width: BridgePayload.GetInt(request.Payload, "Width", 160),
            Height: BridgePayload.GetInt(request.Payload, "Height", 40),
            CreateSheet: BridgePayload.GetBool(request.Payload, "CreateSheet"),
            VerifyMacro: BridgePayload.GetBool(request.Payload, "VerifyMacro"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");
        return _service.Execute(request, args, cancellationToken);
    }
}
