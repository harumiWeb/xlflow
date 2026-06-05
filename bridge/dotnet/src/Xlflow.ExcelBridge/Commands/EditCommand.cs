using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class EditCommand : ICommandHandler
{
    private readonly IEditService _service;

    public EditCommand(IEditService? service = null)
    {
        _service = service ?? new ExcelEditService();
    }

    public string CommandName => "edit";

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
                Code: "edit_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "edit",
                Source: "xlflow-excel-bridge"));
        }

        var args = new EditCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: workbookPath,
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            Sheet: BridgePayload.GetString(request.Payload, "Sheet") ?? "",
            Cell: BridgePayload.GetString(request.Payload, "Cell") ?? "",
            RangeAddress: BridgePayload.GetString(request.Payload, "RangeAddress") ?? "",
            Rows: BridgePayload.GetString(request.Payload, "Rows") ?? "",
            Columns: BridgePayload.GetString(request.Payload, "Columns") ?? "",
            Value: BridgePayload.GetString(request.Payload, "Value") ?? "",
            ValueSpecified: BridgePayload.HasProperty(request.Payload, "Value"),
            Formula: BridgePayload.GetString(request.Payload, "Formula") ?? "",
            FormulaSpecified: BridgePayload.HasProperty(request.Payload, "Formula"),
            Fill: BridgePayload.GetString(request.Payload, "Fill") ?? "",
            Clear: BridgePayload.GetString(request.Payload, "Clear") ?? "",
            Height: BridgePayload.GetString(request.Payload, "Height") ?? "",
            Width: BridgePayload.GetString(request.Payload, "Width") ?? "",
            Events: BridgePayload.GetString(request.Payload, "Events") ?? "",
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");
        return _service.Execute(request, args, cancellationToken);
    }
}
