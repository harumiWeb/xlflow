using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class InspectFormCommand : ICommandHandler
{
    private readonly IInspectFormService _service;

    public InspectFormCommand(IInspectFormService? service = null)
    {
        _service = service ?? new ExcelFormInspectionService();
    }

    public string CommandName => "inspect-form";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new InspectFormCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            FormName: BridgePayload.GetString(request.Payload, "FormName") ?? "",
            Basis: BridgePayload.GetString(request.Payload, "Basis") ?? "runtime",
            Initializer: BridgePayload.GetString(request.Payload, "Initializer") ?? "",
            StrictDesigner: BridgePayload.GetBool(request.Payload, "StrictDesigner"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "inspect_form_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "inspect-form",
                Source: "xlflow"));
        }

        if (string.IsNullOrWhiteSpace(args.FormName))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "inspect_form_args_invalid",
                Message: "FormName is required",
                Phase: "inspect-form",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
