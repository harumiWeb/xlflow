using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class FormExportImageCommand : ICommandHandler
{
    private readonly IFormExportImageService _service;

    public FormExportImageCommand(IFormExportImageService? service = null)
    {
        _service = service ?? new ExcelFormExportImageService();
    }

    public string CommandName => "form-export-image";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new FormExportImageCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            FormName: BridgePayload.GetString(request.Payload, "FormName") ?? "",
            OutputPath: BridgePayload.GetString(request.Payload, "OutputPath") ?? "",
            Initializer: BridgePayload.GetString(request.Payload, "Initializer") ?? "",
            Overwrite: BridgePayload.GetBool(request.Payload, "Overwrite"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return Failure(request, "form_export_image_args_invalid", "WorkbookPath is required");
        }
        if (string.IsNullOrWhiteSpace(args.FormName))
        {
            return Failure(request, "form_export_image_args_invalid", "FormName is required");
        }
        if (string.IsNullOrWhiteSpace(args.OutputPath))
        {
            return Failure(request, "form_export_image_args_invalid", "OutputPath is required");
        }

        return _service.Execute(request, args, cancellationToken);
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: code,
            Message: message,
            Phase: "form-export-image",
            Source: "xlflow"));
    }
}
