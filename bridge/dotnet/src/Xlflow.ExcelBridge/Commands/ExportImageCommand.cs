using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class ExportImageCommand : ICommandHandler
{
    private readonly IExportImageService _service;

    public ExportImageCommand(IExportImageService? service = null)
    {
        _service = service ?? new ExcelExportImageService();
    }

    public string CommandName => "export-image";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new ExportImageCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            Sheet: BridgePayload.GetString(request.Payload, "Sheet") ?? "",
            RangeAddress: BridgePayload.GetString(request.Payload, "RangeAddress") ?? "",
            OutputPath: BridgePayload.GetString(request.Payload, "OutputPath") ?? "",
            OutputIsDefault: BridgePayload.GetBool(request.Payload, "OutputIsDefault"),
            ImageFormat: BridgePayload.GetString(request.Payload, "ImageFormat") ?? "png",
            Overwrite: BridgePayload.GetBool(request.Payload, "Overwrite"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return Failure(request, "export_image_args_invalid", "WorkbookPath is required");
        }
        if (string.IsNullOrWhiteSpace(args.Sheet))
        {
            return Failure(request, "export_image_args_invalid", "Sheet is required");
        }
        if (string.IsNullOrWhiteSpace(args.RangeAddress))
        {
            return Failure(request, "export_image_args_invalid", "RangeAddress is required");
        }
        if (string.IsNullOrWhiteSpace(args.OutputPath))
        {
            return Failure(request, "export_image_args_invalid", "OutputPath is required");
        }

        return _service.Execute(request, args, cancellationToken);
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: code,
            Message: message,
            Phase: "export-image",
            Source: "xlflow"));
    }
}
