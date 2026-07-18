using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class PullCommand : ICommandHandler
{
    private readonly IPullService _service;

    public PullCommand(IPullService? service = null)
    {
        _service = service ?? new ExcelPullService();
    }

    public string CommandName => "pull";

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
                Code: "pull_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "pull",
                Source: "xlflow-excel-bridge"));
        }

        var args = new PullCommandArguments(
            WorkbookPath: workbookPath,
            ModulesDir: BridgePayload.GetString(request.Payload, "ModulesDir") ?? "",
            ClassesDir: BridgePayload.GetString(request.Payload, "ClassesDir") ?? "",
            FormsDir: BridgePayload.GetString(request.Payload, "FormsDir") ?? "",
            WorkbookDir: BridgePayload.GetString(request.Payload, "WorkbookDir") ?? "",
            CodeSource: BridgePayload.GetString(request.Payload, "CodeSource") ?? "",
            Folders: BridgePayload.GetBool(request.Payload, "Folders"),
            FolderAnnotation: BridgePayload.GetString(request.Payload, "FolderAnnotation") ?? "",
            DefaultComponentFolders: BridgePayload.GetBool(request.Payload, "DefaultComponentFolders"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            LineNumbersEnabled: BridgePayload.GetBool(request.Payload, "LineNumbersEnabled"));

        return _service.Execute(request, args, cancellationToken);
    }
}
