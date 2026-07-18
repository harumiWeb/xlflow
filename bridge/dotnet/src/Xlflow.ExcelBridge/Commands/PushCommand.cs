using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class PushCommand : ICommandHandler
{
    private readonly IPushService _service;

    public PushCommand(IPushService? service = null)
    {
        _service = service ?? new ExcelPushService();
    }

    public string CommandName => "push";

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
                Code: "push_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "push",
                Source: "xlflow-excel-bridge"));
        }

        var args = new PushCommandArguments(
            WorkbookPath: workbookPath,
            ModulesDir: BridgePayload.GetString(request.Payload, "ModulesDir") ?? "",
            ClassesDir: BridgePayload.GetString(request.Payload, "ClassesDir") ?? "",
            FormsDir: BridgePayload.GetString(request.Payload, "FormsDir") ?? "",
            WorkbookDir: BridgePayload.GetString(request.Payload, "WorkbookDir") ?? "",
            CodeSource: BridgePayload.GetString(request.Payload, "CodeSource") ?? "",
            BackupRoot: BridgePayload.GetString(request.Payload, "BackupRoot") ?? "",
            Folders: BridgePayload.GetBool(request.Payload, "Folders"),
            FolderAnnotation: BridgePayload.GetString(request.Payload, "FolderAnnotation") ?? "",
            DefaultComponentFolders: BridgePayload.GetBool(request.Payload, "DefaultComponentFolders"),
            StatePath: BridgePayload.GetString(request.Payload, "StatePath") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            BackupMode: BridgePayload.GetString(request.Payload, "BackupMode") ?? "",
            ChangedOnly: BridgePayload.GetBool(request.Payload, "ChangedOnly"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            NoSave: BridgePayload.GetBool(request.Payload, "NoSave"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            LineNumbersEnabled: BridgePayload.GetBool(request.Payload, "LineNumbersEnabled"));

        return _service.Execute(request, args, cancellationToken);
    }
}
