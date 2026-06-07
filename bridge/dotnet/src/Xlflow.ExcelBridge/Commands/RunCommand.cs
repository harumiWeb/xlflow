using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class RunCommand : ICommandHandler
{
    private readonly IRunService _service;

    public RunCommand(IRunService? service = null)
    {
        _service = service ?? new ExcelRunService();
    }

    public string CommandName => "run";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new RunCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            MacroName: BridgePayload.GetString(request.Payload, "MacroName") ?? "",
            MacroArgsJSON: BridgePayload.GetString(request.Payload, "MacroArgsJSON") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            DisplayAlerts: BridgePayload.GetBool(request.Payload, "DisplayAlerts"),
            SaveWorkbook: BridgePayload.GetBool(request.Payload, "SaveWorkbook"),
            Direct: BridgePayload.GetBool(request.Payload, "Direct"),
            Diagnostic: BridgePayload.GetBool(request.Payload, "Diagnostic"),
            SuppressModalErrors: BridgePayload.GetBool(request.Payload, "SuppressModalErrors"),
            MsgBoxResponsesJSON: BridgePayload.GetString(request.Payload, "MsgBoxResponsesJSON") ?? "",
            InputResponsesJSON: BridgePayload.GetString(request.Payload, "InputResponsesJSON") ?? "",
            FileDialogResponsesJSON: BridgePayload.GetString(request.Payload, "FileDialogResponsesJSON") ?? "",
            DebugStreamEnabled: BridgePayload.GetBool(request.Payload, "DebugStreamEnabled"),
            DebugStreamPipeName: BridgePayload.GetString(request.Payload, "DebugStreamPipeName") ?? "",
            UIStreamEnabled: BridgePayload.GetBool(request.Payload, "UIStreamEnabled"),
            UIStreamPipeName: BridgePayload.GetString(request.Payload, "UIStreamPipeName") ?? "",
            UIStreamRedactInput: BridgePayload.GetBool(request.Payload, "UIStreamRedactInput"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            RuntimeMode: BridgePayload.GetString(request.Payload, "RuntimeMode") ?? "",
            RuntimeSource: BridgePayload.GetString(request.Payload, "RuntimeSource") ?? "",
            SaveAsPath: BridgePayload.GetString(request.Payload, "SaveAsPath") ?? "",
            TimeoutSeconds: BridgePayload.GetInt(request.Payload, "TimeoutSeconds"),
            ModulesDir: BridgePayload.GetString(request.Payload, "ModulesDir") ?? "",
            ClassesDir: BridgePayload.GetString(request.Payload, "ClassesDir") ?? "",
            FormsDir: BridgePayload.GetString(request.Payload, "FormsDir") ?? "",
            WorkbookDir: BridgePayload.GetString(request.Payload, "WorkbookDir") ?? "",
            CodeSource: BridgePayload.GetString(request.Payload, "CodeSource") ?? "",
            Folders: BridgePayload.GetBool(request.Payload, "Folders"),
            FolderAnnotation: BridgePayload.GetString(request.Payload, "FolderAnnotation") ?? "update",
            DefaultComponentFolders: BridgePayload.GetBool(request.Payload, "DefaultComponentFolders"));

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "run_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "run",
                Source: "xlflow"));
        }

        if (string.IsNullOrWhiteSpace(args.MacroName))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "run_args_invalid",
                Message: "MacroName is required",
                Phase: "run",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
