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
            TraceEnabled: BridgePayload.GetBool(request.Payload, "TraceEnabled"),
            TraceFile: BridgePayload.GetString(request.Payload, "TraceFile") ?? "",
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
            TimeoutSeconds: BridgePayload.GetInt(request.Payload, "TimeoutSeconds"));

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

        if (args.TraceEnabled)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "run_args_invalid",
                Message: "--trace is not supported by xlflow run --bridge dotnet in this partial parity release; use --bridge powershell for trace collection.",
                Phase: "run",
                Source: "xlflow"));
        }

        if (!string.IsNullOrWhiteSpace(args.MsgBoxResponsesJSON) ||
            !string.IsNullOrWhiteSpace(args.InputResponsesJSON) ||
            !string.IsNullOrWhiteSpace(args.FileDialogResponsesJSON) ||
            args.UIStreamEnabled)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "run_args_invalid",
                Message: "Scripted XlflowUI responses and UI stream are not supported by xlflow run --bridge dotnet in this partial parity release; use --bridge powershell for those options.",
                Phase: "run",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
