using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class TestCommand : ICommandHandler
{
    private readonly ITestService _service;

    public TestCommand(ITestService? service = null)
    {
        _service = service ?? new ExcelTestService();
    }

    public string CommandName => "test";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new TestCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            Filter: BridgePayload.GetString(request.Payload, "Filter") ?? "",
            ModuleFilter: BridgePayload.GetString(request.Payload, "ModuleFilter") ?? "",
            TagFilter: BridgePayload.GetString(request.Payload, "TagFilter") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            RuntimeMode: BridgePayload.GetString(request.Payload, "RuntimeMode") ?? "test",
            RuntimeSource: BridgePayload.GetString(request.Payload, "RuntimeSource") ?? "command",
            MsgBoxResponsesJSON: BridgePayload.GetString(request.Payload, "MsgBoxResponsesJSON") ?? "",
            InputResponsesJSON: BridgePayload.GetString(request.Payload, "InputResponsesJSON") ?? "",
            FileDialogResponsesJSON: BridgePayload.GetString(request.Payload, "FileDialogResponsesJSON") ?? "",
            DebugStreamEnabled: BridgePayload.GetBool(request.Payload, "DebugStreamEnabled"),
            DebugStreamPipeName: BridgePayload.GetString(request.Payload, "DebugStreamPipeName") ?? "",
            UIStreamEnabled: BridgePayload.GetBool(request.Payload, "UIStreamEnabled"),
            UIStreamPipeName: BridgePayload.GetString(request.Payload, "UIStreamPipeName") ?? "",
            UIStreamRedactInput: BridgePayload.GetBool(request.Payload, "UIStreamRedactInput", true),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            Isolation: BridgePayload.GetString(request.Payload, "Isolation") ?? "none",
            NoSave: BridgePayload.GetBool(request.Payload, "NoSave"),
            FailFast: BridgePayload.GetBool(request.Payload, "FailFast"),
            MaxFailures: BridgePayload.GetInt(request.Payload, "MaxFailures"),
            RerunFailed: BridgePayload.GetInt(request.Payload, "RerunFailed"),
            DisableAutoSession: BridgePayload.GetBool(request.Payload, "DisableAutoSession", true),
            SourceWorkbookPath: BridgePayload.GetString(request.Payload, "SourceWorkbookPath") ?? "",
            ProjectRoot: BridgePayload.GetString(request.Payload, "ProjectRoot") ?? "",
            TempRunRoot: BridgePayload.GetString(request.Payload, "TempRunRoot") ?? "",
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "test",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
