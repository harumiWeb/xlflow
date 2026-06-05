using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class TraceCommand : ICommandHandler
{
    private readonly ITraceService _service;

    public TraceCommand(ITraceService? service = null)
    {
        _service = service ?? new ExcelTraceService();
    }

    public string CommandName => "trace";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new TraceCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "enable",
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            ModulesDir: BridgePayload.GetString(request.Payload, "ModulesDir") ?? "",
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            Force: BridgePayload.GetBool(request.Payload, "Force"),
            TraceDir: BridgePayload.GetString(request.Payload, "TraceDir") ?? "",
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        var action = args.Action.Trim().ToLowerInvariant();
        if (action == "inject")
        {
            action = "enable";
        }
        if (action is not ("enable" or "disable" or "status" or "clean"))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "trace_args_invalid",
                Message: $"unsupported trace action: {args.Action}",
                Phase: "trace",
                Source: "xlflow"));
        }
        if (action != "clean" && string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "trace_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "trace",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
