using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class ProcessCommand : ICommandHandler
{
    private readonly IProcessService _service;

    public ProcessCommand(IProcessService? service = null)
    {
        _service = service ?? new ExcelProcessService();
    }

    public string CommandName => "process";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var action = BridgePayload.GetString(request.Payload, "Action");
        action = string.IsNullOrWhiteSpace(action) ? "list" : action.Trim().ToLowerInvariant();

        if (action is not ("list" or "cleanup"))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "process_args_invalid",
                Message: "Action must be list or cleanup",
                Phase: "process",
                Source: "xlflow"));
        }

        var args = new ProcessCommandArguments(
            Action: action,
            TargetPid: BridgePayload.GetNullableInt(request.Payload, "TargetPid"),
            Auto: BridgePayload.GetBool(request.Payload, "Auto"),
            All: BridgePayload.GetBool(request.Payload, "All"),
            SkipWorkbookProbePids: BridgePayload.GetIntSet(request.Payload, "SkipWorkbookProbePids"));

        return _service.Execute(request, args, cancellationToken);
    }
}
