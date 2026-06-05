using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class RunnerCommand : ICommandHandler
{
    private readonly IRunnerService _service;

    public RunnerCommand(IRunnerService? service = null)
    {
        _service = service ?? new ExcelRunnerService();
    }

    public string CommandName => "runner";

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
                Code: "runner_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "runner",
                Source: "xlflow-excel-bridge"));
        }

        var args = new RunnerCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: workbookPath,
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");
        return _service.Execute(request, args, cancellationToken);
    }
}
