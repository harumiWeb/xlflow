using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

// Internal bridge command. The public Go CLI deliberately does not call this
// until the publication and recovery stages are implemented.
public sealed class BuildCommand : ICommandHandler
{
    private readonly IBuildService _service;

    public BuildCommand(IBuildService? service = null) => _service = service ?? new ExcelBuildService();

    public string CommandName => "build";

    public bool Supports(BridgeRequest request) => string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var baseWorkbookPath = BridgePayload.GetString(request.Payload, "BaseWorkbookPath") ?? "";
        var temporaryDirectory = BridgePayload.GetString(request.Payload, "TemporaryDirectory") ?? "";
        var planJson64 = BridgePayload.GetString(request.Payload, "PlanJson64") ?? "";
        if (string.IsNullOrWhiteSpace(baseWorkbookPath) || string.IsNullOrWhiteSpace(temporaryDirectory) || string.IsNullOrWhiteSpace(planJson64))
        {
            return BridgeResponse.Failed(request, new BridgeError("build_args_invalid", "BaseWorkbookPath, TemporaryDirectory, and PlanJson64 are required", "build", "xlflow-excel-bridge"));
        }

        return _service.Execute(request, new BuildCommandArguments(
            baseWorkbookPath,
            temporaryDirectory,
            planJson64,
            BridgePayload.GetString(request.Payload, "CodeSource") ?? "",
            BridgePayload.GetBool(request.Payload, "Visible")), cancellationToken);
    }
}
