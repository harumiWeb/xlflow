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
        var projectRoot = BridgePayload.GetString(request.Payload, "ProjectRoot") ?? "";
        var baseWorkbookPath = BridgePayload.GetString(request.Payload, "BaseWorkbookPath") ?? "";
        var planJson64 = BridgePayload.GetString(request.Payload, "PlanJson64") ?? "";
        var outputWorkbookPath = BridgePayload.GetString(request.Payload, "OutputWorkbookPath") ?? "";
        if (string.IsNullOrWhiteSpace(projectRoot) || string.IsNullOrWhiteSpace(baseWorkbookPath) || string.IsNullOrWhiteSpace(outputWorkbookPath) || string.IsNullOrWhiteSpace(planJson64))
        {
            return BridgeResponse.Failed(request, new BridgeError("build_args_invalid", "ProjectRoot, BaseWorkbookPath, OutputWorkbookPath, and PlanJson64 are required", "build", "xlflow-excel-bridge"));
        }

        return _service.Execute(request, new BuildCommandArguments(
            projectRoot,
            baseWorkbookPath,
            outputWorkbookPath,
            planJson64,
            BridgePayload.GetString(request.Payload, "CodeSource") ?? "",
            BridgePayload.GetBool(request.Payload, "Visible"),
            BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            BridgePayload.GetString(request.Payload, "SessionWorkbookPath") ?? ""), cancellationToken);
    }
}
