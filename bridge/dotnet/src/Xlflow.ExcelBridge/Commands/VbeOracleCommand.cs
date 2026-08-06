using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

/// <summary>
/// Developer-only bridge command used by the local VBE oracle harness.
/// The public xlflow command registry intentionally does not expose this
/// command; the standalone oracle executable calls the bridge directly.
/// </summary>
public sealed class VbeOracleCommand : ICommandHandler
{
    private readonly IVbeOracleService _service;

    public VbeOracleCommand(IVbeOracleService? service = null)
    {
        _service = service ?? new VbeOracleService();
    }

    public string CommandName => "vbe-oracle";

    public bool Supports(BridgeRequest request) =>
        string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var planJson64 = BridgePayload.GetString(request.Payload, "FixtureJSON64")
            ?? BridgePayload.GetString(request.Payload, "PlanJson64")
            ?? "";
        if (string.IsNullOrWhiteSpace(planJson64))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                "vbe_oracle_args_invalid",
                "PlanJson64 is required",
                "vbe-oracle",
                "xlflow-excel-bridge"));
        }

        var timeoutMilliseconds = BridgePayload.GetInt(
            request.Payload,
            "TimeoutMS",
            BridgePayload.GetInt(request.Payload, "TimeoutMs", 300_000));
        if (timeoutMilliseconds <= 0)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                "vbe_oracle_args_invalid",
                "TimeoutMs must be greater than zero",
                "vbe-oracle",
                "xlflow-excel-bridge"));
        }

        return _service.Execute(
            request,
            new VbeOracleCommandArguments(planJson64, TimeSpan.FromMilliseconds(timeoutMilliseconds)),
            cancellationToken);
    }
}
