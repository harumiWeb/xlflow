namespace Xlflow.ExcelBridge.Services;

public sealed record VbeOracleCommandArguments(string PlanJson64, TimeSpan Timeout)
{
    // FixtureJSON64 is the name used by the standalone Go oracle runner;
    // PlanJson64 remains accepted for compatibility with early prototypes.
    public string FixtureJson64 => PlanJson64;
}

public interface IVbeOracleService
{
    Contract.BridgeResponse Execute(Contract.BridgeRequest request, VbeOracleCommandArguments args, CancellationToken cancellationToken);
}
