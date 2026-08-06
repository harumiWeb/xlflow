namespace Xlflow.ExcelBridge.Services;

public sealed record VbeOracleCommandArguments(string PlanJson64, TimeSpan Timeout)
{
    // FixtureJson64 remains exposed for compatibility with the bridge service;
    // the standalone Go runner sends PlanJson64.
    public string FixtureJson64 => PlanJson64;
}

public interface IVbeOracleService
{
    Contract.BridgeResponse Execute(Contract.BridgeRequest request, VbeOracleCommandArguments args, CancellationToken cancellationToken);
}
