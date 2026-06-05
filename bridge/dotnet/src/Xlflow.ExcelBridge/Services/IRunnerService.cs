using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IRunnerService
{
    BridgeResponse Execute(BridgeRequest request, RunnerCommandArguments args, CancellationToken cancellationToken);
}
