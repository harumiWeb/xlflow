using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IRunService
{
    BridgeResponse Execute(BridgeRequest request, RunCommandArguments args, CancellationToken cancellationToken);
}
