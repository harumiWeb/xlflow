using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IPullService
{
    BridgeResponse Execute(BridgeRequest request, PullCommandArguments args, CancellationToken cancellationToken);
}
