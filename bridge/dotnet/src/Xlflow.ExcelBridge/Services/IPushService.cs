using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IPushService
{
    BridgeResponse Execute(BridgeRequest request, PushCommandArguments args, CancellationToken cancellationToken);
}
