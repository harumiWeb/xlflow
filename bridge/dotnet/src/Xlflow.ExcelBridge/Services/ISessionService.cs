using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface ISessionService
{
    BridgeResponse Execute(BridgeRequest request, SessionCommandArguments args, CancellationToken cancellationToken);
}
