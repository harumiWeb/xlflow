using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IUIService
{
    BridgeResponse Execute(BridgeRequest request, UICommandArguments args, CancellationToken cancellationToken);
}
