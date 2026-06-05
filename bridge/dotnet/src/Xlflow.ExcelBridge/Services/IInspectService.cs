using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IInspectService
{
    BridgeResponse Execute(BridgeRequest request, InspectCommandArguments args, CancellationToken cancellationToken);
}
