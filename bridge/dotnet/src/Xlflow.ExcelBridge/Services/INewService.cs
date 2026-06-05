using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface INewService
{
    BridgeResponse Execute(BridgeRequest request, NewCommandArguments args, CancellationToken cancellationToken);
}
