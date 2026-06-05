using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IListService
{
    BridgeResponse Execute(BridgeRequest request, ListCommandArguments args, CancellationToken cancellationToken);
}
