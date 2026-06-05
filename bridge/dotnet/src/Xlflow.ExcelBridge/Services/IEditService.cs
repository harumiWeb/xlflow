using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IEditService
{
    BridgeResponse Execute(BridgeRequest request, EditCommandArguments args, CancellationToken cancellationToken);
}
