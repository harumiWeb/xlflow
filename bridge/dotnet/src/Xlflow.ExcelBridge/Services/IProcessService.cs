using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IProcessService
{
    BridgeResponse Execute(BridgeRequest request, ProcessCommandArguments args, CancellationToken cancellationToken);
}
