using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IBuildService
{
    BridgeResponse Execute(BridgeRequest request, BuildCommandArguments args, CancellationToken cancellationToken);
}
