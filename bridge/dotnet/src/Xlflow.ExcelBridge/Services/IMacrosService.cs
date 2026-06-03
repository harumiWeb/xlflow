using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IMacrosService
{
    BridgeResponse Execute(BridgeRequest request, MacrosCommandArguments args, CancellationToken cancellationToken);
}
