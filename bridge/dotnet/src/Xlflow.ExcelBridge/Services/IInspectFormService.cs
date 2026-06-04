using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IInspectFormService
{
    BridgeResponse Execute(BridgeRequest request, InspectFormCommandArguments args, CancellationToken cancellationToken);
}
