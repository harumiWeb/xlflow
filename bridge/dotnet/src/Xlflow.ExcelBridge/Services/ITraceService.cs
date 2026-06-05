using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface ITraceService
{
    BridgeResponse Execute(BridgeRequest request, TraceCommandArguments args, CancellationToken cancellationToken);
}
