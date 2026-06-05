using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface ITestService
{
    BridgeResponse Execute(BridgeRequest request, TestCommandArguments args, CancellationToken cancellationToken);
}
