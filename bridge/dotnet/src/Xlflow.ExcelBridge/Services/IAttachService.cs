using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IAttachService
{
    BridgeResponse Execute(BridgeRequest request, AttachCommandArguments args, CancellationToken cancellationToken);
}
