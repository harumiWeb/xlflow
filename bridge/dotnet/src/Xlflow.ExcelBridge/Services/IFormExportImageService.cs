using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IFormExportImageService
{
    BridgeResponse Execute(BridgeRequest request, FormExportImageCommandArguments args, CancellationToken cancellationToken);
}
