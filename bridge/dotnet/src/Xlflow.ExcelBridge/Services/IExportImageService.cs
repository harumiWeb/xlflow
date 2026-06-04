using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

public interface IExportImageService
{
    BridgeResponse Execute(BridgeRequest request, ExportImageCommandArguments args, CancellationToken cancellationToken);
}
