using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Commands;

public interface ICommandHandler
{
    string CommandName { get; }

    bool Supports(BridgeRequest request);

    BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken);
}
