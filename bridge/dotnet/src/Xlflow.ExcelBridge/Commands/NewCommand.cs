using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class NewCommand : ICommandHandler
{
    private readonly INewService _service;

    public NewCommand(INewService? service = null)
    {
        _service = service ?? new ExcelNewService();
    }

    public string CommandName => "new";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var workbookPath = BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "";
        if (string.IsNullOrWhiteSpace(workbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "new_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "new",
                Source: "xlflow-excel-bridge"));
        }

        return _service.Execute(request, new NewCommandArguments(workbookPath), cancellationToken);
    }
}
