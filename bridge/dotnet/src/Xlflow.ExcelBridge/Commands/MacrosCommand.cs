using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class MacrosCommand : ICommandHandler
{
    private readonly IMacrosService _service;

    public MacrosCommand(IMacrosService? service = null)
    {
        _service = service ?? new ExcelMacrosService();
    }

    public string CommandName => "macros";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new MacrosCommandArguments(
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "",
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            Entry: BridgePayload.GetString(request.Payload, "Entry") ?? "",
            RunnableOnly: BridgePayload.GetBool(request.Payload, "RunnableOnly"));

        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "macros_args_invalid",
                Message: "WorkbookPath is required",
                Phase: "macros",
                Source: "xlflow"));
        }

        return _service.Execute(request, args, cancellationToken);
    }
}
