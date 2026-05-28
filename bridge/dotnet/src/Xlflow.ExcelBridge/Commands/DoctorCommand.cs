using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Commands;

public sealed class DoctorCommand : ICommandHandler
{
    public string CommandName => "doctor";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        return BridgeResponse.Ok(request, new Dictionary<string, object?>
        {
            ["diagnostics"] = new Dictionary<string, object?>
            {
                ["bridge"] = BridgeInfo.Current(),
                ["os"] = Environment.OSVersion.ToString(),
                ["architecture"] = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
                ["runtime"] = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            },
        });
    }
}
