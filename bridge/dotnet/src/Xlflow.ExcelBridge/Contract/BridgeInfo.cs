using System.Reflection;

namespace Xlflow.ExcelBridge.Contract;

public sealed record BridgeInfo(
    string Name,
    string Version,
    int ProtocolVersion,
    string Commit,
    string Runtime,
    string Architecture)
{
    public static BridgeInfo Current()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version?.ToString(3) ?? "0.1.0";

        return new BridgeInfo(
            "xlflow-excel-bridge",
            version,
            Xlflow.ExcelBridge.Contract.ProtocolVersion.Current,
            ThisAssemblyCommit.Value,
            System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString());
    }
}
