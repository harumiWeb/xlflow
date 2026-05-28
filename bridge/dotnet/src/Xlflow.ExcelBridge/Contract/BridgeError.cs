namespace Xlflow.ExcelBridge.Contract;

public sealed record BridgeError(
    string Code,
    string Message,
    string Phase,
    string Source,
    int? Number = null,
    string? HResult = null,
    IReadOnlyDictionary<string, object?>? Details = null)
{
    public static BridgeError Create(string code, string message, string phase)
    {
        return new BridgeError(code, message, phase, "xlflow-excel-bridge");
    }
}
