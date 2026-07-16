using System.Text.Json.Serialization;

namespace Xlflow.ExcelBridge.Contract;

public sealed record BridgeResponse
{
    [JsonPropertyName("protocol_version")]
    public int ProtocolVersion { get; init; } = Contract.ProtocolVersion.Current;

    [JsonPropertyName("request_id")]
    public string RequestId { get; init; } = "";

    [JsonPropertyName("status")]
    public string Status { get; init; } = BridgeStatus.Ok;

    [JsonPropertyName("command")]
    public string Command { get; init; } = "";

    [JsonPropertyName("logs")]
    public IReadOnlyList<string> Logs { get; init; } = Array.Empty<string>();

    [JsonPropertyName("error")]
    public BridgeError? Error { get; init; }

    [JsonPropertyName("bridge")]
    public BridgeInfo Bridge { get; init; } = BridgeInfo.Current();

    [JsonPropertyName("recovery")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BridgeRecovery? Recovery { get; init; }

    [JsonExtensionData]
    public IDictionary<string, object?> Extensions { get; init; } = new Dictionary<string, object?>();

    public static BridgeResponse Ok(BridgeRequest request, IDictionary<string, object?>? extensions = null)
    {
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Extensions = extensions ?? new Dictionary<string, object?>(),
        };
    }

    public static BridgeResponse Failed(BridgeRequest request, BridgeError error)
    {
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = error,
        };
    }
}
