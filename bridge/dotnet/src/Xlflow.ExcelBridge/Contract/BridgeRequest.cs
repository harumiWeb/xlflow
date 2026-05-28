using System.Text.Json;
using System.Text.Json.Serialization;

namespace Xlflow.ExcelBridge.Contract;

public sealed record BridgeRequest
{
    public static BridgeRequest Unknown { get; } = new()
    {
        ProtocolVersion = Contract.ProtocolVersion.Current,
        RequestId = "",
        Command = "unknown",
    };

    [JsonPropertyName("protocol_version")]
    public int ProtocolVersion { get; init; }

    [JsonPropertyName("request_id")]
    public string RequestId { get; init; } = "";

    [JsonPropertyName("command")]
    public string Command { get; init; } = "";

    [JsonPropertyName("timeout_ms")]
    public int? TimeoutMs { get; init; }

    [JsonPropertyName("payload")]
    public JsonElement Payload { get; init; }
}
