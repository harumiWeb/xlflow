using System.Text.Json.Serialization;

namespace Xlflow.ExcelBridge.Contract;

public sealed record BridgeRecovery
{
    [JsonPropertyName("required")]
    public bool Required { get; init; }

    [JsonPropertyName("reason")]
    public string Reason { get; init; } = "";

    [JsonPropertyName("operation")]
    public string Operation { get; init; } = "";

    [JsonPropertyName("excel_pid")]
    public int? ExcelProcessId { get; init; }

    [JsonPropertyName("worker_pid")]
    public int? WorkerProcessId { get; init; }

    [JsonPropertyName("cleanup_confirmed")]
    public bool CleanupConfirmed { get; init; }

    [JsonPropertyName("session")]
    public BridgeRecoverySession Session { get; init; } = new();
}

public sealed record BridgeRecoverySession
{
    [JsonPropertyName("active")]
    public bool Active { get; init; }

    [JsonPropertyName("owner")]
    public string Owner { get; init; } = "none";
}
