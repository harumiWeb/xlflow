using System.Text.Json;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Tests;

public sealed class BridgeHostTests
{
    [Fact]
    public void VersionJsonWritesBridgeInfo()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run(["--version-json"], TextReader.Null, stdout, stderr);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("xlflow-excel-bridge", json.RootElement.GetProperty("name").GetString());
        Assert.Equal(ProtocolVersion.Current, json.RootElement.GetProperty("protocol_version").GetInt32());
    }

    [Fact]
    public void CapabilitiesJsonIncludesDoctor()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run(["--capabilities-json"], TextReader.Null, stdout, stderr);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        var commands = json.RootElement.GetProperty("commands").EnumerateArray().Select(item => item.GetString()).ToArray();
        Assert.Contains("doctor", commands);
    }

    [Fact]
    public void StdinDoctorRequestWritesBridgeResponse()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-1",
              "command": "doctor",
              "payload": {}
            }
            """;
        using var stdin = new StringReader(request);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("req-1", json.RootElement.GetProperty("request_id").GetString());
        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("doctor", json.RootElement.GetProperty("command").GetString());
        Assert.True(json.RootElement.TryGetProperty("diagnostics", out _));
    }

    [Fact]
    public void UnsupportedCommandReturnsStructuredErrorWithoutFallback()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-2",
              "command": "run",
              "payload": {}
            }
            """;
        using var stdin = new StringReader(request);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.Equal(3, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("BRIDGE_COMMAND_UNSUPPORTED", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void EmptyStdinReturnsStructuredRequestEmptyError()
    {
        using var stdin = new StringReader(string.Empty);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.Equal(2, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("BRIDGE_REQUEST_EMPTY", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void InvalidJsonReturnsStructuredRequestInvalidJsonError()
    {
        using var stdin = new StringReader("{");
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.Equal(2, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("BRIDGE_REQUEST_INVALID_JSON", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }
}
