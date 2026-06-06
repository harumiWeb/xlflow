using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Diagnostics;

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
    public void CapabilitiesJsonIncludesSupportedCommands()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run(["--capabilities-json"], TextReader.Null, stdout, stderr);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        var commands = json.RootElement.GetProperty("commands").EnumerateArray().Select(item => item.GetString()).ToArray();
        Assert.Contains("attach", commands);
        Assert.Contains("doctor", commands);
        Assert.Contains("edit", commands);
        Assert.Contains("export-image", commands);
        Assert.Contains("form-export-image", commands);
        Assert.Contains("form-write", commands);
        Assert.Contains("inspect", commands);
        Assert.Contains("inspect-form", commands);
        Assert.Contains("list", commands);
        Assert.Contains("macros", commands);
        Assert.Contains("new", commands);
        Assert.Contains("pull", commands);
        Assert.Contains("process", commands);
        Assert.Contains("push", commands);
        Assert.Contains("run", commands);
        Assert.Contains("runner", commands);
        Assert.Contains("session", commands);
        Assert.Contains("test", commands);
        Assert.Contains("ui", commands);
    }

    [Fact]
    public void StdinDoctorRequestWritesBridgeResponse()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null);

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

        var registry = CommandRegistry.Create(() => probeResult);
        var code = BridgeHost.Run([], stdin, stdout, stderr, registry);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("req-1", json.RootElement.GetProperty("request_id").GetString());
        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("doctor", json.RootElement.GetProperty("command").GetString());
        Assert.Equal(ProtocolVersion.Current, json.RootElement.GetProperty("bridge").GetProperty("protocol_version").GetInt32());
        var diagnostics = json.RootElement.GetProperty("diagnostics");
        Assert.Equal("dotnet", diagnostics.GetProperty("requested_bridge").GetString());
        Assert.Equal("dotnet", diagnostics.GetProperty("selected_bridge").GetString());
        Assert.False(diagnostics.GetProperty("fallback").GetBoolean());
        Assert.False(diagnostics.GetProperty("legacy").GetBoolean());
        Assert.Equal(ProtocolVersion.Current, diagnostics.GetProperty("protocol_version").GetInt32());
        var excel = diagnostics.GetProperty("excel");
        Assert.True(excel.GetProperty("com_activation").GetBoolean());
        Assert.Equal("16.0", excel.GetProperty("version").GetString());
        Assert.Equal("17726.20160", excel.GetProperty("build").GetString());
        Assert.True(excel.GetProperty("vbide_access").GetBoolean());
        Assert.Equal(1, excel.GetProperty("automation_security").GetInt32());
        Assert.Equal(JsonValueKind.Null, excel.GetProperty("trust_vba_access").ValueKind);
    }

    [Fact]
    public void UnsupportedCommandReturnsStructuredErrorWithoutFallback()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-2",
              "command": "unknown-command",
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
    public void RunRequestIsHandledByDefaultRegistryWithoutUnsupportedFallback()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-run",
              "command": "run",
              "payload": {}
            }
            """;
        using var stdin = new StringReader(request);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.NotEqual(3, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("run", json.RootElement.GetProperty("command").GetString());
        if (json.RootElement.GetProperty("status").GetString() == "failed")
        {
            Assert.NotEqual("BRIDGE_COMMAND_UNSUPPORTED", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        }
    }

    [Fact]
    public void MacrosRequestIsHandledByDefaultRegistryWithoutUnsupportedFallback()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-macros",
              "command": "macros",
              "payload": {}
            }
            """;
        using var stdin = new StringReader(request);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.NotEqual(3, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("macros", json.RootElement.GetProperty("command").GetString());
        if (json.RootElement.GetProperty("status").GetString() == "failed")
        {
            Assert.NotEqual("BRIDGE_COMMAND_UNSUPPORTED", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        }
    }

    [Fact]
    public void TestRequestIsHandledByDefaultRegistryWithoutUnsupportedFallback()
    {
        const string request = """
            {
              "protocol_version": 1,
              "request_id": "req-test",
              "command": "test",
              "payload": {}
            }
            """;
        using var stdin = new StringReader(request);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeHost.Run([], stdin, stdout, stderr);

        Assert.NotEqual(3, code);
        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("test", json.RootElement.GetProperty("command").GetString());
        if (json.RootElement.GetProperty("status").GetString() == "failed")
        {
            Assert.NotEqual("BRIDGE_COMMAND_UNSUPPORTED", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        }
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
