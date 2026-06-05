using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Diagnostics;
using Xlflow.ExcelBridge.Serialization;

namespace Xlflow.ExcelBridge.Tests;

public sealed class DoctorCommandTests
{
    [Fact]
    public void HandleReturnsSelectedBridgeAndRuntimeDiagnostics()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null);

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("doctor", json.RootElement.GetProperty("command").GetString());

        var diagnostics = json.RootElement.GetProperty("diagnostics");
        Assert.Equal("dotnet", diagnostics.GetProperty("requested_bridge").GetString());
        Assert.Equal("dotnet", diagnostics.GetProperty("selected_bridge").GetString());
        Assert.False(diagnostics.GetProperty("fallback").GetBoolean());
        Assert.False(diagnostics.GetProperty("legacy").GetBoolean());
        Assert.Equal(ProtocolVersion.Current, diagnostics.GetProperty("protocol_version").GetInt32());

        var runtime = diagnostics.GetProperty("runtime");
        Assert.False(string.IsNullOrWhiteSpace(runtime.GetProperty("os").GetString()));
        Assert.Equal(System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(), runtime.GetProperty("process_architecture").GetString());
        Assert.Equal(System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription, runtime.GetProperty("dotnet_runtime").GetString());

        var excel = diagnostics.GetProperty("excel");
        Assert.True(excel.GetProperty("com_activation").GetBoolean());
        Assert.Equal("16.0", excel.GetProperty("version").GetString());
        Assert.Equal("17726.20160", excel.GetProperty("build").GetString());
        Assert.True(excel.GetProperty("vbide_access").GetBoolean());
    }

    [Fact]
    public void HandleReturnsFailedResponseWhenComActivationFails()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: false,
            Version: null,
            Build: null,
            VbideAccess: null,
            AutomationSecurity: null,
            TrustVbaAccess: null,
            Error: "COM activation failed: HRESULT 0x80040154: Class not registered",
            ComErrorNumber: -2147221164,
            ComHResult: "0x80040154",
            ComDetails: new Dictionary<string, object?>
            {
                ["source"] = "test",
                ["stack_trace"] = "at TestMethod",
            });

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-com-failure",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("doctor", json.RootElement.GetProperty("command").GetString());

        var error = json.RootElement.GetProperty("error");
        Assert.Equal("excel_com_failure", error.GetProperty("code").GetString());
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("message").GetString()));
        Assert.Equal("doctor", error.GetProperty("phase").GetString());
        Assert.Equal("xlflow-excel-bridge", error.GetProperty("source").GetString());
        Assert.Equal(-2147221164, error.GetProperty("number").GetInt32());
        Assert.Equal("0x80040154", error.GetProperty("h_result").GetString());
        Assert.Equal(JsonValueKind.Object, error.GetProperty("details").ValueKind);
    }

    [Fact]
    public void HandleIncludesNonFatalDiagnosticWarningWhenPresent()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: false,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: "VBIDE access unavailable");

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-warning",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());

        var excel = json.RootElement.GetProperty("diagnostics").GetProperty("excel");
        Assert.True(excel.GetProperty("com_activation").GetBoolean());
        Assert.False(excel.GetProperty("vbide_access").GetBoolean());
        Assert.Equal("VBIDE access unavailable", excel.GetProperty("error").GetString());
    }

    [Fact]
    public void BridgeErrorSerializesStructuredFields()
    {
        var error = new BridgeError(
            Code: "excel_com_failure",
            Message: "COM activation failed",
            Phase: "doctor",
            Source: "xlflow-excel-bridge",
            Number: -2147221164,
            HResult: "0x80040154",
            Details: new Dictionary<string, object?>
            {
                ["source"] = "test",
            });

        var json = JsonSerializer.SerializeToDocument(error, JsonOptions.Default);
        var root = json.RootElement;

        Assert.Equal("excel_com_failure", root.GetProperty("code").GetString());
        Assert.Equal("doctor", root.GetProperty("phase").GetString());
        Assert.Equal("xlflow-excel-bridge", root.GetProperty("source").GetString());
        Assert.Equal(-2147221164, root.GetProperty("number").GetInt32());
        Assert.Equal("0x80040154", root.GetProperty("h_result").GetString());
        Assert.Equal(JsonValueKind.Object, root.GetProperty("details").ValueKind);
        Assert.Equal("test", root.GetProperty("details").GetProperty("source").GetString());
    }
}
