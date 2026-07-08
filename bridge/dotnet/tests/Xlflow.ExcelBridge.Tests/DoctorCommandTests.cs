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
            VbProjectAccess: true,
            VbProjectAccessError: null,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: true, SysWow64: true));

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
        Assert.True(excel.GetProperty("vbproject_access").GetBoolean());
        Assert.Equal(JsonValueKind.Null, excel.GetProperty("trust_vba_access").ValueKind);
        var systemProfileDesktop = excel.GetProperty("systemprofile_desktop");
        Assert.Equal("exists", systemProfileDesktop.GetProperty("system32").GetProperty("status").GetString());
        Assert.Equal("exists", systemProfileDesktop.GetProperty("syswow64").GetProperty("status").GetString());
        Assert.True(systemProfileDesktop.GetProperty("ok").GetBoolean());
    }

    [Fact]
    public void HandleReturnsFailedResponseWhenComActivationFails()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: false,
            Version: null,
            Build: null,
            VbideAccess: null,
            VbProjectAccess: null,
            VbProjectAccessError: null,
            AutomationSecurity: null,
            TrustVbaAccess: null,
            Error: "COM activation failed: HRESULT 0x80040154: Class not registered",
            ComErrorNumber: -2147221164,
            ComHResult: "0x80040154",
            ComDetails: new Dictionary<string, object?>
            {
                ["source"] = "test",
                ["stack_trace"] = "at TestMethod",
            },
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: true, SysWow64: true));

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
        var systemProfileDesktop = json.RootElement.GetProperty("diagnostics").GetProperty("excel").GetProperty("systemprofile_desktop");
        Assert.True(systemProfileDesktop.GetProperty("ok").GetBoolean());
    }

    [Fact]
    public void HandleIncludesNonFatalDiagnosticWarningWhenPresent()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: false,
            VbProjectAccess: false,
            VbProjectAccessError: "VBProject access is denied.",
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: "VBIDE access unavailable",
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: true, SysWow64: true));

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
        Assert.False(excel.GetProperty("vbproject_access").GetBoolean());
        Assert.Equal("VBProject access is denied.", excel.GetProperty("vbproject_access_error").GetString());
        Assert.Equal("VBIDE access unavailable", excel.GetProperty("error").GetString());
    }

    [Fact]
    public void HandleReturnsFailedResponseWhenSystemProfileDesktopDirectoriesAreMissing()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            VbProjectAccess: true,
            VbProjectAccessError: null,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(
                System32: new SystemProfileDesktopPathDiagnostics(@"D:\Windows\System32\config\systemprofile\Desktop", SystemProfileDesktopStatus.Missing),
                SysWow64: new SystemProfileDesktopPathDiagnostics(@"D:\Windows\SysWOW64\config\systemprofile\Desktop", SystemProfileDesktopStatus.Exists)));

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-systemprofile",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        var error = json.RootElement.GetProperty("error");
        Assert.Equal("systemprofile_desktop_missing", error.GetProperty("code").GetString());
        Assert.Equal("doctor", error.GetProperty("phase").GetString());
        Assert.Contains("systemprofile Desktop directories are missing", error.GetProperty("message").GetString());
        Assert.Contains(@"D:\Windows\System32\config\systemprofile\Desktop", error.GetProperty("message").GetString());
        Assert.Contains(@"D:\Windows\SysWOW64\config\systemprofile\Desktop", error.GetProperty("message").GetString());

        var systemProfileDesktop = json.RootElement.GetProperty("diagnostics").GetProperty("excel").GetProperty("systemprofile_desktop");
        Assert.Equal("missing", systemProfileDesktop.GetProperty("system32").GetProperty("status").GetString());
        Assert.Equal("exists", systemProfileDesktop.GetProperty("syswow64").GetProperty("status").GetString());
        Assert.False(systemProfileDesktop.GetProperty("ok").GetBoolean());
    }

    [Fact]
    public void HandleDoesNotFailWhenSystemProfileDesktopCannotBeInspected()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            VbProjectAccess: true,
            VbProjectAccessError: null,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(
                System32: new SystemProfileDesktopPathDiagnostics(@"C:\Windows\System32\config\systemprofile\Desktop", SystemProfileDesktopStatus.Exists),
                SysWow64: new SystemProfileDesktopPathDiagnostics(@"C:\Windows\SysWOW64\config\systemprofile\Desktop", SystemProfileDesktopStatus.AccessDenied, "Access is denied.")));

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-systemprofile-access-denied",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        var systemProfileDesktop = json.RootElement.GetProperty("diagnostics").GetProperty("excel").GetProperty("systemprofile_desktop");
        Assert.Equal("exists", systemProfileDesktop.GetProperty("system32").GetProperty("status").GetString());
        Assert.Equal("access_denied", systemProfileDesktop.GetProperty("syswow64").GetProperty("status").GetString());
        Assert.True(systemProfileDesktop.GetProperty("access_denied").GetBoolean());
        Assert.False(systemProfileDesktop.GetProperty("missing").GetBoolean());
    }

    [Fact]
    public void HandleChecksWorkbookWhenRequested()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            VbProjectAccess: true,
            VbProjectAccessError: null,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: true, SysWow64: true));

        var command = new DoctorCommand(() => probeResult);
        var missingWorkbookPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "Book.xlsm");
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-workbook",
            Command = "doctor",
            Payload = JsonSerializer.SerializeToElement(new Dictionary<string, string>
            {
                ["CheckWorkbook"] = "true",
                ["WorkbookPath"] = missingWorkbookPath,
            }),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        var error = json.RootElement.GetProperty("error");
        Assert.Equal("workbook_open_failed", error.GetProperty("code").GetString());
        Assert.Equal("doctor.open_workbook", error.GetProperty("phase").GetString());
        Assert.Contains("Configured workbook could not be opened", error.GetProperty("message").GetString());

        var diagnostics = json.RootElement.GetProperty("diagnostics");
        Assert.False(diagnostics.GetProperty("workbook_openable").GetBoolean());
        Assert.Equal(Path.GetFullPath(missingWorkbookPath), json.RootElement.GetProperty("workbook").GetProperty("path").GetString());
    }

    [Fact]
    public void HandleStillChecksWorkbookWhenSystemProfileDesktopDirectoriesAreMissing()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: true,
            VbProjectAccess: true,
            VbProjectAccessError: null,
            AutomationSecurity: 1,
            TrustVbaAccess: null,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: false, SysWow64: false));

        var command = new DoctorCommand(() => probeResult);
        var missingWorkbookPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "Book.xlsm");
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-systemprofile-workbook",
            Command = "doctor",
            Payload = JsonSerializer.SerializeToElement(new Dictionary<string, string>
            {
                ["CheckWorkbook"] = "true",
                ["WorkbookPath"] = missingWorkbookPath,
            }),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("systemprofile_desktop_missing", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(Path.GetFullPath(missingWorkbookPath), json.RootElement.GetProperty("workbook").GetProperty("path").GetString());

        var diagnostics = json.RootElement.GetProperty("diagnostics");
        Assert.False(diagnostics.GetProperty("workbook_openable").GetBoolean());
        var systemProfileDesktop = diagnostics.GetProperty("excel").GetProperty("systemprofile_desktop");
        Assert.False(systemProfileDesktop.GetProperty("ok").GetBoolean());
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

    [Fact]
    public void HandleReportsTrustVbaAccessWhenProbeProvidesValue()
    {
        var probeResult = new ExcelDiagnosticsResult(
            ComActivation: true,
            Version: "16.0",
            Build: "17726.20160",
            VbideAccess: false,
            VbProjectAccess: false,
            VbProjectAccessError: "VBProject access is denied.",
            AutomationSecurity: 1,
            TrustVbaAccess: false,
            Error: null,
            SystemProfileDesktop: new SystemProfileDesktopDiagnostics(System32: true, SysWow64: true));

        var command = new DoctorCommand(() => probeResult);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-doctor-trust-vba-access",
            Command = "doctor",
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        var excel = json.RootElement.GetProperty("diagnostics").GetProperty("excel");
        Assert.False(excel.GetProperty("trust_vba_access").GetBoolean());
        Assert.False(excel.GetProperty("vbide_access").GetBoolean());
    }

    [Fact]
    public void ReadTrustVbaAccessFromRegistryPrefersPolicyValue()
    {
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            [@"Software\Policies\Microsoft\Office\16.0\Excel\Security:AccessVBOM"] = 0,
            [@"Software\Microsoft\Office\16.0\Excel\Security:AccessVBOM"] = 1,
        };

        var access = ExcelDiagnostics.ReadTrustVbaAccessFromRegistry(
            "16.0",
            (key, name) => values.TryGetValue(key + ":" + name, out var value) ? value : null);

        Assert.False(access);
    }

    [Fact]
    public void ReadTrustVbaAccessFromRegistryFallsBackToUserValueAndUnknown()
    {
        var enabled = ExcelDiagnostics.ReadTrustVbaAccessFromRegistry(
            "15.1",
            (key, name) => key == @"Software\Microsoft\Office\15.0\Excel\Security" && name == "AccessVBOM" ? "1" : null);
        var unknown = ExcelDiagnostics.ReadTrustVbaAccessFromRegistry("16.0", (_, _) => null);
        var inaccessible = ExcelDiagnostics.ReadTrustVbaAccessFromRegistry("16.0", (_, _) => throw new UnauthorizedAccessException());

        Assert.True(enabled);
        Assert.Null(unknown);
        Assert.Null(inaccessible);
    }
}
