using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Diagnostics;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This bridge only runs on Windows where Excel COM is available.")]
public sealed class DoctorCommand : ICommandHandler
{
    private readonly Func<ExcelDiagnosticsResult>? _probeExcel;

    public DoctorCommand(Func<ExcelDiagnosticsResult>? probeExcel = null)
    {
        _probeExcel = probeExcel;
    }

    public string CommandName => "doctor";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var excelProbe = (_probeExcel ?? ExcelDiagnostics.Probe)();
        var diagnostics = BuildDiagnostics(excelProbe);
        var checkWorkbook = BridgePayload.GetBool(request.Payload, "CheckWorkbook");
        var workbookPath = BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "";
        var visible = BridgePayload.GetBool(request.Payload, "Visible");

        if (!excelProbe.ComActivation)
        {
            return FailedWithDiagnostics(request, diagnostics, new BridgeError(
                Code: "excel_com_failure",
                Message: excelProbe.Error ?? "Excel COM activation failed",
                Phase: "doctor",
                Source: "xlflow-excel-bridge",
                Number: excelProbe.ComErrorNumber,
                HResult: excelProbe.ComHResult,
                Details: excelProbe.ComDetails));
        }

        var extensions = new Dictionary<string, object?>
        {
            ["diagnostics"] = diagnostics,
        };

        WorkbookOpenDiagnostics? workbookProbe = null;
        if (checkWorkbook)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                return FailedWithDiagnostics(request, diagnostics, new BridgeError(
                    Code: "doctor_args_invalid",
                    Message: "WorkbookPath is required when CheckWorkbook is true",
                    Phase: "doctor",
                    Source: "xlflow-excel-bridge"));
            }

            extensions["workbook"] = new Dictionary<string, object?>
            {
                ["path"] = ExcelBridgeSupport.NormalizePath(workbookPath),
            };
            workbookProbe = ProbeWorkbookOpen(workbookPath, visible);
            diagnostics["workbook_openable"] = workbookProbe.Openable;
        }

        if (excelProbe.SystemProfileDesktop is { Missing: true } systemProfileDesktop)
        {
            return FailedWithDiagnostics(request, diagnostics, new BridgeError(
                Code: "systemprofile_desktop_missing",
                Message: MissingSystemProfileDesktopMessage(systemProfileDesktop),
                Phase: "doctor",
                Source: "xlflow-excel-bridge",
                Details: new Dictionary<string, object?>
                {
                    ["system32"] = BuildSystemProfileDesktopPathPayload(systemProfileDesktop.System32),
                    ["syswow64"] = BuildSystemProfileDesktopPathPayload(systemProfileDesktop.SysWow64),
                }),
                extensions);
        }

        if (workbookProbe is { Openable: false })
        {
            return FailedWithDiagnostics(request, diagnostics, new BridgeError(
                Code: "workbook_open_failed",
                Message: "Configured workbook could not be opened: " + workbookProbe.Message,
                Phase: "doctor.open_workbook",
                Source: "xlflow-excel-bridge",
                Number: workbookProbe.Number,
                HResult: workbookProbe.HResult,
                Details: new Dictionary<string, object?>
                {
                    ["path"] = ExcelBridgeSupport.NormalizePath(workbookPath),
                }),
                extensions);
        }

        return BridgeResponse.Ok(request, extensions);
    }

    private static Dictionary<string, object?> BuildDiagnostics(ExcelDiagnosticsResult excelProbe)
    {
        var excel = new Dictionary<string, object?>
        {
            ["com_activation"] = excelProbe.ComActivation,
            ["version"] = excelProbe.Version,
            ["build"] = excelProbe.Build,
            ["vbide_access"] = excelProbe.VbideAccess,
            ["automation_security"] = excelProbe.AutomationSecurity,
            ["trust_vba_access"] = excelProbe.TrustVbaAccess,
        };
        if (excelProbe.Error is not null)
        {
            excel["error"] = excelProbe.Error;
        }
        if (excelProbe.SystemProfileDesktop is not null)
        {
            excel["systemprofile_desktop"] = new Dictionary<string, object?>
            {
                ["system32"] = BuildSystemProfileDesktopPathPayload(excelProbe.SystemProfileDesktop.System32),
                ["syswow64"] = BuildSystemProfileDesktopPathPayload(excelProbe.SystemProfileDesktop.SysWow64),
                ["ok"] = excelProbe.SystemProfileDesktop.Ok,
                ["missing"] = excelProbe.SystemProfileDesktop.Missing,
                ["access_denied"] = excelProbe.SystemProfileDesktop.AccessDenied,
            };
        }

        return new Dictionary<string, object?>
        {
            ["requested_bridge"] = "dotnet",
            ["selected_bridge"] = "dotnet",
            ["fallback"] = false,
            ["legacy"] = false,
            ["protocol_version"] = ProtocolVersion.Current,
            ["runtime"] = new Dictionary<string, object?>
            {
                ["os"] = Environment.OSVersion.ToString(),
                ["process_architecture"] = RuntimeInformation.ProcessArchitecture.ToString(),
                ["dotnet_runtime"] = RuntimeInformation.FrameworkDescription,
            },
            ["excel"] = excel,
        };
    }

    private static BridgeResponse FailedWithDiagnostics(
        BridgeRequest request,
        Dictionary<string, object?> diagnostics,
        BridgeError error,
        Dictionary<string, object?>? extensions = null)
    {
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = error,
            Extensions = extensions ?? new Dictionary<string, object?>
            {
                ["diagnostics"] = diagnostics,
            },
        };
    }

    private static Dictionary<string, object?> BuildSystemProfileDesktopPathPayload(SystemProfileDesktopPathDiagnostics path)
    {
        var payload = new Dictionary<string, object?>
        {
            ["path"] = path.Path,
            ["status"] = path.Status,
        };
        if (!string.IsNullOrWhiteSpace(path.Message))
        {
            payload["message"] = path.Message;
        }
        return payload;
    }

    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Workbook doctor diagnostics normalize all open/close failures into structured output.")]
    private static WorkbookOpenDiagnostics ProbeWorkbookOpen(string workbookPath, bool visible)
    {
        ExcelSessionAttachment? attachment = null;
        try
        {
            attachment = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
            return new WorkbookOpenDiagnostics(Openable: true, Message: null, Number: null, HResult: null);
        }
        catch (Exception ex)
        {
            var failure = ExcelBridgeSupport.ClassifyComFailure(ex);
            return new WorkbookOpenDiagnostics(
                Openable: false,
                Message: failure.Message,
                Number: failure.Number,
                HResult: failure.HResult);
        }
        finally
        {
            if (attachment is not null)
            {
                try { ExcelBridgeSupport.InvokeViaDynamic(attachment.Workbook, "Close", false); } catch { }
                try { ExcelBridgeSupport.InvokeViaDynamic(attachment.Excel, "Quit"); } catch { }
                ExcelBridgeSupport.ReleaseComObject(attachment.Workbook);
                ExcelBridgeSupport.ReleaseComObject(attachment.Excel);
            }
        }
    }

    private static string MissingSystemProfileDesktopMessage(SystemProfileDesktopDiagnostics systemProfileDesktop)
    {
        return $"""
            systemprofile Desktop directories are missing.
            Create both directories:
            - {systemProfileDesktop.System32.Path}
            - {systemProfileDesktop.SysWow64.Path}

            This is required for Excel COM automation in non-interactive sessions such as SSH, services, or CI.
            """;
    }

    private sealed record WorkbookOpenDiagnostics(
        bool Openable,
        string? Message,
        int? Number,
        string? HResult);
}
