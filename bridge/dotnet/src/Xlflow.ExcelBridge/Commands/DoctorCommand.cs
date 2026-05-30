using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Diagnostics;

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

        if (!excelProbe.ComActivation)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "excel_com_failure",
                Message: excelProbe.Error ?? "Excel COM activation failed",
                Phase: "doctor",
                Source: "xlflow-excel-bridge",
                Number: excelProbe.ComErrorNumber,
                HResult: excelProbe.ComHResult,
                Details: excelProbe.ComDetails));
        }

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

        return BridgeResponse.Ok(request, new Dictionary<string, object?>
        {
            ["diagnostics"] = new Dictionary<string, object?>
            {
                ["selected_bridge"] = "dotnet",
                ["protocol_version"] = ProtocolVersion.Current,
                ["runtime"] = new Dictionary<string, object?>
                {
                    ["os"] = Environment.OSVersion.ToString(),
                    ["process_architecture"] = RuntimeInformation.ProcessArchitecture.ToString(),
                    ["dotnet_runtime"] = RuntimeInformation.FrameworkDescription,
                },
                ["excel"] = excel,
            },
        });
    }
}
