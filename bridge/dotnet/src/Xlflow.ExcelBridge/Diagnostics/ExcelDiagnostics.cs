using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Xlflow.ExcelBridge.Diagnostics;

public sealed record ExcelDiagnosticsResult(
    bool ComActivation,
    string? Version,
    string? Build,
    bool? VbideAccess,
    object? AutomationSecurity,
    bool? TrustVbaAccess,
    string? Error,
    int? ComErrorNumber = null,
    string? ComHResult = null,
    IReadOnlyDictionary<string, object?>? ComDetails = null);

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This bridge only runs on Windows where Excel COM is available.")]
public static class ExcelDiagnostics
{
    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "COM diagnostics must not propagate exceptions to the caller.")]
    public static ExcelDiagnosticsResult Probe()
    {
        object? app = null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType is null)
            {
                return new ExcelDiagnosticsResult(
                    ComActivation: false,
                    Version: null,
                    Build: null,
                    VbideAccess: null,
                    AutomationSecurity: null,
                    TrustVbaAccess: null,
                    Error: "Excel.Application ProgID not registered");
            }

            try
            {
                app = Activator.CreateInstance(excelType);
            }
            catch (COMException ex)
            {
                return new ExcelDiagnosticsResult(
                    ComActivation: false,
                    Version: null,
                    Build: null,
                    VbideAccess: null,
                    AutomationSecurity: null,
                    TrustVbaAccess: null,
                    Error: $"COM activation failed: {FormatComError(ex)}",
                    ComErrorNumber: ex.ErrorCode,
                    ComHResult: $"0x{ex.HResult:X8}",
                    ComDetails: new Dictionary<string, object?>
                    {
                        ["source"] = ex.Source,
                        ["stack_trace"] = ex.StackTrace,
                    });
            }
            catch (Exception ex)
            {
                return new ExcelDiagnosticsResult(
                    ComActivation: false,
                    Version: null,
                    Build: null,
                    VbideAccess: null,
                    AutomationSecurity: null,
                    TrustVbaAccess: null,
                    Error: $"COM activation failed: {ex.Message}");
            }

            if (app is null)
            {
                return new ExcelDiagnosticsResult(
                    ComActivation: false,
                    Version: null,
                    Build: null,
                    VbideAccess: null,
                    AutomationSecurity: null,
                    TrustVbaAccess: null,
                    Error: "COM activation returned null");
            }

            string? version = null;
            string? build = null;
            try
            {
                version = InvokeStringProperty(app, "Version");
                build = InvokeProperty(app, "Build")?.ToString();
            }
            catch (COMException)
            {
                // version/build unavailable
            }

            object? automationSecurity = null;
            try
            {
                automationSecurity = InvokeProperty(app, "AutomationSecurity");
            }
            catch (COMException)
            {
                // AutomationSecurity unavailable
            }

            bool? vbideAccess = null;
            bool? trustVbaAccess = null;
            try
            {
                var vbe = InvokeProperty(app, "VBE");
                vbideAccess = vbe is not null;
                // trustVbaAccess is left null: "Trust access to VBA project object model"
                // is a separate Trust Center setting that cannot be observed via VBE property.
                if (vbe is not null)
                {
                    Marshal.ReleaseComObject(vbe);
                }
            }
            catch (COMException)
            {
                vbideAccess = false;
            }
            catch (UnauthorizedAccessException)
            {
                vbideAccess = false;
            }

            return new ExcelDiagnosticsResult(
                ComActivation: true,
                Version: version,
                Build: build,
                VbideAccess: vbideAccess,
                AutomationSecurity: automationSecurity,
                TrustVbaAccess: trustVbaAccess,
                Error: null);
        }
        finally
        {
            if (app is not null)
            {
                try
                {
                    app.GetType().InvokeMember("Quit",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null, app, null, CultureInfo.InvariantCulture);
                }
                catch
                {
                    // best-effort cleanup
                }
                try
                {
                    Marshal.ReleaseComObject(app);
                }
                catch
                {
                    // best-effort cleanup
                }
            }
        }
    }

    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Reflection on COM objects may throw various exceptions.")]
    private static object? InvokeProperty(object comObject, string name)
    {
        try
        {
            return comObject.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.InvokeMethod,
                null,
                comObject,
                null,
                CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Reflection on COM objects may throw various exceptions.")]
    private static string? InvokeStringProperty(object comObject, string name)
    {
        try
        {
            return comObject.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.InvokeMethod,
                null,
                comObject,
                null,
                CultureInfo.InvariantCulture) as string;
        }
        catch
        {
            return null;
        }
    }

    private static string FormatComError(COMException ex)
    {
        if (ex.HResult != 0)
        {
            return $"HRESULT 0x{ex.HResult:X8}: {ex.Message}";
        }
        return ex.Message;
    }
}
