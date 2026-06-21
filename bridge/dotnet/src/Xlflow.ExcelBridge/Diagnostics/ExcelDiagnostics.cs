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
    IReadOnlyDictionary<string, object?>? ComDetails = null,
    SystemProfileDesktopDiagnostics? SystemProfileDesktop = null);

public sealed record SystemProfileDesktopDiagnostics(
    SystemProfileDesktopPathDiagnostics System32,
    SystemProfileDesktopPathDiagnostics SysWow64)
{
    public SystemProfileDesktopDiagnostics(bool System32, bool SysWow64)
        : this(
            new SystemProfileDesktopPathDiagnostics("C:\\Windows\\System32\\config\\systemprofile\\Desktop", System32 ? SystemProfileDesktopStatus.Exists : SystemProfileDesktopStatus.Missing),
            new SystemProfileDesktopPathDiagnostics("C:\\Windows\\SysWOW64\\config\\systemprofile\\Desktop", SysWow64 ? SystemProfileDesktopStatus.Exists : SystemProfileDesktopStatus.Missing))
    {
    }

    public bool Ok => System32.Status == SystemProfileDesktopStatus.Exists && SysWow64.Status == SystemProfileDesktopStatus.Exists;
    public bool Missing => System32.Status == SystemProfileDesktopStatus.Missing || SysWow64.Status == SystemProfileDesktopStatus.Missing;
    public bool AccessDenied => System32.Status == SystemProfileDesktopStatus.AccessDenied || SysWow64.Status == SystemProfileDesktopStatus.AccessDenied;
}

public sealed record SystemProfileDesktopPathDiagnostics(
    string Path,
    string Status,
    string? Message = null);

public static class SystemProfileDesktopStatus
{
    public const string Exists = "exists";
    public const string Missing = "missing";
    public const string AccessDenied = "access_denied";
    public const string Unknown = "unknown";
}

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This bridge only runs on Windows where Excel COM is available.")]
public static class ExcelDiagnostics
{
    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "COM diagnostics must not propagate exceptions to the caller.")]
    public static ExcelDiagnosticsResult Probe()
    {
        object? app = null;
        var systemProfileDesktop = ProbeSystemProfileDesktop();
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
                    Error: "Excel.Application ProgID not registered",
                    SystemProfileDesktop: systemProfileDesktop);
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
                    },
                    SystemProfileDesktop: systemProfileDesktop);
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
                    Error: $"COM activation failed: {ex.Message}",
                    SystemProfileDesktop: systemProfileDesktop);
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
                    Error: "COM activation returned null",
                    SystemProfileDesktop: systemProfileDesktop);
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
                Error: null,
                SystemProfileDesktop: systemProfileDesktop);
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

    private static SystemProfileDesktopDiagnostics ProbeSystemProfileDesktop()
    {
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            windows = "C:\\Windows";
        }

        var system32 = Path.Combine(windows, "System32", "config", "systemprofile", "Desktop");
        var sysWow64 = Path.Combine(windows, "SysWOW64", "config", "systemprofile", "Desktop");

        return new SystemProfileDesktopDiagnostics(
            System32: ProbeSystemProfileDesktopPath(system32),
            SysWow64: ProbeSystemProfileDesktopPath(sysWow64));
    }

    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Systemprofile Desktop diagnostics must classify inaccessible or unusual filesystem states without failing doctor.")]
    private static SystemProfileDesktopPathDiagnostics ProbeSystemProfileDesktopPath(string path)
    {
        try
        {
            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.Directory) == FileAttributes.Directory)
            {
                return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.Exists);
            }

            return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.Missing, "Path exists but is not a directory.");
        }
        catch (UnauthorizedAccessException ex)
        {
            return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.AccessDenied, ex.Message);
        }
        catch (DirectoryNotFoundException ex)
        {
            return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.Missing, ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.Missing, ex.Message);
        }
        catch (Exception ex)
        {
            return new SystemProfileDesktopPathDiagnostics(path, SystemProfileDesktopStatus.Unknown, ex.Message);
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
