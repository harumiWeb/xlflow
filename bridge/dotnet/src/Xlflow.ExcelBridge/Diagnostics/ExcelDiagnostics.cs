using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using Microsoft.Win32;

namespace Xlflow.ExcelBridge.Diagnostics;

public sealed record ExcelDiagnosticsResult(
    bool ComActivation,
    string? Version,
    string? Build,
    bool? VbideAccess,
    bool? VbProjectAccess,
    string? VbProjectAccessError,
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
                    VbProjectAccess: null,
                    VbProjectAccessError: null,
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
                    VbProjectAccess: null,
                    VbProjectAccessError: null,
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
                    VbProjectAccess: null,
                    VbProjectAccessError: null,
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
                    VbProjectAccess: null,
                    VbProjectAccessError: null,
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
            bool? trustVbaAccess = ReadTrustVbaAccessFromRegistry(version, ReadCurrentUserRegistryValue);
            try
            {
                var vbe = InvokeProperty(app, "VBE");
                vbideAccess = vbe is not null;
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

            var vbProjectProbe = ProbeTemporaryWorkbookVbProjectAccess(app);

            return new ExcelDiagnosticsResult(
                ComActivation: true,
                Version: version,
                Build: build,
                VbideAccess: vbideAccess,
                VbProjectAccess: vbProjectProbe.Accessible,
                VbProjectAccessError: vbProjectProbe.Message,
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

    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "VBProject diagnostics must normalize localized COM/reflection failures into a non-fatal probe result.")]
    private static VbProjectAccessProbe ProbeTemporaryWorkbookVbProjectAccess(object app)
    {
        object? workbooks = null;
        object? workbook = null;
        object? vbProject = null;
        object? components = null;
        try
        {
            dynamic excel = app;
            workbooks = (object?)excel.Workbooks;
            if (workbooks is null)
            {
                return new VbProjectAccessProbe(null, "Excel Workbooks collection is unavailable.");
            }
            dynamic books = workbooks;
            workbook = (object?)books.Add();
            if (workbook is null)
            {
                return new VbProjectAccessProbe(null, "Temporary workbook could not be created.");
            }
            dynamic book = workbook;
            vbProject = (object?)book.VBProject;
            if (vbProject is null)
            {
                return new VbProjectAccessProbe(false, "Temporary workbook VBProject is unavailable.");
            }
            dynamic project = vbProject;
            components = (object?)project.VBComponents;
            if (components is null)
            {
                return new VbProjectAccessProbe(false, "Temporary workbook VBComponents is unavailable.");
            }
            dynamic vbComponents = components;
            _ = vbComponents.Count;
            return new VbProjectAccessProbe(true, null);
        }
        catch (Exception ex)
        {
            return new VbProjectAccessProbe(false, UnwrapExceptionMessage(ex));
        }
        finally
        {
            if (workbook is not null)
            {
                try
                {
                    dynamic book = workbook;
                    book.Close(false);
                }
                catch
                {
                    // best-effort cleanup
                }
            }
            TryReleaseComObject(components);
            TryReleaseComObject(vbProject);
            TryReleaseComObject(workbook);
            TryReleaseComObject(workbooks);
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

    internal static bool? ReadTrustVbaAccessFromRegistry(string? excelVersion, Func<string, string, object?> readValue)
    {
        ArgumentNullException.ThrowIfNull(readValue);

        foreach (var version in CandidateOfficeVersions(excelVersion))
        {
            foreach (var root in new[] { @"Software\Policies\Microsoft\Office", @"Software\Microsoft\Office" })
            {
                object? value;
                try
                {
                    value = readValue($@"{root}\{version}\Excel\Security", "AccessVBOM");
                }
                catch (UnauthorizedAccessException)
                {
                    continue;
                }
                catch (SecurityException)
                {
                    continue;
                }
                catch (IOException)
                {
                    continue;
                }
                var parsed = ParseAccessVbomValue(value);
                if (parsed is not null)
                {
                    return parsed;
                }
            }
        }
        return null;
    }

    private static IEnumerable<string> CandidateOfficeVersions(string? excelVersion)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var normalized = NormalizeOfficeVersion(excelVersion);
        if (normalized is not null && seen.Add(normalized))
        {
            yield return normalized;
            yield break;
        }

        foreach (var fallback in new[] { "16.0", "15.0", "14.0", "12.0" })
        {
            if (seen.Add(fallback))
            {
                yield return fallback;
            }
        }
    }

    private static string? NormalizeOfficeVersion(string? excelVersion)
    {
        if (string.IsNullOrWhiteSpace(excelVersion))
        {
            return null;
        }
        var parts = excelVersion.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var major))
        {
            return null;
        }
        return major.ToString(CultureInfo.InvariantCulture) + ".0";
    }

    private static bool? ParseAccessVbomValue(object? value)
    {
        switch (value)
        {
            case null:
                return null;
            case int i:
                return i != 0;
            case long l:
                return l != 0;
            case string s when int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed):
                return parsed != 0;
            default:
                return null;
        }
    }

    private static object? ReadCurrentUserRegistryValue(string keyPath, string valueName)
    {
        using var key = Registry.CurrentUser.OpenSubKey(keyPath);
        return key?.GetValue(valueName);
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

    private static string UnwrapExceptionMessage(Exception ex)
    {
        return ex.InnerException is null ? ex.Message : $"{ex.Message} ({ex.InnerException.Message})";
    }

    private static void TryReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }
        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch (ArgumentException)
        {
            // best-effort cleanup
        }
        catch (InvalidComObjectException)
        {
            // best-effort cleanup
        }
        catch (COMException)
        {
            // best-effort cleanup
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

    private sealed record VbProjectAccessProbe(bool? Accessible, string? Message);
}
