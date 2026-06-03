using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Xlflow.ExcelBridge.Services;

internal static class BridgePayload
{
    public static string? GetString(JsonElement payload, string name)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        foreach (var property in payload.EnumerateObject())
        {
            if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return property.Value.ValueKind switch
            {
                JsonValueKind.Null => null,
                JsonValueKind.String => property.Value.GetString(),
                _ => property.Value.ToString(),
            };
        }

        return null;
    }

    public static bool GetBool(JsonElement payload, string name, bool defaultValue = false)
    {
        var value = GetString(payload, name);
        return bool.TryParse(value, out var parsed) ? parsed : defaultValue;
    }

    public static int GetInt(JsonElement payload, string name, int defaultValue = 0)
    {
        var value = GetString(payload, name);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : defaultValue;
    }

    public static int? GetNullableInt(JsonElement payload, string name)
    {
        var value = GetString(payload, name);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
    }
}

internal sealed record SessionMetadata(long Hwnd, int Pid, string WorkbookPath);

internal sealed record ExcelSessionAttachment(object Excel, object Workbook, string SessionMode);

internal sealed record ExcelProcessInfo(int ProcessId, bool? HasWorkbook);

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge support intentionally treats COM inspection as best-effort and normalizes failures to null/false or structured bridge errors.")]
[SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance", Justification = "These helpers favor simple iteration-focused signatures.")]
internal static class ExcelBridgeSupport
{
    private const int ObjIdNativeOm = unchecked((int)0xFFFFFFF0);
    private static readonly Guid DispatchGuid = new("00020400-0000-0000-C000-000000000046");

    public static string NormalizePath(string path)
    {
        return Path.GetFullPath(path);
    }

    public static bool PathsEqual(string left, string right)
    {
        return string.Equals(NormalizePath(left), NormalizePath(right), StringComparison.OrdinalIgnoreCase);
    }

    public static SessionMetadata? ReadSessionMetadata(string metadataPath)
    {
        if (string.IsNullOrWhiteSpace(metadataPath) || !File.Exists(metadataPath))
        {
            return null;
        }

        using var json = JsonDocument.Parse(File.ReadAllText(metadataPath));
        var root = json.RootElement;
        var hwnd = TryGetInt64(root, "hwnd");
        var pid = TryGetInt32(root, "pid");
        var workbookPath = BridgePayload.GetString(root, "workbook_path") ?? "";
        return new SessionMetadata(hwnd, pid, workbookPath);
    }

    public static bool SessionMetadataMatchesWorkbook(string metadataPath, string workbookPath)
    {
        var metadata = ReadSessionMetadata(metadataPath);
        if (metadata is null || string.IsNullOrWhiteSpace(metadata.WorkbookPath))
        {
            return false;
        }

        return PathsEqual(metadata.WorkbookPath, workbookPath);
    }

    public static ExcelSessionAttachment AttachToSessionWorkbook(string workbookPath, string metadataPath, bool useSession)
    {
        workbookPath = NormalizePath(workbookPath);

        if (useSession)
        {
            var excel = RunPhase("get_session_excel", () => GetSessionExcel(metadataPath))
                ?? throw new InvalidOperationException("xlflow session is not running");
            var workbook = RunPhase("get_session_workbook", () => GetOpenWorkbook(excel, workbookPath));
            return new ExcelSessionAttachment(excel, workbook, "explicit");
        }

        if (SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            var excel = RunPhase("get_auto_session_excel", () => GetExcelFromSessionMetadata(metadataPath));
            if (excel is not null)
            {
                try
                {
                    var workbook = RunPhase("get_auto_session_workbook", () => GetOpenWorkbook(excel, workbookPath));
                    return new ExcelSessionAttachment(excel, workbook, "auto");
                }
                catch
                {
                    ReleaseComObject(excel);
                }
            }
        }

        throw new InvalidOperationException("no matching xlflow session workbook is running; run xlflow session start or use the configured workbook session");
    }

    public static ExcelSessionAttachment OpenWorkbookDirect(string workbookPath, bool visible)
    {
        workbookPath = NormalizePath(workbookPath);

        if (!IsExcelFile(workbookPath))
        {
            throw new InvalidOperationException($"bridge_file_not_openable: File does not appear to be an Excel workbook: {workbookPath}");
        }

        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel.Application COM class is not registered");
        var excel = Activator.CreateInstance(excelType)
            ?? throw new InvalidOperationException("Failed to create Excel.Application COM instance");

        try
        {
            dynamic app = excel;
            app.Visible = visible;
            app.DisplayAlerts = false;
            app.EnableEvents = false;

            object workbook = app.Workbooks.Open(workbookPath);
            return new ExcelSessionAttachment(excel, workbook, "none");
        }
        catch (InvalidOperationException)
        {
            ReleaseComObject(excel);
            throw;
        }
        catch (Exception ex)
        {
            ReleaseComObject(excel);
            var detail = UnwrapComErrorMessage(ex);
            throw new InvalidOperationException($"bridge_file_not_openable: {detail}", ex);
        }
    }

    public static bool IsExcelFile(string path)
    {
        if (!File.Exists(path))
        {
            return false;
        }

        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".xlsm" or ".xlsx" or ".xls" or ".xlt" or ".xla" or ".xlam" or ".xltx" or ".xltm";
    }

    private static string UnwrapComErrorMessage(Exception ex)
    {
        if (ex.InnerException is not null)
        {
            return $"{ex.Message} ({ex.InnerException.Message})";
        }
        return ex.Message;
    }

    public static object? GetSessionExcel(string metadataPath)
    {
        var excel = GetExcelFromSessionMetadata(metadataPath);
        if (excel is not null)
        {
            return excel;
        }
        return null;
    }

    public static object? GetExcelFromSessionMetadata(string metadataPath)
    {
        var metadata = ReadSessionMetadata(metadataPath);
        if (metadata is null)
        {
            return null;
        }

        if (metadata.Hwnd != 0)
        {
            var excel = TryGetExcelByHwnd(metadata.Hwnd);
            if (excel is not null)
            {
                return excel;
            }
        }

        if (metadata.Pid > 0)
        {
            return TryGetExcelByProcessId(metadata.Pid);
        }

        return null;
    }

    public static object GetOpenWorkbook(object excel, string workbookPath)
    {
        workbookPath = NormalizePath(workbookPath);
        object? workbooks = null;
        try
        {
            dynamic app = excel;
            workbooks = (object)app.Workbooks;
            dynamic books = workbooks;
            var count = ToInt(books.Count);
            for (var index = 1; index <= count; index++)
            {
                object? candidate = null;
                var matched = false;
                try
                {
                    candidate = (object)books.Item(index);
                    dynamic workbook = candidate;
                    var fullName = (string?)workbook.FullName;
                    if (!string.IsNullOrWhiteSpace(fullName) && PathsEqual(fullName, workbookPath))
                    {
                        matched = true;
                        return candidate!;
                    }
                }
                finally
                {
                    if (!matched)
                    {
                        ReleaseComObject(candidate);
                    }
                }
            }
        }
        finally
        {
            ReleaseComObject(workbooks);
        }

        throw new InvalidOperationException($"xlflow session workbook is not open: {workbookPath}");
    }

    public static bool TryGetWorkbookDirtyState(object workbook, out bool dirty)
    {
        try
        {
            dirty = !ToBool(Get(workbook, "Saved"));
            return true;
        }
        catch
        {
            dirty = false;
            return false;
        }
    }

    public static bool? GetWorkbookStateByProcessId(int processId)
    {
        if (processId <= 0)
        {
            return null;
        }

        var sawWorkbookFreeState = false;
        foreach (var hwnd in GetWindowsForProcess(processId))
        {
            foreach (var candidateHwnd in EnumerateWindowAndChildren(hwnd))
            {
                var excel = TryGetExcelFromWindow(candidateHwnd);
                if (excel is null)
                {
                    continue;
                }

                object? workbooks = null;
                try
                {
                    workbooks = Get(excel, "Workbooks");
                    if (ToInt(Get(workbooks!, "Count")) > 0)
                    {
                        return true;
                    }

                    sawWorkbookFreeState = true;
                }
                catch
                {
                    // ignore inaccessible candidate
                }
                finally
                {
                    ReleaseComObject(workbooks);
                    ReleaseComObject(excel);
                }
            }
        }

        return sawWorkbookFreeState ? false : null;
    }

    public static List<ExcelProcessInfo> GetExcelProcesses()
    {
        return Process.GetProcessesByName("EXCEL")
            .Select(process =>
            {
                try
                {
                    return new ExcelProcessInfo(process.Id, GetWorkbookStateByProcessId(process.Id));
                }
                finally
                {
                    process.Dispose();
                }
            })
            .OrderBy(process => process.ProcessId)
            .ToList();
    }

    public static object? TryGetExcelByProcessId(int processId)
    {
        if (processId <= 0)
        {
            return null;
        }

        foreach (var hwnd in GetWindowsForProcess(processId))
        {
            foreach (var candidateHwnd in EnumerateWindowAndChildren(hwnd))
            {
                var excel = TryGetExcelFromWindow(candidateHwnd);
                if (excel is not null)
                {
                    return excel;
                }
            }
        }

        return null;
    }

    public static object? TryGetExcelByHwnd(long hwnd)
    {
        if (hwnd == 0)
        {
            return null;
        }

        foreach (var candidateHwnd in EnumerateWindowAndChildren(new IntPtr(hwnd)))
        {
            var excel = TryGetExcelFromWindow(candidateHwnd);
            if (excel is not null)
            {
                return excel;
            }
        }

        return null;
    }

    public static long GetExcelMainHwnd(object excel)
    {
        try
        {
            return Convert.ToInt64(Get(excel, "Hwnd"), CultureInfo.InvariantCulture);
        }
        catch
        {
            return 0;
        }
    }

    public static int GetExcelProcessId(object excel)
    {
        var hwnd = GetExcelMainHwnd(excel);
        if (hwnd == 0)
        {
            return 0;
        }
        _ = NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd), out var processId);
        return processId;
    }

    public static string GetRangeAddress(object range)
    {
        try
        {
            return GetString(range, "Address", false, false, 1, false) ?? "";
        }
        catch
        {
            return "";
        }
    }

    public static string? TryGetCellText(object? cell)
    {
        if (cell is null)
        {
            return null;
        }

        try
        {
            var text = GetString(cell, "Text");
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    public static string? TryGetFormula(object? cell)
    {
        if (cell is null)
        {
            return null;
        }

        try
        {
            var formula = GetString(cell, "Formula");
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith('='))
            {
                return null;
            }

            return formula;
        }
        catch
        {
            return null;
        }
    }

    public static bool VisibleToBool(object? value)
    {
        try
        {
            return Convert.ToInt32(value, CultureInfo.InvariantCulture) == -1;
        }
        catch
        {
            return false;
        }
    }

    public static string? ColorToHex(object? value)
    {
        if (value is null)
        {
            return null;
        }

        try
        {
            var color = Convert.ToInt64(value, CultureInfo.InvariantCulture);
            if (color < 0)
            {
                return null;
            }

            var red = color & 255;
            var green = (color >> 8) & 255;
            var blue = (color >> 16) & 255;
            return $"#{red:X2}{green:X2}{blue:X2}";
        }
        catch
        {
            return null;
        }
    }

    public static string BorderStyleName(object? lineStyle, object? weight)
    {
        try
        {
            var line = Convert.ToInt32(lineStyle, CultureInfo.InvariantCulture);
            if (line == -4142)
            {
                return "none";
            }

            return line switch
            {
                -4119 => "double",
                -4115 => "dashed",
                -4118 => "dotted",
                4 => "dashDot",
                5 => "dashDotDot",
                13 => "slantDashDot",
                _ => Convert.ToInt32(weight, CultureInfo.InvariantCulture) switch
                {
                    1 => "hair",
                    2 => "thin",
                    -4138 => "medium",
                    4 => "thick",
                    _ => "thin",
                },
            };
        }
        catch
        {
            return "none";
        }
    }

    public static string? AlignmentName(object? value, string axis)
    {
        try
        {
            var alignment = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            return alignment switch
            {
                -4131 => "left",
                -4108 => "center",
                -4152 => "right",
                -4130 => "justify",
                -4117 => "distributed",
                -4160 when axis == "vertical" => "top",
                -4107 when axis == "vertical" => "bottom",
                _ => null,
            };
        }
        catch
        {
            return null;
        }
    }

    public static string FillType(object? pattern)
    {
        try
        {
            var fill = Convert.ToInt32(pattern, CultureInfo.InvariantCulture);
            return fill switch
            {
                -4142 => "none",
                1 => "solid",
                _ => $"pattern:{fill}",
            };
        }
        catch
        {
            return "none";
        }
    }

    public static object? Get(object comObject, string memberName, params object?[] args)
    {
        return comObject.GetType().InvokeMember(
            memberName,
            System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.InvokeMethod,
            null,
            comObject,
            args,
            CultureInfo.InvariantCulture);
    }

    public static object? InvokeMethod(object comObject, string memberName, params object?[] args)
    {
        return comObject.GetType().InvokeMember(
            memberName,
            System.Reflection.BindingFlags.InvokeMethod,
            null,
            comObject,
            args,
            CultureInfo.InvariantCulture);
    }

    public static object? InvokeViaDynamic(object comObject, string memberName, params object?[] args)
    {
        dynamic dyn = comObject;
        return memberName switch
        {
            "Open" => dyn.Open(args[0]),
            "Save" => dyn.Save(),
            "SaveCopyAs" => dyn.SaveCopyAs(args[0]),
            "Import" => dyn.Import(args[0]),
            "Export" => dyn.Export(args[0]),
            "Remove" => dyn.Remove(args[0]),
            "Close" => args.Length > 0 ? dyn.Close(args[0]) : dyn.Close(),
            "Quit" => dyn.Quit(),
            "DeleteLines" => dyn.DeleteLines(args[0], args[1]),
            "InsertLines" => dyn.InsertLines(args[0], args[1]),
            _ => throw new InvalidOperationException($"Unsupported COM method for dynamic dispatch: {memberName}"),
        };
    }

    public static T RunPhase<T>(string phase, Func<T> action)
    {
        try
        {
            return action();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            throw new InvalidOperationException(
                $"{phase} failed: {FormatExceptionDetail(ex)}",
                ex);
        }
    }

    public static void RunPhase(string phase, Action action)
    {
        RunPhase(phase, () =>
        {
            action();
            return true;
        });
    }

    public static string FormatExceptionDetail(Exception ex)
    {
        if (ex.InnerException is not null)
        {
            if (ex.Message.Contains(ex.InnerException.Message, StringComparison.Ordinal))
            {
                return ex.Message;
            }
            return $"{ex.Message} ({ex.InnerException.Message})";
        }
        return ex.Message;
    }

    public static void Set(object comObject, string memberName, object? value)
    {
        comObject.GetType().InvokeMember(
            memberName,
            System.Reflection.BindingFlags.SetProperty,
            null,
            comObject,
            [value],
            CultureInfo.InvariantCulture);
    }

    public static string? GetString(object comObject, string memberName, params object?[] args)
    {
        var value = Get(comObject, memberName, args);
        return value?.ToString();
    }

    public static int ToInt(object? value)
    {
        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    public static bool ToBool(object? value)
    {
        return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    public static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch
        {
            // best-effort COM cleanup
        }
    }

    private static List<IntPtr> GetWindowsForProcess(int processId)
    {
        var windows = new List<IntPtr>();
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            var threadId = NativeMethods.GetWindowThreadProcessId(hwnd, out var candidatePid);
            if (threadId == 0)
            {
                return true;
            }
            if (candidatePid == processId)
            {
                windows.Add(hwnd);
            }

            return true;
        }, IntPtr.Zero);
        return windows;
    }

    private static IEnumerable<IntPtr> EnumerateWindowAndChildren(IntPtr hwnd)
    {
        yield return hwnd;

        var children = new List<IntPtr>();
        NativeMethods.EnumChildWindows(hwnd, (childHwnd, _) =>
        {
            children.Add(childHwnd);
            return true;
        }, IntPtr.Zero);

        foreach (var child in children)
        {
            yield return child;
        }
    }

    private static object? TryGetExcelFromWindow(IntPtr hwnd)
    {
        object? dispatch = null;
        try
        {
            var dispatchGuid = DispatchGuid;
            var hr = NativeMethods.AccessibleObjectFromWindow(hwnd, ObjIdNativeOm, ref dispatchGuid, out dispatch);
            if (hr != 0 || dispatch is null)
            {
                return null;
            }

            object? candidate = dispatch;
            try
            {
                var application = Get(dispatch, "Application");
                if (application is not null)
                {
                    candidate = application;
                }
            }
            catch
            {
                candidate = dispatch;
            }

            object? workbooks = null;
            try
            {
                workbooks = Get(candidate!, "Workbooks");
                _ = ToInt(Get(workbooks!, "Count"));
                ReleaseComObject(workbooks);
                workbooks = null;
                if (!ReferenceEquals(candidate, dispatch))
                {
                    ReleaseComObject(dispatch);
                }

                return candidate;
            }
            catch
            {
                ReleaseComObject(workbooks);
                if (!ReferenceEquals(candidate, dispatch))
                {
                    ReleaseComObject(candidate);
                }

                ReleaseComObject(dispatch);
                return null;
            }
        }
        catch
        {
            ReleaseComObject(dispatch);
            return null;
        }
    }

    private static int TryGetInt32(JsonElement element, string name)
    {
        var value = BridgePayload.GetString(element, name);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
    }

    private static long TryGetInt64(JsonElement element, string name)
    {
        var value = BridgePayload.GetString(element, name);
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0L;
    }

    private static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            int dwId,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object? ppvObject);
    }
}
