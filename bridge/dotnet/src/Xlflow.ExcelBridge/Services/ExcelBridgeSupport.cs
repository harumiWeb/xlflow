using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Xlflow.ExcelBridge.Services;

internal static class BridgePayload
{
    public static bool HasProperty(JsonElement payload, string name)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        foreach (var property in payload.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

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

internal sealed record SessionMetadata(
    long Hwnd,
    int Pid,
    string WorkbookPath,
    string Owner = "managed",
    bool Poisoned = false,
    string PoisonedAt = "",
    string PoisonReason = "",
    string HResult = "",
    string LastCommand = "");

internal sealed record ExcelSessionAttachment(object Excel, object Workbook, string SessionMode);

internal sealed record ExcelProcessInfo(int ProcessId, bool? HasWorkbook);

internal sealed record OwnedExcelProcess(int ProcessId, DateTime? StartTime = null)
{
    public static OwnedExcelProcess None { get; } = new(0);
}

internal sealed class SessionPoisonedException(SessionMetadata metadata)
    : InvalidOperationException("xlflow session is poisoned after a fatal Excel COM/RPC failure")
{
    public SessionMetadata Metadata { get; } = metadata;
}

internal sealed record ComFailureInfo(bool Fatal, string HResult, int Number, string Message);

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge support intentionally treats COM inspection as best-effort and normalizes failures to null/false or structured bridge errors.")]
[SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance", Justification = "These helpers favor simple iteration-focused signatures.")]
internal static class ExcelBridgeSupport
{
    private const int ObjIdNativeOm = unchecked((int)0xFFFFFFF0);
    private static readonly Guid DispatchGuid = new("00020400-0000-0000-C000-000000000046");
    internal static CultureInfo ComInvokeCulture => CultureInfo.CurrentCulture;

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
        var owner = BridgePayload.GetString(root, "owner") ?? "managed";
        if (string.IsNullOrWhiteSpace(owner))
        {
            owner = "managed";
        }
        var poisoned = TryGetBool(root, "poisoned");
        var poisonedAt = BridgePayload.GetString(root, "poisoned_at") ?? "";
        var poisonReason = BridgePayload.GetString(root, "poison_reason") ?? "";
        var hResult = BridgePayload.GetString(root, "h_result") ?? "";
        var lastCommand = BridgePayload.GetString(root, "last_command") ?? "";
        return new SessionMetadata(hwnd, pid, workbookPath, owner, poisoned, poisonedAt, poisonReason, hResult, lastCommand);
    }

    public static void WriteSessionMetadata(string metadataPath, object excel, string workbookPath, string owner = "managed")
    {
        if (string.IsNullOrWhiteSpace(metadataPath))
        {
            return;
        }

        var parent = Path.GetDirectoryName(metadataPath);
        if (!string.IsNullOrWhiteSpace(parent))
        {
            Directory.CreateDirectory(parent);
        }

        var payload = new Dictionary<string, object?>
        {
            ["hwnd"] = GetExcelMainHwnd(excel),
            ["pid"] = GetExcelProcessId(excel),
            ["workbook_path"] = NormalizePath(workbookPath),
            ["owner"] = string.IsNullOrWhiteSpace(owner) ? "managed" : owner,
        };

        File.WriteAllText(metadataPath, JsonSerializer.Serialize(payload));
    }

    public static void MarkSessionPoisoned(string metadataPath, string workbookPath, string reason, string hResult, string lastCommand)
    {
        if (string.IsNullOrWhiteSpace(metadataPath))
        {
            return;
        }

        var metadata = ReadSessionMetadata(metadataPath);
        if (metadata is null)
        {
            return;
        }

        workbookPath = NormalizePath(workbookPath);
        if (!string.IsNullOrWhiteSpace(metadata.WorkbookPath) && !PathsEqual(metadata.WorkbookPath, workbookPath))
        {
            return;
        }

        var parent = Path.GetDirectoryName(metadataPath);
        if (!string.IsNullOrWhiteSpace(parent))
        {
            Directory.CreateDirectory(parent);
        }

        var payload = new Dictionary<string, object?>
        {
            ["hwnd"] = metadata.Hwnd,
            ["pid"] = metadata.Pid,
            ["workbook_path"] = string.IsNullOrWhiteSpace(metadata.WorkbookPath) ? workbookPath : NormalizePath(metadata.WorkbookPath),
            ["owner"] = string.IsNullOrWhiteSpace(metadata.Owner) ? "managed" : metadata.Owner,
            ["poisoned"] = true,
            ["poisoned_at"] = DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            ["poison_reason"] = reason,
            ["h_result"] = hResult,
            ["last_command"] = lastCommand,
        };

        File.WriteAllText(metadataPath, JsonSerializer.Serialize(payload));
    }

    public static void DeleteSessionMetadata(string metadataPath)
    {
        if (!string.IsNullOrWhiteSpace(metadataPath) && File.Exists(metadataPath))
        {
            File.Delete(metadataPath);
        }
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
            ThrowIfSessionPoisoned(metadataPath, workbookPath);
            var metadata = ReadSessionMetadata(metadataPath);
            var sessionMode = string.Equals(metadata?.Owner, "external", StringComparison.OrdinalIgnoreCase) ? "external" : "explicit";
            var excel = RunPhase("get_session_excel", () => GetSessionExcel(metadataPath))
                ?? throw new InvalidOperationException("xlflow session is not running");
            var workbook = RunPhase("get_session_workbook", () => GetOpenWorkbook(excel, workbookPath));
            return new ExcelSessionAttachment(excel, workbook, sessionMode);
        }

        if (SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            ThrowIfSessionPoisoned(metadataPath, workbookPath);
            var metadata = ReadSessionMetadata(metadataPath);
            var sessionMode = string.Equals(metadata?.Owner, "external", StringComparison.OrdinalIgnoreCase) ? "external" : "auto";
            var excel = RunPhase("get_auto_session_excel", () => GetExcelFromSessionMetadata(metadataPath));
            if (excel is not null)
            {
                try
                {
                    var workbook = RunPhase("get_auto_session_workbook", () => GetOpenWorkbook(excel, workbookPath));
                    return new ExcelSessionAttachment(excel, workbook, sessionMode);
                }
                catch
                {
                    ReleaseComObject(excel);
                }
            }
        }

        throw new InvalidOperationException("no matching xlflow session workbook is running; run xlflow session start or use the configured workbook session");
    }

    public static void ThrowIfSessionPoisoned(string metadataPath, string workbookPath)
    {
        var metadata = ReadSessionMetadata(metadataPath);
        if (metadata is null || !metadata.Poisoned)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(metadata.WorkbookPath) && !PathsEqual(metadata.WorkbookPath, workbookPath))
        {
            return;
        }

        throw new SessionPoisonedException(metadata);
    }

    public static ExcelSessionAttachment OpenWorkbookDirect(string workbookPath, bool visible, bool disableAutomationMacros = true)
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
            ConfigureExcelForAutomation(excel, visible, disableAutomationMacros);
            dynamic app = excel;
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

    public static ExcelSessionAttachment AttachToAlreadyOpenWorkbook(string workbookPath)
    {
        workbookPath = NormalizePath(workbookPath);
        string? activePath = null;

        var running = TryGetRunningExcelApplication();
        if (running is not null)
        {
            if (TryGetMatchingOpenWorkbook(running, workbookPath, out var workbook))
            {
                return new ExcelSessionAttachment(running, workbook!, "external");
            }
            activePath ??= TryGetActiveWorkbookPath(running);
            ReleaseComObject(running);
        }

        var foreground = TryGetForegroundExcel();
        if (foreground is not null)
        {
            if (TryGetMatchingOpenWorkbook(foreground, workbookPath, out var workbook))
            {
                return new ExcelSessionAttachment(foreground, workbook!, "external");
            }
            activePath ??= TryGetActiveWorkbookPath(foreground);
            ReleaseComObject(foreground);
        }

        foreach (var process in GetExcelProcesses())
        {
            var excel = TryGetExcelByProcessId(process.ProcessId);
            if (excel is null)
            {
                continue;
            }

            if (TryGetMatchingOpenWorkbook(excel, workbookPath, out var workbook))
            {
                return new ExcelSessionAttachment(excel, workbook!, "external");
            }

            activePath ??= TryGetActiveWorkbookPath(excel);
            ReleaseComObject(excel);
        }

        if (!string.IsNullOrWhiteSpace(activePath))
        {
            throw new InvalidOperationException("active_workbook_mismatch: Active workbook does not match configured workbook: " + activePath);
        }

        throw new InvalidOperationException("active_workbook_not_found: No open Excel workbook matches the configured workbook.");
    }

    private static bool TryGetMatchingOpenWorkbook(object excel, string workbookPath, out object? workbook)
    {
        try
        {
            workbook = GetOpenWorkbook(excel, workbookPath);
            return true;
        }
        catch
        {
            workbook = null;
            return false;
        }
    }

    private static string? TryGetActiveWorkbookPath(object excel)
    {
        object? workbook = null;
        try
        {
            workbook = TryGetActiveWorkbook(excel);
            return TryGetWorkbookFullName(workbook);
        }
        catch
        {
            return null;
        }
        finally
        {
            ReleaseComObject(workbook);
        }
    }

    private static void ConfigureExcelForAutomation(object excel, bool visible, bool disableAutomationMacros)
    {
        dynamic app = excel;
        app.Visible = visible;
        app.DisplayAlerts = false;
        app.EnableEvents = false;
        try
        {
            app.AutomationSecurity = disableAutomationMacros ? 3 : 1;
        }
        catch
        {
            // best-effort: some hosts may not expose AutomationSecurity
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

        var running = TryGetRunningExcelApplication();
        if (running is not null)
        {
            try
            {
                _ = GetOpenWorkbook(running, metadata.WorkbookPath);
                return running;
            }
            catch
            {
                ReleaseComObject(running);
            }
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
            dynamic wb = workbook;
            dirty = !Convert.ToBoolean(wb.Saved, CultureInfo.InvariantCulture);
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
            dynamic dyn = range;
            return Convert.ToString(dyn.Address(false, false, 1, false), CultureInfo.InvariantCulture) ?? "";
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
            dynamic dyn = cell;
            var text = Convert.ToString(dyn.Text, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    public static object? TryGetCellValue(object? cell)
    {
        if (cell is null)
        {
            return null;
        }

        try
        {
            dynamic dyn = cell;
            return NormalizeCellValue(dyn.Value2);
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
            dynamic dyn = cell;
            var formula = Convert.ToString(dyn.Formula, CultureInfo.InvariantCulture);
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

    public static object? NormalizeCellValue(object? value)
    {
        return value switch
        {
            null => null,
            DBNull => null,
            string text => text,
            bool boolean => boolean,
            byte number => number,
            short number => number,
            int number => number,
            long number => number,
            float number => number,
            double number => number,
            decimal number => number,
            DateTime dateTime => dateTime.ToString("O", CultureInfo.InvariantCulture),
            _ => value is Array ? value.ToString() : value,
        };
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
            ComInvokeCulture);
    }

    public static object? InvokeMethod(object comObject, string memberName, params object?[] args)
    {
        return comObject.GetType().InvokeMember(
            memberName,
            System.Reflection.BindingFlags.InvokeMethod,
            null,
            comObject,
            args,
            ComInvokeCulture);
    }

    public static object? InvokeViaDynamic(object comObject, string memberName, params object?[] args)
    {
        dynamic dyn = comObject;
        return memberName switch
        {
            "Buttons" => dyn.Buttons(),
            "CopyPicture" => InvokeCopyPicture(dyn, args),
            "Open" => dyn.Open(args[0]),
            "Save" => dyn.Save(),
            "SaveAs" => InvokeSaveAs(dyn, args),
            "SaveCopyAs" => dyn.SaveCopyAs(args[0]),
            "Import" => dyn.Import(args[0]),
            "Export" => dyn.Export(args[0]),
            "Remove" => dyn.Remove(args[0]),
            "Paste" => InvokePaste(dyn),
            "Close" => args.Length > 0 ? dyn.Close(args[0]) : dyn.Close(),
            "Quit" => dyn.Quit(),
            "DeleteLines" => dyn.DeleteLines(args[0], args[1]),
            "InsertLines" => dyn.InsertLines(args[0], args[1]),
            _ => throw new InvalidOperationException($"Unsupported COM method for dynamic dispatch: {memberName}"),
        };
    }

    private static object? InvokeSaveAs(dynamic dyn, object?[] args)
    {
        switch (args.Length)
        {
            case 1:
                dyn.SaveAs(args[0]);
                return null;
            case 2:
                dyn.SaveAs(args[0], args[1]);
                return null;
            default:
                throw new InvalidOperationException("SaveAs currently supports 1 or 2 arguments.");
        }
    }

    private static object? InvokeCopyPicture(dynamic dyn, object?[] args)
    {
        switch (args.Length)
        {
            case 0:
                dyn.CopyPicture();
                return null;
            case 1:
                dyn.CopyPicture(args[0]);
                return null;
            case 2:
                dyn.CopyPicture(args[0], args[1]);
                return null;
            default:
                throw new InvalidOperationException("CopyPicture currently supports up to 2 arguments.");
        }
    }

    private static object? InvokePaste(dynamic dyn)
    {
        dyn.Paste();
        return null;
    }

    public static T RunPhase<T>(string phase, Func<T> action)
    {
        try
        {
            return action();
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not SessionPoisonedException)
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

    public static ComFailureInfo ClassifyComFailure(Exception ex)
    {
        var detail = ex.InnerException ?? ex;
        var number = detail.HResult;
        var hResult = FormatHResult(number);
        return new ComFailureInfo(IsFatalComFailure(number), hResult, number, FormatExceptionDetail(ex));
    }

    public static bool IsFatalComFailure(int number)
    {
        return unchecked((uint)number) switch
        {
            0x800706BE => true, // RPC_S_CALL_FAILED
            0x800706BA => true, // RPC_S_SERVER_UNAVAILABLE
            0x80010108 => true, // RPC_E_DISCONNECTED
            0x80010105 => true, // RPC_E_SERVERFAULT
            0x80010001 => true, // RPC_E_CALL_REJECTED
            _ => false,
        };
    }

    public static string FormatHResult(int number)
    {
        return "0x" + unchecked((uint)number).ToString("X8", CultureInfo.InvariantCulture);
    }

    public static void StabilizeExcelForMacroRun(object excel, string workbookPath, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow.Add(timeout);
        Exception? lastError = null;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var workbook = GetOpenWorkbook(excel, workbookPath);
                try
                {
                    dynamic wb = workbook;
                    wb.Activate();
                }
                finally
                {
                    ReleaseComObject(workbook);
                }

                dynamic app = excel;
                try
                {
                    _ = app.Ready;
                }
                catch
                {
                    // Some Excel hosts reject Ready during startup; DoEvents and retry below.
                }
                TryDoEvents(excel);
                PumpWaitingMessages();
                return;
            }
            catch (Exception ex)
            {
                lastError = ex;
                if (ClassifyComFailure(ex).Fatal)
                {
                    throw;
                }
                SleepAndPump(TimeSpan.FromMilliseconds(100));
            }
        }

        if (lastError is not null)
        {
            throw new InvalidOperationException("Excel did not become ready for macro execution.", lastError);
        }
    }

    public static void TryDoEvents(object? excel)
    {
        if (excel is null)
        {
            return;
        }

        try
        {
            dynamic app = excel;
            app.Run("DoEvents");
        }
        catch
        {
            // DoEvents is best-effort only.
        }
    }

    public static void SleepAndPump(TimeSpan duration)
    {
        var deadline = DateTime.UtcNow.Add(duration);
        while (DateTime.UtcNow < deadline)
        {
            PumpWaitingMessages();
            Thread.Sleep(Math.Min(25, Math.Max(1, (int)(deadline - DateTime.UtcNow).TotalMilliseconds)));
        }
        PumpWaitingMessages();
    }

    public static void PumpWaitingMessages()
    {
        while (NativeMethods.PeekMessage(out var message, IntPtr.Zero, 0, 0, NativeMethods.PmRemove))
        {
            NativeMethods.TranslateMessage(ref message);
            NativeMethods.DispatchMessage(ref message);
        }
    }

    public static IDisposable RegisterOleMessageFilter()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
        {
            return NullDisposable.Instance;
        }

        try
        {
            var filter = new OleMessageFilter();
            var hr = NativeMethods.CoRegisterMessageFilter(filter, out var previous);
            return hr == 0 ? new OleMessageFilterRegistration(filter, previous) : NullDisposable.Instance;
        }
        catch
        {
            return NullDisposable.Instance;
        }
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

    public static object? TryGetWorkbookVbProject(object? workbook)
    {
        if (workbook is null)
        {
            return null;
        }
        try
        {
            dynamic wb = workbook;
            return (object?)wb.VBProject;
        }
        catch
        {
            return null;
        }
    }

    public static string? TryGetWorkbookName(object? workbook)
    {
        if (workbook is null)
        {
            return null;
        }
        try
        {
            dynamic wb = workbook;
            return Convert.ToString(wb.Name, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    public static string? TryGetWorkbookFullName(object? workbook)
    {
        if (workbook is null)
        {
            return null;
        }
        try
        {
            dynamic wb = workbook;
            return Convert.ToString(wb.FullName, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    public static object? TryGetActiveSheet(object? excel)
    {
        if (excel is null)
        {
            return null;
        }
        try
        {
            dynamic app = excel;
            return (object?)app.ActiveSheet;
        }
        catch
        {
            return null;
        }
    }

    public static object? TryGetForegroundExcel()
    {
        var hwnd = NativeMethods.GetForegroundWindow();
        if (hwnd == IntPtr.Zero)
        {
            return TryGetRunningExcelApplication();
        }
        return TryGetExcelByHwnd(hwnd.ToInt64()) ?? TryGetRunningExcelApplication();
    }

    public static Dictionary<string, object?> BuildSessionPayload(string workbookPath, bool active, string mode, bool? dirty, bool saveRequired)
    {
        var payload = new Dictionary<string, object?>
        {
            ["active"] = active,
            ["workbook_path"] = NormalizePath(workbookPath),
            ["dirty"] = dirty,
            ["save_required"] = saveRequired,
            ["live_newer_than_disk"] = saveRequired,
            ["mode"] = mode,
            ["source_of_truth"] = saveRequired ? "live_workbook" : "saved_workbook",
        };
        if (string.Equals(mode, "external", StringComparison.OrdinalIgnoreCase))
        {
            payload["owner"] = "external";
        }
        return payload;
    }

    public static Dictionary<string, object?> BuildWorkbookPayload(string workbookPath, bool sessionAttached, string sessionMode, bool sessionRequested, bool saved, bool? dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["path"] = NormalizePath(workbookPath),
            ["session"] = sessionAttached,
            ["session_mode"] = sessionMode,
            ["session_requested"] = sessionRequested,
            ["auto_session"] = sessionAttached && !sessionRequested,
            ["saved"] = saved,
            ["dirty"] = dirty,
            ["needs_save"] = needsSave,
        };
    }

    public static Dictionary<string, object?> BuildTargetPayload(string kind, string workbookPath)
    {
        return new Dictionary<string, object?>
        {
            ["kind"] = kind,
            ["path"] = NormalizePath(workbookPath),
        };
    }

    public static string? GetSessionUsageLog(string sessionMode)
    {
        return sessionMode switch
        {
            "explicit" => "attached to explicit xlflow session workbook",
            "auto" => "attached to matching xlflow session workbook",
            "external" => "attached to external xlflow session workbook",
            _ => null,
        };
    }

    public static object? TryGetSelection(object? excel)
    {
        if (excel is null)
        {
            return null;
        }
        try
        {
            dynamic app = excel;
            return (object?)app.Selection;
        }
        catch
        {
            return null;
        }
    }

    public static bool TryGetVisible(object? excel, out bool visible)
    {
        if (excel is null)
        {
            visible = false;
            return false;
        }
        try
        {
            dynamic app = excel;
            visible = Convert.ToBoolean(app.Visible, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            visible = false;
            return false;
        }
    }

    public static object? RunExcelMacro(object excel, string macroName, params object?[] args)
    {
        dynamic app = excel;
        return args.Length switch
        {
            0 => app.Run(macroName),
            1 => app.Run(macroName, args[0]),
            2 => app.Run(macroName, args[0], args[1]),
            3 => app.Run(macroName, args[0], args[1], args[2]),
            4 => app.Run(macroName, args[0], args[1], args[2], args[3]),
            5 => app.Run(macroName, args[0], args[1], args[2], args[3], args[4]),
            _ => throw new InvalidOperationException("RunExcelMacro currently supports up to 5 macro arguments."),
        };
    }

    public static int ToInt(object? value)
    {
        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    public static bool ToBool(object? value)
    {
        return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    public static double ToDouble(object? value)
    {
        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    public static object? TryGetRunningExcelApplication()
    {
        try
        {
            if (NativeMethods.CLSIDFromProgID("Excel.Application", out var clsid) != 0)
            {
                return null;
            }
            var hr = NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out var app);
            if (hr != 0)
            {
                return null;
            }
            return app;
        }
        catch
        {
            return null;
        }
    }

    public static object? TryGetActiveWorkbook(object? excel)
    {
        if (excel is null)
        {
            return null;
        }
        try
        {
            dynamic app = excel;
            return (object?)app.ActiveWorkbook;
        }
        catch
        {
            return null;
        }
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

    public static void CloseWorkbookAndQuitApplication(object? workbook, object? excel, int ownedProcessId = 0)
    {
        CloseWorkbookAndQuitApplication(workbook, excel, CaptureOwnedExcelProcess(ownedProcessId));
    }

    public static void CloseWorkbookAndQuitApplication(object? workbook, object? excel, OwnedExcelProcess ownedProcess)
    {
        if (workbook is not null)
        {
            try
            {
                InvokeViaDynamic(workbook, "Close", false);
            }
            catch
            {
                // best-effort close
            }
            ReleaseComObject(workbook);
        }

        if (excel is not null)
        {
            try
            {
                dynamic app = excel;
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                app.Quit();
            }
            catch
            {
                // best-effort quit
            }
            ReleaseComObject(excel);
        }

        CollectComGarbage();
        EnsureOwnedExcelProcessExited(ownedProcess);
    }

    public static void CollectComGarbage()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public static OwnedExcelProcess CaptureOwnedExcelProcess(int processId)
    {
        if (processId <= 0)
        {
            return OwnedExcelProcess.None;
        }

        try
        {
            using var process = Process.GetProcessById(processId);
            if (!IsExcelProcess(process))
            {
                return OwnedExcelProcess.None;
            }

            DateTime? startTime = null;
            try
            {
                startTime = process.StartTime;
            }
            catch
            {
                // Some environments deny StartTime; process name still guards non-Excel PID reuse.
            }
            return new OwnedExcelProcess(processId, startTime);
        }
        catch
        {
            return OwnedExcelProcess.None;
        }
    }

    private static void EnsureOwnedExcelProcessExited(OwnedExcelProcess ownedProcess)
    {
        if (ownedProcess.ProcessId <= 0)
        {
            return;
        }

        try
        {
            using var process = Process.GetProcessById(ownedProcess.ProcessId);
            if (process.WaitForExit(1500))
            {
                return;
            }

            if (!IsSameOwnedExcelProcess(process, ownedProcess))
            {
                return;
            }

            process.Kill(entireProcessTree: true);
            _ = process.WaitForExit(3000);
        }
        catch
        {
            // The process may already have exited, or Windows may refuse termination.
        }
    }

    private static bool IsSameOwnedExcelProcess(Process process, OwnedExcelProcess ownedProcess)
    {
        if (!IsExcelProcess(process))
        {
            return false;
        }
        if (ownedProcess.StartTime is null)
        {
            return true;
        }
        try
        {
            return process.StartTime == ownedProcess.StartTime.Value;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsExcelProcess(Process process)
    {
        try
        {
            return string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
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

    private static bool TryGetBool(JsonElement element, string name)
    {
        var value = BridgePayload.GetString(element, name);
        return bool.TryParse(value, out var parsed) && parsed;
    }

    private sealed class NullDisposable : IDisposable
    {
        public static readonly NullDisposable Instance = new();
        public void Dispose()
        {
        }
    }

    private sealed class OleMessageFilterRegistration(IOleMessageFilter current, IOleMessageFilter? previous) : IDisposable
    {
        private bool _disposed;
        private IOleMessageFilter? _current = current;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                _ = NativeMethods.CoRegisterMessageFilter(previous, out _);
            }
            catch
            {
                // best-effort restoration
            }
            _current = null;
        }
    }

    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType);
    }

    private sealed class OleMessageFilter : IOleMessageFilter
    {
        public int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0;
        }

        public int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType)
        {
            const int serverCallRetryLater = 2;
            return dwRejectType == serverCallRetryLater && dwTickCount < 10_000 ? 99 : -1;
        }

        public int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2;
        }
    }

    private static class NativeMethods
    {
        public const uint PmRemove = 0x0001;

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct Message
        {
            public IntPtr Hwnd;
            public uint Msg;
            public IntPtr WParam;
            public IntPtr LParam;
            public uint Time;
            public int PtX;
            public int PtY;
        }

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        public static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32.dll")]
        public static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            int dwId,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object? ppvObject);

        [DllImport("ole32.dll")]
        public static extern int CoRegisterMessageFilter(
            [MarshalAs(UnmanagedType.Interface)] IOleMessageFilter? newFilter,
            [MarshalAs(UnmanagedType.Interface)] out IOleMessageFilter? oldFilter);

        [DllImport("user32.dll")]
        public static extern bool PeekMessage(out Message lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage(ref Message lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref Message lpMsg);
    }
}
