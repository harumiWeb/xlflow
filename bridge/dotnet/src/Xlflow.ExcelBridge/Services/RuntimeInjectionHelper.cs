using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge helpers intentionally normalize COM failures into best-effort results.")]
internal static class RuntimeInjectionHelper
{
    private static readonly JsonSerializerOptions CachedJsonOptions = new() { PropertyNameCaseInsensitive = true };

    public static string NormalizeRuntimeMode(string mode)
    {
        var normalized = (mode ?? "").Trim().ToLowerInvariant();
        return normalized switch
        {
            "headless" => "headless",
            "ci" => "ci",
            "agent" => "agent",
            "test" => "test",
            _ => "interactive",
        };
    }

    public static string ConvertToUIResponseId(string value)
    {
        var lower = (value ?? "").Trim().ToLowerInvariant();
        var builder = new StringBuilder(lower.Length);
        var lastSeparator = false;
        foreach (var c in lower)
        {
            var isValid = (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9');
            if (isValid)
            {
                builder.Append(c);
                lastSeparator = false;
                continue;
            }
            if (!lastSeparator && builder.Length > 0)
            {
                builder.Append('_');
                lastSeparator = true;
            }
        }
        return builder.ToString().Trim('_');
    }

    public static string GetUIResponseDefinedName(string kind, string id)
    {
        var normalizedKind = (kind ?? "").Trim().ToLowerInvariant();
        if (normalizedKind is not ("msgbox" or "input"))
        {
            throw new InvalidOperationException($"unsupported xlflow UI response kind: {kind}");
        }
        var normalizedId = ConvertToUIResponseId(id);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            throw new InvalidOperationException("xlflow UI response id cannot be empty");
        }
        return "__XLFLOW_UI_" + normalizedKind.ToUpperInvariant() + "_" + normalizedId + "__";
    }

    public static string GetFileDialogResponseDefinedName(string kind, string id)
    {
        var normalizedKind = (kind ?? "").Trim().ToLowerInvariant();
        var kindToken = normalizedKind switch
        {
            "get-open" => "GET_OPEN",
            "file-open" => "FILE_OPEN",
            "save-as" => "SAVE_AS",
            "folder" => "FOLDER",
            _ => throw new InvalidOperationException($"unsupported xlflow file dialog kind: {kind}"),
        };
        var normalizedId = ConvertToUIResponseId(id);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            throw new InvalidOperationException("xlflow file dialog response id cannot be empty");
        }
        return "__XLFLOW_UI_FILEDIALOG_" + kindToken + "_" + normalizedId + "__";
    }

    public static Dictionary<string, string> DecodeUIResponsesJson(string base64Json)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(base64Json))
        {
            return result;
        }
        try
        {
            var json = Encoding.UTF8.GetString(Convert.FromBase64String(base64Json));
            using var doc = JsonDocument.Parse(json);
            foreach (var property in doc.RootElement.EnumerateObject())
            {
                result[property.Name] = property.Value.ValueKind == JsonValueKind.String
                    ? property.Value.GetString() ?? ""
                    : property.Value.ToString();
            }
        }
        catch
        {
            // best-effort decode
        }
        return result;
    }

    public static List<FileDialogResponse> DecodeFileDialogResponsesJson(string base64Json)
    {
        var result = new List<FileDialogResponse>();
        if (string.IsNullOrWhiteSpace(base64Json))
        {
            return result;
        }
        try
        {
            var json = Encoding.UTF8.GetString(Convert.FromBase64String(base64Json));
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    var kind = BridgePayload.GetString(item, "kind") ?? "";
                    var dialogId = BridgePayload.GetString(item, "dialog_id") ?? "";
                    var cancelled = BridgePayload.GetBool(item, "cancelled");
                    var values = new List<string>();
                    if (item.TryGetProperty("values", out var valuesElement) && valuesElement.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var v in valuesElement.EnumerateArray())
                        {
                            values.Add(v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : v.ToString());
                        }
                    }
                    result.Add(new FileDialogResponse(kind, dialogId, cancelled, values));
                }
            }
        }
        catch
        {
            // best-effort decode
        }
        return result;
    }

    public static string ConvertToFileDialogMarkerValue(FileDialogResponse response)
    {
        if (response.Cancelled)
        {
            return "@cancel";
        }
        return string.Join("\n", response.Values);
    }

    public static string BuildDefinedNameRefersTo(string value)
    {
        return "=\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }

    public static bool IsTransientRuntimeDefinedName(string name)
    {
        var normalized = (name ?? "").Trim();
        var separator = normalized.LastIndexOf('!');
        if (separator >= 0 && separator + 1 < normalized.Length)
        {
            normalized = normalized[(separator + 1)..];
        }

        if (normalized.Equals("__XLFLOW_MODE__", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("__XLFLOW_RUNTIME_VERSION__", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("__XLFLOW_DEBUG_PIPE__", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("__XLFLOW_RUN_HELPER__", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("__XLFLOW_UI_STREAM_HELPER__", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("__XLFLOW_UI_STREAM_REDACT_INPUT__", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return normalized.EndsWith("__", StringComparison.Ordinal) &&
            (normalized.StartsWith("__XLFLOW_UI_MSGBOX_", StringComparison.OrdinalIgnoreCase) ||
             normalized.StartsWith("__XLFLOW_UI_INPUT_", StringComparison.OrdinalIgnoreCase) ||
             normalized.StartsWith("__XLFLOW_UI_FILEDIALOG_", StringComparison.OrdinalIgnoreCase));
    }

    public static bool IsTemporaryRuntimeComponentName(string name)
    {
        var normalized = (name ?? "").Trim();
        return normalized.StartsWith("XlflowRun_", StringComparison.OrdinalIgnoreCase) ||
            normalized.StartsWith("XlflowUIStream_", StringComparison.OrdinalIgnoreCase);
    }

    public static int RemoveTransientRuntimeArtifacts(object workbook)
    {
        var transientNamesPresent = ContainsTransientRuntimeDefinedNames(workbook);
        var removedComponents = RemoveTemporaryRuntimeComponents(workbook, transientNamesPresent);
        var removedNames = RemoveTransientRuntimeDefinedNames(workbook);
        return removedComponents + removedNames;
    }

    private static bool ContainsTransientRuntimeDefinedNames(object workbook)
    {
        var names = ExcelBridgeSupport.Get(workbook, "Names")
            ?? throw new InvalidOperationException("Workbook names are unavailable.");
        try
        {
            var count = Convert.ToInt32(ExcelBridgeSupport.Get(names, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var definedName = ExcelBridgeSupport.Get(names, "Item", index);
                if (definedName is null)
                {
                    continue;
                }
                try
                {
                    var definedNameText = Convert.ToString(
                        ExcelBridgeSupport.Get(definedName, "Name"),
                        CultureInfo.InvariantCulture) ?? "";
                    if (IsTransientRuntimeDefinedName(definedNameText))
                    {
                        return true;
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(definedName);
                }
            }
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(names);
        }
    }

    private static int RemoveTemporaryRuntimeComponents(object workbook, bool accessRequired)
    {
        object? vbProject;
        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
        }
        catch when (!accessRequired)
        {
            return 0;
        }
        if (vbProject is null)
        {
            if (accessRequired)
            {
                throw new InvalidOperationException("VBProject is unavailable while cleaning transient xlflow runtime state.");
            }
            return 0;
        }
        object? components = null;
        var temporaryComponents = new List<object>();
        try
        {
            try
            {
                components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            }
            catch when (!accessRequired)
            {
                return 0;
            }
            if (components is null)
            {
                if (accessRequired)
                {
                    throw new InvalidOperationException("VBComponents are unavailable while cleaning transient xlflow runtime state.");
                }
                return 0;
            }
            var count = Convert.ToInt32(ExcelBridgeSupport.Get(components, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var component = ExcelBridgeSupport.Get(components, "Item", index);
                if (component is null)
                {
                    continue;
                }
                var componentName = Convert.ToString(
                    ExcelBridgeSupport.Get(component, "Name"),
                    CultureInfo.InvariantCulture) ?? "";
                if (IsTemporaryRuntimeComponentName(componentName))
                {
                    temporaryComponents.Add(component);
                }
                else
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }

            foreach (var component in temporaryComponents)
            {
                ExcelBridgeSupport.InvokeViaDynamic(components, "Remove", component);
            }
            return temporaryComponents.Count;
        }
        finally
        {
            foreach (var component in temporaryComponents)
            {
                ExcelBridgeSupport.ReleaseComObject(component);
            }
            ExcelBridgeSupport.ReleaseComObject(components);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }
    }

    public static int RemoveTransientRuntimeDefinedNames(object workbook)
    {
        var names = ExcelBridgeSupport.Get(workbook, "Names")
            ?? throw new InvalidOperationException("Workbook names are unavailable.");
        var transientNames = new List<object>();
        try
        {
            var count = Convert.ToInt32(ExcelBridgeSupport.Get(names, "Count"), CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                var definedName = ExcelBridgeSupport.Get(names, "Item", index);
                if (definedName is null)
                {
                    continue;
                }

                var definedNameText = Convert.ToString(
                    ExcelBridgeSupport.Get(definedName, "Name"),
                    CultureInfo.InvariantCulture) ?? "";
                if (IsTransientRuntimeDefinedName(definedNameText))
                {
                    transientNames.Add(definedName);
                }
                else
                {
                    ExcelBridgeSupport.ReleaseComObject(definedName);
                }
            }

            foreach (var definedName in transientNames)
            {
                ExcelBridgeSupport.InvokeMethod(definedName, "Delete");
            }

            return transientNames.Count;
        }
        finally
        {
            foreach (var definedName in transientNames)
            {
                ExcelBridgeSupport.ReleaseComObject(definedName);
            }
            ExcelBridgeSupport.ReleaseComObject(names);
        }
    }

    public static void SetDefinedName(object workbook, string name, string value)
    {
        var names = ExcelBridgeSupport.Get(workbook, "Names")
            ?? throw new InvalidOperationException("Workbook names are unavailable.");
        object? existing = null;
        try
        {
            existing = ExcelBridgeSupport.Get(names, "Item", name);
        }
        catch
        {
            // missing name
        }
        if (existing is not null)
        {
            ExcelBridgeSupport.InvokeMethod(existing, "Delete");
            ExcelBridgeSupport.ReleaseComObject(existing);
        }
        ExcelBridgeSupport.InvokeMethod(names, "Add", name, BuildDefinedNameRefersTo(value), false);
        ExcelBridgeSupport.ReleaseComObject(names);
    }

    public static DefinedNameSnapshot CaptureDefinedNameState(object workbook, string name)
    {
        var names = ExcelBridgeSupport.Get(workbook, "Names");
        if (names is null)
        {
            return new DefinedNameSnapshot(name, false, "");
        }
        try
        {
            object? existing = null;
            try
            {
                existing = ExcelBridgeSupport.Get(names, "Item", name);
            }
            catch
            {
                // missing name
            }
            if (existing is not null)
            {
                var refersTo = Convert.ToString(ExcelBridgeSupport.Get(existing, "RefersTo"), CultureInfo.InvariantCulture) ?? "";
                ExcelBridgeSupport.ReleaseComObject(existing);
                return new DefinedNameSnapshot(name, true, refersTo);
            }
            return new DefinedNameSnapshot(name, false, "");
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(names);
        }
    }

    public static void EnsureDefinedNameRestoration(
        object workbook,
        RuntimeInjectionState state,
        string name)
    {
        if (state.NameSnapshots.Any(snapshot =>
            snapshot.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }
        state.NameSnapshots.Add(CaptureDefinedNameState(workbook, name));
        state.RestoreRequired = true;
    }

    public static void RestoreDefinedName(object workbook, DefinedNameSnapshot snapshot)
    {
        try
        {
            var names = ExcelBridgeSupport.Get(workbook, "Names");
            if (names is null)
            {
                return;
            }
            object? existing = null;
            try
            {
                existing = ExcelBridgeSupport.Get(names, "Item", snapshot.Name);
            }
            catch
            {
                // missing name
            }
            if (existing is not null)
            {
                ExcelBridgeSupport.InvokeMethod(existing, "Delete");
                ExcelBridgeSupport.ReleaseComObject(existing);
            }
            if (snapshot.Existed)
            {
                ExcelBridgeSupport.InvokeMethod(names, "Add", snapshot.Name, snapshot.RefersTo, false);
            }
            ExcelBridgeSupport.ReleaseComObject(names);
        }
        catch
        {
            // best-effort restore
        }
    }

    public static string BuildUIStreamModuleCode(string pipeName)
    {
        var builder = new StringBuilder();
        builder.AppendLine("Option Explicit");
        builder.AppendLine();
        builder.AppendLine("#If VBA7 Then");
        builder.AppendLine("  Private Declare PtrSafe Function CreateFileW Lib \"kernel32\" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr");
        builder.AppendLine("  Private Declare PtrSafe Function WriteFile Lib \"kernel32\" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long");
        builder.AppendLine("  Private Declare PtrSafe Function CloseHandle Lib \"kernel32\" (ByVal hObject As LongPtr) As Long");
        builder.AppendLine("  Private Const INVALID_HANDLE_VALUE As LongPtr = -1");
        builder.AppendLine("#Else");
        builder.AppendLine("  Private Declare Function CreateFileW Lib \"kernel32\" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long");
        builder.AppendLine("  Private Declare Function WriteFile Lib \"kernel32\" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long");
        builder.AppendLine("  Private Declare Function CloseHandle Lib \"kernel32\" (ByVal hObject As Long) As Long");
        builder.AppendLine("  Private Const INVALID_HANDLE_VALUE As Long = -1");
        builder.AppendLine("#End If");
        builder.AppendLine();
        builder.AppendLine("Private Const GENERIC_WRITE As Long = &H40000000");
        builder.AppendLine("Private Const OPEN_EXISTING As Long = 3");
        builder.AppendLine("Private Const mPipeName As String = \"" + pipeName.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"");
        builder.AppendLine();
        builder.AppendLine("Public Sub EmitEvent(ByVal jsonText As String)");
        builder.AppendLine("  SendPipeText jsonText & vbLf");
        builder.AppendLine("End Sub");
        builder.AppendLine();
        builder.AppendLine("Private Sub SendPipeText(ByVal payload As String)");
        builder.AppendLine("  Dim bytesWritten As Long");
        builder.AppendLine("#If VBA7 Then");
        builder.AppendLine("  Dim pipeHandle As LongPtr");
        builder.AppendLine("#Else");
        builder.AppendLine("  Dim pipeHandle As Long");
        builder.AppendLine("#End If");
        builder.AppendLine("  pipeHandle = CreateFileW(StrPtr(mPipeName), GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)");
        builder.AppendLine("  If pipeHandle = INVALID_HANDLE_VALUE Then");
        builder.AppendLine("    Exit Sub");
        builder.AppendLine("  End If");
        builder.AppendLine("  On Error GoTo Cleanup");
        builder.AppendLine("  Call WriteFile(pipeHandle, StrPtr(payload), Len(payload) * 2, bytesWritten, 0)");
        builder.AppendLine("Cleanup:");
        builder.AppendLine("  On Error Resume Next");
        builder.AppendLine("  If pipeHandle <> INVALID_HANDLE_VALUE Then CloseHandle pipeHandle");
        builder.AppendLine("  On Error GoTo 0");
        builder.AppendLine("End Sub");
        return builder.ToString();
    }

    public static string AddUIStreamModule(object vbProject, string pipeName)
    {
        var moduleName = "XlflowUIStream_" + Guid.NewGuid().ToString("N")[..8];
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents")
            ?? throw new InvalidOperationException("VBComponents is unavailable.");
        var component = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
            ?? throw new InvalidOperationException("Could not add a temporary UI stream module.");
        SetProperty(component, "Name", moduleName);
        var codeModule = ExcelBridgeSupport.Get(component, "CodeModule")
            ?? throw new InvalidOperationException("CodeModule is unavailable.");
        ExcelBridgeSupport.InvokeMethod(codeModule, "AddFromString", BuildUIStreamModuleCode(pipeName));
        ExcelBridgeSupport.ReleaseComObject(codeModule);
        ExcelBridgeSupport.ReleaseComObject(component);
        ExcelBridgeSupport.ReleaseComObject(components);
        return moduleName;
    }

    public static void RemoveUIStreamModule(object vbProject, string moduleName)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            return;
        }
        try
        {
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is null)
            {
                return;
            }
            object? existing = null;
            try
            {
                existing = ExcelBridgeSupport.Get(components, "Item", moduleName);
            }
            catch
            {
                // module not found
            }
            if (existing is not null)
            {
                ExcelBridgeSupport.InvokeMethod(components, "Remove", existing);
                ExcelBridgeSupport.ReleaseComObject(existing);
            }
            ExcelBridgeSupport.ReleaseComObject(components);
        }
        catch
        {
            // best-effort cleanup
        }
    }

    public static RuntimeInjectionState ApplyRuntimeInjection(
        object workbook,
        string runtimeMode,
        string runtimeSource,
        string msgBoxResponsesJson,
        string inputResponsesJson,
        string fileDialogResponsesJson,
        bool debugStreamEnabled,
        string debugStreamPipeName,
        bool uiStreamEnabled,
        string uiStreamPipeName,
        bool uiStreamRedactInput)
    {
        var state = new RuntimeInjectionState();
        if (string.IsNullOrWhiteSpace(runtimeMode))
        {
            return state;
        }

        try
        {
            var modeName = NormalizeRuntimeMode(runtimeMode);
            var msgBoxResponses = DecodeUIResponsesJson(msgBoxResponsesJson);
            var inputResponses = DecodeUIResponsesJson(inputResponsesJson);
            var fileDialogResponses = DecodeFileDialogResponsesJson(fileDialogResponsesJson);

            state.Mode = modeName;
            state.Source = (runtimeSource ?? "default").Trim().ToLowerInvariant();

            // Capture existing state for all names we'll touch
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_MODE__"));
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_RUNTIME_VERSION__"));
            foreach (var entry in msgBoxResponses)
            {
                var name = GetUIResponseDefinedName("msgbox", entry.Key);
                state.NameSnapshots.Add(CaptureDefinedNameState(workbook, name));
            }
            foreach (var entry in inputResponses)
            {
                var name = GetUIResponseDefinedName("input", entry.Key);
                state.NameSnapshots.Add(CaptureDefinedNameState(workbook, name));
            }
            foreach (var entry in fileDialogResponses)
            {
                var name = GetFileDialogResponseDefinedName(entry.Kind, entry.DialogId);
                state.NameSnapshots.Add(CaptureDefinedNameState(workbook, name));
            }
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_DEBUG_PIPE__"));
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_RUN_HELPER__"));
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_UI_STREAM_HELPER__"));
            state.NameSnapshots.Add(CaptureDefinedNameState(workbook, "__XLFLOW_UI_STREAM_REDACT_INPUT__"));

            // Set all defined names
            state.RestoreRequired = true;
            SetDefinedName(workbook, "__XLFLOW_MODE__", modeName);
            SetDefinedName(workbook, "__XLFLOW_RUNTIME_VERSION__", "1");
            foreach (var entry in msgBoxResponses)
            {
                SetDefinedName(workbook, GetUIResponseDefinedName("msgbox", entry.Key), entry.Value);
            }
            foreach (var entry in inputResponses)
            {
                SetDefinedName(workbook, GetUIResponseDefinedName("input", entry.Key), entry.Value);
            }
            foreach (var entry in fileDialogResponses)
            {
                var markerValue = ConvertToFileDialogMarkerValue(entry);
                SetDefinedName(workbook, GetFileDialogResponseDefinedName(entry.Kind, entry.DialogId), markerValue);
            }
            if (debugStreamEnabled && !string.IsNullOrWhiteSpace(debugStreamPipeName))
            {
                SetDefinedName(workbook, "__XLFLOW_DEBUG_PIPE__", debugStreamPipeName);
            }

            state.Applied = true;
            state.DebugStreamEnabled = debugStreamEnabled && !string.IsNullOrWhiteSpace(debugStreamPipeName);
            state.DebugStreamPipeName = debugStreamPipeName;
            state.UIStreamEnabled = uiStreamEnabled && !string.IsNullOrWhiteSpace(uiStreamPipeName);
            state.UIStreamPipeName = uiStreamPipeName;
            state.UIStreamRedactInput = uiStreamRedactInput;
        }
        catch
        {
            RestoreRuntimeInjection(workbook, state);
            throw;
        }

        return state;
    }

    public static void EnableUIStreamInjection(object workbook, object vbProject, RuntimeInjectionState state)
    {
        if (!state.Applied || !state.UIStreamEnabled)
        {
            return;
        }
        if (vbProject is null || string.IsNullOrWhiteSpace(state.UIStreamPipeName))
        {
            return;
        }

        state.UIStreamModuleName = AddUIStreamModule(vbProject, state.UIStreamPipeName);
        SetDefinedName(workbook, "__XLFLOW_UI_STREAM_HELPER__", state.UIStreamModuleName + ".EmitEvent");
        SetDefinedName(workbook, "__XLFLOW_UI_STREAM_REDACT_INPUT__", state.UIStreamRedactInput.ToString().ToLowerInvariant());
    }

    public static void RestoreRuntimeInjection(object workbook, RuntimeInjectionState state)
    {
        if (!state.RestoreRequired)
        {
            return;
        }

        foreach (var snapshot in state.NameSnapshots.AsEnumerable().Reverse())
        {
            RestoreDefinedName(workbook, snapshot);
        }

        if (!string.IsNullOrWhiteSpace(state.UIStreamModuleName) && workbook is not null)
        {
            try
            {
                var vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
                if (vbProject is not null)
                {
                    RemoveUIStreamModule(vbProject, state.UIStreamModuleName);
                    ExcelBridgeSupport.ReleaseComObject(vbProject);
                }
            }
            catch
            {
                // best-effort
            }
        }

        state.Applied = false;
        state.RestoreRequired = false;
    }

    private static void SetProperty(object comObject, string name, object value)
    {
        comObject.GetType().InvokeMember(
            name,
            System.Reflection.BindingFlags.SetProperty,
            null,
            comObject,
            [value],
            CultureInfo.InvariantCulture);
    }

    public sealed record FileDialogResponse(string Kind, string DialogId, bool Cancelled, List<string> Values);

    public sealed record DefinedNameSnapshot(string Name, bool Existed, string RefersTo);

    public sealed class RuntimeInjectionState
    {
        public bool Applied { get; set; }
        public bool RestoreRequired { get; set; }
        public string Mode { get; set; } = "";
        public string Source { get; set; } = "";
        public bool DebugStreamEnabled { get; set; }
        public string DebugStreamPipeName { get; set; } = "";
        public bool UIStreamEnabled { get; set; }
        public string UIStreamPipeName { get; set; } = "";
        public bool UIStreamRedactInput { get; set; } = true;
        public string UIStreamModuleName { get; set; } = "";
        public List<DefinedNameSnapshot> NameSnapshots { get; } = [];
    }
}
