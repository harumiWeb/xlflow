using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelFormInspectionService : IInspectFormService
{
    public BridgeResponse Execute(BridgeRequest request, InspectFormCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        object? runtimeExcel = null;
        object? runtimeWorkbook = null;
        object? runtimeVBProject = null;
        object? helperComponent = null;
        string? runtimeWorkbookPath = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var basis = NormalizeBasis(args.Basis);
            if (basis is null)
            {
                return Failure(request, "inspect_form_args_invalid", $"unsupported inspect form basis: {args.Basis}", "xlflow");
            }
            if (basis == "designer" && !string.IsNullOrWhiteSpace(args.Initializer))
            {
                return Failure(request, "inspect_form_args_invalid", "--initializer can only be used with runtime or both inspection", "xlflow");
            }

            var openResult = OpenWorkbookForInspect(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible);
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            vbProject = ExcelBridgeSupport.TryGetWorkbookVbProject(workbook);
            if (vbProject is null)
            {
                return Failure(request, "vbproject_access_denied", "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.", "Excel");
            }

            if (!HasUserForm(vbProject, args.FormName))
            {
                return Failure(request, "form_not_found", $"UserForm '{args.FormName}' was not found in the workbook.", "xlflow");
            }

            var targetKind = sessionAttached ? "live_session" : "file";
            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var saveStateKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
            var dirty = sessionAttached ? saveStateKnown ? dirtyState : true : false;
            var needsSave = sessionAttached ? dirty : false;

            object formsPayload;
            if (basis == "runtime")
            {
                var runtimeOpen = OpenRuntimeWorkbookCopy(workbook, args.Visible);
                runtimeExcel = runtimeOpen.Excel;
                runtimeWorkbook = runtimeOpen.Workbook;
                runtimeWorkbookPath = runtimeOpen.Path;
                runtimeVBProject = ExcelBridgeSupport.TryGetWorkbookVbProject(runtimeWorkbook)
                    ?? throw new InvalidOperationException("runtime VBProject is unavailable");
                helperComponent = InstallInspectHelper(runtimeVBProject);
                var helperModuleName = ExcelBridgeSupport.GetString(helperComponent, "Name") ?? "";
                formsPayload = InvokeInspectHelper(runtimeExcel, runtimeWorkbook, helperModuleName, args.FormName, "runtime", args.Initializer);
            }
            else if (basis == "both")
            {
                var runtimeOpen = OpenRuntimeWorkbookCopy(workbook, args.Visible);
                runtimeExcel = runtimeOpen.Excel;
                runtimeWorkbook = runtimeOpen.Workbook;
                runtimeWorkbookPath = runtimeOpen.Path;
                runtimeVBProject = ExcelBridgeSupport.TryGetWorkbookVbProject(runtimeWorkbook)
                    ?? throw new InvalidOperationException("runtime VBProject is unavailable");
                helperComponent = InstallInspectHelper(runtimeVBProject);
                var helperModuleName = ExcelBridgeSupport.GetString(helperComponent, "Name") ?? "";
                var runtimeForm = InvokeInspectHelper(runtimeExcel, runtimeWorkbook, helperModuleName, args.FormName, "runtime", args.Initializer);
                var designerForm = args.StrictDesigner
                    ? InvokeInspectHelper(runtimeExcel, runtimeWorkbook, helperModuleName, args.FormName, "designer", "")
                    : ReadDesignerFormSnapshot(vbProject, args.FormName);
                formsPayload = new Dictionary<string, object?>
                {
                    ["runtime"] = runtimeForm,
                    ["designer"] = designerForm,
                };
            }
            else
            {
                if (args.StrictDesigner)
                {
                    var runtimeOpen = OpenRuntimeWorkbookCopy(workbook, args.Visible);
                    runtimeExcel = runtimeOpen.Excel;
                    runtimeWorkbook = runtimeOpen.Workbook;
                    runtimeWorkbookPath = runtimeOpen.Path;
                    runtimeVBProject = ExcelBridgeSupport.TryGetWorkbookVbProject(runtimeWorkbook)
                        ?? throw new InvalidOperationException("runtime VBProject is unavailable");
                    helperComponent = InstallInspectHelper(runtimeVBProject);
                    var helperModuleName = ExcelBridgeSupport.GetString(helperComponent, "Name") ?? "";
                    formsPayload = InvokeInspectHelper(runtimeExcel, runtimeWorkbook, helperModuleName, args.FormName, "designer", "");
                }
                else
                {
                    formsPayload = ReadDesignerFormSnapshot(vbProject, args.FormName);
                }
            }

            var warnings = new List<Dictionary<string, string>>();
            if (basis is "runtime" or "both")
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "runtime_form_loads_initialize",
                    ["message"] = "Runtime inspection loads the form and executes UserForm_Initialize.",
                });
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "runtime_form_temp_copy",
                    ["message"] = "Runtime inspection executed against a temporary workbook copy so the source workbook and live session are not mutated.",
                });
                if (!string.IsNullOrWhiteSpace(args.Initializer))
                {
                    warnings.Add(new Dictionary<string, string>
                    {
                        ["code"] = "runtime_form_initializer_invoked",
                        ["message"] = $"Runtime inspection also invoked {args.Initializer}(ThisWorkbook).",
                    });
                }
            }
            if (needsSave)
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                });
            }

            var target = new Dictionary<string, object?>
            {
                ["kind"] = targetKind,
                ["path"] = workbookPath,
            };
            if (basis is "runtime" or "both")
            {
                target["note"] = "Runtime inspection used a temporary workbook copy.";
            }
            else if (args.StrictDesigner)
            {
                target["note"] = "Strict designer inspection used a temporary workbook copy plus helper module to recover concrete control types.";
            }

            var extensions = new Dictionary<string, object?>
            {
                ["target"] = target,
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = sessionAttached,
                    ["workbook_path"] = workbookPath,
                    ["dirty"] = dirty,
                    ["save_required"] = needsSave,
                    ["live_newer_than_disk"] = needsSave,
                    ["mode"] = sessionMode,
                    ["source_of_truth"] = needsSave ? "live_workbook" : "saved_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = workbookPath,
                    ["session"] = sessionAttached,
                    ["session_mode"] = sessionMode,
                    ["session_requested"] = sessionAttached && args.UseSession,
                    ["auto_session"] = sessionAttached && !args.UseSession,
                    ["saved"] = false,
                    ["dirty"] = dirty,
                    ["needs_save"] = needsSave,
                },
                ["forms"] = formsPayload,
            };
            if (warnings.Count > 0)
            {
                extensions["warnings"] = warnings;
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = [sessionAttached ? $"attached to xlflow session ({sessionMode})" : $"inspected {basis} UserForm {args.FormName}"],
                Extensions = extensions,
            };
        }
        catch (JsonException ex)
        {
            return Failure(request, "inspect_form_failed", ex.Message, "xlflow-excel-bridge");
        }
        catch (InvalidOperationException ex)
        {
            var code = ClassifyInspectErrorCode(ex.Message);
            return Failure(request, code, ex.Message, "xlflow-excel-bridge");
        }
        catch (Exception ex)
        {
            return Failure(request, "inspect_form_failed", ExcelBridgeSupport.FormatExceptionDetail(ex), "xlflow-excel-bridge");
        }
        finally
        {
            RemoveTemporaryComponent(runtimeVBProject, helperComponent);
            CloseWorkbook(runtimeWorkbook, runtimeExcel);
            if (!string.IsNullOrWhiteSpace(runtimeWorkbookPath) && File.Exists(runtimeWorkbookPath))
            {
                try
                {
                    File.Delete(runtimeWorkbookPath);
                }
                catch
                {
                    // best-effort cleanup
                }
            }

            ExcelBridgeSupport.ReleaseComObject(runtimeVBProject);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
            if (sessionAttached)
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }
            else
            {
                CloseWorkbook(workbook, excel);
            }
        }
    }

    private static string? NormalizeBasis(string basis)
    {
        var normalized = (basis ?? "").Trim().ToLowerInvariant();
        return normalized switch
        {
            "runtime" => "runtime",
            "designer" => "designer",
            "both" => "both",
            _ => null,
        };
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message, string source)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: code,
            Message: message,
            Phase: "inspect-form",
            Source: source));
    }

    private static string ClassifyInspectErrorCode(string message)
    {
        if (message.Contains("runtime_load", StringComparison.OrdinalIgnoreCase))
        {
            return "runtime_form_load_failed";
        }
        if (message.Contains("initializer", StringComparison.OrdinalIgnoreCase))
        {
            return "form_initializer_failed";
        }
        if (message.Contains("enumerate", StringComparison.OrdinalIgnoreCase))
        {
            return "control_enumeration_failed";
        }
        if (message.Contains("designer", StringComparison.OrdinalIgnoreCase))
        {
            return "designer_access_failed";
        }
        if (message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            return "session_required";
        }
        return "inspect_form_failed";
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForInspect(
        string workbookPath, string metadataPath, bool useSession, bool visible)
    {
        if (useSession)
        {
            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, true);
            return (attachment.Excel, attachment.Workbook, true, attachment.SessionMode);
        }

        if (ExcelBridgeSupport.SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            try
            {
                var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, false);
                return (attachment.Excel, attachment.Workbook, true, attachment.SessionMode);
            }
            catch
            {
                // fall through to direct open
            }
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, "none");
    }

    private static (object Excel, object Workbook, string Path) OpenRuntimeWorkbookCopy(object sourceWorkbook, bool visible)
    {
        var sourceFullName = ExcelBridgeSupport.TryGetWorkbookFullName(sourceWorkbook) ?? "";
        var extension = Path.GetExtension(sourceFullName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".xlsm";
        }

        var tempPath = Path.Combine(Path.GetTempPath(), $"xlflow-inspect-form-{Guid.NewGuid():N}{extension}");
        ExcelBridgeSupport.InvokeViaDynamic(sourceWorkbook, "SaveCopyAs", tempPath);
        var direct = ExcelBridgeSupport.OpenWorkbookDirect(tempPath, visible, disableAutomationMacros: false);
        return (direct.Excel, direct.Workbook, tempPath);
    }

    private static bool HasUserForm(object vbProject, string formName)
    {
        object? components = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components!, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(components!, "Item", index);
                    var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component!, "Type"));
                    var name = ExcelBridgeSupport.GetString(component!, "Name") ?? "";
                    if (type == 3 && string.Equals(name, formName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
        }

        return false;
    }

    private static Dictionary<string, object?> ReadDesignerFormSnapshot(object vbProject, string formName)
    {
        object? components = null;
        object? component = null;
        object? designer = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            component = ExcelBridgeSupport.Get(components!, "Item", formName)
                ?? throw new InvalidOperationException($"XlflowInspectFormJson.designer: UserForm '{formName}' was not found.");
            designer = ExcelBridgeSupport.Get(component!, "Designer")
                ?? throw new InvalidOperationException($"XlflowInspectFormJson.designer: failed to access Designer for '{formName}'.");

            var rootControls = GetChildControls(designer, formName);
            return new Dictionary<string, object?>
            {
                ["name"] = formName,
                ["basis"] = "designer",
                ["caption"] = ExcelBridgeSupport.GetString(designer, "Caption") ?? "",
                ["width"] = TryGetDouble(designer, "Width"),
                ["height"] = TryGetDouble(designer, "Height"),
                ["coordinate_system"] = "parent-relative",
                ["controls"] = rootControls,
            };
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(designer);
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static object InstallInspectHelper(object runtimeVBProject)
    {
        object? components = null;
        object? component = null;
        object? codeModule = null;
        try
        {
            components = ExcelBridgeSupport.Get(runtimeVBProject, "VBComponents")
                ?? throw new InvalidOperationException("VBComponents is unavailable.");
            component = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
                ?? throw new InvalidOperationException("could not add inspect helper module.");
            SetProperty(component, "Name", "XlflowForm_" + Guid.NewGuid().ToString("N")[..20]);
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule")
                ?? throw new InvalidOperationException("inspect helper CodeModule is unavailable.");
            ExcelBridgeSupport.InvokeMethod(codeModule, "AddFromString", BuildInspectHelperCode());
            return component;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static object InvokeInspectHelper(object runtimeExcel, object runtimeWorkbook, string helperModuleName, string formName, string basis, string initializer)
    {
        var workbookName = (ExcelBridgeSupport.TryGetWorkbookName(runtimeWorkbook) ?? "").Replace("'", "''", StringComparison.Ordinal);
        var moduleName = string.IsNullOrWhiteSpace(helperModuleName) ? "XlflowForm" : helperModuleName;
        var macroName = $"'{workbookName}'!{moduleName}.XlflowInspectFormJson";
        try
        {
            var json = Convert.ToString(ExcelBridgeSupport.RunExcelMacro(runtimeExcel, macroName, formName, basis, initializer), CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new InvalidOperationException("inspect form helper returned no JSON");
            }

            using var document = JsonDocument.Parse(json);
            return JsonSerializer.Deserialize<object>(document.RootElement.GetRawText()) ?? new Dictionary<string, object?>();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(ExcelBridgeSupport.FormatExceptionDetail(ex), ex);
        }
    }

    internal static List<Dictionary<string, object?>> GetChildControls(object parent, string expectedParentName)
    {
        var controls = new List<Dictionary<string, object?>>();
        object? children = null;
        try
        {
            try
            {
                children = ExcelBridgeSupport.Get(parent, "Controls");
            }
            catch
            {
                return controls;
            }
            if (children is null)
            {
                return controls;
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(children, "Count"));
            for (var index = 0; index < count; index++)
            {
                object? control = null;
                try
                {
                    control = ExcelBridgeSupport.Get(children, "Item", index);
                    if (control is null || !HasExpectedParent(control, expectedParentName))
                    {
                        continue;
                    }
                    controls.Add(SerializeControl(control));
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(control);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(children);
        }

        return controls;
    }

    private static Dictionary<string, object?> SerializeControl(object control)
    {
        var progId = TryGetStringMember(control, "ProgId") ?? "";
        var fallbackTypeName = TryGetComControlTypeName(control) ?? control.GetType().Name;
        var result = new Dictionary<string, object?>
        {
            ["name"] = TryGetStringMember(control, "Name") ?? "",
            ["type"] = ResolveDesignerControlType(progId, fallbackTypeName),
        };

        if (!string.IsNullOrWhiteSpace(progId))
        {
            result["prog_id"] = progId;
        }
        AddStringMember(control, result, "Caption", "caption");
        AddStringMember(control, result, "Text", "text");
        AddStringMember(control, result, "Value", "value");
        AddDoubleMember(control, result, "Left", "left");
        AddDoubleMember(control, result, "Top", "top");
        AddDoubleMember(control, result, "Width", "width");
        AddDoubleMember(control, result, "Height", "height");
        AddIntMember(control, result, "TabIndex", "tab_index");
        AddBoolMember(control, result, "Enabled", "enabled");
        AddBoolMember(control, result, "Visible", "visible");
        AddIntMember(control, result, "ListIndex", "selected_index");

        var list = TryGetList(control);
        if (list.Count > 0)
        {
            result["list"] = list;
        }

        if (ControlCanContainChildren(Convert.ToString(result["type"], CultureInfo.InvariantCulture) ?? ""))
        {
            var name = Convert.ToString(result["name"], CultureInfo.InvariantCulture) ?? "";
            var children = GetChildControls(control, name);
            if (children.Count > 0)
            {
                result["controls"] = children;
            }
        }

        return result;
    }

    internal static string ResolveDesignerControlType(string? progId, string? fallbackTypeName)
    {
        if (!string.IsNullOrWhiteSpace(progId))
        {
            var segments = progId.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (segments.Length >= 2 && !string.IsNullOrWhiteSpace(segments[1]))
            {
                return segments[1];
            }
        }

        return string.IsNullOrWhiteSpace(fallbackTypeName) ? "Control" : fallbackTypeName;
    }

    internal static string? TryGetComControlTypeName(object control)
    {
        return NormalizeComTypeName(TryGetComTypeInfoName(control))
            ?? NormalizeComTypeName(TryGetTypeDescriptorClassName(control));
    }

    internal static string? NormalizeComTypeName(string? value)
    {
        var normalized = (value ?? "").Trim();
        if (normalized.Length == 0 || string.Equals(normalized, "__ComObject", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (normalized.StartsWith("MSForms.", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized["MSForms.".Length..];
        }
        if (normalized.EndsWith("Class", StringComparison.OrdinalIgnoreCase) && normalized.Length > "Class".Length)
        {
            normalized = normalized[..^"Class".Length];
        }
        while (normalized.StartsWith('_') && normalized.Length > 1)
        {
            normalized = normalized[1..];
        }
        if (KnownMsFormsInterfaceTypeNames.TryGetValue(normalized, out var controlType))
        {
            return controlType;
        }
        if (normalized.StartsWith('I') && normalized.Length > 1 && char.IsUpper(normalized[1]))
        {
            normalized = normalized[1..];
        }

        return normalized.Length == 0 ? null : normalized;
    }

    private static readonly Dictionary<string, string> KnownMsFormsInterfaceTypeNames = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ICommandButton"] = "CommandButton",
        ["IImage"] = "Image",
        ["ILabelControl"] = "Label",
        ["IMdcCheckBox"] = "CheckBox",
        ["IMdcCombo"] = "ComboBox",
        ["IMdcList"] = "ListBox",
        ["IMdcOptionButton"] = "OptionButton",
        ["IMdcText"] = "TextBox",
        ["IMdcToggleButton"] = "ToggleButton",
        ["IMultiPage"] = "MultiPage",
        ["IOptionFrame"] = "Frame",
        ["IPage"] = "Page",
        ["IScrollbar"] = "ScrollBar",
        ["ISpinbutton"] = "SpinButton",
        ["ITabStrip"] = "TabStrip",
    };

    private static string? TryGetTypeDescriptorClassName(object control)
    {
        try
        {
            return TypeDescriptor.GetClassName(control);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetComTypeInfoName(object control)
    {
        if (!Marshal.IsComObject(control))
        {
            return null;
        }

        IntPtr dispatchPtr = IntPtr.Zero;
        object? dispatchObject = null;
        ITypeInfo? typeInfo = null;
        try
        {
            dispatchPtr = Marshal.GetIDispatchForObject(control);
            if (dispatchPtr == IntPtr.Zero)
            {
                return null;
            }

            dispatchObject = Marshal.GetTypedObjectForIUnknown(dispatchPtr, typeof(IDispatch));
            if (dispatchObject is not IDispatch dispatch)
            {
                return null;
            }

            if (dispatch.GetTypeInfo(0, 0, out typeInfo) != 0 || typeInfo is null)
            {
                return null;
            }

            typeInfo.GetDocumentation(-1, out var name, out _, out _, out _);
            return name;
        }
        catch
        {
            return null;
        }
        finally
        {
            if (typeInfo is not null && Marshal.IsComObject(typeInfo))
            {
                Marshal.ReleaseComObject(typeInfo);
            }
            if (dispatchObject is not null && Marshal.IsComObject(dispatchObject))
            {
                Marshal.ReleaseComObject(dispatchObject);
            }
            if (dispatchPtr != IntPtr.Zero)
            {
                Marshal.Release(dispatchPtr);
            }
        }
    }

    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount(out uint pctinfo);

        [PreserveSig]
        int GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo);
    }

    private static bool HasExpectedParent(object control, string expectedParentName)
    {
        object? parent = null;
        try
        {
            parent = ExcelBridgeSupport.Get(control, "Parent");
            var parentName = ExcelBridgeSupport.GetString(parent!, "Name") ?? "";
            return string.Equals(parentName.Trim(), expectedParentName.Trim(), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(parent);
        }
    }

    private static bool ControlCanContainChildren(string controlType)
    {
        return controlType.Trim().ToLowerInvariant() is "frame" or "multipage" or "page" or "tabstrip";
    }

    private static List<string> TryGetList(object control)
    {
        var items = new List<string>();
        int listCount;
        try
        {
            listCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(control, "ListCount"));
        }
        catch
        {
            return items;
        }

        for (var index = 0; index < listCount; index++)
        {
            try
            {
                var item = ExcelBridgeSupport.Get(control, "List", index);
                items.Add(item?.ToString() ?? "");
            }
            catch
            {
                break;
            }
        }

        return items;
    }

    private static void AddStringMember(object source, Dictionary<string, object?> target, string memberName, string jsonKey)
    {
        var value = TryGetStringMember(source, memberName);
        if (value is not null)
        {
            target[jsonKey] = value;
        }
    }

    private static string? TryGetStringMember(object source, string memberName)
    {
        try
        {
            var value = ExcelBridgeSupport.Get(source, memberName);
            if (value is null)
            {
                return null;
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
        catch
        {
            return null;
        }
    }

    private static void AddDoubleMember(object source, Dictionary<string, object?> target, string memberName, string jsonKey)
    {
        var value = TryGetDouble(source, memberName);
        if (value is not null)
        {
            target[jsonKey] = value.Value;
        }
    }

    private static void AddIntMember(object source, Dictionary<string, object?> target, string memberName, string jsonKey)
    {
        try
        {
            var raw = ExcelBridgeSupport.Get(source, memberName);
            if (raw is null)
            {
                return;
            }
            target[jsonKey] = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
        }
        catch
        {
            // best-effort read
        }
    }

    private static void AddBoolMember(object source, Dictionary<string, object?> target, string memberName, string jsonKey)
    {
        try
        {
            var raw = ExcelBridgeSupport.Get(source, memberName);
            if (raw is null)
            {
                return;
            }
            target[jsonKey] = Convert.ToBoolean(raw, CultureInfo.InvariantCulture);
        }
        catch
        {
            // best-effort read
        }
    }

    private static double? TryGetDouble(object source, string memberName)
    {
        try
        {
            var raw = ExcelBridgeSupport.Get(source, memberName);
            if (raw is null)
            {
                return null;
            }
            return Convert.ToDouble(raw, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
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

    private static void RemoveTemporaryComponent(object? vbProject, object? component)
    {
        object? components = null;
        try
        {
            if (vbProject is null || component is null)
            {
                return;
            }
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is not null)
            {
                ExcelBridgeSupport.InvokeViaDynamic(components, "Remove", component);
            }
        }
        catch
        {
            // best-effort cleanup
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void CloseWorkbook(object? workbook, object? excel)
    {
        if (workbook is not null)
        {
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            }
            catch
            {
                // best-effort close
            }
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }

        if (excel is not null)
        {
            try
            {
                dynamic app = excel;
                app.Quit();
            }
            catch
            {
                // best-effort quit
            }
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static string BuildInspectHelperCode()
    {
        return """""
Option Explicit

Private Const xlflowBasisDesigner As String = "designer"
Private Const xlflowBasisRuntime As String = "runtime"
Private Const xlflowCoordinateSystem As String = "parent-relative"

Public Function XlflowInspectFormJson(ByVal formName As String, ByVal basis As String, Optional ByVal initializer As String = "") As String
  Dim normalizedBasis As String
  normalizedBasis = LCase$(Trim$(basis))

  Select Case normalizedBasis
    Case xlflowBasisDesigner
      XlflowInspectFormJson = InspectDesignerFormJson(formName)
    Case xlflowBasisRuntime
      XlflowInspectFormJson = InspectRuntimeFormJson(formName, initializer)
    Case Else
      Err.Raise vbObjectError + 7300, "XlflowInspectFormJson.args", "Unsupported inspect basis: " & basis
  End Select
End Function

Private Function InspectDesignerFormJson(ByVal formName As String) As String
  On Error GoTo ErrHandler

  Dim component As Object
  Dim designer As Object

  Set component = ThisWorkbook.VBProject.VBComponents.Item(formName)
  Set designer = component.Designer
  InspectDesignerFormJson = SerializeFormSnapshot(formName, xlflowBasisDesigner, designer)
  Exit Function

ErrHandler:
  Err.Raise Err.Number, "XlflowInspectFormJson.designer", Err.Description
End Function

Private Function InspectRuntimeFormJson(ByVal formName As String, ByVal initializer As String) As String
  Dim formInstance As Object
  Dim loaded As Boolean
  Dim initializerRan As Boolean
  Dim errorNumber As Long
  Dim errorDescription As String
  Dim errorSource As String

  On Error GoTo ErrHandler

  Set formInstance = UserForms.Add(formName)
  loaded = True

  If Len(Trim$(initializer)) > 0 Then
    CallByName formInstance, Trim$(initializer), VbMethod, ThisWorkbook
    initializerRan = True
  End If

  InspectRuntimeFormJson = SerializeFormSnapshot(formName, xlflowBasisRuntime, formInstance)

Cleanup:
  On Error Resume Next
  If loaded Then
    Unload formInstance
  End If
  Set formInstance = Nothing
  On Error GoTo 0

  If errorNumber <> 0 Then
    Err.Raise errorNumber, errorSource, errorDescription
  End If
  Exit Function

ErrHandler:
  errorNumber = Err.Number
  errorDescription = Err.Description
  If Not loaded Then
    errorSource = "XlflowInspectFormJson.runtime_load"
  ElseIf Len(Trim$(initializer)) > 0 And Not initializerRan Then
    errorSource = "XlflowInspectFormJson.initializer"
  Else
    errorSource = "XlflowInspectFormJson.enumerate"
  End If
  Resume Cleanup
End Function

Private Function SerializeFormSnapshot(ByVal formName As String, ByVal basis As String, ByVal formObject As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim controls As Object

  json = "{"

  JsonAddString json, "name", formName, hasFields
  JsonAddString json, "basis", basis, hasFields
  JsonAddStringFromMember json, formObject, "Caption", "caption", hasFields
  JsonAddNumberFromMember json, formObject, "Width", "width", hasFields
  JsonAddNumberFromMember json, formObject, "Height", "height", hasFields
  JsonAddString json, "coordinate_system", xlflowCoordinateSystem, hasFields

  Set controls = GetObjectControls(formObject)
  JsonAddRaw json, "controls", SerializeControls(controls, formName), hasFields

  json = json & "}"
  SerializeFormSnapshot = json
End Function

Private Function SerializeControls(ByVal controls As Object, ByVal expectedParentName As String) As String
  Dim json As String
  Dim first As Boolean
  Dim control As Object

  json = "["
  first = True

  If Not controls Is Nothing Then
    For Each control In controls
      If Not ControlHasExpectedParent(control, expectedParentName) Then
        GoTo ContinueLoop
      End If
      If Not first Then
        json = json & ","
      End If
      json = json & SerializeControl(control)
      first = False
ContinueLoop:
    Next control
  End If

  json = json & "]"
  SerializeControls = json
End Function

Private Function SerializeControl(ByVal control As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim children As Object
  Dim listCount As Long
  Dim selectedIndex As Long
  Dim typeNameValue As String

  json = "{"

  JsonAddStringFromMember json, control, "Name", "name", hasFields
  typeNameValue = TypeName(control)
  JsonAddString json, "type", typeNameValue, hasFields
  JsonAddStringFromMember json, control, "ProgId", "prog_id", hasFields
  JsonAddStringFromMember json, control, "Caption", "caption", hasFields
  JsonAddFocusSafeText json, control, typeNameValue, hasFields
  JsonAddStringFromMember json, control, "Value", "value", hasFields
  JsonAddNumberFromMember json, control, "Left", "left", hasFields
  JsonAddNumberFromMember json, control, "Top", "top", hasFields
  JsonAddNumberFromMember json, control, "Width", "width", hasFields
  JsonAddNumberFromMember json, control, "Height", "height", hasFields
  JsonAddLongFromMember json, control, "TabIndex", "tab_index", hasFields
  JsonAddBoolFromMember json, control, "Enabled", "enabled", hasFields
  JsonAddBoolFromMember json, control, "Visible", "visible", hasFields

  If TryGetLongMember(control, "ListIndex", selectedIndex) Then
    JsonAddLong json, "selected_index", selectedIndex, hasFields
  End If
  If TryGetLongMember(control, "ListCount", listCount) Then
    JsonAddRaw json, "list", SerializeControlList(control, listCount), hasFields
  End If

  If ControlCanContainChildren(typeNameValue) Then
    Set children = GetObjectControls(control)
  End If
  If Not children Is Nothing Then
    JsonAddRaw json, "controls", SerializeControls(children, SafeControlName(control)), hasFields
  End If

  json = json & "}"
  SerializeControl = json
End Function

Private Function ControlHasExpectedParent(ByVal control As Object, ByVal expectedParentName As String) As Boolean
  On Error GoTo Missing

  Dim parentObject As Object
  Dim parentName As String
  Set parentObject = CallByName(control, "Parent", VbGet)
  If parentObject Is Nothing Then
    GoTo Missing
  End If

  parentName = CallByName(parentObject, "Name", VbGet)
  ControlHasExpectedParent = (StrComp(Trim$(parentName), Trim$(expectedParentName), vbTextCompare) = 0)
  Exit Function

Missing:
  ControlHasExpectedParent = False
End Function

Private Function SafeControlName(ByVal control As Object) As String
  On Error GoTo Missing

  SafeControlName = CStr(CallByName(control, "Name", VbGet))
  Exit Function

Missing:
  SafeControlName = vbNullString
End Function

Private Function ControlCanContainChildren(ByVal controlType As String) As Boolean
  Select Case LCase$(Trim$(controlType))
    Case "frame", "multipage", "page", "tabstrip"
      ControlCanContainChildren = True
    Case Else
      ControlCanContainChildren = False
  End Select
End Function

Private Function SerializeControlList(ByVal control As Object, ByVal listCount As Long) As String
  Dim json As String
  Dim i As Long
  Dim first As Boolean
  Dim itemValue As String

  json = "["
  first = True

  If listCount > 0 Then
    For i = 0 To listCount - 1
      If TryGetListItem(control, i, itemValue) Then
        If Not first Then
          json = json & ","
        End If
        json = json & JsonQuote(itemValue)
        first = False
      End If
    Next i
  End If

  json = json & "]"
  SerializeControlList = json
End Function

Private Function TryGetListItem(ByVal control As Object, ByVal index As Long, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim itemValue As Variant
  itemValue = control.List(index)
  If IsNull(itemValue) Or IsEmpty(itemValue) Then
    valueOut = vbNullString
  Else
    valueOut = CStr(itemValue)
  End If
  TryGetListItem = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetListItem = False
End Function

Private Function GetObjectControls(ByVal target As Object) As Object
  On Error Resume Next
  Set GetObjectControls = target.Controls
  On Error GoTo 0
End Function

Private Sub JsonAddStringFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As String
  If TryGetStringMember(target, memberName, value) Then
    JsonAddString json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddFocusSafeText(ByRef json As String, ByVal target As Object, ByVal controlType As String, ByRef hasFields As Boolean)
  Dim value As String
  Select Case LCase$(Trim$(controlType))
    Case "textbox", "combobox"
      If TryGetStringMember(target, "Value", value) Then
        JsonAddString json, "text", value, hasFields
      ElseIf TryGetStringMember(target, "Text", value) Then
        JsonAddString json, "text", value, hasFields
      End If
  End Select
End Sub

Private Sub JsonAddNumberFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Double
  If TryGetNumberMember(target, memberName, value) Then
    JsonAddNumber json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddLongFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Long
  If TryGetLongMember(target, memberName, value) Then
    JsonAddLong json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddBoolFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Boolean
  If TryGetBoolMember(target, memberName, value) Then
    JsonAddBool json, jsonKey, value, hasFields
  End If
End Sub

Private Function TryGetStringMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CStr(rawValue)
  TryGetStringMember = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetStringMember = False
End Function

Private Function TryGetNumberMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Double) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CDbl(rawValue)
  TryGetNumberMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetNumberMember = False
End Function

Private Function TryGetLongMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Long) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CLng(rawValue)
  TryGetLongMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetLongMember = False
End Function

Private Function TryGetBoolMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Boolean) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CBool(rawValue)
  TryGetBoolMember = True
  Exit Function

Missing:
  valueOut = False
  TryGetBoolMember = False
End Function

Private Sub JsonAddRaw(ByRef json As String, ByVal key As String, ByVal rawValue As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & rawValue
  hasFields = True
End Sub

Private Sub JsonAddString(ByRef json As String, ByVal key As String, ByVal value As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & JsonQuote(value)
  hasFields = True
End Sub

Private Sub JsonAddNumber(ByRef json As String, ByVal key As String, ByVal value As Double, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & Trim$(Str$(value))
  hasFields = True
End Sub

Private Sub JsonAddLong(ByRef json As String, ByVal key As String, ByVal value As Long, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & CStr(value)
  hasFields = True
End Sub

Private Sub JsonAddBool(ByRef json As String, ByVal key As String, ByVal value As Boolean, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & IIf(value, "true", "false")
  hasFields = True
End Sub

Private Function JsonQuote(ByVal value As String) As String
  JsonQuote = """" & JsonEscape(value) & """"
End Function

Private Function JsonEscape(ByVal value As String) As String
  Dim text As String
  text = value
  text = Replace(text, "\", "\\")
  text = Replace(text, """", Chr$(92) & Chr$(34))
  text = Replace(text, vbCrLf, "\n")
  text = Replace(text, vbCr, "\n")
  text = Replace(text, vbLf, "\n")
  text = Replace(text, vbTab, "\t")
  JsonEscape = text
End Function
""""";
    }
}
