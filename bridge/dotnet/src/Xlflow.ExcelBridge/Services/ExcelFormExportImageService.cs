using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Windows;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelFormExportImageService : IFormExportImageService
{
    public BridgeResponse Execute(BridgeRequest request, FormExportImageCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        object? runtimeExcel = null;
        object? runtimeWorkbook = null;
        object? runtimeVbProject = null;
        object? helperComponent = null;
        var runtimeWorkbookPath = "";
        var outputPath = "";
        var temporaryOutputPath = "";
        var helperModuleName = "";
        var createdParentDirs = false;
        var sessionAttached = false;
        var sessionMode = "none";
        var dirty = false;
        var needsSave = false;
        var phase = "validate_args";

        try
        {
            outputPath = PrepareOutputPath(args.OutputPath, args.Overwrite, out createdParentDirs, out temporaryOutputPath);

            phase = "open_source_workbook";
            var open = OpenWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible);
            excel = open.Excel;
            workbook = open.Workbook;
            sessionAttached = open.SessionAttached;
            sessionMode = open.SessionMode;
            if (sessionAttached && ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out dirty))
            {
                needsSave = dirty;
            }

            phase = "read_vbproject";
            vbProject = ExcelBridgeSupport.TryGetWorkbookVbProject(workbook);
            if (vbProject is null)
            {
                return Failure(request, "vbproject_access_denied", "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.", "Excel", phase, args, sessionAttached, sessionMode, dirty, needsSave);
            }
            if (!HasUserForm(vbProject, args.FormName))
            {
                return Failure(request, "form_not_found", $"UserForm '{args.FormName}' was not found in the workbook.", "xlflow", phase, args, sessionAttached, sessionMode, dirty, needsSave);
            }

            phase = "open_runtime_copy";
            var runtimeOpen = OpenRuntimeWorkbookCopy(workbook);
            runtimeExcel = runtimeOpen.Excel;
            runtimeWorkbook = runtimeOpen.Workbook;
            runtimeWorkbookPath = runtimeOpen.Path;
            runtimeVbProject = ExcelBridgeSupport.TryGetWorkbookVbProject(runtimeWorkbook)
                ?? throw new InvalidOperationException("vbproject_access_denied: runtime VBProject is unavailable");
            helperComponent = InstallHelper(runtimeVbProject);
            helperModuleName = ExcelBridgeSupport.GetString(helperComponent, "Name") ?? "";

            var processId = ExcelBridgeSupport.GetExcelProcessId(runtimeExcel);
            var token = $"xlflow-capture-{Guid.NewGuid():N}";

            phase = "schedule_form_capture";
            InvokePrepareCapture(runtimeExcel, runtimeWorkbook, helperModuleName, args.FormName, token, args.Initializer);

            phase = "find_form_window";
            var capture = WaitForCaptureWindow(runtimeExcel, runtimeWorkbook, helperModuleName, processId, token, cancellationToken);
            var window = WindowCapture.MoveWindowIntoCaptureBounds(capture)
                ?? throw new InvalidOperationException($"window_not_found: could not find a visible UserForm window for capture token {token}");

            phase = "capture_window_image";
            var imageInfo = WindowCapture.CaptureWindowImage(window.Hwnd, outputPath);
            if (!string.IsNullOrWhiteSpace(temporaryOutputPath))
            {
                File.Move(outputPath, args.OutputPath, true);
                outputPath = args.OutputPath;
                temporaryOutputPath = "";
            }

            var targetPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var output = new Dictionary<string, object?>
            {
                ["path"] = outputPath,
                ["format"] = "png",
                ["width_px"] = imageInfo.WidthPx,
                ["height_px"] = imageInfo.HeightPx,
            };
            if (createdParentDirs)
            {
                output["created_parent_dirs"] = true;
            }

            var target = new Dictionary<string, object?>
            {
                ["kind"] = sessionAttached ? "live_session" : "file",
                ["path"] = targetPath,
                ["form"] = args.FormName,
                ["capture_state"] = "temporary_copy",
                ["note"] = "Runtime export used a temporary workbook copy.",
                ["capture_window"] = new Dictionary<string, object?>
                {
                    ["hwnd"] = window.Hwnd,
                    ["title"] = window.Title,
                    ["class_name"] = window.ClassName,
                    ["left"] = window.Left,
                    ["top"] = window.Top,
                    ["width"] = window.Width,
                    ["height"] = window.Height,
                },
            };

            var warnings = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["code"] = "runtime_form_loads_initialize",
                    ["message"] = "Form image export loads the form at runtime and executes UserForm_Initialize.",
                },
                new()
                {
                    ["code"] = "runtime_form_temp_copy",
                    ["message"] = "Form image export executed against a temporary workbook copy so the source workbook and live session are not mutated.",
                },
                new()
                {
                    ["code"] = "userform_image_export_experimental",
                    ["message"] = "UserForm image export is experimental and currently supports Windows desktop Excel only.",
                },
            };
            if (!string.IsNullOrWhiteSpace(args.Initializer))
            {
                warnings.Add(new Dictionary<string, object?>
                {
                    ["code"] = "runtime_form_initializer_invoked",
                    ["message"] = $"Form image export also invoked {args.Initializer}(ThisWorkbook).",
                });
            }
            if (needsSave)
            {
                warnings.Add(new Dictionary<string, object?>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live workbook is newer than disk. `form export-image` used the live workbook state, not the saved workbook file.",
                });
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs =
                [
                    sessionAttached ? $"attached to xlflow session ({sessionMode})" : $"opened workbook {targetPath}",
                    $"exported runtime UserForm {args.FormName} to {outputPath}",
                ],
                Extensions = new Dictionary<string, object?>
                {
                    ["target"] = target,
                    ["session"] = SessionPayload(targetPath, sessionAttached, sessionMode, dirty, needsSave),
                    ["workbook"] = WorkbookPayload(targetPath, sessionAttached, sessionMode, dirty, needsSave),
                    ["forms"] = new Dictionary<string, object?>
                    {
                        ["name"] = args.FormName,
                        ["basis"] = "runtime",
                        ["initializer"] = string.IsNullOrWhiteSpace(args.Initializer) ? null : args.Initializer,
                    },
                    ["output"] = output,
                    ["warnings"] = warnings,
                },
            };
        }
        catch (InvalidOperationException ex)
        {
            return Failure(request, ClassifyErrorCode(ex.Message), ex.Message, "xlflow-excel-bridge", phase, args, sessionAttached, sessionMode, dirty, needsSave);
        }
        catch (Exception ex)
        {
            return Failure(request, "form_export_image_failed", ExcelBridgeSupport.FormatExceptionDetail(ex), "xlflow-excel-bridge", phase, args, sessionAttached, sessionMode, dirty, needsSave);
        }
        finally
        {
            CleanupCapture(runtimeExcel, runtimeWorkbook, helperModuleName);
            RemoveTemporaryComponent(runtimeVbProject, helperComponent);
            CloseWorkbook(runtimeExcel, runtimeWorkbook, false);
            if (!string.IsNullOrWhiteSpace(runtimeWorkbookPath) && File.Exists(runtimeWorkbookPath))
            {
                try
                {
                    File.Delete(runtimeWorkbookPath);
                }
                catch
                {
                }
            }
            if (!string.IsNullOrWhiteSpace(temporaryOutputPath) && File.Exists(temporaryOutputPath))
            {
                try
                {
                    File.Delete(temporaryOutputPath);
                }
                catch
                {
                }
            }
            ExcelBridgeSupport.ReleaseComObject(runtimeVbProject);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
            CloseWorkbook(excel, workbook, sessionAttached);
        }
    }

    private static string PrepareOutputPath(string outputPath, bool overwrite, out bool createdParentDirs, out string temporaryOutputPath)
    {
        createdParentDirs = false;
        temporaryOutputPath = "";
        var resolved = Path.GetFullPath(outputPath);
        if (Directory.Exists(resolved))
        {
            throw new InvalidOperationException($"form_export_image_args_invalid: Output path '{resolved}' is a directory.");
        }
        var extension = Path.GetExtension(resolved);
        if (string.IsNullOrWhiteSpace(extension) || !string.Equals(extension, ".png", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"unsupported_image_format: Image format '{extension.TrimStart('.')}' is not supported. Supported formats: png.");
        }
        var parent = Path.GetDirectoryName(resolved);
        if (!string.IsNullOrWhiteSpace(parent) && !Directory.Exists(parent))
        {
            Directory.CreateDirectory(parent);
            createdParentDirs = true;
        }
        if (File.Exists(resolved) && !overwrite)
        {
            throw new InvalidOperationException($"output_file_exists: Output file '{resolved}' already exists. Use --overwrite to replace it.");
        }
        if (File.Exists(resolved) && overwrite)
        {
            temporaryOutputPath = Path.Combine(parent ?? Path.GetTempPath(), $"xlflow-form-export-{Guid.NewGuid():N}.png");
            return temporaryOutputPath;
        }
        return resolved;
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbook(string workbookPath, string metadataPath, bool useSession, bool visible)
    {
        if (useSession)
        {
            var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, true);
            return (attached.Excel, attached.Workbook, true, attached.SessionMode);
        }
        if (ExcelBridgeSupport.SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            try
            {
                var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, false);
                return (attached.Excel, attached.Workbook, true, attached.SessionMode);
            }
            catch
            {
            }
        }
        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, direct.SessionMode);
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

    private static (object Excel, object Workbook, string Path) OpenRuntimeWorkbookCopy(object workbook)
    {
        var fullName = ExcelBridgeSupport.TryGetWorkbookFullName(workbook) ?? "";
        var extension = Path.GetExtension(fullName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".xlsm";
        }
        var tempPath = Path.Combine(Path.GetTempPath(), $"xlflow-form-export-image-{Guid.NewGuid():N}{extension}");
        ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", tempPath);
        var direct = ExcelBridgeSupport.OpenWorkbookDirect(tempPath, true, disableAutomationMacros: false);
        return (direct.Excel, direct.Workbook, tempPath);
    }

    private static object InstallHelper(object runtimeVbProject)
    {
        object? components = null;
        object? component = null;
        object? codeModule = null;
        try
        {
            components = ExcelBridgeSupport.Get(runtimeVbProject, "VBComponents")
                ?? throw new InvalidOperationException("vbproject_access_denied: VBComponents is unavailable.");
            component = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
                ?? throw new InvalidOperationException("vba_compile_failed: could not add helper module.");
            ExcelBridgeSupport.Set(component, "Name", "XlflowCap_" + Guid.NewGuid().ToString("N")[..20]);
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule")
                ?? throw new InvalidOperationException("vba_compile_failed: helper CodeModule is unavailable.");
            ExcelBridgeSupport.InvokeMethod(codeModule, "AddFromString", BuildHelperCode());
            return component;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void InvokePrepareCapture(object runtimeExcel, object runtimeWorkbook, string helperModuleName, string formName, string token, string initializer)
    {
        var workbookName = (ExcelBridgeSupport.TryGetWorkbookName(runtimeWorkbook) ?? "").Replace("'", "''", StringComparison.Ordinal);
        var moduleName = string.IsNullOrWhiteSpace(helperModuleName) ? "XlflowCap" : helperModuleName;
        var macroName = $"'{workbookName}'!{moduleName}.XlflowPrepareFormImageCapture";
        ExcelBridgeSupport.RunExcelMacro(runtimeExcel, macroName, formName, token, initializer);
    }

    private static WindowInfo WaitForCaptureWindow(object runtimeExcel, object runtimeWorkbook, string helperModuleName, int processId, string token, CancellationToken cancellationToken)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(7);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var status = ReadCaptureStatus(runtimeExcel, runtimeWorkbook, helperModuleName);
            if (!string.IsNullOrWhiteSpace(status.Source) || !string.IsNullOrWhiteSpace(status.Message))
            {
                var source = string.IsNullOrWhiteSpace(status.Source) ? "XlflowFormImageCapture.capture_prepare" : status.Source;
                throw new InvalidOperationException($"{source}: {status.Message}");
            }

            WindowInfo? window = null;
            if (status.Hwnd != 0)
            {
                window = WindowCapture.WaitForStableWindow(new IntPtr(status.Hwnd), TimeSpan.FromMilliseconds(1200), TimeSpan.FromMilliseconds(100), 2);
                if (WindowCapture.IsLikelyUserFormWindow(window))
                {
                    return window!;
                }
            }
            if (!string.IsNullOrWhiteSpace(status.Caption))
            {
                window = WindowCapture.FindWindowByTitle(processId, status.Caption, exactMatch: true);
                if (WindowCapture.IsLikelyUserFormWindow(window))
                {
                    return WindowCapture.WaitForStableWindow(new IntPtr(window!.Hwnd), TimeSpan.FromMilliseconds(1200), TimeSpan.FromMilliseconds(100), 2) ?? window;
                }
            }
            window = WindowCapture.FindWindowByTitle(processId, token, exactMatch: false);
            if (WindowCapture.IsLikelyUserFormWindow(window))
            {
                return WindowCapture.WaitForStableWindow(new IntPtr(window!.Hwnd), TimeSpan.FromMilliseconds(1200), TimeSpan.FromMilliseconds(100), 2) ?? window;
            }
            Thread.Sleep(100);
        }
        throw new InvalidOperationException($"window_not_found: could not find a visible UserForm window for capture token {token}");
    }

    private static CaptureStatus ReadCaptureStatus(object runtimeExcel, object runtimeWorkbook, string helperModuleName)
    {
        var workbookName = (ExcelBridgeSupport.TryGetWorkbookName(runtimeWorkbook) ?? "").Replace("'", "''", StringComparison.Ordinal);
        var moduleName = string.IsNullOrWhiteSpace(helperModuleName) ? "XlflowCap" : helperModuleName;
        var macroName = $"'{workbookName}'!{moduleName}.XlflowReadFormImageCaptureStatus";
        var value = Convert.ToString(ExcelBridgeSupport.RunExcelMacro(runtimeExcel, macroName), CultureInfo.InvariantCulture) ?? "";
        var parts = value.Split('\t');
        long hwnd = 0;
        _ = parts.Length >= 5 && long.TryParse(parts[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out hwnd);
        return new CaptureStatus(
            parts.Length >= 1 ? parts[0] : "",
            parts.Length >= 2 ? parts[1] : "",
            parts.Length >= 3 && string.Equals(parts[2], "True", StringComparison.OrdinalIgnoreCase),
            parts.Length >= 4 ? parts[3] : "",
            hwnd);
    }

    private static void CleanupCapture(object? runtimeExcel, object? runtimeWorkbook, string helperModuleName)
    {
        if (runtimeExcel is null || runtimeWorkbook is null)
        {
            return;
        }
        try
        {
            var workbookName = (ExcelBridgeSupport.TryGetWorkbookName(runtimeWorkbook) ?? "").Replace("'", "''", StringComparison.Ordinal);
            var moduleName = string.IsNullOrWhiteSpace(helperModuleName) ? "XlflowCap" : helperModuleName;
            var macroName = $"'{workbookName}'!{moduleName}.XlflowCleanupFormImageCapture";
            ExcelBridgeSupport.RunExcelMacro(runtimeExcel, macroName);
        }
        catch
        {
        }
    }

    private static Dictionary<string, object?> SessionPayload(string path, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["active"] = sessionAttached,
            ["workbook_path"] = path,
            ["dirty"] = dirty,
            ["save_required"] = needsSave,
            ["live_newer_than_disk"] = needsSave,
            ["mode"] = sessionMode,
            ["source_of_truth"] = needsSave ? "live_workbook" : "saved_workbook",
        };
    }

    private static Dictionary<string, object?> WorkbookPayload(string path, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["path"] = path,
            ["session"] = sessionAttached,
            ["session_mode"] = sessionMode,
            ["session_requested"] = sessionAttached,
            ["auto_session"] = sessionAttached && string.Equals(sessionMode, "auto", StringComparison.OrdinalIgnoreCase),
            ["saved"] = false,
            ["dirty"] = dirty,
            ["needs_save"] = needsSave,
        };
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message, string source, string phase, FormExportImageCommandArguments args, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        var path = string.IsNullOrWhiteSpace(args.WorkbookPath) ? null : ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
        var extensions = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(path))
        {
            extensions["target"] = new Dictionary<string, object?>
            {
                ["kind"] = sessionAttached ? "live_session" : "file",
                ["path"] = path,
                ["form"] = args.FormName,
                ["capture_state"] = "temporary_copy",
                ["note"] = "Runtime export used a temporary workbook copy.",
            };
            extensions["session"] = SessionPayload(path, sessionAttached, sessionMode, dirty, needsSave);
            extensions["workbook"] = WorkbookPayload(path, sessionAttached, sessionMode, dirty, needsSave);
        }
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(code, message, phase, source),
            Extensions = extensions,
        };
    }

    private static string ClassifyErrorCode(string message)
    {
        if (message.Contains("initializer", StringComparison.OrdinalIgnoreCase))
        {
            return "form_initializer_failed";
        }
        if (message.Contains("runtime_load", StringComparison.OrdinalIgnoreCase))
        {
            return "runtime_form_load_failed";
        }
        if (message.Contains("compile", StringComparison.OrdinalIgnoreCase))
        {
            return "vba_compile_failed";
        }
        if (message.Contains("window_not_found", StringComparison.OrdinalIgnoreCase))
        {
            return "window_not_found";
        }
        if (message.Contains("image_capture_failed", StringComparison.OrdinalIgnoreCase))
        {
            return "image_capture_failed";
        }
        if (message.Contains("output_file_exists", StringComparison.OrdinalIgnoreCase))
        {
            return "output_file_exists";
        }
        if (message.Contains("unsupported_image_format", StringComparison.OrdinalIgnoreCase))
        {
            return "unsupported_image_format";
        }
        if (message.Contains("vbproject_access_denied", StringComparison.OrdinalIgnoreCase))
        {
            return "vbproject_access_denied";
        }
        if (message.Contains("form_not_found", StringComparison.OrdinalIgnoreCase))
        {
            return "form_not_found";
        }
        return "form_export_image_failed";
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
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void CloseWorkbook(object? excel, object? workbook, bool sessionAttached)
    {
        if (sessionAttached)
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
            return;
        }
        if (workbook is not null)
        {
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            }
            catch
            {
            }
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }
        if (excel is not null)
        {
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit");
            }
            catch
            {
            }
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static void TrySetVisible(object excel, bool visible)
    {
        try
        {
            ExcelBridgeSupport.Set(excel, "Visible", visible);
        }
        catch
        {
        }
    }

    private static string BuildHelperCode()
    {
        return """""
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
#Else
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
#End If

Private xlflowCapturedForm As Object
Private xlflowCaptureFormName As String
Private xlflowCaptureToken As String
Private xlflowCaptureInitializer As String
Private xlflowCaptureScheduledAt As Date
Private xlflowLastErrorSource As String
Private xlflowLastErrorMessage As String
Private xlflowCaptureReady As Boolean
Private xlflowCaptureWindowCaption As String
Private xlflowCaptureWindowHandle As String

Private Function XlflowFindFormWindowHandle(ByVal caption As String) As String
#If VBA7 Then
  Dim hwnd As LongPtr
#Else
  Dim hwnd As Long
#End If

  hwnd = 0
  If Len(caption) > 0 Then
    hwnd = FindWindowW(0, StrPtr(caption))
  End If
  XlflowFindFormWindowHandle = CStr(hwnd)
End Function

Public Sub XlflowPrepareFormImageCapture(ByVal formName As String, ByVal token As String, Optional ByVal initializer As String = "")
  xlflowCaptureFormName = formName
  xlflowCaptureToken = token
  xlflowCaptureInitializer = Trim$(initializer)
  xlflowLastErrorSource = ""
  xlflowLastErrorMessage = ""
  xlflowCaptureReady = False
  xlflowCaptureWindowCaption = ""
  xlflowCaptureWindowHandle = "0"
  xlflowCaptureScheduledAt = Now + TimeSerial(0, 0, 1)
  Application.OnTime xlflowCaptureScheduledAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!XlflowExecuteFormImageCapture"
End Sub

Public Sub XlflowExecuteFormImageCapture()
  Dim loaded As Boolean
  Dim initializerRan As Boolean
  Dim caption As String
  On Error GoTo ErrHandler

  Set xlflowCapturedForm = UserForms.Add(xlflowCaptureFormName)
  loaded = True

  If Len(xlflowCaptureInitializer) > 0 Then
    CallByName xlflowCapturedForm, xlflowCaptureInitializer, VbMethod, ThisWorkbook
    initializerRan = True
  End If

  caption = ""
  On Error Resume Next
  caption = CStr(xlflowCapturedForm.Caption)
  On Error GoTo ErrHandler
  PrimeFocusableControl xlflowCapturedForm
  xlflowCapturedForm.Caption = caption & " [xlflow-capture-" & xlflowCaptureToken & "]"
  xlflowCapturedForm.Show vbModeless
  DoEvents
  xlflowCaptureWindowCaption = CStr(xlflowCapturedForm.Caption)
  xlflowCaptureWindowHandle = XlflowFindFormWindowHandle(xlflowCaptureWindowCaption)
  xlflowCaptureReady = True
  Exit Sub

ErrHandler:
  If Not loaded Then
    xlflowLastErrorSource = "XlflowFormImageCapture.runtime_load"
  ElseIf Len(xlflowCaptureInitializer) > 0 And Not initializerRan Then
    xlflowLastErrorSource = "XlflowFormImageCapture.initializer"
  Else
    xlflowLastErrorSource = "XlflowFormImageCapture.capture_prepare"
  End If
  xlflowLastErrorMessage = Err.Description
End Sub

Private Sub PrimeFocusableControl(ByVal formObject As Object)
  On Error Resume Next

  Dim focusTarget As Object
  Set focusTarget = FindFirstFocusableControl(GetObjectControls(formObject))
  If Not focusTarget Is Nothing Then
    CallByName focusTarget, "SetFocus", VbMethod
  End If
End Sub

Private Function FindFirstFocusableControl(ByVal controls As Object) As Object
  On Error GoTo Missing

  Dim control As Object
  Dim childControls As Object
  Dim nested As Object

  If controls Is Nothing Then
    Exit Function
  End If

  For Each control In controls
    If ControlCanReceiveFocus(control) Then
      Set FindFirstFocusableControl = control
      Exit Function
    End If

    Set childControls = GetObjectControls(control)
    If Not childControls Is Nothing Then
      Set nested = FindFirstFocusableControl(childControls)
      If Not nested Is Nothing Then
        Set FindFirstFocusableControl = nested
        Exit Function
      End If
    End If
  Next control

  Exit Function

Missing:
  Set FindFirstFocusableControl = Nothing
End Function

Private Function ControlCanReceiveFocus(ByVal control As Object) As Boolean
  On Error GoTo Missing

  Dim visibleValue As Variant
  Dim enabledValue As Variant
  Dim tabStopValue As Variant

  visibleValue = CallByName(control, "Visible", VbGet)
  enabledValue = CallByName(control, "Enabled", VbGet)
  If Not CBool(visibleValue) Or Not CBool(enabledValue) Then
    GoTo Missing
  End If

  On Error Resume Next
  tabStopValue = CallByName(control, "TabStop", VbGet)
  If Err.Number = 0 Then
    If Not CBool(tabStopValue) Then
      GoTo Missing
    End If
  End If
  Err.Clear
  On Error GoTo Missing

  Select Case LCase$(TypeName(control))
    Case "textbox", "combobox", "listbox", "optionbutton", "togglebutton", "checkbox", "commandbutton", "spinbutton", "scrollbar", "tabstrip", "multipage"
      ControlCanReceiveFocus = True
      Exit Function
  End Select

Missing:
  ControlCanReceiveFocus = False
End Function

Private Function GetObjectControls(ByVal target As Object) As Object
  On Error Resume Next
  Set GetObjectControls = target.Controls
  On Error GoTo 0
End Function

Public Sub XlflowCleanupFormImageCapture()
  On Error Resume Next
  If xlflowCaptureScheduledAt <> 0 Then
    Application.OnTime xlflowCaptureScheduledAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!XlflowExecuteFormImageCapture", , False
  End If
  If Not xlflowCapturedForm Is Nothing Then
    Unload xlflowCapturedForm
  End If
  Set xlflowCapturedForm = Nothing
  xlflowCaptureScheduledAt = 0
  xlflowCaptureReady = False
  xlflowCaptureWindowCaption = ""
  xlflowCaptureWindowHandle = "0"
  xlflowLastErrorSource = ""
  xlflowLastErrorMessage = ""
  On Error GoTo 0
End Sub

Public Function XlflowReadFormImageCaptureStatus() As String
  XlflowReadFormImageCaptureStatus = xlflowLastErrorSource & vbTab & xlflowLastErrorMessage & vbTab & CStr(xlflowCaptureReady) & vbTab & Replace(xlflowCaptureWindowCaption, vbTab, " ") & vbTab & xlflowCaptureWindowHandle
End Function
""""";
    }

    private sealed record CaptureStatus(string Source, string Message, bool Ready, string Caption, long Hwnd);
}
