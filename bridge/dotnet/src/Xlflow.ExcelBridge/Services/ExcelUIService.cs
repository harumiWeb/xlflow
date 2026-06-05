using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelUIService : IUIService
{
    public BridgeResponse Execute(BridgeRequest request, UICommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";
        bool? dirty = false;
        var needsSave = false;

        try
        {
            if (string.IsNullOrWhiteSpace(args.Action))
            {
                return ValidationFailure(request, "ui_button_args_invalid", "-Action is required.");
            }

            var action = args.Action.Trim().ToLowerInvariant();
            if (action is not ("add" or "list" or "remove"))
            {
                return ValidationFailure(request, "ui_button_args_invalid", "Unsupported action: " + args.Action);
            }

            (excel, workbook, sessionAttached, sessionMode) = OpenWorkbook(args);
            UpdateSaveState(workbook!, sessionAttached, ref dirty, ref needsSave);

            return action switch
            {
                "add" => AddButton(request, args, excel!, workbook!, sessionAttached, sessionMode, dirty, needsSave),
                "list" => ListButtons(request, args, workbook!, sessionAttached, sessionMode, dirty, needsSave),
                _ => RemoveButton(request, args, workbook!, sessionAttached, sessionMode, dirty, needsSave),
            };
        }
        catch (Exception ex)
        {
            var response = BridgeResponse.Failed(request, new BridgeError(
                Code: "ui_button_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "ui",
                Source: ex.Source ?? "xlflow-excel-bridge",
                Number: ex.HResult));
            AttachWorkbookState(response, args.WorkbookPath, sessionAttached, sessionMode, args.UseSession, dirty, needsSave, sessionAttached ? "live_session" : "file");
            return response;
        }
        finally
        {
            CloseWorkbook(excel, workbook, sessionAttached);
        }
    }

    private static BridgeResponse AddButton(BridgeRequest request, UICommandArguments args, object excel, object workbook, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave)
    {
        if (string.IsNullOrWhiteSpace(args.Sheet))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Sheet is required.");
        }

        if (string.IsNullOrWhiteSpace(args.Cell))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Cell is required.");
        }

        if (string.IsNullOrWhiteSpace(args.Text))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Text is required.");
        }

        if (string.IsNullOrWhiteSpace(args.Macro))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Macro is required.");
        }

        if (args.Width <= 0)
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Width must be greater than 0.");
        }

        if (args.Height <= 0)
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Height must be greater than 0.");
        }

        var buttonId = NormalizeButtonId(args.Id);
        if (string.IsNullOrWhiteSpace(buttonId))
        {
            buttonId = NormalizeButtonId(args.Macro);
        }
        if (string.IsNullOrWhiteSpace(buttonId))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Id could not be derived from -Macro.");
        }

        object? worksheet = null;
        object? range = null;
        object? button = null;
        try
        {
            worksheet = GetWorksheet(workbook, args.Sheet);
            if (worksheet is null)
            {
                if (args.CreateSheet)
                {
                    dynamic wb = workbook;
                    worksheet = (object?)wb.Worksheets.Add();
                    SetMember(worksheet!, "Name", args.Sheet);
                }
                else
                {
                    return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "sheet_not_found", "Worksheet not found: " + args.Sheet, "Excel");
                }
            }

            if (args.VerifyMacro && !MacroExists(workbook, args.Macro))
            {
                return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "macro_not_found", "Macro not found: " + args.Macro, "Excel", "verify_macro");
            }

            try
            {
                dynamic ws = worksheet!;
                range = (object?)ws.Range(args.Cell);
            }
            catch (Exception)
            {
                return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "ui_button_args_invalid", "Invalid cell address: " + args.Cell, "Excel");
            }

            var buttonName = ButtonName(buttonId);
            button = GetShape(worksheet!, buttonName);
            var updated = button is not null;
            if (button is null)
            {
                dynamic ws = worksheet!;
                dynamic buttons = ws.Buttons();
                button = (object?)buttons.Add(
                    GetDouble(range!, "Left"),
                    GetDouble(range!, "Top"),
                    args.Width,
                    args.Height);
                ExcelBridgeSupport.ReleaseComObject(buttons);
                SetMember(button!, "Name", buttonName);
            }

            SetMember(button!, "Caption", args.Text);
            SetMember(button!, "OnAction", args.Macro);
            SetMember(button!, "Left", GetDouble(range!, "Left"));
            SetMember(button!, "Top", GetDouble(range!, "Top"));
            SetMember(button!, "Width", args.Width);
            SetMember(button!, "Height", args.Height);

            return SaveAndReturnUiResponse(
                request,
                args,
                workbook,
                sessionAttached,
                sessionMode,
                dirty,
                needsSave,
                new Dictionary<string, object?>
                {
                    ["ui"] = new Dictionary<string, object?>
                    {
                        ["button"] = ButtonInfo(button!, args.Sheet, buttonId, updated),
                    },
                },
                [updated ? "updated workbook button " + buttonName : "added workbook button " + buttonName]);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(button);
            ExcelBridgeSupport.ReleaseComObject(range);
            ExcelBridgeSupport.ReleaseComObject(worksheet);
        }
    }

    private static BridgeResponse ListButtons(BridgeRequest request, UICommandArguments args, object workbook, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave)
    {
        var buttons = new List<Dictionary<string, object?>>();
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
        var worksheets = GetWorksheets(workbook);
        var foundSheet = string.IsNullOrWhiteSpace(args.Sheet);
        try
        {
            var count = GetCollectionCount(worksheets!);
            for (var i = 1; i <= count; i++)
            {
                object? worksheet = null;
                object? buttonCollection = null;
                try
                {
                    worksheet = GetCollectionItem(worksheets!, i);
                    var sheetName = GetString(worksheet!, "Name") ?? "";
                    if (!string.IsNullOrWhiteSpace(args.Sheet) && !string.Equals(sheetName, args.Sheet, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    foundSheet = true;
                    buttonCollection = GetShapes(worksheet!);
                    var buttonCount = GetCollectionCount(buttonCollection!);
                    for (var j = 1; j <= buttonCount; j++)
                    {
                        object? button = null;
                        try
                        {
                            button = GetCollectionItem(buttonCollection!, j);
                            var buttonName = GetString(button!, "Name") ?? "";
                            if (!buttonName.StartsWith("xlflow.button.", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }
                            var buttonId = buttonName["xlflow.button.".Length..];
                            buttons.Add(ButtonInfo(button!, sheetName, buttonId, null));
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(button);
                        }
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(buttonCollection);
                    ExcelBridgeSupport.ReleaseComObject(worksheet);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(worksheets);
        }

        if (!foundSheet)
        {
            return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "sheet_not_found", "Worksheet not found: " + args.Sheet, "Excel");
        }

        var response = new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = [$"found {buttons.Count} xlflow-managed button(s)"],
            Extensions = new Dictionary<string, object?>
            {
                ["ui"] = new Dictionary<string, object?> { ["buttons"] = buttons },
            },
        };
        AttachWorkbookState(response, workbookPath, sessionAttached, sessionMode, args.UseSession, dirty, needsSave, sessionAttached ? "live_session" : "file");
        if (needsSave)
        {
            response.Extensions["warnings"] = SaveRequiredWarnings();
        }
        return response;
    }

    private static BridgeResponse RemoveButton(BridgeRequest request, UICommandArguments args, object workbook, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave)
    {
        var buttonId = NormalizeButtonId(args.Id);
        if (string.IsNullOrWhiteSpace(buttonId))
        {
            return ValidationFailure(request, "ui_button_args_invalid", "-Id is required.");
        }

        var buttonName = ButtonName(buttonId);
        var worksheets = GetWorksheets(workbook);
        var foundSheet = string.IsNullOrWhiteSpace(args.Sheet);
        Dictionary<string, object?>? removed = null;
        try
        {
            var count = GetCollectionCount(worksheets!);
            for (var i = 1; i <= count; i++)
            {
                object? worksheet = null;
                object? button = null;
                try
                {
                    worksheet = GetCollectionItem(worksheets!, i);
                    var sheetName = GetString(worksheet!, "Name") ?? "";
                    if (!string.IsNullOrWhiteSpace(args.Sheet) && !string.Equals(sheetName, args.Sheet, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    foundSheet = true;
                    button = GetShape(worksheet!, buttonName);
                    if (button is null)
                    {
                        continue;
                    }
                    removed = ButtonInfo(button!, sheetName, buttonId, null);
                    ExcelBridgeSupport.InvokeMethod(button, "Delete");
                    break;
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(button);
                    ExcelBridgeSupport.ReleaseComObject(worksheet);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(worksheets);
        }

        if (!foundSheet)
        {
            return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "sheet_not_found", "Worksheet not found: " + args.Sheet, "Excel");
        }
        if (removed is null)
        {
            return ErrorWithWorkbookState(request, args, sessionAttached, sessionMode, dirty, needsSave, "button_not_found", "Button not found: " + buttonName, "Excel");
        }

        return SaveAndReturnUiResponse(
            request,
            args,
            workbook,
            sessionAttached,
            sessionMode,
            dirty,
            needsSave,
            new Dictionary<string, object?>
            {
                ["ui"] = new Dictionary<string, object?> { ["button"] = removed },
            },
            ["removed workbook button " + buttonName]);
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbook(UICommandArguments args)
    {
        if (args.UseSession)
        {
            var attached = ExcelBridgeSupport.AttachToSessionWorkbook(args.WorkbookPath, args.MetadataPath, true);
            return (attached.Excel, attached.Workbook, true, attached.SessionMode);
        }

        if (ExcelBridgeSupport.SessionMetadataMatchesWorkbook(args.MetadataPath, args.WorkbookPath))
        {
            try
            {
                var attached = ExcelBridgeSupport.AttachToSessionWorkbook(args.WorkbookPath, args.MetadataPath, false);
                return (attached.Excel, attached.Workbook, true, attached.SessionMode);
            }
            catch (Exception)
            {
            }
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(args.WorkbookPath, args.Visible);
        return (direct.Excel, direct.Workbook, false, direct.SessionMode);
    }

    private static void UpdateSaveState(object workbook, bool sessionAttached, ref bool? dirty, ref bool needsSave)
    {
        if (sessionAttached)
        {
            if (!ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var workbookDirty))
            {
                dirty = null;
                needsSave = true;
                return;
            }
            dirty = workbookDirty;
            needsSave = workbookDirty;
        }
        else
        {
            dirty = false;
            needsSave = false;
        }
    }

    private static BridgeResponse SaveAndReturnUiResponse(BridgeRequest request, UICommandArguments args, object workbook, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave, Dictionary<string, object?> extensions, IReadOnlyList<string> logs)
    {
        try
        {
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
        }
        catch (Exception ex)
        {
            var failure = BridgeResponse.Failed(request, new BridgeError(
                Code: "save_failed",
                Message: "Failed to save workbook: " + ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "ui",
                Source: "Excel"));
            foreach (var pair in extensions)
            {
                failure.Extensions[pair.Key] = pair.Value;
            }
            AttachWorkbookState(failure, args.WorkbookPath, sessionAttached, sessionMode, args.UseSession, true, true, sessionAttached ? "live_session" : "file");
            failure.Extensions["warnings"] = SaveRequiredWarnings();
            return failure;
        }

        UpdateSaveState(workbook, sessionAttached, ref dirty, ref needsSave);
        var response = new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = extensions,
        };
        AttachWorkbookState(response, args.WorkbookPath, sessionAttached, sessionMode, args.UseSession, dirty, needsSave, sessionAttached ? "live_session" : "file");
        if (needsSave)
        {
            response.Extensions["warnings"] = SaveRequiredWarnings();
        }
        return response;
    }

    private static BridgeResponse ValidationFailure(BridgeRequest request, string code, string message)
    {
        return BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: "ui", Source: "xlflow"));
    }

    private static BridgeResponse ErrorWithWorkbookState(BridgeRequest request, UICommandArguments args, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave, string code, string message, string source, string? phase = null)
    {
        var response = BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: phase ?? "ui", Source: source));
        AttachWorkbookState(response, args.WorkbookPath, sessionAttached, sessionMode, args.UseSession, dirty, needsSave, sessionAttached ? "live_session" : "file");
        if (needsSave)
        {
            response.Extensions["warnings"] = SaveRequiredWarnings();
        }
        return response;
    }

    private static void AttachWorkbookState(BridgeResponse response, string workbookPath, bool sessionAttached, string sessionMode, bool sessionRequested, bool? dirty, bool needsSave, string targetKind)
    {
        response.Extensions["target"] = ExcelBridgeSupport.BuildTargetPayload(targetKind, workbookPath);
        response.Extensions["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, sessionAttached, sessionMode, dirty, needsSave);
        response.Extensions["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, sessionAttached, sessionMode, sessionRequested, false, dirty, needsSave);
    }

    private static object? GetWorksheet(object workbook, string sheet)
    {
        if (string.IsNullOrWhiteSpace(sheet))
        {
            return null;
        }

        try
        {
            var worksheets = GetWorksheets(workbook);
            try
            {
                var count = GetCollectionCount(worksheets!);
                for (var i = 1; i <= count; i++)
                {
                    var worksheet = GetCollectionItem(worksheets!, i);
                    try
                    {
                        if (string.Equals(GetString(worksheet!, "Name"), sheet, StringComparison.OrdinalIgnoreCase))
                        {
                            return worksheet;
                        }
                    }
                    catch
                    {
                    }
                    ExcelBridgeSupport.ReleaseComObject(worksheet);
                }
                return null;
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(worksheets);
            }
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static bool MacroExists(object workbook, string macroName)
    {
        object? project = null;
        object? components = null;
        try
        {
            project = ExcelBridgeSupport.Get(workbook, "VBProject");
            components = ExcelBridgeSupport.Get(project!, "VBComponents");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components!, "Count"));
            var normalized = macroName.Trim();
            for (var i = 1; i <= count; i++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(components!, "Item", i);
                    var moduleName = ExcelBridgeSupport.GetString(component!, "Name") ?? "";
                    var codeModule = ExcelBridgeSupport.Get(component!, "CodeModule");
                    try
                    {
                        var code = VbaSourceHelper.GetCodeModuleText(codeModule!);
                        var entries = DiscoverProcedures(moduleName, code);
                        if (entries.Any(entry => string.Equals(entry, normalized, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(entry, normalized[(normalized.Contains('.') ? normalized.IndexOf('.') + 1 : 0)..], StringComparison.OrdinalIgnoreCase)))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ExcelBridgeSupport.ReleaseComObject(codeModule);
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
            return false;
        }
        catch (Exception)
        {
            throw;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
            ExcelBridgeSupport.ReleaseComObject(project);
        }
    }

    private static IEnumerable<string> DiscoverProcedures(string moduleName, string code)
    {
        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);
        foreach (var raw in lines)
        {
            var line = raw.Trim();
            var match = System.Text.RegularExpressions.Regex.Match(
                line,
                @"^(?:(Public)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(([^)]*)\))?",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            yield return moduleName + "." + match.Groups[3].Value;
            yield return match.Groups[3].Value;
        }
    }

    private static object? GetShape(object worksheet, string buttonName)
    {
        object? shapes = null;
        try
        {
            shapes = GetShapes(worksheet);
            return GetCollectionItem(shapes!, buttonName);
        }
        catch (Exception)
        {
            return null;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(shapes);
        }
    }

    private static Dictionary<string, object?> ButtonInfo(object button, string sheetName, string buttonId, bool? updated)
    {
        object? topLeftCell = null;
        try
        {
            dynamic dynButton = button;
            topLeftCell = (object?)dynButton.TopLeftCell;
            var result = new Dictionary<string, object?>
            {
                ["id"] = buttonId,
                ["name"] = GetString(button, "Name") ?? ButtonName(buttonId),
                ["sheet"] = sheetName,
                ["text"] = GetButtonText(button),
                ["macro"] = GetString(button, "OnAction") ?? "",
                ["cell"] = topLeftCell is null ? "" : GetAddress(topLeftCell),
                ["left"] = GetDouble(button, "Left"),
                ["top"] = GetDouble(button, "Top"),
                ["width"] = GetDouble(button, "Width"),
                ["height"] = GetDouble(button, "Height"),
            };
            if (updated.HasValue)
            {
                result["updated"] = updated.Value;
            }
            return result;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(topLeftCell);
        }
    }

    private static string NormalizeButtonId(string value)
    {
        var trimmed = (value ?? "").Trim();
        if (trimmed.Length == 0)
        {
            return "";
        }

        var normalized = new string(trimmed
            .Select(ch => char.IsLetterOrDigit(ch) || ch == '.' || ch == '_' || ch == '-' ? char.ToLowerInvariant(ch) : '-')
            .ToArray());
        while (normalized.Contains("--", StringComparison.Ordinal))
        {
            normalized = normalized.Replace("--", "-", StringComparison.Ordinal);
        }
        return normalized.Trim('-');
    }

    private static string ButtonName(string buttonId) => "xlflow.button." + NormalizeButtonId(buttonId);

    private static string? GetString(object comObject, string memberName)
    {
        try
        {
            dynamic dyn = comObject;
            return memberName switch
            {
                "Name" => Convert.ToString(dyn.Name, CultureInfo.InvariantCulture),
                "Caption" => Convert.ToString(dyn.Caption, CultureInfo.InvariantCulture),
                "OnAction" => Convert.ToString(dyn.OnAction, CultureInfo.InvariantCulture),
                _ => Convert.ToString(ExcelBridgeSupport.Get(comObject, memberName), CultureInfo.InvariantCulture),
            };
        }
        catch
        {
            return null;
        }
    }

    private static double GetDouble(object comObject, string memberName)
    {
        dynamic dyn = comObject;
        return memberName switch
        {
            "Left" => Convert.ToDouble(dyn.Left, CultureInfo.InvariantCulture),
            "Top" => Convert.ToDouble(dyn.Top, CultureInfo.InvariantCulture),
            "Width" => Convert.ToDouble(dyn.Width, CultureInfo.InvariantCulture),
            "Height" => Convert.ToDouble(dyn.Height, CultureInfo.InvariantCulture),
            _ => Convert.ToDouble(ExcelBridgeSupport.Get(comObject, memberName), CultureInfo.InvariantCulture),
        };
    }

    private static string GetAddress(object cell)
    {
        dynamic dyn = cell;
        return Convert.ToString(dyn.Address(false, false), CultureInfo.InvariantCulture) ?? "";
    }

    private static object GetShapes(object worksheet)
    {
        dynamic ws = worksheet;
        return ws.Shapes;
    }

    private static object GetWorksheets(object workbook)
    {
        dynamic wb = workbook;
        return wb.Worksheets;
    }

    private static int GetCollectionCount(object collection)
    {
        dynamic dyn = collection;
        return Convert.ToInt32(dyn.Count, CultureInfo.InvariantCulture);
    }

    private static object? GetCollectionItem(object collection, object index)
    {
        dynamic dyn = collection;
        return (object?)dyn.Item(index);
    }


    private static string GetButtonText(object button)
    {
        var caption = GetString(button, "Caption");
        if (!string.IsNullOrWhiteSpace(caption))
        {
            return caption;
        }

        try
        {
            dynamic dyn = button;
            return Convert.ToString(dyn.TextFrame.Characters().Text, CultureInfo.InvariantCulture) ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static void SetMember(object comObject, string memberName, object? value)
    {
        dynamic dyn = comObject;
        switch (memberName)
        {
            case "Name":
                dyn.Name = value;
                return;
            case "Caption":
                dyn.Caption = value;
                return;
            case "OnAction":
                dyn.OnAction = value;
                return;
            case "Left":
                dyn.Left = value;
                return;
            case "Top":
                dyn.Top = value;
                return;
            case "Width":
                dyn.Width = value;
                return;
            case "Height":
                dyn.Height = value;
                return;
            default:
                ExcelBridgeSupport.Set(comObject, memberName, value);
                return;
        }
    }

    private static object[] SaveRequiredWarnings()
    {
        return
        [
            new Dictionary<string, object?>
            {
                ["code"] = "save_required",
                ["message"] = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
            },
        ];
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
            try { ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false); } catch (Exception) { }
        }
        if (excel is not null)
        {
            try { ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit"); } catch (Exception) { }
        }
        ExcelBridgeSupport.ReleaseComObject(workbook);
        ExcelBridgeSupport.ReleaseComObject(excel);
    }
}
