using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelListService : IListService
{
    public BridgeResponse Execute(BridgeRequest request, ListCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!string.Equals(args.Action, "forms", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "list_args_invalid",
                Message: "-Action must be forms.",
                Phase: "list",
                Source: "xlflow"));
        }

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        var sessionAttached = false;
        var sessionMode = "none";
        bool? dirty = false;
        var needsSave = false;

        try
        {
            (excel, workbook, sessionAttached, sessionMode) = OpenWorkbook(args);
            var workbookDirty = false;
            if (sessionAttached && !ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook!, out workbookDirty))
            {
                dirty = null;
                needsSave = true;
            }
            else
            {
                dirty = sessionAttached ? workbookDirty : false;
                needsSave = sessionAttached && workbookDirty;
            }

            vbProject = ExcelBridgeSupport.TryGetWorkbookVbProject(workbook);
            if (vbProject is null)
            {
                return Failure(request, args, sessionAttached, sessionMode, dirty, needsSave,
                    "vbproject_access_denied",
                    "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.",
                    "Excel");
            }

            var forms = GetForms(vbProject, args);
            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var logs = new List<string>();
            var sessionLog = ExcelBridgeSupport.GetSessionUsageLog(sessionMode);
            if (!string.IsNullOrWhiteSpace(sessionLog))
            {
                logs.Add(sessionLog);
            }
            logs.Add($"listed {forms.Count} UserForm component(s)");

            var extensions = new Dictionary<string, object?>
            {
                ["target"] = ExcelBridgeSupport.BuildTargetPayload(sessionAttached ? "live_session" : "file", workbookPath),
                ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, sessionAttached, sessionMode, dirty, needsSave),
                ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, sessionAttached, sessionMode, args.UseSession, false, dirty, needsSave),
                ["forms"] = forms,
            };
            if (needsSave)
            {
                extensions["warnings"] = new[]
                {
                    new Dictionary<string, object?>
                    {
                        ["code"] = "save_required",
                        ["message"] = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                    },
                };
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = extensions,
            };
        }
        catch (Exception ex)
        {
            return Failure(request, args, sessionAttached, sessionMode, dirty, needsSave,
                "form_list_failed",
                ExcelBridgeSupport.FormatExceptionDetail(ex),
                ex.Source ?? "xlflow-excel-bridge");
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbProject);
            if (sessionAttached)
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }
            else
            {
                if (workbook is not null)
                {
                    try { ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false); } catch { }
                }
                if (excel is not null)
                {
                    try { ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit"); } catch { }
                }
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }
        }
    }

    private static BridgeResponse Failure(BridgeRequest request, ListCommandArguments args, bool sessionAttached, string sessionMode, bool? dirty, bool needsSave, string code, string message, string source)
    {
        var response = BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: "list", Source: source));
        response.Extensions["target"] = ExcelBridgeSupport.BuildTargetPayload(sessionAttached ? "live_session" : "file", args.WorkbookPath);
        response.Extensions["session"] = ExcelBridgeSupport.BuildSessionPayload(args.WorkbookPath, sessionAttached, sessionMode, dirty, needsSave);
        response.Extensions["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(args.WorkbookPath, sessionAttached, sessionMode, args.UseSession, false, dirty, needsSave);
        return response;
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbook(ListCommandArguments args)
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

    private static List<Dictionary<string, object?>> GetForms(object vbProject, ListCommandArguments args)
    {
        var forms = new List<Dictionary<string, object?>>();
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
        if (components is null)
        {
            return forms;
        }

        try
        {
            var count = Convert.ToInt32(ExcelBridgeSupport.Get(components, "Count"), CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(components, "Item", i);
                    var type = Convert.ToInt32(ExcelBridgeSupport.Get(component!, "Type"), CultureInfo.InvariantCulture);
                    if (type != 3)
                    {
                        continue;
                    }

                    var name = ExcelBridgeSupport.GetString(component!, "Name") ?? $"UserForm{i}";
                    var sourcePath = Path.Combine(args.FormsDir, name + ".frm");
                    var relativeSourcePath = string.IsNullOrWhiteSpace(args.ProjectRoot)
                        ? sourcePath
                        : Path.GetRelativePath(args.ProjectRoot, sourcePath).Replace('\\', '/');
                    var frxPath = Path.ChangeExtension(sourcePath, ".frx")!;
                    var hasFrx = File.Exists(frxPath);
                    var item = new Dictionary<string, object?>
                    {
                        ["name"] = name,
                        ["component_type"] = "MSForm",
                        ["has_frx"] = hasFrx,
                        ["source_path"] = relativeSourcePath,
                    };
                    if (hasFrx)
                    {
                        item["frx_path"] = string.IsNullOrWhiteSpace(args.ProjectRoot)
                            ? frxPath
                            : Path.GetRelativePath(args.ProjectRoot, frxPath).Replace('\\', '/');
                    }
                    forms.Add(item);
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

        return forms.OrderBy(form => Convert.ToString(form["name"], System.Globalization.CultureInfo.InvariantCulture), StringComparer.OrdinalIgnoreCase).ToList();
    }
}
