using System.Diagnostics.CodeAnalysis;
using System.Text;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelRunnerService : IRunnerService
{
    public BridgeResponse Execute(BridgeRequest request, RunnerCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!IsSupportedAction(args.Action))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "runner_args_invalid",
                Message: "-Action must be install, remove, or status.",
                Phase: "runner",
                Source: "xlflow"));
        }

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var attachment = OpenWorkbook(args);
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            sessionAttached = attachment.SessionAttached;
            sessionMode = attachment.SessionMode;
            vbProject = ExcelBridgeSupport.TryGetWorkbookVbProject(workbook)
                ?? throw new InvalidOperationException("VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.");

            Dictionary<string, object?> runner;
            List<string> logs;
            var saved = false;

            switch (args.Action.Trim().ToLowerInvariant())
            {
                case "install":
                    InstallRunnerModule(vbProject);
                    ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                    saved = true;
                    runner = new Dictionary<string, object?> { ["installed"] = true, ["module"] = "XlflowRunner" };
                    logs = ["installed XlflowRunner"];
                    break;
                case "remove":
                    var removed = RemoveRunnerModule(vbProject);
                    if (removed)
                    {
                        ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                        saved = true;
                    }
                    runner = new Dictionary<string, object?> { ["installed"] = false, ["removed"] = removed, ["module"] = "XlflowRunner" };
                    logs = [removed ? "removed XlflowRunner" : "XlflowRunner was not installed"];
                    break;
                default:
                    var installed = TestRunnerInstalled(vbProject);
                    runner = new Dictionary<string, object?> { ["installed"] = installed, ["module"] = "XlflowRunner" };
                    logs = [installed ? "XlflowRunner is installed" : "XlflowRunner is not installed"];
                    break;
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = new Dictionary<string, object?>
                {
                    ["runner"] = runner,
                    ["workbook"] = new Dictionary<string, object?>
                    {
                        ["path"] = ExcelBridgeSupport.NormalizePath(args.WorkbookPath),
                        ["saved"] = saved,
                        ["session"] = sessionAttached,
                        ["session_mode"] = sessionMode,
                    },
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "runner_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "runner",
                Source: ex.Source ?? "xlflow-excel-bridge"));
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

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbook(RunnerCommandArguments args)
    {
        if (ExcelBridgeSupport.SessionMetadataMatchesWorkbook(args.MetadataPath, args.WorkbookPath))
        {
            try
            {
                var attached = ExcelBridgeSupport.AttachToSessionWorkbook(args.WorkbookPath, args.MetadataPath, false);
                return (attached.Excel, attached.Workbook, true, attached.SessionMode);
            }
            catch
            {
                // fall through to direct open
            }
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(args.WorkbookPath, args.Visible);
        return (direct.Excel, direct.Workbook, false, direct.SessionMode);
    }

    private static bool IsSupportedAction(string action)
    {
        return action.Equals("install", StringComparison.OrdinalIgnoreCase)
            || action.Equals("remove", StringComparison.OrdinalIgnoreCase)
            || action.Equals("status", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TestRunnerInstalled(object vbProject)
    {
        try
        {
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            var component = ExcelBridgeSupport.Get(components!, "Item", "XlflowRunner");
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
            return component is not null;
        }
        catch
        {
            return false;
        }
    }

    private static bool RemoveRunnerModule(object vbProject)
    {
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
        object? existing = null;
        try
        {
            existing = ExcelBridgeSupport.Get(components!, "Item", "XlflowRunner");
            if (existing is null)
            {
                return false;
            }
            ExcelBridgeSupport.InvokeMethod(components!, "Remove", existing);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(existing);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void InstallRunnerModule(object vbProject)
    {
        _ = RemoveRunnerModule(vbProject);
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents")
            ?? throw new InvalidOperationException("VBComponents is unavailable.");
        object? component = null;
        object? codeModule = null;
        try
        {
            component = ExcelBridgeSupport.InvokeMethod(components, "Add", 1);
            ExcelBridgeSupport.Set(component!, "Name", "XlflowRunner");
            codeModule = ExcelBridgeSupport.Get(component!, "CodeModule");
            ExcelBridgeSupport.InvokeMethod(codeModule!, "AddFromString", BuildRunnerModuleCode());
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    internal static string BuildRunnerModuleCode()
    {
        var builder = new StringBuilder();
        builder.AppendLine("Option Explicit");
        builder.AppendLine();
        builder.AppendLine("Public Function XlflowRunnerVersion() As String");
        builder.AppendLine("  XlflowRunnerVersion = \"1\"");
        builder.AppendLine("End Function");
        return builder.ToString();
    }
}
