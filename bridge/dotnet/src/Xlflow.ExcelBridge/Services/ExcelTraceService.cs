using System.Diagnostics.CodeAnalysis;
using System.Text;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelTraceService : ITraceService
{
    public BridgeResponse Execute(BridgeRequest request, TraceCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var action = (args.Action ?? "").Trim().ToLowerInvariant();
        if (action == "inject") action = "enable";

        // clean does not need Excel COM
        if (action == "clean")
        {
            return HandleClean(request, args);
        }

        if (action is not ("enable" or "disable" or "status"))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "trace_args_invalid",
                Message: $"unsupported trace action: {args.Action}",
                Phase: "trace",
                Source: "xlflow"));
        }

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForTrace(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            object? vbProject = null;
            try
            {
                vbProject = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"));
            }
            catch (Exception ex)
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "vbide_access_denied",
                    Message: ex.Message,
                    Phase: "trace",
                    Source: "vbide"));
            }

            if (vbProject is null)
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "vbide_access_denied",
                    Message: "VBIDE access is not available.",
                    Phase: "trace",
                    Source: "vbide"));
            }

            try
            {
                return action switch
                {
                    "enable" => HandleEnable(request, args, workbook, vbProject, sessionAttached, sessionMode),
                    "disable" => HandleDisable(request, args, workbook, vbProject, sessionAttached, sessionMode),
                    "status" => HandleStatus(request, args, workbook, vbProject, sessionAttached, sessionMode),
                    _ => BridgeResponse.Failed(request, new BridgeError(
                        Code: "trace_args_invalid",
                        Message: $"unsupported trace action: {args.Action}",
                        Phase: "trace",
                        Source: "xlflow")),
                };
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(vbProject);
            }
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "trace_failed",
                Message: detail,
                Phase: "trace",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            if (sessionAttached)
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
            }
            else
            {
                CloseWorkbook(workbook, excel);
            }
        }
    }

    private static BridgeResponse HandleEnable(BridgeRequest request, TraceCommandArguments args,
        object workbook, object vbProject, bool sessionAttached, string sessionMode)
    {
        TraceHelper.RemoveTraceModule(vbProject);
        TraceHelper.InstallTraceModule(vbProject);
        ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));

        var logs = new List<string>();
        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null) logs.Add(sessionLog);
        logs.Add($"enabled XlflowTrace in {args.WorkbookPath}");

        var extensions = new Dictionary<string, object?>();

        if (!string.IsNullOrWhiteSpace(args.ModulesDir))
        {
            var sourcePath = TraceHelper.WriteTraceModuleSource(args.ModulesDir);
            if (sourcePath is not null)
            {
                extensions["source"] = new { path = sourcePath, updated = true };
                logs.Add($"wrote {sourcePath}");
            }
        }

        extensions["workbook"] = new
        {
            path = args.WorkbookPath,
            session = sessionAttached,
            session_mode = sessionMode,
            session_requested = args.UseSession,
            auto_session = sessionAttached && !args.UseSession,
            saved = true,
            dirty = false,
            needs_save = false,
        };
        extensions["trace"] = new
        {
            lifecycle = "enabled",
            workbook_injected = true,
            log_dir = args.TraceDir,
        };

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = extensions,
        };
    }

    private static BridgeResponse HandleDisable(BridgeRequest request, TraceCommandArguments args,
        object workbook, object vbProject, bool sessionAttached, string sessionMode)
    {
        var sourceRemoved = false;

        if (!string.IsNullOrWhiteSpace(args.ModulesDir))
        {
            var sourcePath = Path.Combine(args.ModulesDir, "XlflowTrace.bas");
            if (File.Exists(sourcePath) && !CanRemoveTraceSource(args.ModulesDir, args.Force))
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "trace_source_modified",
                    Message: "XlflowTrace.bas differs from the bundled helper. Use --force to remove it.",
                    Phase: "trace",
                    Source: "xlflow"));
            }
        }

        var removedWorkbook = TraceHelper.RemoveTraceModule(vbProject);

        if (!string.IsNullOrWhiteSpace(args.ModulesDir))
        {
            var sourcePath = Path.Combine(args.ModulesDir, "XlflowTrace.bas");
            if (File.Exists(sourcePath))
            {
                File.Delete(sourcePath);
                sourceRemoved = true;
            }
        }

        if (removedWorkbook)
        {
            ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));
        }

        var logs = new List<string>();
        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null) logs.Add(sessionLog);
        logs.Add($"disabled XlflowTrace in {args.WorkbookPath}");

        var extensions = new Dictionary<string, object?>
        {
            ["workbook"] = new
            {
                path = args.WorkbookPath,
                session = sessionAttached,
                session_mode = sessionMode,
                session_requested = args.UseSession,
                auto_session = sessionAttached && !args.UseSession,
                saved = removedWorkbook,
                dirty = false,
                needs_save = false,
            },
            ["trace"] = new
            {
                lifecycle = "disabled",
                workbook_removed = removedWorkbook,
                source_removed = sourceRemoved,
                log_dir = args.TraceDir,
            },
        };

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = extensions,
        };
    }

    private static BridgeResponse HandleStatus(BridgeRequest request, TraceCommandArguments args,
        object workbook, object vbProject, bool sessionAttached, string sessionMode)
    {
        string? sourcePath = null;
        var sourceExists = false;
        var sourceMatches = false;
        if (!string.IsNullOrWhiteSpace(args.ModulesDir))
        {
            sourcePath = Path.Combine(args.ModulesDir, "XlflowTrace.bas");
            sourceExists = File.Exists(sourcePath);
            sourceMatches = TraceHelper.TraceModuleSourceMatches(args.ModulesDir);
        }

        var workbookInjected = TraceHelper.HasTraceModule(vbProject);

        var logs = new List<string>();
        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null) logs.Add(sessionLog);
        logs.Add($"reported XlflowTrace status for {args.WorkbookPath}");

        var extensions = new Dictionary<string, object?>
        {
            ["workbook"] = new
            {
                path = args.WorkbookPath,
                session = sessionAttached,
                session_mode = sessionMode,
                session_requested = args.UseSession,
                auto_session = sessionAttached && !args.UseSession,
                saved = false,
                dirty = false,
                needs_save = false,
            },
            ["source"] = new
            {
                path = sourcePath,
                exists = sourceExists,
                matches_bundled = sourceMatches,
            },
            ["trace"] = new
            {
                status = "ok",
                workbook_injected = workbookInjected,
                source_exists = sourceExists,
                source_matches_bundled = sourceMatches,
                log_dir = args.TraceDir,
            },
        };

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = extensions,
        };
    }

    internal static bool CanRemoveTraceSource(string modulesDir, bool force)
    {
        if (force)
        {
            return true;
        }
        return TraceHelper.TraceModuleSourceMatches(modulesDir);
    }

    private static BridgeResponse HandleClean(BridgeRequest request, TraceCommandArguments args)
    {
        var traceDir = args.TraceDir;
        if (string.IsNullOrWhiteSpace(traceDir))
        {
            traceDir = Path.Combine(Directory.GetCurrentDirectory(), ".xlflow", "traces");
        }

        var removed = 0;
        if (Directory.Exists(traceDir))
        {
            var files = Directory.GetFiles(traceDir);
            removed = files.Length;
            Directory.Delete(traceDir, true);
        }

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = [$"cleaned trace logs from {traceDir}"],
            Extensions = new Dictionary<string, object?>
            {
                ["trace"] = new
                {
                    cleaned = true,
                    path = traceDir,
                    files_removed = removed,
                },
            },
        };
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForTrace(
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
        // Trace direct-open disables automation macros for safety
        var direct = OpenWorkbookDirectWithAutomationDisabled(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, "none");
    }

    private static (object Excel, object Workbook) OpenWorkbookDirectWithAutomationDisabled(string workbookPath, bool visible)
    {
        workbookPath = ExcelBridgeSupport.NormalizePath(workbookPath);
        if (!ExcelBridgeSupport.IsExcelFile(workbookPath))
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
            try { app.AutomationSecurity = 3; } catch { /* best-effort */ }
            object workbook = app.Workbooks.Open(workbookPath);
            return (excel, workbook);
        }
        catch (InvalidOperationException)
        {
            ExcelBridgeSupport.ReleaseComObject(excel);
            throw;
        }
        catch (Exception ex)
        {
            ExcelBridgeSupport.ReleaseComObject(excel);
            throw new InvalidOperationException($"bridge_file_not_openable: {ExcelBridgeSupport.FormatExceptionDetail(ex)}", ex);
        }
    }

    private static void CloseWorkbook(object? workbook, object? excel)
    {
        if (workbook is not null)
        {
            try { ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false); }
            catch { /* best-effort */ }
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }
        if (excel is not null)
        {
            try { dynamic app = excel; app.Quit(); }
            catch { /* best-effort */ }
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static string? GetSessionUsageLog(string sessionMode)
    {
        return sessionMode switch
        {
            "explicit" => "using xlflow session workbook (--session)",
            "auto" => "auto-reused matching xlflow session workbook",
            _ => null,
        };
    }
}
