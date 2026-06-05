using System.Diagnostics.CodeAnalysis;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelSessionService : ISessionService
{
    public BridgeResponse Execute(BridgeRequest request, SessionCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var action = (args.Action ?? "").Trim().ToLowerInvariant();
        return action switch
        {
            "start" => Start(request, args),
            "status" => Status(request, args),
            "save" => Save(request, args),
            "stop" => Stop(request, args),
            _ => BridgeResponse.Failed(request, new BridgeError(
                Code: "session_args_invalid",
                Message: "-Action must be start, status, stop, or save.",
                Phase: "session",
                Source: "xlflow")),
        };
    }

    private static BridgeResponse Start(BridgeRequest request, SessionCommandArguments args)
    {
        object? excel = null;
        object? workbook = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);

        try
        {
            ExcelBridgeSupport.DeleteSessionMetadata(args.MetadataPath);

            var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, true);
            excel = direct.Excel;
            workbook = direct.Workbook;
            ExcelBridgeSupport.WriteSessionMetadata(args.MetadataPath, excel, workbookPath);

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = ["started xlflow Excel session"],
                Extensions = new Dictionary<string, object?>
                {
                    ["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", workbookPath),
                    ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, true, "managed", false, false),
                    ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, true, "managed", false, false, false, false),
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "session",
                Source: ex.Source ?? "xlflow-excel-bridge"));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse Status(BridgeRequest request, SessionCommandArguments args)
    {
        object? excel = null;
        object? workbook = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);

        try
        {
            var metadata = ExcelBridgeSupport.ReadSessionMetadata(args.MetadataPath);
            var running = false;
            var open = false;
            bool? dirty = false;
            var mode = "none";

            if (metadata is not null)
            {
                excel = ExcelBridgeSupport.GetExcelFromSessionMetadata(args.MetadataPath);
                if (excel is not null)
                {
                    running = true;
                    try
                    {
                        workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, workbookPath);
                        open = workbook is not null;
                        if (open && workbook is not null)
                        {
                            mode = "managed";
                            if (!ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var workbookDirty))
                            {
                                dirty = null;
                            }
                            else
                            {
                                dirty = workbookDirty;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        open = false;
                    }
                }
            }

            var needsSave = running && open && (dirty is null || dirty == true);
            var logs = new[] { running && open ? "xlflow session is running" : "xlflow session is not running" };
            var extensions = new Dictionary<string, object?>
            {
                ["target"] = ExcelBridgeSupport.BuildTargetPayload(running && open ? "live_session" : "file", workbookPath),
                ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, running && open, mode, dirty, needsSave),
                ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, running && open, mode, false, false, dirty, needsSave),
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
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "session",
                Source: ex.Source ?? "xlflow-excel-bridge"));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse Save(BridgeRequest request, SessionCommandArguments args)
    {
        object? excel = null;
        object? workbook = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);

        try
        {
            var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, args.MetadataPath, args.UseSession);
            excel = attached.Excel;
            workbook = attached.Workbook;
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
            ExcelBridgeSupport.WriteSessionMetadata(args.MetadataPath, excel, workbookPath);

            var sessionLog = ExcelBridgeSupport.GetSessionUsageLog(attached.SessionMode);
            var logs = new List<string>();
            if (!string.IsNullOrWhiteSpace(sessionLog))
            {
                logs.Add(sessionLog);
            }
            logs.Add("saved xlflow session workbook");

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = new Dictionary<string, object?>
                {
                    ["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", workbookPath),
                    ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, true, attached.SessionMode, false, false),
                    ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, true, attached.SessionMode, args.UseSession, true, false, false),
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "session",
                Source: ex.Source ?? "xlflow-excel-bridge"));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse Stop(BridgeRequest request, SessionCommandArguments args)
    {
        object? excel = null;
        object? workbook = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
        var removedStaleMetadata = false;

        try
        {
            var metadata = ExcelBridgeSupport.ReadSessionMetadata(args.MetadataPath);
            if (metadata is null)
            {
                throw new InvalidOperationException("xlflow session is not running");
            }

            excel = ExcelBridgeSupport.GetSessionExcel(args.MetadataPath);
            if (excel is null)
            {
                ExcelBridgeSupport.DeleteSessionMetadata(args.MetadataPath);
                removedStaleMetadata = true;
                return BuildStopResponse(request, workbookPath, false, removedStaleMetadata);
            }

            try
            {
                workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, workbookPath);
            }
            catch (InvalidOperationException)
            {
                ExcelBridgeSupport.DeleteSessionMetadata(args.MetadataPath);
                removedStaleMetadata = true;
                return BuildStopResponse(request, workbookPath, false, removedStaleMetadata);
            }

            var wasDirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var wasDirtyValue);
            var wasDirty = wasDirtyKnown && wasDirtyValue;
            if (wasDirty)
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
            }

            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit");
            ExcelBridgeSupport.DeleteSessionMetadata(args.MetadataPath);
            return BuildStopResponse(request, workbookPath, wasDirty, false);
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "session",
                Source: ex.Source ?? "xlflow-excel-bridge"));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse BuildStopResponse(BridgeRequest request, string workbookPath, bool autoSavedOnStop, bool removedStaleMetadata)
    {
        var logs = new List<string>();
        if (autoSavedOnStop)
        {
            logs.Add("warning: session workbook had unsaved changes before stop");
            logs.Add("auto-saved workbook while stopping xlflow session; prefer xlflow save before stop");
        }
        if (removedStaleMetadata)
        {
            logs.Add("cleaned stale xlflow session metadata");
        }
        logs.Add("stopped xlflow Excel session");

        var extensions = new Dictionary<string, object?>
        {
            ["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", workbookPath),
            ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, false, "none", false, false),
            ["workbook"] = new Dictionary<string, object?>
            {
                ["path"] = workbookPath,
                ["session"] = false,
                ["session_mode"] = "none",
                ["saved"] = true,
                ["dirty"] = false,
                ["needs_save"] = false,
                ["dirty_before_stop"] = autoSavedOnStop,
                ["auto_saved_on_stop"] = autoSavedOnStop,
            },
        };
        if (autoSavedOnStop)
        {
            extensions["warnings"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live session workbook had unsaved changes and was saved while stopping the session.",
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

}
