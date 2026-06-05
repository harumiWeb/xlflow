using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelRunService : IRunService
{
    private static readonly JsonSerializerOptions CachedJsonOptions = new() { PropertyNameCaseInsensitive = true };
    public BridgeResponse Execute(BridgeRequest request, RunCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var commandStopwatch = Stopwatch.StartNew();

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        object? runnerComponent = null;
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState = null;
        var sessionAttached = false;
        var sessionMode = "none";
        var skipComCleanup = false;
        var traceTemporaryInjected = false;

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForRun(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            var dirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
            var dirty = sessionAttached ? dirtyKnown ? dirtyState : true : false;
            var needsSave = sessionAttached ? dirty : false;

            if (!string.IsNullOrWhiteSpace(args.DisplayAlerts.ToString()))
            {
                try
                {
                    dynamic app = excel;
                    app.DisplayAlerts = args.DisplayAlerts;
                }
                catch
                {
                    // best-effort
                }
            }

            var macroArgs = DecodeMacroArgs(args.MacroArgsJSON);
            var macroName = args.MacroName;
            var logs = new List<string>();
            if (sessionAttached)
            {
                logs.Add($"attached to xlflow session ({sessionMode})");
            }

            var excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            var excelHwnd = ExcelBridgeSupport.GetExcelMainHwnd(excel);
            if (excelProcessId <= 0)
            {
                throw new InvalidOperationException("invoke_macro failed: could not resolve the Excel process id.");
            }

            if (args.Diagnostic)
            {
                var compileInvocation = InvokeWithWorker(
                    new MacroRunWorkerRequest(
                        excelProcessId,
                        excelHwnd,
                        "",
                        Operation: "compile",
                        WorkbookPath: args.WorkbookPath),
                    excelHwnd,
                    DialogKind.Compile,
                    args.SuppressModalErrors,
                    ResolveTimeout(request, args, commandStopwatch.Elapsed),
                    cancellationToken);
                if (compileInvocation.Dialog is not null ||
                    compileInvocation.TimedOut ||
                    compileInvocation.Result is null ||
                    !compileInvocation.Result.Ok)
                {
                    return BuildCompileFailureResponse(
                        request,
                        args,
                        macroArgs,
                        sessionAttached,
                        sessionMode,
                        dirty,
                        needsSave,
                        logs,
                        compileInvocation);
                }
            }

            runtimeState = ApplyRuntimeMarkers(workbook, args);
            var runtimeInjected = runtimeState.Applied;

            // Inject trace helper temporarily if trace is enabled and module doesn't exist
            var traceLifecycle = "none";
            if (args.TraceEnabled)
            {
                object? vbProjectForTrace = null;
                try
                {
                    vbProjectForTrace = ExcelBridgeSupport.RunPhase("get_vbproject_for_trace", () => ExcelBridgeSupport.Get(workbook, "VBProject"));
                    if (vbProjectForTrace is not null)
                    {
                        if (TraceHelper.HasTraceModule(vbProjectForTrace))
                        {
                            traceLifecycle = "existing";
                        }
                        else
                        {
                            TraceHelper.InstallTraceModule(vbProjectForTrace);
                            traceLifecycle = "temporary";
                            traceTemporaryInjected = true;
                        }
                    }
                }
                catch
                {
                    // best-effort trace injection
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(vbProjectForTrace);
                }
            }

            var macroReference = BuildWorkbookQualifiedMacroReference(workbook, macroName);
            if (!args.Direct)
            {
                vbProject = ExcelBridgeSupport.RunPhase("prepare_vbide", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                    ?? throw new InvalidOperationException("prepare_vbide failed: VBProject is unavailable.");
                RuntimeInjectionHelper.EnableUIStreamInjection(workbook, vbProject, runtimeState);
                var components = ExcelBridgeSupport.Get(vbProject, "VBComponents")
                    ?? throw new InvalidOperationException("prepare_vbide failed: VBComponents is unavailable.");
                runnerComponent = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
                    ?? throw new InvalidOperationException("inject_harness failed: could not add a temporary module.");
                var runnerName = "XlflowRun_" + Guid.NewGuid().ToString("N")[..8];
                SetProperty(runnerComponent, "Name", runnerName);
                var codeModule = ExcelBridgeSupport.Get(runnerComponent, "CodeModule")
                    ?? throw new InvalidOperationException("inject_harness failed: CodeModule is unavailable.");
                ExcelBridgeSupport.InvokeMethod(codeModule, "AddFromString", BuildRunHarnessCode(macroName, macroArgs, args.TraceEnabled, args.TraceFile));
                macroReference = runnerName + ".RunMacro";
            }

            var sw = Stopwatch.StartNew();
            var invocation = InvokeWithWorker(
                new MacroRunWorkerRequest(
                    excelProcessId,
                    excelHwnd,
                    macroReference,
                    (args.Direct ? macroArgs : [])
                        .Select(argument => new MacroRunWorkerArgument(argument.Type, argument.Value))
                        .ToArray(),
                    WorkbookPath: args.WorkbookPath),
                excelHwnd,
                DialogKind.Any,
                args.SuppressModalErrors,
                ResolveTimeout(request, args, commandStopwatch.Elapsed),
                cancellationToken);
            sw.Stop();

            var durationMs = sw.ElapsedMilliseconds;
            string? runError = null;
            int? runErrorNumber = null;
            int? runErrorLine = null;
            var runTimedOut = invocation.TimedOut;
            skipComCleanup = runTimedOut;
            if (invocation.Dialog is not null)
            {
                runError = DialogMessage(invocation.Dialog);
                runErrorNumber = ParseRuntimeErrorNumber(invocation.Dialog);
            }
            else if (invocation.Result is null)
            {
                runError = runTimedOut ? "Macro execution timed out." : "The macro worker did not return a result.";
            }
            else if (!invocation.Result.Ok)
            {
                runError = invocation.Result.Error?.Message ?? "Macro execution failed.";
                runErrorNumber = invocation.Result.Error?.Number;
            }
            else if (!args.Direct)
            {
                var harnessResult = ParseHarnessResult(invocation.Result.Value);
                if (harnessResult is not null)
                {
                    durationMs = harnessResult.DurationMs > 0 ? harnessResult.DurationMs : durationMs;
                    if (!harnessResult.Success)
                    {
                        runError = harnessResult.Description;
                        runErrorNumber = harnessResult.Number;
                        runErrorLine = harnessResult.Line;
                    }
                }
            }

            if (!runTimedOut)
            {
                RemoveTemporaryComponent(vbProject, runnerComponent);
                runnerComponent = null;
                RestoreRuntimeMarkers(workbook, runtimeState);
                runtimeState = null;
            }

            // Read trace events and clean up trace helper
            var traceEvents = new List<object>();
            if (args.TraceEnabled && !string.IsNullOrWhiteSpace(args.TraceFile))
            {
                try
                {
                    var events = TraceHelper.ReadTraceEvents(args.TraceFile);
                    foreach (var evt in events)
                    {
                        traceEvents.Add(new { timestamp = evt.Timestamp, message = evt.Message });
                        logs.Add(evt.Raw);
                    }
                }
                catch
                {
                    // best-effort trace read
                }
            }

            // Revert temporary trace helper
            if (traceTemporaryInjected && !runTimedOut)
            {
                object? vbProjectForTraceCleanup = null;
                try
                {
                    vbProjectForTraceCleanup = ExcelBridgeSupport.Get(workbook, "VBProject");
                    if (vbProjectForTraceCleanup is not null)
                    {
                        TraceHelper.RemoveTraceModule(vbProjectForTraceCleanup);
                    }
                }
                catch
                {
                    // best-effort
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(vbProjectForTraceCleanup);
                }
            }

            var targetKind = sessionAttached ? "live_session" : "file";
            var warnings = new List<object>();
            var suggestions = new List<object>();
            var extensions = new Dictionary<string, object?>
            {
                ["target"] = new { kind = targetKind, path = args.WorkbookPath },
            };

            if (runError is not null)
            {
                var compileFailure = !runTimedOut && IsLikelyVbaCompileFailure(runError, runErrorNumber, invocation.Dialog);
                var errorCode = runTimedOut ? "macro_timeout" : compileFailure ? "vba_compile_failed" : ClassifyRunError(runError, runErrorNumber);
                var errorPhase = compileFailure ? "compile_vba" : "invoke_macro";
                logs.Add(compileFailure ? $"VBA compile failed: {runError}" : $"macro execution failed: {runError}");
                if (sessionAttached)
                {
                    dirty = true;
                    needsSave = true;
                    warnings.Add(new
                    {
                        code = "save_required",
                        message = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                    });
                }
                extensions["session"] = new { active = sessionAttached, workbook_path = args.WorkbookPath, dirty, save_required = needsSave, live_newer_than_disk = needsSave, mode = sessionMode, source_of_truth = needsSave ? "live_workbook" : "saved_workbook" };
                extensions["workbook"] = new
                {
                    path = args.WorkbookPath,
                    session = sessionAttached,
                    session_mode = sessionMode,
                    session_requested = args.UseSession,
                    auto_session = sessionAttached && !args.UseSession,
                    saved = false,
                    dirty,
                    needs_save = needsSave,
                };

                extensions["macro"] = new
                {
                    name = macroName,
                    duration_ms = durationMs,
                    arguments = macroArgs.Select(a => new { type = a.Type, value = a.Value }).ToArray(),
                    error = new
                    {
                        message = runError,
                        number = runErrorNumber,
                        line = runErrorLine,
                    },
                };

                if (!string.IsNullOrEmpty(args.RuntimeMode))
                {
                    extensions["runtime"] = new { mode = args.RuntimeMode, source = args.RuntimeSource, injected = runtimeInjected };
                }

                if (args.TraceEnabled)
                {
                    extensions["trace"] = new
                    {
                        lifecycle = traceLifecycle,
                        temporary_injected = traceTemporaryInjected,
                        temporary_reverted = false,
                        events = traceEvents.ToArray(),
                        hint = (string?)null,
                        read_error = (string?)null,
                    };
                }

                if (invocation.Dialog is not null || args.Diagnostic || runTimedOut || compileFailure)
                {
                    extensions["run_diagnostic"] = new
                    {
                        kind = runTimedOut ? "timeout" : compileFailure ? "compile" : "runtime",
                        location = new
                        {
                            macro = macroName,
                            line = runErrorLine,
                        },
                        dialog = invocation.Dialog,
                        dialogs = invocation.Dialogs,
                        worker = new { pid = invocation.WorkerProcessId, completed = invocation.Result?.Completed ?? false, timed_out = runTimedOut },
                    };
                }

                if (runTimedOut)
                {
                    suggestions.Add(new { code = "check_dialog", message = "Inspect dialog diagnostics for an unresolved Excel or VBE modal window." });
                    suggestions.Add(new { code = "use_interactive", message = "Use xlflow run --interactive when a human must complete workbook UI." });
                }
                if (suggestions.Count > 0)
                {
                    extensions["suggestions"] = suggestions;
                }

                if (warnings.Count > 0)
                {
                    extensions["warnings"] = warnings;
                }

                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Status = BridgeStatus.Failed,
                    Error = new BridgeError(
                        Code: errorCode,
                        Message: runError,
                        Phase: errorPhase,
                        Source: "xlflow-excel-bridge",
                        Number: runErrorNumber),
                    Logs = logs,
                    Extensions = extensions,
                };
            }

            var saved = false;
            var saveAsCopy = false;
            if (!string.IsNullOrWhiteSpace(args.SaveAsPath))
            {
                var saveAsPath = ExcelBridgeSupport.NormalizePath(args.SaveAsPath);
                AssertSaveAsExtension(args.WorkbookPath, saveAsPath);
                ExcelBridgeSupport.RunPhase("save_as", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", saveAsPath));
                saveAsCopy = true;
            }
            else if (args.SaveWorkbook)
            {
                ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));
                saved = true;
            }

            (dirty, needsSave) = ComputePostRunSaveState(sessionAttached, saved);

            if (needsSave)
            {
                warnings.Add(new
                {
                    code = "save_required",
                    message = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                });
                suggestions.Add(new { code = "save_session", message = "Run xlflow save --session before session stop." });
            }

            extensions["session"] = new { active = sessionAttached, workbook_path = args.WorkbookPath, dirty, save_required = needsSave, live_newer_than_disk = needsSave, mode = sessionMode, source_of_truth = needsSave ? "live_workbook" : "saved_workbook" };
            var workbookResult = new Dictionary<string, object?>
            {
                ["path"] = args.WorkbookPath,
                ["session"] = sessionAttached,
                ["session_mode"] = sessionMode,
                ["session_requested"] = args.UseSession,
                ["auto_session"] = sessionAttached && !args.UseSession,
                ["saved"] = saved,
                ["dirty"] = dirty,
                ["needs_save"] = needsSave,
            };
            if (saveAsCopy)
            {
                workbookResult["save_as"] = ExcelBridgeSupport.NormalizePath(args.SaveAsPath);
            }
            extensions["workbook"] = workbookResult;

            logs.Add($"ran {macroName} in {durationMs}ms");

            if (saveAsCopy)
            {
                logs.Add($"wrote workbook copy to {ExcelBridgeSupport.NormalizePath(args.SaveAsPath)}");
            }

            if (sessionAttached && !saved)
            {
                logs.Add("SAVE REQUIRED: live workbook is newer than disk; run xlflow save before session stop");
            }

            extensions["macro"] = new
            {
                name = macroName,
                duration_ms = durationMs,
                arguments = macroArgs.Select(a => new { type = a.Type, value = a.Value }).ToArray(),
            };

            if (!string.IsNullOrEmpty(args.RuntimeMode))
            {
                extensions["runtime"] = new { mode = args.RuntimeMode, source = args.RuntimeSource, injected = runtimeInjected };
            }

            if (args.TraceEnabled)
            {
                extensions["trace"] = new
                {
                    lifecycle = traceLifecycle,
                    temporary_injected = traceTemporaryInjected,
                    temporary_reverted = traceTemporaryInjected && !runTimedOut,
                    events = traceEvents.ToArray(),
                    hint = traceEvents.Count == 0 && string.IsNullOrWhiteSpace(runError) ? "trace file produced no events; check that the macro calls XlflowTrace.XlflowLog" : null,
                    read_error = (string?)null,
                };
            }

            if (args.Diagnostic)
            {
                extensions["run_diagnostic"] = new
                {
                    kind = "success",
                    location = new
                    {
                        macro = macroName,
                    },
                };
            }

            if (suggestions.Count > 0)
            {
                extensions["suggestions"] = suggestions;
            }

            if (warnings.Count > 0)
            {
                extensions["warnings"] = warnings;
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = extensions,
            };
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("bridge_file_not_openable", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "bridge_file_not_openable",
                Message: ex.Message.Replace("bridge_file_not_openable: ", "", StringComparison.OrdinalIgnoreCase),
                Phase: "run",
                Source: "xlflow-excel-bridge"));
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "macro_failed",
                Message: detail,
                Phase: "run",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            if (!skipComCleanup)
            {
                RemoveTemporaryComponent(vbProject, runnerComponent);
                RestoreRuntimeMarkers(workbook, runtimeState);
                if (traceTemporaryInjected)
                {
                    try
                    {
                        object? vbProjectForTraceCleanup = null;
                        try
                        {
                            if (workbook is not null)
                            {
                                vbProjectForTraceCleanup = ExcelBridgeSupport.Get(workbook, "VBProject");
                            }
                            if (vbProjectForTraceCleanup is not null)
                            {
                                TraceHelper.RemoveTraceModule(vbProjectForTraceCleanup);
                            }
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(vbProjectForTraceCleanup);
                        }
                    }
                    catch
                    {
                        // best-effort
                    }
                }
                ExcelBridgeSupport.ReleaseComObject(vbProject);
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
    }

    internal static string BuildRunHarnessCode(string macroName, IReadOnlyList<MacroArg> args, bool traceEnabled, string traceFile)
    {
        var moduleName = MacroModuleName(macroName);
        var invocation = new StringBuilder("  Application.Run targetMacro");
        foreach (var argument in args)
        {
            invocation.Append(", ");
            invocation.Append(ToVbaLiteral(argument));
        }

        var builder = new StringBuilder();
        builder.AppendLine("Option Explicit");
        builder.AppendLine();
        builder.AppendLine("Public Function RunMacro() As Variant");
        builder.AppendLine("  Dim startedAt As Double");
        builder.AppendLine("  Dim targetMacro As String");
        builder.AppendLine("  startedAt = Timer");
        builder.AppendLine("  targetMacro = \"'\" & ThisWorkbook.Name & \"'!\" & " + ToVbaString(macroName));
        builder.AppendLine("  On Error GoTo Handler");
        if (traceEnabled)
        {
            builder.AppendLine("  XlflowTrace.XlflowSetTraceFile " + ToVbaString(traceFile));
        }
        builder.AppendLine(invocation.ToString());
        builder.AppendLine("  RunMacro = Array(True, " + ToVbaString(moduleName) + ", CLng(0), \"\", CLng(0), CLng((Timer - startedAt) * 1000))");
        builder.AppendLine("  Exit Function");
        builder.AppendLine("Handler:");
        builder.AppendLine("  RunMacro = Array(False, " + ToVbaString(moduleName) + ", CLng(Err.Number), CStr(Err.Description), CLng(Erl), CLng((Timer - startedAt) * 1000))");
        builder.AppendLine("  Err.Clear");
        builder.AppendLine("End Function");
        return builder.ToString();
    }

    private static string ToVbaLiteral(MacroArg argument)
    {
        return argument.Type.ToLowerInvariant() switch
        {
            "int" => $"CLng({argument.Value})",
            "double" => $"CDbl({argument.Value})",
            "bool" => string.Equals(argument.Value, "true", StringComparison.OrdinalIgnoreCase) ? "CBool(True)" : "CBool(False)",
            _ => ToVbaString(argument.Value),
        };
    }

    private static string ToVbaString(string value)
    {
        return "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }

    private static string MacroModuleName(string macroName)
    {
        var index = macroName.LastIndexOf('.');
        return index > 0 ? macroName[..index] : macroName;
    }

    private static string BuildWorkbookQualifiedMacroReference(object workbook, string macroName)
    {
        var workbookName = Convert.ToString(ExcelBridgeSupport.Get(workbook, "Name"), CultureInfo.InvariantCulture) ?? "";
        return "'" + workbookName.Replace("'", "''", StringComparison.Ordinal) + "'!" + macroName;
    }

    private static WorkerInvocationResult InvokeWithWorker(
        MacroRunWorkerRequest workerRequest,
        long excelHwnd,
        DialogKind dialogKind,
        bool suppressModalErrors,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        using var worker = MacroRunWorkerProcess.Start(workerRequest);
        var watcher = new DialogWatcher();
        var watchRequest = new DialogWatchRequest(
            workerRequest.ExcelProcessId,
            excelHwnd,
            dialogKind,
            suppressModalErrors ? DialogActionPolicy.SuppressVbaError : DialogActionPolicy.ObserveOnly,
            timeout,
            TimeSpan.FromMilliseconds(50));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var watcherTask = Task.Run(() => watcher.WaitForDialog(watchRequest, linked.Token), linked.Token);
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
            {
                worker.Stop();
                linked.Cancel();
                return new WorkerInvocationResult(null, watcherTask.Result, [watcherTask.Result], false, worker.ProcessId);
            }
            if (worker.HasExited)
            {
                var result = worker.WaitForResult(TimeSpan.FromSeconds(1));
                var postDialog = WaitForPostWorkerDialog(
                    watcher,
                    watcherTask,
                    watchRequest,
                    workerRequest.Operation,
                    result,
                    linked.Token);
                linked.Cancel();
                if (postDialog is not null)
                {
                    return new WorkerInvocationResult(result, postDialog, [postDialog], false, worker.ProcessId);
                }
                return new WorkerInvocationResult(result, null, [], false, worker.ProcessId);
            }
            Thread.Sleep(25);
        }

        worker.Stop();
        linked.Cancel();
        var dialogs = watcher.CaptureCurrentDialogs(watchRequest, includeUia: false).ToArray();
        return new WorkerInvocationResult(null, dialogs.FirstOrDefault(), dialogs, true, worker.ProcessId);
    }

    internal static DialogSnapshot? WaitForPostWorkerDialog(
        DialogWatcher watcher,
        Task<DialogSnapshot?> watcherTask,
        DialogWatchRequest watchRequest,
        string operation,
        MacroRunWorkerResult? result,
        CancellationToken cancellationToken)
    {
        var shouldWait =
            string.Equals(operation, "compile", StringComparison.OrdinalIgnoreCase) ||
            result is null ||
            !result.Ok ||
            result.Error is not null;
        if (!shouldWait)
        {
            return null;
        }

        if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
        {
            return watcherTask.Result;
        }

        var deadline = DateTime.UtcNow.AddMilliseconds(900);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
            {
                return watcherTask.Result;
            }
            var dialog = watcher.TryCaptureCurrentDialog(watchRequest, includeUia: true, executeAction: true);
            if (dialog is not null)
            {
                return dialog;
            }
            Thread.Sleep(25);
        }
        return null;
    }

    private static BridgeResponse BuildCompileFailureResponse(
        BridgeRequest request,
        RunCommandArguments args,
        IReadOnlyList<MacroArg> macroArgs,
        bool sessionAttached,
        string sessionMode,
        bool dirty,
        bool needsSave,
        List<string> logs,
        WorkerInvocationResult invocation)
    {
        var message = invocation.Dialog is not null
            ? DialogMessage(invocation.Dialog)
            : invocation.TimedOut
                ? "VBE Compile timed out."
                : invocation.Result?.Error?.Message ?? "VBE Compile failed.";
        logs.Add($"VBE Compile failed: {message}");
        var extensions = new Dictionary<string, object?>
        {
            ["target"] = new { kind = sessionAttached ? "live_session" : "file", path = args.WorkbookPath },
            ["session"] = new { active = sessionAttached, workbook_path = args.WorkbookPath, dirty, save_required = needsSave, live_newer_than_disk = needsSave, mode = sessionMode, source_of_truth = needsSave ? "live_workbook" : "saved_workbook" },
            ["workbook"] = new
            {
                path = args.WorkbookPath,
                session = sessionAttached,
                session_mode = sessionMode,
                session_requested = args.UseSession,
                auto_session = sessionAttached && !args.UseSession,
                saved = false,
                dirty,
                needs_save = needsSave,
            },
            ["macro"] = new
            {
                name = args.MacroName,
                duration_ms = 0,
                arguments = macroArgs.Select(argument => new { type = argument.Type, value = argument.Value }).ToArray(),
            },
            ["run_diagnostic"] = new
            {
                kind = "compile",
                location = new { macro = args.MacroName },
                dialog = invocation.Dialog,
                dialogs = invocation.Dialogs,
                worker = new { pid = invocation.WorkerProcessId, completed = invocation.Result?.Completed ?? false, timed_out = invocation.TimedOut },
            },
        };
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(
                Code: "vba_compile_failed",
                Message: message,
                Phase: "compile_vba",
                Source: "xlflow-excel-bridge",
                Number: invocation.Result?.Error?.Number),
            Logs = logs,
            Extensions = extensions,
        };
    }

    private static TimeSpan ResolveTimeout(BridgeRequest request, RunCommandArguments args, TimeSpan elapsed)
    {
        var values = new List<int>();
        if (args.TimeoutSeconds > 0)
        {
            values.Add(Math.Max(1, checked(args.TimeoutSeconds * 1000) - 1000));
        }
        if (request.TimeoutMs is > 0)
        {
            values.Add(Math.Max(1, request.TimeoutMs.Value - (int)elapsed.TotalMilliseconds - 1000));
        }
        return values.Count == 0 ? TimeSpan.FromMinutes(5) : TimeSpan.FromMilliseconds(values.Min());
    }

    private static HarnessResult? ParseHarnessResult(object? value)
    {
        if (value is not JsonElement element || element.ValueKind != JsonValueKind.Array || element.GetArrayLength() < 6)
        {
            return null;
        }
        return new HarnessResult(
            element[0].GetBoolean(),
            element[1].GetString() ?? "",
            element[2].GetInt32(),
            element[3].GetString() ?? "",
            element[4].GetInt32(),
            element[5].GetInt64());
    }

    private static string DialogMessage(DialogSnapshot dialog)
    {
        var lines = dialog.Text.Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
        return lines.Length > 0 ? string.Join(Environment.NewLine, lines) : dialog.Title;
    }

    private static int? ParseRuntimeErrorNumber(DialogSnapshot dialog)
    {
        var message = DialogMessage(dialog);
        var match = Regex.Match(message, "(?i)(?:run-?time error|runtime error|実行時エラー)\\s*'?(?<number>-?\\d+)'?");
        return match.Success && int.TryParse(match.Groups["number"].Value, CultureInfo.InvariantCulture, out var number) ? number : null;
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
        if (vbProject is null || component is null)
        {
            return;
        }
        try
        {
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is not null)
            {
                ExcelBridgeSupport.InvokeMethod(components, "Remove", component);
            }
        }
        catch
        {
            // best-effort temporary component cleanup
        }
        ExcelBridgeSupport.ReleaseComObject(component);
    }

    private static RuntimeInjectionHelper.RuntimeInjectionState ApplyRuntimeMarkers(object workbook, RunCommandArguments args)
    {
        return RuntimeInjectionHelper.ApplyRuntimeInjection(
            workbook,
            args.RuntimeMode,
            args.RuntimeSource,
            args.MsgBoxResponsesJSON,
            args.InputResponsesJSON,
            args.FileDialogResponsesJSON,
            args.DebugStreamEnabled,
            args.DebugStreamPipeName,
            args.UIStreamEnabled,
            args.UIStreamPipeName,
            args.UIStreamRedactInput);
    }

    private static void RestoreRuntimeMarkers(object? workbook, RuntimeInjectionHelper.RuntimeInjectionState? state)
    {
        if (workbook is null || state is null)
        {
            return;
        }
        RuntimeInjectionHelper.RestoreRuntimeInjection(workbook, state);
    }

    private sealed record WorkerInvocationResult(
        MacroRunWorkerResult? Result,
        DialogSnapshot? Dialog,
        DialogSnapshot[] Dialogs,
        bool TimedOut,
        int WorkerProcessId);

    private sealed record HarnessResult(bool Success, string Source, int Number, string Description, int Line, long DurationMs);

    internal sealed class MacroArg
    {
        public string Type { get; set; } = "string";
        public string Value { get; set; } = "";
    }

    internal static (bool Dirty, bool NeedsSave) ComputePostRunSaveState(bool sessionAttached, bool saved)
    {
        if (saved)
        {
            return (false, false);
        }

        if (sessionAttached)
        {
            return (true, true);
        }

        return (false, false);
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForRun(
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

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible, disableAutomationMacros: false);
        return (direct.Excel, direct.Workbook, false, "none");
    }

    private static List<MacroArg> DecodeMacroArgs(string encoded)
    {
        if (string.IsNullOrWhiteSpace(encoded))
        {
            return [];
        }

        try
        {
            var json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            return System.Text.Json.JsonSerializer.Deserialize<List<MacroArg>>(json, CachedJsonOptions) ?? [];
        }
        catch
        {
            return [];
        }
    }

    internal static string ClassifyRunError(string message, int? number)
    {
        if (IsMacroNotFoundError(message, number))
        {
            return "macro_not_found";
        }

        if (IsMacroDisabledError(message, number))
        {
            return "macro_disabled";
        }

        return "macro_failed";
    }

    internal static bool IsLikelyVbaCompileFailure(string message, int? number, DialogSnapshot? dialog = null)
    {
        if (dialog is not null && string.Equals(dialog.Kind, "compile", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        const int vbeCompileDialogHResult = unchecked((int)0x800A9C68);
        if (number == vbeCompileDialogHResult)
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("0x800A9C68", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMacroNotFoundError(string message, int? number)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        var upper = message.ToUpperInvariant();
        if (upper.Contains("CANNOT RUN THE MACRO") ||
            upper.Contains("SUB OR FUNCTION NOT DEFINED") ||
            upper.Contains("MACRO MAY NOT BE AVAILABLE") ||
            upper.Contains("UNABLE TO RUN"))
        {
            return true;
        }
        if (number == 1004 && upper.Contains("MACRO"))
        {
            return true;
        }
        return false;
    }

    private static bool IsMacroDisabledError(string message, int? number)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        var upper = message.ToUpperInvariant();
        if (upper.Contains("SECURITY SETTINGS") && upper.Contains("MACRO"))
        {
            return true;
        }
        if (upper.Contains("MACRO") && upper.Contains("DISABLED"))
        {
            return true;
        }
        if (number == 1004 && upper.Contains("SECURITY"))
        {
            return true;
        }
        return false;
    }

    internal static void AssertSaveAsExtension(string workbookPath, string saveAsPath)
    {
        var workbookExt = Path.GetExtension(workbookPath);
        var saveAsExt = Path.GetExtension(saveAsPath);
        if (!string.Equals(workbookExt, saveAsExt, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"save-as extension {saveAsExt} does not match workbook extension {workbookExt}");
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
}
