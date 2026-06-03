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
        RuntimeMarkerState? runtimeState = null;
        var sessionAttached = false;
        var sessionMode = "none";
        var skipComCleanup = false;

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

            var macroReference = BuildWorkbookQualifiedMacroReference(workbook, macroName);
            if (!args.Direct)
            {
                vbProject = ExcelBridgeSupport.RunPhase("prepare_vbide", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                    ?? throw new InvalidOperationException("prepare_vbide failed: VBProject is unavailable.");
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

            var targetKind = sessionAttached ? "live_session" : "file";
            var warnings = new List<object>();
            var suggestions = new List<object>();
            var extensions = new Dictionary<string, object?>
            {
                ["target"] = new { kind = targetKind, path = args.WorkbookPath },
            };

            if (runError is not null)
            {
                var errorCode = runTimedOut ? "macro_timeout" : ClassifyRunError(runError, runErrorNumber);
                logs.Add($"macro execution failed: {runError}");
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

                if (invocation.Dialog is not null || args.Diagnostic || runTimedOut)
                {
                    extensions["run_diagnostic"] = new
                    {
                        kind = runTimedOut ? "timeout" : "runtime",
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
                        Phase: "invoke_macro",
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

            if (saved)
            {
                needsSave = false;
                dirty = false;
            }
            else if (sessionAttached)
            {
                dirty = true;
                needsSave = true;
            }
            else if (ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var postDirty))
            {
                dirty = postDirty;
                needsSave = postDirty;
            }

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
                linked.Cancel();
                var result = worker.WaitForResult(TimeSpan.FromSeconds(1));
                return new WorkerInvocationResult(result, null, [], false, worker.ProcessId);
            }
            Thread.Sleep(25);
        }

        worker.Stop();
        linked.Cancel();
        var dialogs = watcher.CaptureCurrentDialogs(watchRequest, includeUia: false).ToArray();
        return new WorkerInvocationResult(null, dialogs.FirstOrDefault(), dialogs, true, worker.ProcessId);
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

    private static RuntimeMarkerState ApplyRuntimeMarkers(object workbook, RunCommandArguments args)
    {
        var state = new RuntimeMarkerState();
        if (string.IsNullOrWhiteSpace(args.RuntimeMode))
        {
            return state;
        }
        try
        {
            state.Applied = true;
            state.Names.Add(SetDefinedName(workbook, "__XLFLOW_MODE__", args.RuntimeMode));
            state.Names.Add(SetDefinedName(workbook, "__XLFLOW_RUNTIME_VERSION__", "1"));
            if (args.DebugStreamEnabled && !string.IsNullOrWhiteSpace(args.DebugStreamPipeName))
            {
                state.Names.Add(SetDefinedName(workbook, "__XLFLOW_DEBUG_PIPE__", args.DebugStreamPipeName));
            }
        }
        catch
        {
            RestoreRuntimeMarkers(workbook, state);
            state.Applied = false;
        }
        return state;
    }

    private static DefinedNameState SetDefinedName(object workbook, string name, string value)
    {
        var names = ExcelBridgeSupport.Get(workbook, "Names") ?? throw new InvalidOperationException("Workbook names are unavailable.");
        object? existing = null;
        try
        {
            existing = ExcelBridgeSupport.Get(names, "Item", name);
        }
        catch
        {
            // missing name
        }
        var state = new DefinedNameState(name, existing is not null, existing is null ? "" : Convert.ToString(ExcelBridgeSupport.Get(existing, "RefersTo"), CultureInfo.InvariantCulture) ?? "");
        if (existing is not null)
        {
            ExcelBridgeSupport.InvokeMethod(existing, "Delete");
            ExcelBridgeSupport.ReleaseComObject(existing);
        }
        ExcelBridgeSupport.InvokeMethod(names, "Add", name, "=\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"", false);
        ExcelBridgeSupport.ReleaseComObject(names);
        return state;
    }

    private static void RestoreRuntimeMarkers(object? workbook, RuntimeMarkerState? state)
    {
        if (workbook is null || state is null || !state.Applied)
        {
            return;
        }
        foreach (var item in state.Names.AsEnumerable().Reverse())
        {
            try
            {
                var names = ExcelBridgeSupport.Get(workbook, "Names");
                if (names is null)
                {
                    continue;
                }
                object? existing = null;
                try
                {
                    existing = ExcelBridgeSupport.Get(names, "Item", item.Name);
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
                if (item.Existed)
                {
                    ExcelBridgeSupport.InvokeMethod(names, "Add", item.Name, item.RefersTo, false);
                }
                ExcelBridgeSupport.ReleaseComObject(names);
            }
            catch
            {
                // best-effort runtime marker cleanup
            }
        }
        state.Applied = false;
    }

    private sealed record WorkerInvocationResult(
        MacroRunWorkerResult? Result,
        DialogSnapshot? Dialog,
        DialogSnapshot[] Dialogs,
        bool TimedOut,
        int WorkerProcessId);

    private sealed record HarnessResult(bool Success, string Source, int Number, string Description, int Line, long DurationMs);

    private sealed class RuntimeMarkerState
    {
        public bool Applied { get; set; }
        public List<DefinedNameState> Names { get; } = [];
    }

    private sealed record DefinedNameState(string Name, bool Existed, string RefersTo);

    private static object? InvokeMacro(object excel, object workbook, string macroName, List<MacroArg> args)
    {
        var invokeArgs = new List<object?> { macroName };
        foreach (var arg in args)
        {
            invokeArgs.Add(ConvertArg(arg));
        }

        return ExcelBridgeSupport.RunPhase("invoke_macro", () =>
        {
            return ExcelBridgeSupport.InvokeMethod(excel, "Run", invokeArgs.ToArray());
        });
    }

    private static object? ConvertArg(MacroArg arg)
    {
        return arg.Type.ToLowerInvariant() switch
        {
            "int" or "integer" or "long" => int.TryParse(arg.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : (object)arg.Value,
            "double" or "float" or "number" => double.TryParse(arg.Value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) ? d : (object)arg.Value,
            "bool" or "boolean" => bool.TryParse(arg.Value, out var b) ? b : (object)arg.Value,
            _ => arg.Value,
        };
    }

    private static List<MacroArg> DecodeMacroArgs(string encoded)
    {
        if (string.IsNullOrWhiteSpace(encoded))
        {
            return [];
        }

        try
        {
            var json = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            return JsonSerializer.Deserialize<List<MacroArg>>(json, CachedJsonOptions) ?? [];
        }
        catch
        {
            return [];
        }
    }

    private static int? CaptureErrorLine(object? excel)
    {
        if (excel is null)
        {
            return null;
        }

        try
        {
            var err = ExcelBridgeSupport.Get(excel, "Err");
            if (err is null)
            {
                return null;
            }

            dynamic errObj = err;
            var source = errObj.Source as string;
            if (!string.IsNullOrWhiteSpace(source) && int.TryParse(source, out var line))
            {
                return line;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static (string Message, int Number)? CaptureExcelError(object? excel)
    {
        if (excel is null)
        {
            return null;
        }

        try
        {
            var err = ExcelBridgeSupport.Get(excel, "Err");
            if (err is null)
            {
                return null;
            }

            dynamic errObj = err;
            var number = ExcelBridgeSupport.ToInt(errObj.Number);
            if (number == 0)
            {
                return null;
            }

            var description = errObj.Description as string;
            return (description ?? $"Excel error {number}", number);
        }
        catch
        {
            return null;
        }
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

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, "none");
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

    internal sealed class MacroArg
    {
        public string Type { get; set; } = "string";
        public string Value { get; set; } = "";
    }
}
