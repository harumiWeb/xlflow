using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Security.Cryptography;
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
        object? runnerComponents = null;
        object? runnerCodeModule = null;
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState = null;
        var sessionAttached = false;
        var sessionMode = "none";
        var skipComCleanup = false;
        var runtimeCleanupSaved = false;
        var excelProcessId = 0;
        var ownedExcelProcess = OwnedExcelProcess.None;

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForRun(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            ownedExcelProcess = ExcelBridgeSupport.CaptureOwnedExcelProcess(excelProcessId);
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;
            ExcelBridgeSupport.StabilizeExcelForMacroRun(excel, args.WorkbookPath, TimeSpan.FromSeconds(3));

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
                    cancellationToken,
                    CreateSelectionLocator(excelProcessId, excelHwnd, args));
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
            WorkbookFileFingerprint? workbookPersistenceBaseline = null;
            if (runtimeState.RestoreRequired || !args.Direct)
            {
                workbookPersistenceBaseline = ExcelBridgeSupport.RunPhase(
                    "capture_workbook_persistence",
                    () => CaptureWorkbookFileFingerprint(args.WorkbookPath));
            }
            ExcelBridgeSupport.SleepAndPump(TimeSpan.FromMilliseconds(150));

            var macroReference = BuildWorkbookQualifiedMacroReference(workbook, macroName);
            if (!args.Direct)
            {
                vbProject = ExcelBridgeSupport.RunPhase("prepare_vbide", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                    ?? throw new InvalidOperationException("prepare_vbide failed: VBProject is unavailable.");
                RuntimeInjectionHelper.EnableUIStreamInjection(workbook, vbProject, runtimeState);
                runnerComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents")
                    ?? throw new InvalidOperationException("prepare_vbide failed: VBComponents is unavailable.");
                runnerComponent = ExcelBridgeSupport.InvokeMethod(runnerComponents, "Add", 1)
                    ?? throw new InvalidOperationException("inject_harness failed: could not add a temporary module.");
                var runnerName = "XlflowRun_" + Guid.NewGuid().ToString("N")[..8];
                SetProperty(runnerComponent, "Name", runnerName);
                RuntimeInjectionHelper.EnsureDefinedNameRestoration(
                    workbook,
                    runtimeState,
                    "__XLFLOW_RUN_HELPER__");
                RuntimeInjectionHelper.SetDefinedName(workbook, "__XLFLOW_RUN_HELPER__", runnerName);
                runnerCodeModule = ExcelBridgeSupport.Get(runnerComponent, "CodeModule")
                    ?? throw new InvalidOperationException("inject_harness failed: CodeModule is unavailable.");
                ExcelBridgeSupport.InvokeMethod(runnerCodeModule, "AddFromString", BuildRunHarnessCode(macroName, macroArgs));
                ExcelBridgeSupport.ReleaseComObject(runnerCodeModule);
                runnerCodeModule = null;
                ExcelBridgeSupport.ReleaseComObject(runnerComponents);
                runnerComponents = null;
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
                MacroInvocationDialogKind(args),
                args.SuppressModalErrors,
                ResolveTimeout(request, args, commandStopwatch.Elapsed),
                cancellationToken,
                CreateSelectionLocator(excelProcessId, excelHwnd, args));
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

            var runtimeStateWasPersistedByMacro = false;
            if (!runTimedOut && runtimeState.RestoreRequired)
            {
                try
                {
                    runtimeStateWasPersistedByMacro = WorkbookFileChanged(
                        args.WorkbookPath,
                        workbookPersistenceBaseline ?? throw new InvalidOperationException(
                            "Workbook persistence baseline is unavailable."));
                }
                catch
                {
                    // If the persisted workbook cannot be inspected after user VBA ran, save the
                    // cleaned live state rather than risk leaving injected runtime artifacts on disk.
                    runtimeStateWasPersistedByMacro = true;
                }
            }

            if (!runTimedOut)
            {
                RemoveTemporaryComponent(vbProject, runnerComponent);
                runnerComponent = null;
                RestoreRuntimeMarkers(workbook, runtimeState);
                runtimeState = null;
                if (runtimeStateWasPersistedByMacro)
                {
                    SaveRuntimeCleanup(workbook);
                    runtimeCleanupSaved = true;
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
                var fatalComFailure = invocation.Result?.Error is not null &&
                    ExcelBridgeSupport.IsFatalComFailure(invocation.Result.Error.Number);
                var fatalHResult = fatalComFailure
                    ? ExcelBridgeSupport.FormatHResult(invocation.Result!.Error!.Number)
                    : null;
                var fatalStage = invocation.Result?.Error?.Stage ?? "invoke_macro";
                var compileFailure = !runTimedOut && IsLikelyVbaCompileFailure(runError, runErrorNumber, invocation.Dialog);
                var errorCode = fatalComFailure ? "excel_com_rpc_failure" : runTimedOut ? "macro_timeout" : compileFailure ? "vba_compile_failed" : ClassifyRunError(runError, runErrorNumber);
                var errorPhase = fatalComFailure ? fatalStage : compileFailure ? "compile_vba" : "invoke_macro";
                logs.Add(compileFailure ? $"VBA compile failed: {runError}" : $"macro execution failed: {runError}");
                if (fatalComFailure && sessionAttached)
                {
                    ExcelBridgeSupport.MarkSessionPoisoned(
                        args.MetadataPath,
                        args.WorkbookPath,
                        runError,
                        fatalHResult ?? "",
                        "run");
                }
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
                    reused = sessionAttached,
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

                if (fatalComFailure)
                {
                    extensions["com"] = BuildComFailureDiagnostics(
                        args,
                        runError,
                        fatalHResult ?? "",
                        fatalStage,
                        excelProcessId,
                        excelHwnd,
                        invocation.WorkerProcessId,
                        sessionAttached,
                        sessionMode);
                }

                if (!string.IsNullOrEmpty(args.RuntimeMode))
                {
                    extensions["runtime"] = new { mode = args.RuntimeMode, source = args.RuntimeSource, injected = runtimeInjected };
                }

                if (invocation.Dialog is not null || args.Diagnostic || runTimedOut || compileFailure)
                {
                    extensions["run_diagnostic"] = BuildRunDiagnostic(
                        runTimedOut ? "timeout" : compileFailure ? "compile" : "runtime",
                        invocation,
                        new
                        {
                            macro = macroName,
                            line = runErrorLine,
                        },
                        runTimedOut);
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
                        Number: runErrorNumber,
                        HResult: fatalHResult,
                        Details: fatalComFailure
                            ? BuildComFailureDetails(args, fatalStage, excelProcessId, excelHwnd, invocation.WorkerProcessId, sessionAttached, sessionMode)
                            : null),
                    Logs = logs,
                    Extensions = extensions,
                };
            }

            var saved = runtimeCleanupSaved;
            var saveAsCopy = false;
            if (!string.IsNullOrWhiteSpace(args.SaveAsPath))
            {
                var saveAsPath = ExcelBridgeSupport.NormalizePath(args.SaveAsPath);
                AssertSaveAsExtension(args.WorkbookPath, saveAsPath);
                ExcelBridgeSupport.RunPhase("save_as", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", saveAsPath));
                saveAsCopy = true;
            }
            else if (args.SaveWorkbook && !saved)
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
        catch (SessionPoisonedException ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_poisoned",
                Message: "The xlflow session was marked poisoned after a fatal Excel COM/RPC failure. Run `xlflow session stop --json` and start a fresh session.",
                Phase: "run",
                Source: "xlflow-excel-bridge",
                HResult: ex.Metadata.HResult,
                Details: new Dictionary<string, object?>
                {
                    ["workbook_path"] = ex.Metadata.WorkbookPath,
                    ["excel_pid"] = ex.Metadata.Pid,
                    ["excel_hwnd"] = ex.Metadata.Hwnd,
                    ["poisoned_at"] = ex.Metadata.PoisonedAt,
                    ["poison_reason"] = ex.Metadata.PoisonReason,
                    ["last_command"] = ex.Metadata.LastCommand,
                }));
        }
        catch (RuntimeCleanupSaveException ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "runtime_cleanup_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "save_result",
                Source: "xlflow-excel-bridge"));
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            var comFailure = ExcelBridgeSupport.ClassifyComFailure(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: comFailure.Fatal ? "excel_com_rpc_failure" : "macro_failed",
                Message: detail,
                Phase: "run",
                Source: "xlflow-excel-bridge",
                Number: comFailure.Fatal ? comFailure.Number : null,
                HResult: comFailure.Fatal ? comFailure.HResult : null));
        }
        finally
        {
            if (!skipComCleanup)
            {
                ExcelBridgeSupport.ReleaseComObject(runnerCodeModule);
                ExcelBridgeSupport.ReleaseComObject(runnerComponents);
                RemoveTemporaryComponent(vbProject, runnerComponent);
                RestoreRuntimeMarkers(workbook, runtimeState);
                ExcelBridgeSupport.ReleaseComObject(vbProject);
                if (sessionAttached)
                {
                    ExcelBridgeSupport.ReleaseComObject(workbook);
                }
                else
                {
                    CloseWorkbook(workbook, excel, ownedExcelProcess);
                }
            }
        }
    }

    internal static WorkbookFileFingerprint CaptureWorkbookFileFingerprint(string workbookPath)
    {
        using var stream = new FileStream(
            workbookPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        return new WorkbookFileFingerprint(stream.Length, hash);
    }

    internal static bool WorkbookFileChanged(
        string workbookPath,
        WorkbookFileFingerprint baseline)
    {
        return !baseline.Equals(CaptureWorkbookFileFingerprint(workbookPath));
    }

    private static void SaveRuntimeCleanup(object workbook)
    {
        try
        {
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
        }
        catch (Exception ex)
        {
            throw new RuntimeCleanupSaveException(ex);
        }
    }

    internal static string BuildRunHarnessCode(string macroName, IReadOnlyList<MacroArg> args)
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
        builder.AppendLine("  DoEvents");
        builder.AppendLine(invocation.ToString());
        builder.AppendLine("  DoEvents");
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
        CancellationToken cancellationToken,
        VbeSelectionLocator? selectionLocator = null)
    {
        return ExcelWorkerInvocation.InvokeWithWorker(
            workerRequest,
            excelHwnd,
            dialogKind,
            suppressModalErrors,
            timeout,
            cancellationToken,
            selectionLocator);
    }

    internal static DialogKind MacroInvocationDialogKind(RunCommandArguments args)
    {
        return string.Equals(args.RuntimeMode, "interactive", StringComparison.OrdinalIgnoreCase)
            ? DialogKind.MacroError
            : DialogKind.Any;
    }

    internal static DialogSnapshot? WaitForPostWorkerDialog(
        DialogWatcher watcher,
        Task<DialogSnapshot?> watcherTask,
        DialogWatchRequest watchRequest,
        string operation,
        MacroRunWorkerResult? result,
        CancellationToken cancellationToken,
        Action<DialogKind>? beforeAction = null,
        Action<DialogKind>? afterAction = null)
    {
        return ExcelWorkerInvocation.WaitForPostWorkerDialog(
            watcher,
            watcherTask,
            watchRequest,
            operation,
            result,
            cancellationToken,
            beforeAction,
            afterAction);
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
        var fatalComFailure = invocation.Result?.Error is not null &&
            ExcelBridgeSupport.IsFatalComFailure(invocation.Result.Error.Number);
        var fatalHResult = fatalComFailure
            ? ExcelBridgeSupport.FormatHResult(invocation.Result!.Error!.Number)
            : null;
        var fatalStage = invocation.Result?.Error?.Stage ?? "compile_vba";
        if (fatalComFailure && sessionAttached)
        {
            ExcelBridgeSupport.MarkSessionPoisoned(
                args.MetadataPath,
                args.WorkbookPath,
                message,
                fatalHResult ?? "",
                "run");
        }
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
            ["run_diagnostic"] = BuildRunDiagnostic("compile", invocation, new { macro = args.MacroName }, invocation.TimedOut),
        };
        if (fatalComFailure)
        {
            extensions["com"] = BuildComFailureDiagnostics(
                args,
                message,
                fatalHResult ?? "",
                fatalStage,
                0,
                0,
                invocation.WorkerProcessId,
                sessionAttached,
                sessionMode);
        }
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(
                Code: fatalComFailure ? "excel_com_rpc_failure" : "vba_compile_failed",
                Message: message,
                Phase: fatalComFailure ? fatalStage : "compile_vba",
                Source: "xlflow-excel-bridge",
                Number: invocation.Result?.Error?.Number,
                HResult: fatalHResult,
                Details: fatalComFailure
                    ? BuildComFailureDetails(args, fatalStage, 0, 0, invocation.WorkerProcessId, sessionAttached, sessionMode)
                    : null),
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

    private static VbeSelectionLocator CreateSelectionLocator(int excelProcessId, long excelHwnd, RunCommandArguments args)
    {
        return new VbeSelectionLocator(excelProcessId, excelHwnd, new VbeSourceMappingOptions(
            args.ModulesDir,
            args.ClassesDir,
            args.FormsDir,
            args.WorkbookDir,
            args.CodeSource,
            args.Folders,
            args.FolderAnnotation,
            args.DefaultComponentFolders));
    }

    internal static object DiagnosticLocation(VbeSelectionCapture capture, object fallback)
    {
        if (capture.Location is null || !capture.HasMeaningfulLocation)
        {
            return fallback;
        }
        return capture.Location;
    }

    internal static object? DiagnosticLocationCapture(VbeSelectionCapture capture)
    {
        if (capture.Attempts.Count == 0)
        {
            return null;
        }
        return new { attempts = capture.Attempts };
    }

    private static Dictionary<string, object?> BuildRunDiagnostic(
        string kind,
        WorkerInvocationResult invocation,
        object fallbackLocation,
        bool timedOut)
    {
        var diagnostic = new Dictionary<string, object?>
        {
            ["kind"] = kind,
            ["location"] = DiagnosticLocation(invocation.LocationCapture, fallbackLocation),
            ["dialog"] = invocation.Dialog,
            ["dialogs"] = invocation.Dialogs,
            ["worker"] = new { pid = invocation.WorkerProcessId, completed = invocation.Result?.Completed ?? false, timed_out = timedOut },
        };
        if (DiagnosticLocationCapture(invocation.LocationCapture) is { } capture)
        {
            diagnostic["location_capture"] = capture;
        }
        return diagnostic;
    }

    internal static Dictionary<string, object?> BuildComFailureDiagnostics(
        RunCommandArguments args,
        string message,
        string hResult,
        string stage,
        int excelProcessId,
        long excelHwnd,
        int workerProcessId,
        bool sessionAttached,
        string sessionMode)
    {
        return new Dictionary<string, object?>
        {
            ["fatal"] = true,
            ["message"] = message,
            ["h_result"] = hResult,
            ["stage"] = stage,
            ["macro"] = args.MacroName,
            ["excel_pid"] = excelProcessId,
            ["excel_hwnd"] = excelHwnd,
            ["worker_pid"] = workerProcessId,
            ["session_id"] = string.IsNullOrWhiteSpace(args.MetadataPath) ? null : args.MetadataPath,
            ["session_mode"] = sessionMode,
            ["visible"] = args.Visible,
            ["headless"] = !args.Visible,
            ["workbook_reused"] = sessionAttached,
            ["workbook_reuse_mode"] = sessionAttached ? sessionMode : "direct",
            ["poisoned"] = sessionAttached,
        };
    }

    internal static IReadOnlyDictionary<string, object?> BuildComFailureDetails(
        RunCommandArguments args,
        string stage,
        int excelProcessId,
        long excelHwnd,
        int workerProcessId,
        bool sessionAttached,
        string sessionMode)
    {
        return new Dictionary<string, object?>
        {
            ["macro"] = args.MacroName,
            ["stage"] = stage,
            ["excel_pid"] = excelProcessId,
            ["excel_hwnd"] = excelHwnd,
            ["worker_pid"] = workerProcessId,
            ["session_id"] = string.IsNullOrWhiteSpace(args.MetadataPath) ? null : args.MetadataPath,
            ["session_mode"] = sessionMode,
            ["visible"] = args.Visible,
            ["headless"] = !args.Visible,
            ["workbook_reused"] = sessionAttached,
            ["workbook_reuse_mode"] = sessionAttached ? sessionMode : "direct",
        };
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
        if (IsDefaultRuntimeWithoutInjectedFeatures(args))
        {
            return new RuntimeInjectionHelper.RuntimeInjectionState
            {
                Mode = args.RuntimeMode,
                Source = args.RuntimeSource,
                Applied = false,
            };
        }

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

    internal static bool IsDefaultRuntimeWithoutInjectedFeatures(RunCommandArguments args)
    {
        return string.Equals(args.RuntimeSource, "default", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(args.MsgBoxResponsesJSON) &&
            string.IsNullOrWhiteSpace(args.InputResponsesJSON) &&
            string.IsNullOrWhiteSpace(args.FileDialogResponsesJSON) &&
            (!args.DebugStreamEnabled || string.IsNullOrWhiteSpace(args.DebugStreamPipeName)) &&
            (!args.UIStreamEnabled || string.IsNullOrWhiteSpace(args.UIStreamPipeName));
    }

    private static void RestoreRuntimeMarkers(object? workbook, RuntimeInjectionHelper.RuntimeInjectionState? state)
    {
        if (workbook is null || state is null)
        {
            return;
        }
        RuntimeInjectionHelper.RestoreRuntimeInjection(workbook, state);
    }

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
            catch (SessionPoisonedException)
            {
                throw;
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

    private static void CloseWorkbook(object? workbook, object? excel, OwnedExcelProcess ownedProcess)
    {
        ExcelBridgeSupport.CloseWorkbookAndQuitApplication(workbook, excel, ownedProcess);
    }
}

internal sealed record WorkbookFileFingerprint(long Length, string Sha256);

internal sealed class RuntimeCleanupSaveException(Exception innerException)
    : Exception(
        $"save_runtime_cleanup failed: {ExcelBridgeSupport.FormatExceptionDetail(innerException)}",
        innerException);
