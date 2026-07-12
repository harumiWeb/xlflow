using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelTestService : ITestService
{
    private const int VbObjectError = unchecked((int)0x80040000);
    private const int InconclusiveErrorNumber = VbObjectError + 516;
    private const string StopReasonMaximumFailures = "maximum failure count reached";
    private static readonly string[] SupportedTestParameterTypes = ["Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "String", "Variant"];

    public BridgeResponse Execute(BridgeRequest request, TestCommandArguments args, CancellationToken cancellationToken)
    {
        if (args.FailFast && args.MaxFailures > 0 && args.MaxFailures != 1)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_args_invalid",
                Message: "--fail-fast is equivalent to --max-failures 1 and cannot be combined with --max-failures",
                Phase: "test",
                Source: "xlflow"));
        }
        if (args.MaxFailures < 0)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_args_invalid",
                Message: "--max-failures must be greater than zero",
                Phase: "test",
                Source: "xlflow"));
        }
        if (args.RerunFailed < 0)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_args_invalid",
                Message: "--rerun-failed must be zero or greater",
                Phase: "test",
                Source: "xlflow"));
        }
        if (args.FailFast && args.MaxFailures == 0)
        {
            args = args with { MaxFailures = 1 };
        }
        var isolation = NormalizeIsolation(args.Isolation);
        if (isolation is null)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_args_invalid",
                Message: $"unsupported isolation mode \"{args.Isolation}\"; expected none, module, or test",
                Phase: "test",
                Source: "xlflow"));
        }
        if (args.UseSession && isolation != "none")
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "unsupported_test_isolation",
                Message: $"isolation mode \"{isolation}\" is not supported with --session",
                Phase: "test",
                Source: "xlflow"));
        }

        if (args.UseSession)
        {
            var response = ExecuteSingleWorkbook(
                request,
                args,
                args.WorkbookPath,
                DisplayWorkbookPath(args),
                explicitTests: null,
                saveAfterRun: !args.NoSave,
                cancellationToken);
            AttachTestRunMetadata(response, args, isolation, session: true, temporaryWorkbook: false, workbookSaved: WorkbookSaved(response, fallback: false), cleanup: null);
            return response;
        }

        return ExecuteIsolated(request, args, isolation, cancellationToken);
    }

    private static BridgeResponse ExecuteIsolated(BridgeRequest request, TestCommandArguments args, string isolation, CancellationToken cancellationToken)
    {
        var sourceWorkbook = SourceWorkbookPath(args);
        var runDir = Path.Combine(TempRunRoot(args), Guid.NewGuid().ToString("N"));
        var runDirCreated = false;
        CleanupInfo? cleanup = null;

        try
        {
            Directory.CreateDirectory(runDir);
            runDirCreated = true;

            if (isolation == "none")
            {
                if (args.RerunFailed > 0)
                {
                    var retryDiscoveryWorkbook = CopyWorkbookForTest(sourceWorkbook, runDir, "discovery");
                    var retryDiscovery = DiscoverSelectedTests(request, args, retryDiscoveryWorkbook, sourceWorkbook, cancellationToken);
                    if (retryDiscovery.Response is not null)
                    {
                        cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
                        AttachTestRunMetadata(retryDiscovery.Response, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
                        return retryDiscovery.Response;
                    }

                    var selectedWithRetries = retryDiscovery.Tests;
                    var tempWorkbookWithRetries = CopyWorkbookForTest(sourceWorkbook, runDir, "run");
                    var firstAttemptArgs = args with { MaxFailures = 0, FailFast = false };
                    var responseWithRetries = ExecuteSingleWorkbook(
                        request,
                        firstAttemptArgs,
                        tempWorkbookWithRetries,
                        sourceWorkbook,
                        explicitTests: selectedWithRetries,
                        saveAfterRun: false,
                        cancellationToken);
                    if (responseWithRetries.Extensions.TryGetValue("tests", out var retryTestsPayload) && retryTestsPayload is IEnumerable<object> retryTests)
                    {
                        var retryLogs = new List<string>(responseWithRetries.Logs);
                        var mergedResults = RetryFailedResults(request, args, sourceWorkbook, runDir, isolation, selectedWithRetries, retryTests.ToList(), retryLogs, cancellationToken);
                        var finalFailed = CountFailedResultObjects(mergedResults);
                        responseWithRetries.Extensions["tests"] = mergedResults.ToArray();
                        responseWithRetries.Extensions["execution"] = BuildExecutionPayload(args, stoppedEarly: false, stopReason: "");
                        var finalError = finalFailed > 0 ? new BridgeError(
                            Code: "test_failed",
                            Message: $"{finalFailed} of {selectedWithRetries.Count} test(s) failed",
                            Phase: "test",
                            Source: "xlflow") : null;
                        if (finalFailed == 0)
                        {
                            responseWithRetries.Extensions.Remove("error");
                            responseWithRetries.Extensions.Remove("status");
                        }
                        responseWithRetries = responseWithRetries with
                        {
                            Logs = retryLogs,
                            Status = finalFailed > 0 ? BridgeStatus.Failed : BridgeStatus.Ok,
                            Error = finalError,
                        };
                    }
                    cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
                    AttachTestRunMetadata(responseWithRetries, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
                    return responseWithRetries;
                }

                var tempWorkbook = CopyWorkbookForTest(sourceWorkbook, runDir, "run");
                var groupArgs = args.RerunFailed > 0 ? args with { MaxFailures = 0, FailFast = false } : args;
                var response = ExecuteSingleWorkbook(
                    request,
                    groupArgs,
                    tempWorkbook,
                    sourceWorkbook,
                    explicitTests: null,
                    saveAfterRun: false,
                    cancellationToken);
                cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
                AttachTestRunMetadata(response, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
                return response;
            }

            var discoveryWorkbook = CopyWorkbookForTest(sourceWorkbook, runDir, "discovery");
            var discovery = DiscoverSelectedTests(request, args, discoveryWorkbook, sourceWorkbook, cancellationToken);
            if (discovery.Response is not null)
            {
                cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
                AttachTestRunMetadata(discovery.Response, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
                return discovery.Response;
            }

            var selected = discovery.Tests;
            var groups = isolation == "module"
                ? selected.GroupBy(t => t.Module, StringComparer.OrdinalIgnoreCase).Select(g => g.ToList()).ToList()
                : selected.Select(t => new List<TestCase> { t }).ToList();

            var logs = new List<string>();
            var results = new List<object>();
            var failed = 0;
            var stoppedEarly = false;
            var stopReason = "";
            foreach (var group in groups)
            {
                if (group.All(IsNonExecutedTest))
                {
                    foreach (var test in group)
                    {
                        results.Add(BuildNonExecutedTestResult(test));
                        logs.Add(NonExecutedLogLine(test));
                    }

                    continue;
                }
                if (ShouldStopAfterFailures(args, failed))
                {
                    stoppedEarly = true;
                    stopReason = StopReasonMaximumFailures;
                    foreach (var test in group)
                    {
                        results.Add(BuildNotRunTestResult(test, stopReason));
                        logs.Add($"NOT RUN {test.QualifiedName}: {stopReason}");
                    }
                    continue;
                }

                var unitName = isolation == "module" ? group[0].Module : group[0].QualifiedName;
                var tempWorkbook = CopyWorkbookForTest(sourceWorkbook, runDir, SanitizeFileSegment(unitName));
                var groupArgs = args.RerunFailed > 0 ? args with { MaxFailures = 0, FailFast = false } : args;
                var response = ExecuteSingleWorkbook(
                    request,
                    groupArgs,
                    tempWorkbook,
                    sourceWorkbook,
                    explicitTests: group,
                    saveAfterRun: false,
                    cancellationToken);
                logs.AddRange(response.Logs);
                if (response.Extensions.TryGetValue("tests", out var testsPayload) && testsPayload is IEnumerable<object> testItems)
                {
                    var groupResults = testItems.ToList();
                    if (args.RerunFailed > 0)
                    {
                        groupResults = RetryFailedResults(request, args, sourceWorkbook, runDir, isolation, group, groupResults, logs, cancellationToken);
                    }
                    results.AddRange(groupResults);
                }

                if (response.Error is not null && response.Error.Code != "test_failed")
                {
                    cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
                    AttachTestRunMetadata(response, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
                    return response;
                }
                failed = CountFailedResultObjects(results);
            }

            cleanup = CleanupRunDirectory(runDir, args.ProjectRoot);
            var aggregate = new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Status = failed > 0 ? BridgeStatus.Failed : BridgeStatus.Ok,
                Error = failed > 0 ? new BridgeError(
                    Code: "test_failed",
                    Message: $"{failed} of {selected.Count} test(s) failed",
                    Phase: "test",
                    Source: "xlflow") : null,
                Logs = logs,
                Extensions = new Dictionary<string, object?>
                {
                    ["workbook"] = new
                    {
                        path = sourceWorkbook,
                        session = false,
                        session_mode = "none",
                        session_requested = false,
                        auto_session = false,
                        saved = false,
                        dirty = false,
                        needs_save = false,
                    },
                    ["tests"] = results.ToArray(),
                    ["execution"] = BuildExecutionPayload(args, stoppedEarly, stopReason),
                },
            };
            AttachTestRunMetadata(aggregate, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
            return aggregate;
        }
        catch (Exception ex)
        {
            cleanup ??= runDirCreated
                ? CleanupRunDirectory(runDir, args.ProjectRoot)
                : new CleanupInfo("failed", DisplayCleanupPath(runDir, args.ProjectRoot), ex.Message);
            var response = BridgeResponse.Failed(request, new BridgeError(
                Code: "test_environment_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "test",
                Source: "xlflow-excel-bridge"));
            AttachTestRunMetadata(response, args, isolation, session: false, temporaryWorkbook: true, workbookSaved: false, cleanup);
            return response;
        }
    }

    private static (List<TestCase> Tests, BridgeResponse? Response) DiscoverSelectedTests(
        BridgeRequest request,
        TestCommandArguments args,
        string workbookPath,
        string displayWorkbookPath,
        CancellationToken cancellationToken)
    {
        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        var excelProcessId = 0;
        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForTest(workbookPath, args.MetadataPath, useSession: false, args.Visible, disableAutoSession: true));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            vbProject = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                ?? throw new InvalidOperationException("VBIDE access is not available.");

            var discovered = DiscoverTests(vbProject, cancellationToken);
            if (discovered.Count == 0)
            {
                return ([], BuildErrorResponse(request, args, "no_tests_found", "no VBA tests found",
                    sessionAttached: false, sessionMode: "none", [], runtimeState: null, runtimeInjected: false, displayWorkbookPath: displayWorkbookPath));
            }

            var duplicates = DuplicateTestQualifiedNames(discovered);
            if (duplicates.Count > 0)
            {
                var names = string.Join(", ", duplicates);
                return ([], BuildErrorResponse(request, args, "duplicate_test_name",
                    $"duplicate VBA test name(s): {names}",
                    sessionAttached: false, sessionMode: "none", discovered, runtimeState: null, runtimeInjected: false, displayWorkbookPath: displayWorkbookPath));
            }

            var selection = SelectTests(discovered, args.Filter, args.ModuleFilter, args.TagFilter);
            if (selection.Ambiguous)
            {
                var ambiguousName = args.Filter.Trim();
                var ambiguousMatches = selection.Matches.Select(t => t.QualifiedName).ToArray();
                return ([], BuildErrorResponse(request, args, "ambiguous_test_name",
                    $"test name is ambiguous: {ambiguousName}{Environment.NewLine}{Environment.NewLine}Matches:{Environment.NewLine}- {string.Join(Environment.NewLine + "- ", ambiguousMatches)}{Environment.NewLine}{Environment.NewLine}Use a qualified test name.",
                    sessionAttached: false, sessionMode: "none", selection.Matches, runtimeState: null, runtimeInjected: false,
                    errorDetails: new Dictionary<string, object?> { ["matches"] = ambiguousMatches },
                    displayWorkbookPath: displayWorkbookPath));
            }

            if (selection.Tests.Count == 0)
            {
                var filterDesc = FirstNonEmpty(args.Filter, args.ModuleFilter, args.TagFilter, "(no filter)");
                return ([], BuildErrorResponse(request, args, "test_not_found",
                    $"test not found: {filterDesc}",
                    sessionAttached: false, sessionMode: "none", [], runtimeState: null, runtimeInjected: false, displayWorkbookPath: displayWorkbookPath));
            }

            return (selection.Tests, null);
        }
        catch (InvalidTestMetadataException ex)
        {
            return ([], BuildErrorResponse(request, args, "invalid_test_metadata", ex.Message,
                sessionAttached: false, sessionMode: "none", [], runtimeState: null, runtimeInjected: false, displayWorkbookPath: displayWorkbookPath));
        }
        catch (InvalidTestCaseException ex)
        {
            return ([], BuildErrorResponse(request, args, "invalid_test_case", ex.Message,
                sessionAttached: false, sessionMode: "none", [], runtimeState: null, runtimeInjected: false, displayWorkbookPath: displayWorkbookPath));
        }
        catch (Exception ex)
        {
            return ([], BridgeResponse.Failed(request, new BridgeError(
                Code: "test_environment_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "test",
                Source: "xlflow-excel-bridge")));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbProject);
            CloseWorkbook(workbook, excel, excelProcessId);
        }
    }

    private static BridgeResponse ExecuteSingleWorkbook(
        BridgeRequest request,
        TestCommandArguments args,
        string workbookPath,
        string displayWorkbookPath,
        List<TestCase>? explicitTests,
        bool saveAfterRun,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        object? vbProject = null;
        object? runnerComponent = null;
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState = null;
        var sessionAttached = false;
        var sessionMode = "none";
        var excelProcessId = 0;
        var runtimeInjected = false;

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForTest(workbookPath, args.MetadataPath, args.UseSession, args.Visible, args.DisableAutoSession));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;
            excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);

            runtimeState = RuntimeInjectionHelper.ApplyRuntimeInjection(
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

            runtimeInjected = runtimeState.Applied;

            vbProject = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                ?? throw new InvalidOperationException("VBIDE access is not available.");

            RuntimeInjectionHelper.EnableUIStreamInjection(workbook, vbProject, runtimeState);

            var discovered = explicitTests is null ? DiscoverTests(vbProject, cancellationToken) : explicitTests;

            if (discovered.Count == 0)
            {
                return BuildErrorResponse(request, args, "no_tests_found", "no VBA tests found",
                    sessionAttached, sessionMode, [], runtimeState, runtimeInjected, displayWorkbookPath: displayWorkbookPath);
            }

            var selected = discovered;
            if (explicitTests is null)
            {
                var duplicates = DuplicateTestQualifiedNames(discovered);
                if (duplicates.Count > 0)
                {
                    var names = string.Join(", ", duplicates);
                    return BuildErrorResponse(request, args, "duplicate_test_name",
                        $"duplicate VBA test name(s): {names}",
                        sessionAttached, sessionMode, discovered, runtimeState, runtimeInjected, displayWorkbookPath: displayWorkbookPath);
                }

                var selection = SelectTests(discovered, args.Filter, args.ModuleFilter, args.TagFilter);
                if (selection.Ambiguous)
                {
                    var ambiguousName = args.Filter.Trim();
                    var ambiguousMatches = selection.Matches.Select(t => t.QualifiedName).ToArray();
                    return BuildErrorResponse(request, args, "ambiguous_test_name",
                        $"test name is ambiguous: {ambiguousName}{Environment.NewLine}{Environment.NewLine}Matches:{Environment.NewLine}- {string.Join(Environment.NewLine + "- ", ambiguousMatches)}{Environment.NewLine}{Environment.NewLine}Use a qualified test name.",
                        sessionAttached, sessionMode, selection.Matches, runtimeState, runtimeInjected,
                        new Dictionary<string, object?> { ["matches"] = ambiguousMatches },
                        displayWorkbookPath: displayWorkbookPath);
                }

                selected = selection.Tests;
                if (selected.Count == 0)
                {
                    var filterDesc = FirstNonEmpty(args.Filter, args.ModuleFilter, args.TagFilter, "(no filter)");

                    return BuildErrorResponse(request, args, "test_not_found",
                        $"test not found: {filterDesc}",
                        sessionAttached, sessionMode, [], runtimeState, runtimeInjected, displayWorkbookPath: displayWorkbookPath);
                }
            }

            if (selected.All(IsNonExecutedTest))
            {
                var nonExecutedLogs = new List<string>();
                var nonExecutedResults = new List<object>();
                foreach (var test in selected)
                {
                    nonExecutedResults.Add(BuildNonExecutedTestResult(test));
                    nonExecutedLogs.Add(NonExecutedLogLine(test));
                }
                var response = BuildSuccessfulTestResponse(request, args, displayWorkbookPath, sessionAttached, sessionMode, saveAfterRun, workbook, nonExecutedLogs, nonExecutedResults, runtimeState);
                runtimeState = null;
                return response;
            }

            var executableSelected = selected.Where(t => !IsNonExecutedTest(t)).ToList();

            // Build hooks map
            var hooksByModule = new Dictionary<string, ModuleHooks>(StringComparer.OrdinalIgnoreCase);
            foreach (var moduleGroup in executableSelected.GroupBy(t => t.Module, StringComparer.OrdinalIgnoreCase))
            {
                var moduleName = moduleGroup.Key;
                var moduleCode = GetModuleCode(vbProject, moduleName);
                hooksByModule[moduleName] = FindModuleHooks(moduleName, moduleCode);
            }

            // Assign sequential index
            for (var i = 0; i < executableSelected.Count; i++)
            {
                executableSelected[i].Index = i;
            }

            // Generate and inject runner module
            var runnerName = "XlflowTestRunner" + Guid.NewGuid().ToString("N")[..8];
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents")
                ?? throw new InvalidOperationException("VBComponents is unavailable.");
            runnerComponent = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
                ?? throw new InvalidOperationException("Could not add test runner module.");
            SetProperty(runnerComponent, "Name", runnerName);
            var runnerCodeModule = ExcelBridgeSupport.Get(runnerComponent, "CodeModule")
                ?? throw new InvalidOperationException("CodeModule is unavailable.");
            ExcelBridgeSupport.InvokeMethod(runnerCodeModule, "AddFromString",
                BuildTestRunnerCode(executableSelected, hooksByModule));
            ExcelBridgeSupport.ReleaseComObject(runnerCodeModule);
            ExcelBridgeSupport.ReleaseComObject(components);

            // Execute tests
            var results = new List<object>();
            var logs = new List<string>();
            var failed = 0;
            var inconclusiveCount = 0;
            var stoppedEarly = false;
            var stopReason = "";
            var scheduledExecutable = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var moduleGroup in selected.GroupBy(t => t.Module, StringComparer.OrdinalIgnoreCase))
            {
                var moduleName = moduleGroup.Key;
                foreach (var test in moduleGroup.Where(IsNonExecutedTest))
                {
                    results.Add(BuildNonExecutedTestResult(test));
                    logs.Add(NonExecutedLogLine(test));
                }

                var executableModuleGroup = moduleGroup.Where(t => !IsNonExecutedTest(t)).ToList();
                if (executableModuleGroup.Count == 0)
                {
                    continue;
                }
                if (ShouldStopAfterFailures(args, failed))
                {
                    stoppedEarly = true;
                    stopReason = StopReasonMaximumFailures;
                    foreach (var test in executableModuleGroup)
                    {
                        results.Add(BuildNotRunTestResult(test, stopReason));
                        logs.Add($"NOT RUN {test.QualifiedName}: {stopReason}");
                    }
                    continue;
                }

                hooksByModule.TryGetValue(moduleName, out var hooks);

                // BeforeAll
                var beforeAllFailed = false;
                if (hooks?.BeforeAll is not null)
                {
                    var sw = Stopwatch.StartNew();
                    try
                    {
                        var runResult = ExcelBridgeSupport.InvokeMethod(excel, "Run", runnerName + ".RunBeforeAll_" + moduleName);
                        sw.Stop();
                        if (runResult is object[] arr && arr.Length >= 5 && !Convert.ToBoolean(arr[0], CultureInfo.InvariantCulture))
                        {
                            beforeAllFailed = true;
                            var message = Convert.ToString(arr[3], CultureInfo.InvariantCulture) ?? "";
                            foreach (var test in executableModuleGroup)
                            {
                                scheduledExecutable.Add(test.QualifiedName);
                                failed++;
                                results.Add(BuildTestResult(test, "failed", (int)sw.ElapsedMilliseconds,
                                    new { code = "before_all_failed", message, source = Convert.ToString(arr[2], CultureInfo.InvariantCulture) ?? "", number = Convert.ToInt32(arr[1], CultureInfo.InvariantCulture) }));
                                logs.Add($"FAIL {test.QualifiedName}: before_all_failed: {message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sw.Stop();
                        beforeAllFailed = true;
                        foreach (var test in executableModuleGroup)
                        {
                            scheduledExecutable.Add(test.QualifiedName);
                            failed++;
                            results.Add(BuildTestResult(test, "failed", (int)sw.ElapsedMilliseconds,
                                new { code = "before_all_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult }));
                            logs.Add($"FAIL {test.QualifiedName}: before_all_failed: {ex.Message}");
                        }
                    }
                }

                if (!beforeAllFailed)
                {
                    foreach (var test in executableModuleGroup)
                    {
                        if (ShouldStopAfterFailures(args, failed))
                        {
                            stoppedEarly = true;
                            stopReason = StopReasonMaximumFailures;
                            results.Add(BuildNotRunTestResult(test, stopReason));
                            logs.Add($"NOT RUN {test.QualifiedName}: {stopReason}");
                            continue;
                        }
                        scheduledExecutable.Add(test.QualifiedName);
                        var sw = Stopwatch.StartNew();
                        try
                        {
                            var runResult = ExcelBridgeSupport.InvokeMethod(excel, "Run", runnerName + ".RunTest", test.Index);
                            sw.Stop();
                            if (runResult is not object[] arr || arr.Length < 6)
                            {
                                throw new InvalidOperationException("Test runner returned unexpected result shape.");
                            }
                            var success = Convert.ToBoolean(arr[0], CultureInfo.InvariantCulture);
                            var errNumber = Convert.ToInt32(arr[1], CultureInfo.InvariantCulture);
                            var errSource = Convert.ToString(arr[2], CultureInfo.InvariantCulture) ?? "";
                            var errDescription = Convert.ToString(arr[3], CultureInfo.InvariantCulture) ?? "";
                            var statusHint = Convert.ToString(arr[4], CultureInfo.InvariantCulture) ?? "";
                            var phaseHint = Convert.ToString(arr[5], CultureInfo.InvariantCulture) ?? "";

                            var classification = ClassifyTestOutcome(test, success, errNumber, errSource, errDescription, statusHint, phaseHint);
                            results.Add(BuildTestResult(test, classification.Status, (int)sw.ElapsedMilliseconds,
                                new { code = classification.ErrorCode, message = classification.ErrorMessage, source = classification.ErrorSource, number = classification.ErrorNumber },
                                classification.ObservedError));

                            if (classification.Status == "passed")
                            {
                                logs.Add($"PASS {test.QualifiedName}");
                            }
                            else if (classification.Status == "inconclusive")
                            {
                                inconclusiveCount++;
                                logs.Add($"? {test.QualifiedName}: inconclusive");
                            }
                            else
                            {
                                failed++;
                                logs.Add($"FAIL {test.QualifiedName}: {classification.ErrorMessage}");
                            }
                        }
                        catch (Exception ex)
                        {
                            sw.Stop();
                            failed++;
                            results.Add(BuildTestResult(test, "failed", (int)sw.ElapsedMilliseconds,
                                new { code = "test_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult }));
                            logs.Add($"FAIL {test.QualifiedName}: {ex.Message}");
                        }
                    }
                }

                // AfterAll
                if (hooks?.AfterAll is not null && !beforeAllFailed)
                {
                    var sw = Stopwatch.StartNew();
                    try
                    {
                        var runResult = ExcelBridgeSupport.InvokeMethod(excel, "Run", runnerName + ".RunAfterAll_" + moduleName);
                        sw.Stop();
                        if (runResult is object[] arr && arr.Length >= 5 && !Convert.ToBoolean(arr[0], CultureInfo.InvariantCulture))
                        {
                            var message = Convert.ToString(arr[3], CultureInfo.InvariantCulture) ?? "";
                            var moduleTestNames = executableModuleGroup.Where(t => scheduledExecutable.Contains(t.QualifiedName)).Select(t => t.QualifiedName).ToHashSet(StringComparer.OrdinalIgnoreCase);
                            for (var i = 0; i < results.Count; i++)
                            {
                                var resultObj = results[i];
                                var resultName = ResultString(resultObj, "qualified_name");
                                var resultModule = ResultString(resultObj, "module");
                                var resultStatus = ResultString(resultObj, "status");
                                if (resultModule == moduleName && moduleTestNames.Contains(resultName))
                                {
                                    AdjustCountsForAfterAllFailure(resultStatus, ref failed, ref inconclusiveCount);
                                    var resultDuration = ResultInt(resultObj, "duration_ms");
                                    var resultQualifiedName = ResultString(resultObj, "qualified_name");
                                    var test = selected.First(t => string.Equals(t.QualifiedName, resultQualifiedName, StringComparison.OrdinalIgnoreCase));
                                    results[i] = BuildTestResult(test, "failed", resultDuration,
                                        new { code = "after_all_failed", message, source = Convert.ToString(arr[2], CultureInfo.InvariantCulture) ?? "", number = Convert.ToInt32(arr[1], CultureInfo.InvariantCulture) });
                                    logs.Add($"FAIL {resultQualifiedName}: after_all_failed: {message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sw.Stop();
                        var moduleTestNames = executableModuleGroup.Where(t => scheduledExecutable.Contains(t.QualifiedName)).Select(t => t.QualifiedName).ToHashSet(StringComparer.OrdinalIgnoreCase);
                        for (var i = 0; i < results.Count; i++)
                        {
                            var resultObj = results[i];
                            var resultName = ResultString(resultObj, "qualified_name");
                            var resultModule = ResultString(resultObj, "module");
                            var resultStatus = ResultString(resultObj, "status");
                            if (resultModule == moduleName && moduleTestNames.Contains(resultName))
                            {
                                AdjustCountsForAfterAllFailure(resultStatus, ref failed, ref inconclusiveCount);
                                var resultDuration = ResultInt(resultObj, "duration_ms");
                                var resultQualifiedName = ResultString(resultObj, "qualified_name");
                                var test = selected.First(t => string.Equals(t.QualifiedName, resultQualifiedName, StringComparison.OrdinalIgnoreCase));
                                results[i] = BuildTestResult(test, "failed", resultDuration,
                                    new { code = "after_all_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult });
                                logs.Add($"FAIL {resultQualifiedName}: after_all_failed: {ex.Message}");
                            }
                        }
                    }
                }
            }

            // Cleanup runner module
            if (runnerComponent is not null)
            {
                try
                {
                    var vbComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents");
                    if (vbComponents is not null)
                    {
                        ExcelBridgeSupport.InvokeMethod(vbComponents, "Remove", runnerComponent);
                        ExcelBridgeSupport.ReleaseComObject(vbComponents);
                    }
                }
                catch
                {
                    // best-effort
                }
                ExcelBridgeSupport.ReleaseComObject(runnerComponent);
                runnerComponent = null;
            }

            // Restore runtime injection
            if (runtimeState is not null)
            {
                RuntimeInjectionHelper.RestoreRuntimeInjection(workbook, runtimeState);
                runtimeState = null;
            }

            var saved = false;
            if (saveAfterRun)
            {
                ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));
                saved = true;
            }

            var sessionLog = GetSessionUsageLog(sessionMode);
            if (sessionLog is not null)
            {
                logs.Insert(0, sessionLog);
            }

            var extensions = new Dictionary<string, object?>
            {
                ["workbook"] = new
                {
                    path = displayWorkbookPath,
                    session = sessionAttached,
                    session_mode = sessionMode,
                    session_requested = args.UseSession,
                    auto_session = sessionAttached && !args.UseSession,
                    saved,
                    dirty = sessionAttached && !saved,
                    needs_save = sessionAttached && !saved,
                },
                ["tests"] = results.ToArray(),
                ["execution"] = BuildExecutionPayload(args, stoppedEarly, stopReason),
            };

            if (failed > 0)
            {
                extensions["error"] = new
                {
                    code = "test_failed",
                    message = $"{failed} of {selected.Count} test(s) failed",
                };
                extensions["status"] = "failed";
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Status = failed > 0 ? BridgeStatus.Failed : BridgeStatus.Ok,
                Error = failed > 0 ? new BridgeError(
                    Code: "test_failed",
                    Message: $"{failed} of {selected.Count} test(s) failed",
                    Phase: "test",
                    Source: "xlflow") : null,
                Logs = logs,
                Extensions = extensions,
            };
        }
        catch (InvalidTestMetadataException ex)
        {
            return BuildErrorResponse(request, args, "invalid_test_metadata", ex.Message,
                sessionAttached, sessionMode, [], runtimeState, runtimeInjected, displayWorkbookPath: displayWorkbookPath);
        }
        catch (InvalidTestCaseException ex)
        {
            return BuildErrorResponse(request, args, "invalid_test_case", ex.Message,
                sessionAttached, sessionMode, [], runtimeState, runtimeInjected, displayWorkbookPath: displayWorkbookPath);
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "test_environment_failed",
                Message: detail,
                Phase: "test",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            if (runnerComponent is not null)
            {
                try
                {
                    if (vbProject is not null)
                    {
                        var vbComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents");
                        if (vbComponents is not null)
                        {
                            ExcelBridgeSupport.InvokeMethod(vbComponents, "Remove", runnerComponent);
                            ExcelBridgeSupport.ReleaseComObject(vbComponents);
                        }
                    }
                }
                catch
                {
                    // best-effort
                }
                ExcelBridgeSupport.ReleaseComObject(runnerComponent);
            }
            if (runtimeState is not null)
            {
                try
                {
                    if (workbook is not null)
                    {
                        RuntimeInjectionHelper.RestoreRuntimeInjection(workbook, runtimeState);
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
                CloseWorkbook(workbook, excel, excelProcessId);
            }
        }
    }

    private static List<TestCase> DiscoverTests(object vbProject, CancellationToken cancellationToken)
    {
        var tests = new List<TestCase>();
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
        if (components is null)
        {
            return tests;
        }
        var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components, "Count"));
        for (var i = 1; i <= count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            object? component = null;
            try
            {
                component = ExcelBridgeSupport.Get(components, "Item", i);
                if (component is null)
                {
                    continue;
                }

                var name = ExcelBridgeSupport.GetString(component, "Name");
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var code = GetCodeModuleText(component);
                var moduleTests = FindTestProcedures(name, code);
                tests.AddRange(moduleTests);
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(component);
            }
        }
        ExcelBridgeSupport.ReleaseComObject(components);
        return tests;
    }

    private static string GetModuleCode(object vbProject, string moduleName)
    {
        object? components = null;
        object? component = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is null)
            {
                return "";
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components, "Count"));
            for (var i = 1; i <= count; i++)
            {
                component = ExcelBridgeSupport.Get(components, "Item", i);
                if (component is null)
                {
                    continue;
                }

                var name = ExcelBridgeSupport.GetString(component, "Name");
                if (string.Equals(name, moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    return GetCodeModuleText(component);
                }
                ExcelBridgeSupport.ReleaseComObject(component);
                component = null;
            }
        }
        catch
        {
            // best-effort
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
        return "";
    }

    private static string GetCodeModuleText(object component)
    {
        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
            if (codeModule is null)
            {
                return "";
            }

            return VbaSourceHelper.GetCodeModuleText(codeModule);
        }
        catch
        {
            return "";
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    internal static List<TestCase> FindTestProcedures(string moduleName, string code)
    {
        var tests = new List<TestCase>();
        if (string.IsNullOrEmpty(code))
        {
            return tests;
        }

        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            var procedureMatch = Regex.Match(line,
                $@"^(?:(?<visibility>Public|Private)\s+)?(?<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>{VbaIdentifierPattern.Identifier})\b",
                RegexOptions.IgnoreCase);
            if (!procedureMatch.Success)
            {
                continue;
            }

            var name = procedureMatch.Groups["name"].Value;
            var visibility = procedureMatch.Groups["visibility"].Value;
            var kind = procedureMatch.Groups["kind"].Value;
            var header = DeclarationHeader(lines, i);
            var metadata = MetadataAbove(lines, i, moduleName);
            var isTest = kind.Equals("Sub", StringComparison.OrdinalIgnoreCase) &&
                !visibility.Equals("Private", StringComparison.OrdinalIgnoreCase) &&
                (name.StartsWith("Test", StringComparison.OrdinalIgnoreCase) || name.EndsWith("_Test", StringComparison.OrdinalIgnoreCase));
            if (!isTest)
            {
                if (metadata.ExpectedError is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, i + 1, metadata.ExpectedErrorLine,
                        "@ExpectedError annotation is only supported on test procedures");
                }
                if (metadata.TestCases.Count > 0)
                {
                    throw new InvalidTestCaseException(moduleName, i + 1, metadata.TestCases[0].Line,
                        "@TestCase annotation is only supported on test procedures");
                }
                continue;
            }

            var parameters = ParseProcedureParameters(header, moduleName, i + 1);
            tests.AddRange(ExpandTestCases(moduleName, name, i + 1, parameters, metadata));
        }
        return tests;
    }

    private static TestMetadata MetadataAbove(string[] lines, int procedureIndex, string moduleName)
    {
        var metadata = new TestMetadata();
        for (var j = procedureIndex - 1; j >= 0; j--)
        {
            var prev = lines[j].Trim();
            if (string.IsNullOrWhiteSpace(prev))
            {
                continue;
            }

            var tagMatch = Regex.Match(prev, @"^'\s*@Tag\s*\(""([^""]+)""\)", RegexOptions.IgnoreCase);
            if (tagMatch.Success)
            {
                metadata.Tags.Add(tagMatch.Groups[1].Value);
                continue;
            }

            if (Regex.IsMatch(prev, @"^'\s*@TestCase\b", RegexOptions.IgnoreCase))
            {
                var testCase = ParseTestCaseAnnotation(prev, moduleName, j + 1);
                testCase.Line = j + 1;
                metadata.TestCases.Add(testCase);
                continue;
            }

            if (Regex.IsMatch(prev, @"^'\s*@ExpectedError\b", RegexOptions.IgnoreCase))
            {
                if (metadata.ExpectedError is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, procedureIndex + 1, j + 1,
                        "multiple @ExpectedError annotations on one test procedure");
                }
                metadata.ExpectedError = ParseExpectedErrorAnnotation(prev, moduleName, j + 1);
                metadata.ExpectedErrorLine = j + 1;
                continue;
            }

            if (Regex.IsMatch(prev, @"^'\s*@Skip\b", RegexOptions.IgnoreCase))
            {
                if (metadata.Skip is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, procedureIndex + 1, j + 1,
                        "multiple @Skip annotations on one test procedure");
                }
                if (metadata.Todo is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, procedureIndex + 1, j + 1,
                        "test cannot be both skipped and todo");
                }
                metadata.Skip = ParseSkipTodoAnnotation(prev, "Skip", moduleName, j + 1);
                metadata.SkipLine = j + 1;
                continue;
            }

            if (Regex.IsMatch(prev, @"^'\s*@Todo\b", RegexOptions.IgnoreCase))
            {
                if (metadata.Todo is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, procedureIndex + 1, j + 1,
                        "multiple @Todo annotations on one test procedure");
                }
                if (metadata.Skip is not null)
                {
                    throw new InvalidTestMetadataException(moduleName, procedureIndex + 1, j + 1,
                        "test cannot be both skipped and todo");
                }
                metadata.Todo = ParseSkipTodoAnnotation(prev, "Todo", moduleName, j + 1);
                metadata.TodoLine = j + 1;
                continue;
            }

            if (prev.StartsWith("''", StringComparison.Ordinal))
            {
                continue;
            }

            break;
        }
        metadata.TestCases.Reverse();
        return metadata;
    }

    internal static StatusReasonMetadata ParseSkipTodoAnnotation(string line, string annotationName, string moduleName = "", int lineNumber = 0)
    {
        var match = Regex.Match(line, @"^'\s*@(Skip|Todo)(?:\s*\((.*)\))?\s*$", RegexOptions.IgnoreCase);
        if (!match.Success || !string.Equals(match.Groups[1].Value, annotationName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, $"malformed @{annotationName} annotation");
        }

        var reasonExpression = match.Groups[2].Success ? match.Groups[2].Value.Trim() : "";
        if (reasonExpression == "" && line.Contains('('))
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, $"malformed @{annotationName} reason: expected a quoted string literal");
        }
        if (reasonExpression == "")
        {
            return new StatusReasonMetadata();
        }

        return new StatusReasonMetadata
        {
            Reason = ParseExpectedErrorStringArg(reasonExpression, moduleName, lineNumber, "reason", "@" + annotationName),
        };
    }

    private static string DeclarationHeader(string[] lines, int startIndex)
    {
        var builder = new StringBuilder();
        for (var i = startIndex; i < lines.Length; i++)
        {
            var line = StripComment(lines[i]).Trim();
            if (line.EndsWith('_'))
            {
                builder.Append(line[..^1].TrimEnd());
                builder.Append(' ');
                continue;
            }
            builder.Append(line);
            break;
        }
        return builder.ToString();
    }

    private static string StripComment(string line)
    {
        var inString = false;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] == '"')
            {
                if (inString && i + 1 < line.Length && line[i + 1] == '"')
                {
                    i++;
                    continue;
                }
                inString = !inString;
            }
            else if (line[i] == '\'' && !inString)
            {
                return line[..i];
            }
        }
        return line;
    }

    private static List<TestParameter> ParseProcedureParameters(string header, string moduleName, int lineNumber)
    {
        var open = header.IndexOf('(');
        if (open < 0)
        {
            return [];
        }
        var close = header.LastIndexOf(')');
        if (close < open)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed test procedure parameter list");
        }
        var body = header[(open + 1)..close].Trim();
        if (body.Length == 0)
        {
            return [];
        }
        var parts = SplitAnnotationArgs(body, "@TestCase", moduleName, lineNumber);
        var parameters = new List<TestParameter>();
        foreach (var part in parts)
        {
            parameters.Add(ParseProcedureParameter(part, moduleName, lineNumber));
        }
        return parameters;
    }

    private static TestParameter ParseProcedureParameter(string input, string moduleName, int lineNumber)
    {
        var text = input.Trim();
        var paramArray = Regex.IsMatch(text, @"\bParamArray\b", RegexOptions.IgnoreCase);
        var optional = Regex.IsMatch(text, @"\bOptional\b", RegexOptions.IgnoreCase);
        var passingMatch = Regex.Match(text, @"\b(ByVal|ByRef)\b", RegexOptions.IgnoreCase);
        var passing = passingMatch.Success ? NormalizeKeyword(passingMatch.Value) : "";
        var cleaned = Regex.Replace(text, @"\b(Optional|ByVal|ByRef|ParamArray)\b", "", RegexOptions.IgnoreCase).Trim();
        var asMatch = Regex.Match(cleaned, $@"^(?<name>{VbaIdentifierPattern.Identifier})(?:\(\))?\s*(?:As\s+(?<type>[A-Za-z][A-Za-z0-9_]*))?(?:\s*=.*)?$", RegexOptions.IgnoreCase);
        if (!asMatch.Success)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, $"unsupported test parameter declaration {input}");
        }
        var type = NormalizeParameterType(asMatch.Groups["type"].Success ? asMatch.Groups["type"].Value : "Variant");
        return new TestParameter(asMatch.Groups["name"].Value, type, passing, optional, paramArray);
    }

    private static List<TestCase> ExpandTestCases(string moduleName, string name, int line, List<TestParameter> parameters, TestMetadata metadata)
    {
        var qualifiedProcedure = moduleName + "." + name;
        if (parameters.Count == 0)
        {
            if (metadata.TestCases.Count > 0)
            {
                var first = metadata.TestCases[0];
                if (first.HasName || first.Arguments.Count > 0)
                {
                    throw new InvalidTestCaseException(moduleName, line, first.Line, "parameterless test must not declare @TestCase arguments");
                }
            }
            return
            [
                new TestCase
                {
                    Name = name,
                    Module = moduleName,
                    Line = line,
                    Tags = metadata.Tags.ToArray(),
                    Skip = metadata.Skip,
                    Todo = metadata.Todo,
                    ExpectedError = metadata.ExpectedError,
                },
            ];
        }

        if (metadata.TestCases.Count == 0)
        {
            throw new InvalidTestCaseException(moduleName, line, line, $"parameterized test {name} requires at least one @TestCase");
        }
        ValidateTestParameters(parameters, moduleName, line);

        var tests = new List<TestCase>();
        foreach (var testCase in metadata.TestCases)
        {
            if (testCase.Arguments.Count != parameters.Count)
            {
                throw new InvalidTestCaseException(moduleName, line, testCase.Line,
                    $"@TestCase provides {testCase.Arguments.Count} arguments, but {name} requires {parameters.Count}");
            }
            var caseId = testCase.HasName
                ? testCase.Name
                : string.Join(",", testCase.Arguments.Select(a => a.Canonical));
            var args = new List<TestArgument>();
            for (var i = 0; i < testCase.Arguments.Count; i++)
            {
                ValidateLiteralForType(testCase.Arguments[i], parameters[i], moduleName, testCase.Line, i + 1, name);
                args.Add(new TestArgument
                {
                    Type = parameters[i].Type,
                    Value = testCase.Arguments[i].Value,
                    VbaLiteral = testCase.Arguments[i].VbaLiteral,
                });
            }
            tests.Add(new TestCase
            {
                Name = name,
                Module = moduleName,
                CaseId = caseId,
                Line = line,
                AnnotationLine = testCase.Line,
                Arguments = args.ToArray(),
                Tags = metadata.Tags.ToArray(),
                Skip = metadata.Skip,
                Todo = metadata.Todo,
                ExpectedError = metadata.ExpectedError,
            });
        }

        var duplicate = tests.GroupBy(t => t.QualifiedName, StringComparer.OrdinalIgnoreCase).FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
        {
            throw new InvalidTestCaseException(moduleName, line, duplicate.Skip(1).First().AnnotationLine,
                $"duplicate generated test case id {duplicate.Key}");
        }
        return tests;
    }

    private static void ValidateTestParameters(List<TestParameter> parameters, string moduleName, int line)
    {
        foreach (var parameter in parameters)
        {
            if (parameter.Optional)
            {
                throw new InvalidTestCaseException(moduleName, line, line, "optional parameters are not supported in parameterized tests");
            }
            if (parameter.ParamArray)
            {
                throw new InvalidTestCaseException(moduleName, line, line, "ParamArray parameters are not supported in parameterized tests");
            }
            if (!parameter.Passing.Equals("ByVal", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidTestCaseException(moduleName, line, line, $"parameter {parameter.Name} must be ByVal");
            }
            if (!IsSupportedParameterType(parameter.Type))
            {
                throw new InvalidTestCaseException(moduleName, line, line, $"unsupported parameter type {parameter.Type} for {parameter.Name}");
            }
        }
    }

    private static ParsedTestCase ParseTestCaseAnnotation(string line, string moduleName = "", int lineNumber = 0)
    {
        var match = Regex.Match(line, @"^'\s*@TestCase\s*\((.*)\)\s*$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed @TestCase annotation");
        }
        var body = match.Groups[1].Value.Trim();
        var (nameExpr, argsExpr, hasName) = SplitTestCaseName(body, moduleName, lineNumber);
        var parsed = new ParsedTestCase { HasName = hasName };
        if (hasName)
        {
            var name = ParseExpectedErrorStringArg(nameExpr, moduleName, lineNumber, "name", "@TestCase");
            name = CanonicalCaseName(name);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "@TestCase name must not be empty");
            }
            parsed.Name = name;
        }
        if (string.IsNullOrWhiteSpace(argsExpr))
        {
            return parsed;
        }
        foreach (var part in SplitAnnotationArgs(argsExpr, "@TestCase", moduleName, lineNumber))
        {
            parsed.Arguments.Add(ParseTestLiteral(part, moduleName, lineNumber));
        }
        return parsed;
    }

    private static (string NameExpr, string ArgsExpr, bool HasName) SplitTestCaseName(string input, string moduleName, int lineNumber)
    {
        var inString = false;
        var inDate = false;
        for (var i = 0; i < input.Length; i++)
        {
            var ch = input[i];
            if (ch == '"' && !inDate)
            {
                if (inString && i + 1 < input.Length && input[i + 1] == '"')
                {
                    i++;
                    continue;
                }
                inString = !inString;
                continue;
            }
            if (ch == '#' && !inString)
            {
                inDate = !inDate;
                continue;
            }
            if (ch == ';' && !inString && !inDate)
            {
                return (input[..i].Trim(), input[(i + 1)..].Trim(), true);
            }
        }
        if (inString)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed string literal");
        }
        if (inDate)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed date literal");
        }
        return ("", input, false);
    }

    private static List<string> SplitAnnotationArgs(string input, string annotationName, string moduleName, int lineNumber)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var inString = false;
        var inDate = false;
        for (var i = 0; i < input.Length; i++)
        {
            var ch = input[i];
            if (ch == '"' && !inDate)
            {
                current.Append(ch);
                if (inString && i + 1 < input.Length && input[i + 1] == '"')
                {
                    i++;
                    current.Append(input[i]);
                    continue;
                }
                inString = !inString;
                continue;
            }
            if (ch == '#' && !inString)
            {
                inDate = !inDate;
                current.Append(ch);
                continue;
            }
            if (ch == ',' && !inString && !inDate)
            {
                args.Add(current.ToString().Trim());
                current.Clear();
                continue;
            }
            current.Append(ch);
        }
        if (inString)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed string literal");
        }
        if (inDate)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed date literal");
        }
        args.Add(current.ToString().Trim());
        if (args.Any(string.IsNullOrWhiteSpace))
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, $"{annotationName} arguments must not be empty");
        }
        return args;
    }

    private static TestLiteral ParseTestLiteral(string input, string moduleName, int lineNumber)
    {
        var raw = input.Trim();
        if (raw.StartsWith('"'))
        {
            var value = ParseExpectedErrorStringArg(raw, moduleName, lineNumber, "argument", "@TestCase");
            return new TestLiteral("string", CanonicalStringLiteral(value), value, CanonicalStringLiteral(value));
        }
        if (raw.StartsWith('#'))
        {
            if (raw.Length < 2 || raw[^1] != '#')
            {
                throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed date literal");
            }
            var value = raw[1..^1].Trim();
            if (value.Length == 0)
            {
                throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, "malformed date literal");
            }
            return new TestLiteral("date", "#" + value + "#", value, "#" + value + "#");
        }
        if (raw.Equals("True", StringComparison.OrdinalIgnoreCase))
        {
            return new TestLiteral("boolean", "True", true, "True");
        }
        if (raw.Equals("False", StringComparison.OrdinalIgnoreCase))
        {
            return new TestLiteral("boolean", "False", false, "False");
        }
        if (raw.Equals("Empty", StringComparison.OrdinalIgnoreCase))
        {
            return new TestLiteral("empty", "Empty", "Empty", "Empty");
        }
        if (raw.Equals("Null", StringComparison.OrdinalIgnoreCase))
        {
            return new TestLiteral("null", "Null", null, "Null");
        }
        var numeric = raw;
        var suffix = '\0';
        if (raw.Length > 0 && "#!@&^%".Contains(raw[^1], StringComparison.Ordinal))
        {
            suffix = raw[^1];
            numeric = raw[..^1];
        }
        if (Regex.IsMatch(numeric, @"^[+-]?\d+$"))
        {
            if (!long.TryParse(numeric, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            {
                throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, $"integer literal {raw} is out of range");
            }
            if (suffix is '#' or '!' or '@')
            {
                return new TestLiteral("float", raw, (double)value, raw);
            }
            return new TestLiteral("integer", raw, value, raw);
        }
        if (Regex.IsMatch(numeric, @"^[+-]?(?:\d+\.\d*|\.\d+|\d+[eE][+-]?\d+|\d+\.\d*[eE][+-]?\d+|\.\d+[eE][+-]?\d+)$"))
        {
            if (!double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, $"floating-point literal {raw} is invalid");
            }
            return new TestLiteral("float", raw, value, raw);
        }
        throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber, $"unsupported @TestCase literal {raw}");
    }

    private static void ValidateLiteralForType(TestLiteral literal, TestParameter parameter, string moduleName, int lineNumber, int argumentIndex, string procedureName)
    {
        if (parameter.Type.Equals("Variant", StringComparison.OrdinalIgnoreCase) || literal.Kind is "empty" or "null")
        {
            return;
        }
        var ok = parameter.Type.ToLowerInvariant() switch
        {
            "boolean" => literal.Kind == "boolean",
            "byte" or "integer" or "long" or "longlong" or "longptr" => literal.Kind == "integer",
            "single" or "double" or "currency" => literal.Kind is "integer" or "float",
            "date" => literal.Kind == "date",
            "string" => literal.Kind == "string",
            _ => false,
        };
        if (!ok)
        {
            throw new InvalidTestCaseException(moduleName, lineNumber, lineNumber,
                $"argument {argumentIndex} for {procedureName}: literal {literal.Canonical} cannot be passed to {parameter.Type}");
        }
    }

    private static string NormalizeParameterType(string input)
    {
        foreach (var type in SupportedTestParameterTypes)
        {
            if (input.Equals(type, StringComparison.OrdinalIgnoreCase))
            {
                return type;
            }
        }
        return input;
    }

    private static bool IsSupportedParameterType(string type)
    {
        return SupportedTestParameterTypes
            .Contains(type, StringComparer.OrdinalIgnoreCase);
    }

    private static string NormalizeKeyword(string value)
    {
        return value.Equals("ByVal", StringComparison.OrdinalIgnoreCase) ? "ByVal" :
            value.Equals("ByRef", StringComparison.OrdinalIgnoreCase) ? "ByRef" : value;
    }

    private static string CanonicalStringLiteral(string value)
    {
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    private static string CanonicalCaseName(string value)
    {
        return value.Trim().Replace("[", "_", StringComparison.Ordinal).Replace("]", "_", StringComparison.Ordinal);
    }

    internal static ExpectedErrorMetadata ParseExpectedErrorAnnotation(string line, string moduleName = "", int lineNumber = 0)
    {
        var match = Regex.Match(line, @"^'\s*@ExpectedError\s*\((.*)\)\s*$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, "malformed @ExpectedError annotation");
        }

        var args = SplitExpectedErrorArgs(match.Groups[1].Value, moduleName, lineNumber);
        if (args.Count is < 1 or > 3)
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, "@ExpectedError supports 1 to 3 arguments");
        }
        if (!int.TryParse(args[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, "@ExpectedError error number must be numeric");
        }

        return new ExpectedErrorMetadata
        {
            Number = number,
            Description = args.Count >= 2 ? ParseExpectedErrorStringArg(args[1], moduleName, lineNumber, "description") : null,
            Source = args.Count >= 3 ? ParseExpectedErrorStringArg(args[2], moduleName, lineNumber, "source") : null,
        };
    }

    private static List<string> SplitExpectedErrorArgs(string input, string moduleName, int lineNumber)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var inString = false;
        for (var i = 0; i < input.Length; i++)
        {
            var ch = input[i];
            if (ch == '"')
            {
                current.Append(ch);
                if (inString && i + 1 < input.Length && input[i + 1] == '"')
                {
                    i++;
                    current.Append(input[i]);
                    continue;
                }
                inString = !inString;
                continue;
            }
            if (ch == ',' && !inString)
            {
                args.Add(current.ToString().Trim());
                current.Clear();
                continue;
            }
            current.Append(ch);
        }
        if (inString)
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, "malformed string literal");
        }
        args.Add(current.ToString().Trim());
        if (args.Any(string.IsNullOrWhiteSpace))
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, "@ExpectedError arguments must not be empty");
        }
        return args;
    }

    private static string ParseExpectedErrorStringArg(string input, string moduleName, int lineNumber, string name, string annotationName = "@ExpectedError")
    {
        input = input.Trim();
        if (input.Length < 2 || input[0] != '"' || input[^1] != '"')
        {
            throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, $"malformed {annotationName} {name}: expected a quoted string literal");
        }
        var body = input[1..^1];
        var builder = new StringBuilder();
        for (var i = 0; i < body.Length; i++)
        {
            if (body[i] == '"')
            {
                if (i + 1 < body.Length && body[i + 1] == '"')
                {
                    builder.Append('"');
                    i++;
                    continue;
                }
                throw new InvalidTestMetadataException(moduleName, lineNumber, lineNumber, $"malformed {annotationName} {name}: unexpected quote");
            }
            builder.Append(body[i]);
        }
        return builder.ToString();
    }

    internal static ModuleHooks FindModuleHooks(string moduleName, string code)
    {
        var hooks = new ModuleHooks();
        if (string.IsNullOrEmpty(code))
        {
            return hooks;
        }

        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            var match = Regex.Match(line,
                @"^(?:Public\s+)?Sub\s+(BeforeAll|AfterAll|BeforeEach|AfterEach)\s*(?:\(\s*\))?\s*(?:'.*)?$",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            var name = match.Groups[1].Value;
            switch (name.ToLowerInvariant())
            {
                case "beforeall": hooks.BeforeAll = new HookInfo(name, moduleName, i + 1); break;
                case "afterall": hooks.AfterAll = new HookInfo(name, moduleName, i + 1); break;
                case "beforeeach": hooks.BeforeEach = new HookInfo(name, moduleName, i + 1); break;
                case "aftereach": hooks.AfterEach = new HookInfo(name, moduleName, i + 1); break;
            }
        }
        return hooks;
    }

    internal static List<string> DuplicateTestQualifiedNames(List<TestCase> tests)
    {
        return tests
            .GroupBy(t => t.QualifiedName, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.First().QualifiedName)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static TestSelection SelectTests(List<TestCase> tests, string filter, string moduleFilter, string tagFilter)
    {
        var candidates = new List<TestCase>();
        foreach (var test in tests)
        {
            var include = true;
            if (!string.IsNullOrWhiteSpace(moduleFilter) && !string.Equals(test.Module, moduleFilter, StringComparison.OrdinalIgnoreCase))
            {
                include = false;
            }
            if (!string.IsNullOrWhiteSpace(tagFilter))
            {
                var tagFound = test.Tags.Any(t => string.Equals(t, tagFilter, StringComparison.OrdinalIgnoreCase));
                if (!tagFound)
                {
                    include = false;
                }
            }
            if (include)
            {
                candidates.Add(test);
            }
        }
        filter = filter.Trim();
        if (string.IsNullOrWhiteSpace(filter))
        {
            return new TestSelection(candidates, [], false);
        }
        if (filter.Contains('.'))
        {
            return new TestSelection(
                candidates.Where(t =>
                    string.Equals(t.QualifiedName, filter, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(t.QualifiedProcedure, filter, StringComparison.OrdinalIgnoreCase)).ToList(),
                [],
                false);
        }

        var matches = candidates.Where(t => string.Equals(t.Name, filter, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count > 1)
        {
            return new TestSelection([], matches.OrderBy(t => t.QualifiedName, StringComparer.OrdinalIgnoreCase).ToList(), true);
        }
        return new TestSelection(matches, [], false);
    }

    internal static void AdjustCountsForAfterAllFailure(string resultStatus, ref int failed, ref int inconclusiveCount)
    {
        if (resultStatus == "passed")
        {
            failed++;
        }
        else if (resultStatus == "inconclusive")
        {
            failed++;
            inconclusiveCount--;
        }
    }

    internal static string BuildTestRunnerCode(List<TestCase> tests, Dictionary<string, ModuleHooks> hooksByModule)
    {
        var builder = new StringBuilder();
        builder.AppendLine("Option Explicit");
        builder.AppendLine();

        // Generate BeforeAll/AfterAll wrappers per module
        var moduleNames = tests.Select(t => t.Module).Distinct(StringComparer.OrdinalIgnoreCase);
        foreach (var mod in moduleNames)
        {
            if (!hooksByModule.TryGetValue(mod, out var hooks))
            {
                continue;
            }

            if (hooks.BeforeAll is not null)
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"Public Function RunBeforeAll_{mod}() As Variant");
                builder.AppendLine("  On Error GoTo Handler");
                builder.AppendLine(CultureInfo.InvariantCulture, $"  {mod}.{hooks.BeforeAll.Name}");
                builder.AppendLine("  RunBeforeAll_" + mod + " = Array(True, CLng(0), \"\", \"\", \"\")");
                builder.AppendLine("  Exit Function");
                builder.AppendLine("Handler:");
                builder.AppendLine("  RunBeforeAll_" + mod + " = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description), \"\")");
                builder.AppendLine("  Err.Clear");
                builder.AppendLine("End Function");
                builder.AppendLine();
            }

            if (hooks.AfterAll is not null)
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"Public Function RunAfterAll_{mod}() As Variant");
                builder.AppendLine("  On Error GoTo Handler");
                builder.AppendLine(CultureInfo.InvariantCulture, $"  {mod}.{hooks.AfterAll.Name}");
                builder.AppendLine("  RunAfterAll_" + mod + " = Array(True, CLng(0), \"\", \"\", \"\")");
                builder.AppendLine("  Exit Function");
                builder.AppendLine("Handler:");
                builder.AppendLine("  RunAfterAll_" + mod + " = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description), \"\")");
                builder.AppendLine("  Err.Clear");
                builder.AppendLine("End Function");
                builder.AppendLine();
            }
        }

        // Generate RunTest dispatch
        builder.AppendLine("Public Function RunTest(ByVal testIndex As Long) As Variant");
        builder.AppendLine("  On Error Resume Next");
        builder.AppendLine("  Err.Clear");
        builder.AppendLine("  Dim beforeEachErr As Variant");
        builder.AppendLine("  Dim testErr As Variant");
        builder.AppendLine("  Dim afterEachErr As Variant");
        builder.AppendLine("  Dim statusHint As String");
        builder.AppendLine("  Dim phaseHint As String");
        builder.AppendLine("  Select Case testIndex");

        for (var i = 0; i < tests.Count; i++)
        {
            var test = tests[i];
            hooksByModule.TryGetValue(test.Module, out var hooks);
            var beforeEachName = hooks?.BeforeEach?.Name ?? "";
            var afterEachName = hooks?.AfterEach?.Name ?? "";

            builder.AppendLine(CultureInfo.InvariantCulture, $"    Case {i}");
            builder.AppendLine("      statusHint = \"\"");
            builder.AppendLine("      phaseHint = \"\"");

            if (!string.IsNullOrEmpty(beforeEachName))
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"      {test.Module}.{beforeEachName}");
                builder.AppendLine("      If Err.Number <> 0 Then");
                builder.AppendLine("        beforeEachErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))");
                builder.AppendLine("        phaseHint = \"before_each\"");
                builder.AppendLine("        Err.Clear");
                builder.AppendLine("      End If");
            }

            builder.AppendLine("      If IsEmpty(beforeEachErr) Then");
            var argumentList = test.Arguments.Length == 0 ? "" : " " + string.Join(", ", test.Arguments.Select(a => a.VbaLiteral));
            builder.AppendLine(CultureInfo.InvariantCulture, $"        {test.Module}.{test.Name}{argumentList}");
            builder.AppendLine("        If Err.Number <> 0 Then");
            builder.AppendLine("          If Err.Number = vbObjectError + 516 Then");
            builder.AppendLine("            statusHint = \"inconclusive\"");
            builder.AppendLine("          End If");
            builder.AppendLine("          testErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))");
            builder.AppendLine("          phaseHint = \"test\"");
            builder.AppendLine("          Err.Clear");
            builder.AppendLine("        End If");
            builder.AppendLine("      End If");

            if (!string.IsNullOrEmpty(afterEachName))
            {
                builder.AppendLine(CultureInfo.InvariantCulture, $"      {test.Module}.{afterEachName}");
                builder.AppendLine("      If Err.Number <> 0 Then");
                builder.AppendLine("        afterEachErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))");
                builder.AppendLine("        If phaseHint = \"\" Then");
                builder.AppendLine("          phaseHint = \"after_each\"");
                builder.AppendLine("        End If");
                builder.AppendLine("        Err.Clear");
                builder.AppendLine("      End If");
            }

            builder.AppendLine("      If Not IsEmpty(afterEachErr) Then");
            builder.AppendLine("        phaseHint = \"after_each\"");
            builder.AppendLine("        statusHint = \"failed\"");
            builder.AppendLine("        RunTest = Array(False, afterEachErr(1), afterEachErr(2), afterEachErr(3), statusHint, phaseHint)");
            builder.AppendLine("      ElseIf Not IsEmpty(testErr) Then");
            builder.AppendLine("        RunTest = Array(False, testErr(1), testErr(2), testErr(3), statusHint, phaseHint)");
            builder.AppendLine("      ElseIf Not IsEmpty(beforeEachErr) Then");
            builder.AppendLine("        RunTest = Array(False, beforeEachErr(1), beforeEachErr(2), beforeEachErr(3), statusHint, phaseHint)");
            builder.AppendLine("      Else");
            builder.AppendLine("        RunTest = Array(True, CLng(0), \"\", \"\", statusHint, phaseHint)");
            builder.AppendLine("      End If");
        }

        builder.AppendLine("  End Select");
        builder.AppendLine("  Err.Clear");
        builder.AppendLine("End Function");
        return builder.ToString();
    }

    internal static TestClassification ClassifyTestOutcome(
        TestCase test,
        bool success,
        int errNumber,
        string errSource,
        string errDescription,
        string statusHint,
        string phaseHint)
    {
        if (!success)
        {
            if (statusHint == "inconclusive")
            {
                return new TestClassification("inconclusive", "test_inconclusive", errDescription, errSource, errNumber, null);
            }

            var hookCode = phaseHint switch
            {
                "before_each" => "before_each_failed",
                "after_each" => "after_each_failed",
                _ => "",
            };
            if (hookCode != "")
            {
                return new TestClassification("failed", hookCode, errDescription, errSource, errNumber, null);
            }

            if (phaseHint == "test" && test.ExpectedError is not null)
            {
                var mismatch = ExpectedErrorMismatch(test.ExpectedError, errNumber, errSource, errDescription);
                if (mismatch == "")
                {
                    return new TestClassification("passed", "", "", "", 0,
                        BuildObservedError(errNumber, errSource, errDescription));
                }
                return new TestClassification("failed", "expected_error_mismatch", mismatch, errSource, errNumber,
                    BuildObservedError(errNumber, errSource, errDescription));
            }

            return new TestClassification("failed", "test_failed", errDescription, errSource, errNumber, null);
        }

        if (test.ExpectedError is not null)
        {
            return new TestClassification("failed", "expected_error_mismatch",
                $"expected VBA error {test.ExpectedError.Number} but no error was raised", "", 0, null);
        }

        return new TestClassification("passed", "", "", "", 0, null);
    }

    private static string ExpectedErrorMismatch(ExpectedErrorMetadata expected, int errNumber, string errSource, string errDescription)
    {
        if (errNumber != expected.Number)
        {
            return $"expected VBA error {expected.Number} but got error {errNumber}: {errDescription}";
        }
        if (expected.Description is not null && errDescription != expected.Description)
        {
            return $"expected error description <{expected.Description}> but got <{errDescription}>";
        }
        if (expected.Source is not null && !string.Equals(errSource, expected.Source, StringComparison.OrdinalIgnoreCase))
        {
            return $"expected error source <{expected.Source}> but got <{errSource}>";
        }
        return "";
    }

    internal static Dictionary<string, object?> BuildTestResult(TestCase test, string status, int durationMs, object error, object? observedError = null)
    {
        var result = new Dictionary<string, object?>
        {
            ["id"] = test.QualifiedName,
            ["qualified_name"] = test.QualifiedName,
            ["name"] = test.Name,
            ["module"] = test.Module,
            ["status"] = status,
            ["duration_ms"] = durationMs,
            ["qualified_procedure"] = test.QualifiedProcedure,
            ["procedure_line"] = test.ProcedureLine,
            ["tags"] = test.Tags,
            ["attempts"] = 1,
            ["flaky"] = false,
        };
        if (!string.IsNullOrEmpty(test.CaseId))
        {
            result["case_id"] = test.CaseId;
            result["annotation_line"] = test.AnnotationLine;
            result["arguments"] = test.Arguments.Select(BuildTestArgument).ToArray();
        }
        if (status != "passed")
        {
            result["error"] = error;
        }
        if (test.ExpectedError is not null)
        {
            result["expected_error"] = BuildExpectedError(test.ExpectedError);
        }
        if (observedError is not null)
        {
            result["observed_error"] = observedError;
        }
        result["attempt_results"] = new[] { BuildAttemptResult(1, status, durationMs, status != "passed" ? error : null) };
        return result;
    }

    private static Dictionary<string, object?> BuildAttemptResult(int attempt, string status, int durationMs, object? error)
    {
        var result = new Dictionary<string, object?>
        {
            ["attempt"] = attempt,
            ["status"] = status,
            ["duration_ms"] = durationMs,
        };
        if (error is not null)
        {
            result["error"] = error;
        }
        return result;
    }

    internal static Dictionary<string, object?> BuildNotRunTestResult(TestCase test, string reason)
    {
        var result = new Dictionary<string, object?>
        {
            ["id"] = test.QualifiedName,
            ["qualified_name"] = test.QualifiedName,
            ["name"] = test.Name,
            ["module"] = test.Module,
            ["status"] = "not_run",
            ["duration_ms"] = 0,
            ["qualified_procedure"] = test.QualifiedProcedure,
            ["procedure_line"] = test.ProcedureLine,
            ["tags"] = test.Tags,
            ["reason"] = reason,
        };
        if (!string.IsNullOrEmpty(test.CaseId))
        {
            result["case_id"] = test.CaseId;
            result["annotation_line"] = test.AnnotationLine;
            result["arguments"] = test.Arguments.Select(BuildTestArgument).ToArray();
        }
        if (test.ExpectedError is not null)
        {
            result["expected_error"] = BuildExpectedError(test.ExpectedError);
        }
        return result;
    }

    internal static Dictionary<string, object?> BuildNonExecutedTestResult(TestCase test)
    {
        var status = test.Skip is not null ? "skipped" : "todo";
        var result = new Dictionary<string, object?>
        {
            ["id"] = test.QualifiedName,
            ["qualified_name"] = test.QualifiedName,
            ["name"] = test.Name,
            ["module"] = test.Module,
            ["status"] = status,
            ["duration_ms"] = 0,
            ["qualified_procedure"] = test.QualifiedProcedure,
            ["procedure_line"] = test.ProcedureLine,
            ["tags"] = test.Tags,
        };
        if (!string.IsNullOrEmpty(test.CaseId))
        {
            result["case_id"] = test.CaseId;
            result["annotation_line"] = test.AnnotationLine;
            result["arguments"] = test.Arguments.Select(BuildTestArgument).ToArray();
        }
        var reason = NonExecutedReason(test);
        if (!string.IsNullOrEmpty(reason))
        {
            result["reason"] = reason;
        }
        if (test.ExpectedError is not null)
        {
            result["expected_error"] = BuildExpectedError(test.ExpectedError);
        }
        return result;
    }

    private static string NonExecutedLogLine(TestCase test)
    {
        var status = test.Skip is not null ? "SKIP" : "TODO";
        var reason = NonExecutedReason(test);
        return string.IsNullOrEmpty(reason)
            ? $"{status} {test.QualifiedName}"
            : $"{status} {test.QualifiedName}: {reason}";
    }

    private static string NonExecutedReason(TestCase test)
    {
        return test.Skip?.Reason ?? test.Todo?.Reason ?? "";
    }

    private static bool IsNonExecutedTest(TestCase test)
    {
        return test.Skip is not null || test.Todo is not null;
    }

    private static bool ShouldStopAfterFailures(TestCommandArguments args, int failed)
    {
        return args.MaxFailures > 0 && failed >= args.MaxFailures;
    }

    private static Dictionary<string, object?> BuildExecutionPayload(TestCommandArguments args, bool stoppedEarly, string stopReason)
    {
        var payload = new Dictionary<string, object?>
        {
            ["fail_fast"] = args.FailFast,
            ["rerun_failed"] = args.RerunFailed,
            ["stopped_early"] = stoppedEarly,
        };
        if (args.MaxFailures > 0)
        {
            payload["max_failures"] = args.MaxFailures;
        }
        if (!string.IsNullOrWhiteSpace(stopReason))
        {
            payload["stop_reason"] = stopReason;
        }
        return payload;
    }

    private static Dictionary<string, object?> BuildExpectedError(ExpectedErrorMetadata expected)
    {
        var payload = new Dictionary<string, object?> { ["number"] = expected.Number };
        if (expected.Description is not null)
        {
            payload["description"] = expected.Description;
        }
        if (expected.Source is not null)
        {
            payload["source"] = expected.Source;
        }
        return payload;
    }

    private static object? BuildObservedError(int number, string source, string message)
    {
        if (number == 0 && source == "" && message == "")
        {
            return null;
        }
        return new { number, source, message };
    }

    internal static BridgeResponse BuildErrorResponse(
        BridgeRequest request,
        TestCommandArguments args,
        string code,
        string message,
        bool sessionAttached,
        string sessionMode,
        List<TestCase> tests,
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState,
        bool runtimeInjected,
        IReadOnlyDictionary<string, object?>? errorDetails = null,
        string? displayWorkbookPath = null)
    {
        var logs = new List<string>();
        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null)
        {
            logs.Add(sessionLog);
        }

        var extensions = new Dictionary<string, object?>
        {
            ["workbook"] = new
            {
                path = displayWorkbookPath ?? DisplayWorkbookPath(args),
                session = sessionAttached,
                session_mode = sessionMode,
                session_requested = args.UseSession,
                auto_session = sessionAttached && !args.UseSession,
            },
            ["tests"] = tests.Select(BuildDiscoveredTestPayload).ToArray(),
        };

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(
                Code: code,
                Message: message,
                Phase: "test",
                Source: "xlflow",
                Details: errorDetails),
            Logs = logs,
            Extensions = extensions,
        };
    }

    private static BridgeResponse BuildSuccessfulTestResponse(
        BridgeRequest request,
        TestCommandArguments args,
        string displayWorkbookPath,
        bool sessionAttached,
        string sessionMode,
        bool saveAfterRun,
        object workbook,
        List<string> logs,
        List<object> results,
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState)
    {
        if (runtimeState is not null)
        {
            RuntimeInjectionHelper.RestoreRuntimeInjection(workbook, runtimeState);
        }

        var saved = false;
        if (saveAfterRun)
        {
            ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));
            saved = true;
        }

        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null)
        {
            logs.Insert(0, sessionLog);
        }

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Ok,
            Logs = logs,
            Extensions = new Dictionary<string, object?>
            {
                ["workbook"] = new
                {
                    path = displayWorkbookPath,
                    session = sessionAttached,
                    session_mode = sessionMode,
                    session_requested = args.UseSession,
                    auto_session = sessionAttached && !args.UseSession,
                    saved,
                    dirty = sessionAttached && !saved,
                    needs_save = sessionAttached && !saved,
                },
                ["tests"] = results.ToArray(),
                ["execution"] = BuildExecutionPayload(args, stoppedEarly: false, stopReason: ""),
            },
        };
    }

    private static object BuildDiscoveredTestPayload(TestCase test)
    {
        var payload = new Dictionary<string, object?>
        {
            ["id"] = test.QualifiedName,
            ["qualified_name"] = test.QualifiedName,
            ["name"] = test.Name,
            ["module"] = test.Module,
            ["line"] = test.Line,
            ["qualified_procedure"] = test.QualifiedProcedure,
            ["procedure_line"] = test.ProcedureLine,
            ["tags"] = test.Tags,
        };
        if (!string.IsNullOrEmpty(test.CaseId))
        {
            payload["case_id"] = test.CaseId;
            payload["annotation_line"] = test.AnnotationLine;
            payload["arguments"] = test.Arguments.Select(BuildTestArgument).ToArray();
        }
        if (test.Skip is not null)
        {
            payload["status_hint"] = "skipped";
            payload["skip"] = BuildStatusReason(test.Skip);
        }
        if (test.Todo is not null)
        {
            payload["status_hint"] = "todo";
            payload["todo"] = BuildStatusReason(test.Todo);
        }
        if (test.ExpectedError is not null)
        {
            payload["expected_error"] = BuildExpectedError(test.ExpectedError);
        }
        return payload;
    }

    private static Dictionary<string, object?> BuildTestArgument(TestArgument argument)
    {
        return new Dictionary<string, object?>
        {
            ["type"] = argument.Type,
            ["value"] = argument.Value,
        };
    }

    private static Dictionary<string, object?> BuildStatusReason(StatusReasonMetadata metadata)
    {
        var payload = new Dictionary<string, object?>();
        if (!string.IsNullOrEmpty(metadata.Reason))
        {
            payload["reason"] = metadata.Reason;
        }
        return payload;
    }

    private static string ResultString(object result, string key)
    {
        if (result is IReadOnlyDictionary<string, object?> readOnly && readOnly.TryGetValue(key, out var readOnlyValue))
        {
            return Convert.ToString(readOnlyValue, CultureInfo.InvariantCulture) ?? "";
        }
        if (result is IDictionary<string, object?> dictionary && dictionary.TryGetValue(key, out var value))
        {
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
        return Convert.ToString(result.GetType().GetProperty(key)?.GetValue(result), CultureInfo.InvariantCulture) ?? "";
    }

    private static int ResultInt(object result, string key)
    {
        if (result is IReadOnlyDictionary<string, object?> readOnly && readOnly.TryGetValue(key, out var readOnlyValue))
        {
            return Convert.ToInt32(readOnlyValue, CultureInfo.InvariantCulture);
        }
        if (result is IDictionary<string, object?> dictionary && dictionary.TryGetValue(key, out var value))
        {
            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        return Convert.ToInt32(result.GetType().GetProperty(key)?.GetValue(result), CultureInfo.InvariantCulture);
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

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForTest(
        string workbookPath, string metadataPath, bool useSession, bool visible, bool disableAutoSession)
    {
        if (useSession)
        {
            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, true);
            return (attachment.Excel, attachment.Workbook, true, attachment.SessionMode);
        }
        if (!disableAutoSession && ExcelBridgeSupport.SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
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

    private static string? NormalizeIsolation(string isolation)
    {
        var normalized = string.IsNullOrWhiteSpace(isolation) ? "none" : isolation.Trim().ToLowerInvariant();
        return normalized is "none" or "module" or "test" ? normalized : null;
    }

    private static string SourceWorkbookPath(TestCommandArguments args)
    {
        return string.IsNullOrWhiteSpace(args.SourceWorkbookPath) ? args.WorkbookPath : args.SourceWorkbookPath;
    }

    private static string DisplayWorkbookPath(TestCommandArguments args)
    {
        return SourceWorkbookPath(args);
    }

    private static string TempRunRoot(TestCommandArguments args)
    {
        if (!string.IsNullOrWhiteSpace(args.TempRunRoot))
        {
            return args.TempRunRoot;
        }
        var root = string.IsNullOrWhiteSpace(args.ProjectRoot)
            ? Path.GetDirectoryName(SourceWorkbookPath(args)) ?? "."
            : args.ProjectRoot;
        return Path.Combine(root, ".xlflow", "test-runs");
    }

    internal static string CopyWorkbookForTest(string sourceWorkbook, string runDir, string segment)
    {
        Directory.CreateDirectory(runDir);
        var extension = Path.GetExtension(sourceWorkbook);
        var fileName = SanitizeFileSegment(segment);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = "workbook";
        }
        var destination = Path.Combine(runDir, fileName + extension);
        var suffix = 1;
        while (File.Exists(destination))
        {
            destination = Path.Combine(runDir, fileName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
            suffix++;
        }
        File.Copy(sourceWorkbook, destination);
        return destination;
    }

    internal static string SanitizeFileSegment(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder();
        foreach (var ch in value)
        {
            builder.Append(invalid.Contains(ch) ? '_' : ch);
        }
        return builder.ToString().Trim();
    }

    private static CleanupInfo CleanupRunDirectory(string runDir, string projectRoot)
    {
        try
        {
            Directory.Delete(runDir, recursive: true);
            return new CleanupInfo("completed", null, null);
        }
        catch (Exception ex)
        {
            return new CleanupInfo("failed", DisplayCleanupPath(runDir, projectRoot), ex.Message);
        }
    }

    private static string DisplayCleanupPath(string path, string projectRoot)
    {
        if (!string.IsNullOrWhiteSpace(projectRoot))
        {
            try
            {
                var relative = Path.GetRelativePath(projectRoot, path);
                if (!relative.StartsWith("..", StringComparison.Ordinal) && !Path.IsPathRooted(relative))
                {
                    return relative;
                }
            }
            catch
            {
                // fall through to the original path
            }
        }
        return path;
    }

    private static void AttachTestRunMetadata(
        BridgeResponse response,
        TestCommandArguments args,
        string isolation,
        bool session,
        bool temporaryWorkbook,
        bool workbookSaved,
        CleanupInfo? cleanup)
    {
        var cleanupPayload = new Dictionary<string, object?>
        {
            ["status"] = cleanup?.Status ?? "not_applicable",
        };
        if (cleanup is not null && cleanup.Status != "completed")
        {
            cleanupPayload["path"] = cleanup.Path;
            cleanupPayload["message"] = cleanup.Message;
        }

        var execution = response.Extensions.TryGetValue("execution", out var executionPayload) && executionPayload is not null
            ? executionPayload
            : BuildExecutionPayload(args, stoppedEarly: false, stopReason: "");

        response.Extensions["test_run"] = new
        {
            isolation,
            session,
            temporary_workbook = temporaryWorkbook,
            source_workbook = SourceWorkbookPath(args),
            workbook_saved = workbookSaved,
            cleanup = cleanupPayload,
            execution,
        };
    }

    private static bool WorkbookSaved(BridgeResponse response, bool fallback)
    {
        if (!response.Extensions.TryGetValue("workbook", out var workbook) || workbook is null)
        {
            return fallback;
        }
        var value = workbook.GetType().GetProperty("saved")?.GetValue(workbook);
        return value is bool saved ? saved : fallback;
    }

    internal static int CountFailedResults(BridgeResponse response)
    {
        if (!response.Extensions.TryGetValue("tests", out var testsPayload) || testsPayload is not IEnumerable<object> tests)
        {
            return response.Error?.Code == "test_failed" ? 1 : 0;
        }

        var failed = 0;
        foreach (var test in tests)
        {
            var status = ResultString(test, "status");
            if (status == "failed")
            {
                failed++;
            }
        }
        return failed;
    }

    private static int CountFailedResultObjects(IEnumerable<object> tests)
    {
        var failed = 0;
        foreach (var test in tests)
        {
            if (ResultString(test, "status") == "failed")
            {
                failed++;
            }
        }
        return failed;
    }

    private static List<object> RetryFailedResults(
        BridgeRequest request,
        TestCommandArguments args,
        string sourceWorkbook,
        string runDir,
        string isolation,
        List<TestCase> selected,
        List<object> initialResults,
        List<string> logs,
        CancellationToken cancellationToken)
    {
        var current = initialResults.ToDictionary(result => ResultString(result, "qualified_name"), result => result, StringComparer.OrdinalIgnoreCase);
        var testsByName = selected.ToDictionary(test => test.QualifiedName, StringComparer.OrdinalIgnoreCase);
        var retryArgs = args with { RerunFailed = 0, MaxFailures = 0, FailFast = false };

        for (var attempt = 2; attempt <= args.RerunFailed + 1; attempt++)
        {
            var failedTests = selected
                .Where(test => current.TryGetValue(test.QualifiedName, out var result) && IsRetryableFinalFailure(result))
                .ToList();
            if (failedTests.Count == 0)
            {
                break;
            }

            var unitName = isolation == "module" && failedTests.Count > 0 ? failedTests[0].Module : "retry-" + attempt.ToString(CultureInfo.InvariantCulture);
            var retryWorkbook = CopyWorkbookForTest(sourceWorkbook, runDir, SanitizeFileSegment(unitName + "-attempt-" + attempt.ToString(CultureInfo.InvariantCulture)));
            var retryResponse = ExecuteSingleWorkbook(
                request,
                retryArgs,
                retryWorkbook,
                sourceWorkbook,
                explicitTests: failedTests,
                saveAfterRun: false,
                cancellationToken);
            logs.AddRange(retryResponse.Logs);

            if (!retryResponse.Extensions.TryGetValue("tests", out var testsPayload) || testsPayload is not IEnumerable<object> retryItems)
            {
                continue;
            }

            foreach (var retryItem in retryItems)
            {
                var qualifiedName = ResultString(retryItem, "qualified_name");
                if (qualifiedName == "" || !testsByName.ContainsKey(qualifiedName) || !current.TryGetValue(qualifiedName, out var previous))
                {
                    continue;
                }
                current[qualifiedName] = MergeAttemptResult(previous, retryItem, attempt);
            }
        }

        return initialResults
            .Select(result => current.TryGetValue(ResultString(result, "qualified_name"), out var merged) ? merged : result)
            .ToList();
    }

    private static bool IsRetryableFinalFailure(object result)
    {
        if (ResultString(result, "status") != "failed")
        {
            return false;
        }
        var code = ResultErrorCode(result);
        return code is not "before_all_failed" and not "after_all_failed";
    }

    private static Dictionary<string, object?> MergeAttemptResult(object previous, object retry, int attempt)
    {
        var merged = CloneResultDictionary(retry);
        var history = ResultAttemptHistory(previous);
        history.Add(BuildAttemptResult(attempt, ResultString(retry, "status"), ResultInt(retry, "duration_ms"), ResultError(retry)));
        var hadFailedAttempt = history.Any(item => ResultString(item, "status") == "failed");
        var finalStatus = ResultString(retry, "status");
        merged["attempts"] = history.Count;
        merged["flaky"] = finalStatus == "passed" && hadFailedAttempt;
        merged["attempt_results"] = history.ToArray();
        return merged;
    }

    private static Dictionary<string, object?> CloneResultDictionary(object result)
    {
        if (result is IReadOnlyDictionary<string, object?> readOnly)
        {
            return readOnly.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        }
        if (result is IDictionary<string, object?> dictionary)
        {
            return dictionary.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        }
        return result.GetType().GetProperties().ToDictionary(property => property.Name, property => property.GetValue(result), StringComparer.OrdinalIgnoreCase);
    }

    private static List<object> ResultAttemptHistory(object result)
    {
        if (ResultValue(result, "attempt_results") is IEnumerable<object> attempts)
        {
            return attempts.ToList();
        }
        return [BuildAttemptResult(1, ResultString(result, "status"), ResultInt(result, "duration_ms"), ResultError(result))];
    }

    private static object? ResultError(object result)
    {
        return ResultValue(result, "error");
    }

    private static string ResultErrorCode(object result)
    {
        var error = ResultError(result);
        if (error is null)
        {
            return "";
        }
        if (error is IReadOnlyDictionary<string, object?> readOnly && readOnly.TryGetValue("code", out var readOnlyValue))
        {
            return Convert.ToString(readOnlyValue, CultureInfo.InvariantCulture) ?? "";
        }
        if (error is IDictionary<string, object?> dictionary && dictionary.TryGetValue("code", out var value))
        {
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
        return Convert.ToString(error.GetType().GetProperty("code")?.GetValue(error) ?? error.GetType().GetProperty("Code")?.GetValue(error), CultureInfo.InvariantCulture) ?? "";
    }

    private static object? ResultValue(object result, string key)
    {
        if (result is IReadOnlyDictionary<string, object?> readOnly && readOnly.TryGetValue(key, out var readOnlyValue))
        {
            return readOnlyValue;
        }
        if (result is IDictionary<string, object?> dictionary && dictionary.TryGetValue(key, out var value))
        {
            return value;
        }
        return result.GetType().GetProperty(key)?.GetValue(result);
    }

    private static string FirstNonEmpty(params string[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }
        return "";
    }

    internal sealed record CleanupInfo(string Status, string? Path, string? Message);

    private static void CloseWorkbook(object? workbook, object? excel, int ownedProcessId)
    {
        ExcelBridgeSupport.CloseWorkbookAndQuitApplication(workbook, excel, ownedProcessId);
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

    internal sealed class TestCase
    {
        public string Name { get; init; } = "";
        public string Module { get; init; } = "";
        public string ProcedureName => Name;
        public string QualifiedProcedure => Module + "." + ProcedureName;
        public string CaseId { get; init; } = "";
        public string QualifiedName => string.IsNullOrEmpty(CaseId) ? QualifiedProcedure : QualifiedProcedure + "[" + CaseId + "]";
        public string Id => QualifiedName;
        public int Line { get; init; }
        public int AnnotationLine { get; init; }
        public int ProcedureLine => Line;
        public TestArgument[] Arguments { get; init; } = [];
        public string[] Tags { get; init; } = [];
        public StatusReasonMetadata? Skip { get; init; }
        public StatusReasonMetadata? Todo { get; init; }
        public ExpectedErrorMetadata? ExpectedError { get; init; }
        public int Index { get; set; }
    }

    internal sealed class TestArgument
    {
        public string Type { get; init; } = "";
        public object? Value { get; init; }
        public string VbaLiteral { get; init; } = "";
    }

    internal sealed class StatusReasonMetadata
    {
        public string? Reason { get; init; }
    }

    internal sealed class ExpectedErrorMetadata
    {
        public int Number { get; init; }
        public string? Description { get; init; }
        public string? Source { get; init; }
    }

    internal sealed record TestClassification(
        string Status,
        string ErrorCode,
        string ErrorMessage,
        string ErrorSource,
        int ErrorNumber,
        object? ObservedError);

    private sealed class TestMetadata
    {
        public List<string> Tags { get; } = [];
        public List<ParsedTestCase> TestCases { get; } = [];
        public StatusReasonMetadata? Skip { get; set; }
        public int SkipLine { get; set; }
        public StatusReasonMetadata? Todo { get; set; }
        public int TodoLine { get; set; }
        public ExpectedErrorMetadata? ExpectedError { get; set; }
        public int ExpectedErrorLine { get; set; }
    }

    private sealed class ParsedTestCase
    {
        public string Name { get; set; } = "";
        public bool HasName { get; set; }
        public List<TestLiteral> Arguments { get; } = [];
        public int Line { get; set; }
    }

    private sealed record TestLiteral(string Kind, string Canonical, object? Value, string VbaLiteral);

    private sealed record TestParameter(string Name, string Type, string Passing, bool Optional, bool ParamArray);

    internal sealed class InvalidTestMetadataException : Exception
    {
        public InvalidTestMetadataException(string module, int procedureLine, int metadataLine, string message)
            : base(BuildMessage(module, procedureLine, metadataLine, message))
        {
            Module = module;
            ProcedureLine = procedureLine;
            MetadataLine = metadataLine;
        }

        public string Module { get; }
        public int ProcedureLine { get; }
        public int MetadataLine { get; }

        private static string BuildMessage(string module, int procedureLine, int metadataLine, string message)
        {
            var locationLine = metadataLine > 0 ? metadataLine : procedureLine;
            if (!string.IsNullOrWhiteSpace(module) && locationLine > 0)
            {
                return $"module {module}:{locationLine}: {message}";
            }
            if (!string.IsNullOrWhiteSpace(module))
            {
                return $"module {module}: {message}";
            }
            return message;
        }
    }

    internal sealed class InvalidTestCaseException : Exception
    {
        public InvalidTestCaseException(string module, int procedureLine, int metadataLine, string message)
            : base(BuildMessage(module, procedureLine, metadataLine, message))
        {
            Module = module;
            ProcedureLine = procedureLine;
            MetadataLine = metadataLine;
        }

        public string Module { get; }
        public int ProcedureLine { get; }
        public int MetadataLine { get; }

        private static string BuildMessage(string module, int procedureLine, int metadataLine, string message)
        {
            var locationLine = metadataLine > 0 ? metadataLine : procedureLine;
            if (!string.IsNullOrWhiteSpace(module) && locationLine > 0)
            {
                return $"module {module}:{locationLine}: {message}";
            }
            if (!string.IsNullOrWhiteSpace(module))
            {
                return $"module {module}: {message}";
            }
            return message;
        }
    }

    internal sealed record TestSelection(List<TestCase> Tests, List<TestCase> Matches, bool Ambiguous);

    internal sealed class ModuleHooks
    {
        public HookInfo? BeforeAll { get; set; }
        public HookInfo? AfterAll { get; set; }
        public HookInfo? BeforeEach { get; set; }
        public HookInfo? AfterEach { get; set; }
    }

    internal sealed record HookInfo(string Name, string Module, int Line);
}
