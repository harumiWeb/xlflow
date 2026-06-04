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

    public BridgeResponse Execute(BridgeRequest request, TestCommandArguments args, CancellationToken cancellationToken)
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

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForTest(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

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

            var runtimeInjected = runtimeState.Applied;

            vbProject = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"))
                ?? throw new InvalidOperationException("VBIDE access is not available.");

            RuntimeInjectionHelper.EnableUIStreamInjection(workbook, vbProject, runtimeState);

            var discovered = DiscoverTests(vbProject, cancellationToken);

            if (discovered.Count == 0)
            {
                return BuildErrorResponse(request, args, "no_tests_found", "no VBA tests found",
                    sessionAttached, sessionMode, [], runtimeState, runtimeInjected);
            }

            var duplicates = discovered
                .GroupBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();
            if (duplicates.Count > 0)
            {
                var names = string.Join(", ", duplicates.Select(g => g.Key));
                return BuildErrorResponse(request, args, "duplicate_test_name",
                    $"duplicate VBA test name(s): {names}",
                    sessionAttached, sessionMode, discovered, runtimeState, runtimeInjected);
            }

            var selected = SelectTests(discovered, args.Filter, args.ModuleFilter, args.TagFilter);
            if (selected.Count == 0)
            {
                var filterDesc = args.Filter;
                if (string.IsNullOrWhiteSpace(filterDesc)) filterDesc = args.ModuleFilter;
                if (string.IsNullOrWhiteSpace(filterDesc)) filterDesc = args.TagFilter;
                if (string.IsNullOrWhiteSpace(filterDesc)) filterDesc = "(no filter)";
                return BuildErrorResponse(request, args, "test_not_found",
                    $"test not found: {filterDesc}",
                    sessionAttached, sessionMode, [], runtimeState, runtimeInjected);
            }

            // Build hooks map
            var hooksByModule = new Dictionary<string, ModuleHooks>(StringComparer.OrdinalIgnoreCase);
            foreach (var moduleGroup in selected.GroupBy(t => t.Module, StringComparer.OrdinalIgnoreCase))
            {
                var moduleName = moduleGroup.Key;
                var moduleCode = GetModuleCode(vbProject, moduleName);
                hooksByModule[moduleName] = FindModuleHooks(moduleName, moduleCode);
            }

            // Assign sequential index
            for (var i = 0; i < selected.Count; i++)
            {
                selected[i].Index = i;
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
                BuildTestRunnerCode(selected, hooksByModule));
            ExcelBridgeSupport.ReleaseComObject(runnerCodeModule);
            ExcelBridgeSupport.ReleaseComObject(components);

            // Execute tests
            var results = new List<object>();
            var logs = new List<string>();
            var failed = 0;
            var inconclusiveCount = 0;

            foreach (var moduleGroup in selected.GroupBy(t => t.Module, StringComparer.OrdinalIgnoreCase))
            {
                var moduleName = moduleGroup.Key;
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
                            foreach (var test in moduleGroup)
                            {
                                failed++;
                                results.Add(new
                                {
                                    name = test.Name,
                                    module = test.Module,
                                    status = "failed",
                                    duration_ms = (int)sw.ElapsedMilliseconds,
                                    error = new { code = "before_all_failed", message, source = Convert.ToString(arr[2], CultureInfo.InvariantCulture) ?? "", number = Convert.ToInt32(arr[1], CultureInfo.InvariantCulture) },
                                });
                                logs.Add($"FAIL {test.Name}: before_all_failed: {message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sw.Stop();
                        beforeAllFailed = true;
                        foreach (var test in moduleGroup)
                        {
                            failed++;
                            results.Add(new
                            {
                                name = test.Name,
                                module = test.Module,
                                status = "failed",
                                duration_ms = (int)sw.ElapsedMilliseconds,
                                error = new { code = "before_all_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult },
                            });
                            logs.Add($"FAIL {test.Name}: before_all_failed: {ex.Message}");
                        }
                    }
                }

                if (!beforeAllFailed)
                {
                    foreach (var test in moduleGroup)
                    {
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

                            var status = "passed";
                            var errorCode = "";
                            var errorMessage = "";
                            var errorSource = "";
                            var errorNumber = 0;

                            if (!success)
                            {
                                status = "failed";
                                errorCode = "test_failed";
                                errorMessage = errDescription;
                                errorSource = errSource;
                                errorNumber = errNumber;
                                if (statusHint == "inconclusive")
                                {
                                    status = "inconclusive";
                                    errorCode = "test_inconclusive";
                                }
                                else
                                {
                                    errorCode = phaseHint switch
                                    {
                                        "before_each" => "before_each_failed",
                                        "after_each" => "after_each_failed",
                                        _ => errorCode,
                                    };
                                }
                            }

                            results.Add(new
                            {
                                name = test.Name,
                                module = test.Module,
                                status,
                                duration_ms = (int)sw.ElapsedMilliseconds,
                                tags = test.Tags,
                                error = new { code = errorCode, message = errorMessage, source = errorSource, number = errorNumber },
                            });

                            if (status == "passed")
                            {
                                logs.Add($"PASS {test.Name}");
                            }
                            else if (status == "inconclusive")
                            {
                                inconclusiveCount++;
                                logs.Add($"? {test.Name}: inconclusive");
                            }
                            else
                            {
                                failed++;
                                logs.Add($"FAIL {test.Name}: {errorMessage}");
                            }
                        }
                        catch (Exception ex)
                        {
                            sw.Stop();
                            failed++;
                            results.Add(new
                            {
                                name = test.Name,
                                module = test.Module,
                                status = "failed",
                                duration_ms = (int)sw.ElapsedMilliseconds,
                                tags = test.Tags,
                                error = new { code = "test_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult },
                            });
                            logs.Add($"FAIL {test.Name}: {ex.Message}");
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
                            var moduleTestNames = moduleGroup.Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                            for (var i = 0; i < results.Count; i++)
                            {
                                // Reconstruct from anonymous type via reflection
                                var resultObj = results[i];
                                var resultName = (string)resultObj.GetType().GetProperty("name")!.GetValue(resultObj)!;
                                var resultModule = (string)resultObj.GetType().GetProperty("module")!.GetValue(resultObj)!;
                                var resultStatus = (string)resultObj.GetType().GetProperty("status")!.GetValue(resultObj)!;
                                if (resultModule == moduleName && moduleTestNames.Contains(resultName))
                                {
                                    if (resultStatus is "passed" or "inconclusive")
                                    {
                                        if (resultStatus == "passed") failed++;
                                        if (resultStatus == "inconclusive") inconclusiveCount--;
                                    }
                                    var resultDuration = (int)resultObj.GetType().GetProperty("duration_ms")!.GetValue(resultObj)!;
                                    results[i] = new
                                    {
                                        name = resultName,
                                        module = resultModule,
                                        status = "failed",
                                        duration_ms = resultDuration,
                                        error = new { code = "after_all_failed", message, source = Convert.ToString(arr[2], CultureInfo.InvariantCulture) ?? "", number = Convert.ToInt32(arr[1], CultureInfo.InvariantCulture) },
                                    };
                                    logs.Add($"FAIL {resultName}: after_all_failed: {message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sw.Stop();
                        var moduleTestNames = moduleGroup.Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                        for (var i = 0; i < results.Count; i++)
                        {
                            var resultObj = results[i];
                            var resultName = (string)resultObj.GetType().GetProperty("name")!.GetValue(resultObj)!;
                            var resultModule = (string)resultObj.GetType().GetProperty("module")!.GetValue(resultObj)!;
                            var resultStatus = (string)resultObj.GetType().GetProperty("status")!.GetValue(resultObj)!;
                            if (resultModule == moduleName && moduleTestNames.Contains(resultName))
                            {
                                if (resultStatus is "passed" or "inconclusive")
                                {
                                    if (resultStatus == "passed") failed++;
                                    if (resultStatus == "inconclusive") inconclusiveCount--;
                                }
                                var resultDuration = (int)resultObj.GetType().GetProperty("duration_ms")!.GetValue(resultObj)!;
                                results[i] = new
                                {
                                    name = resultName,
                                    module = resultModule,
                                    status = "failed",
                                    duration_ms = resultDuration,
                                    error = new { code = "after_all_failed", message = ex.Message, source = ex.Source ?? "", number = ex.HResult },
                                };
                                logs.Add($"FAIL {resultName}: after_all_failed: {ex.Message}");
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

            // Save workbook
            ExcelBridgeSupport.RunPhase("save_workbook", () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));

            var sessionLog = GetSessionUsageLog(sessionMode);
            if (sessionLog is not null)
            {
                logs.Insert(0, sessionLog);
            }

            var extensions = new Dictionary<string, object?>
            {
                ["workbook"] = new
                {
                    path = args.WorkbookPath,
                    session = sessionAttached,
                    session_mode = sessionMode,
                    session_requested = args.UseSession,
                    auto_session = sessionAttached && !args.UseSession,
                    saved = true,
                    dirty = false,
                    needs_save = false,
                },
                ["tests"] = results.ToArray(),
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
                CloseWorkbook(workbook, excel);
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
                if (component is null) continue;
                var name = ExcelBridgeSupport.GetString(component, "Name");
                if (string.IsNullOrWhiteSpace(name)) continue;
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
            if (components is null) return "";
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components, "Count"));
            for (var i = 1; i <= count; i++)
            {
                component = ExcelBridgeSupport.Get(components, "Item", i);
                if (component is null) continue;
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
            if (codeModule is null) return "";
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
        if (string.IsNullOrEmpty(code)) return tests;

        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            var match = Regex.Match(line,
                @"^(?:Public\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(\s*\))?\s*(?:'.*)?$",
                RegexOptions.IgnoreCase);
            if (!match.Success) continue;

            var name = match.Groups[1].Value;
            if (!name.StartsWith("Test", StringComparison.Ordinal) && !name.EndsWith("_Test", StringComparison.Ordinal))
            {
                continue;
            }

            var tags = new List<string>();
            for (var j = i - 1; j >= 0; j--)
            {
                var prev = lines[j].Trim();
                if (string.IsNullOrWhiteSpace(prev)) continue;
                var tagMatch = Regex.Match(prev, @"^'\s*@Tag\s*\(""([^""]+)""\)", RegexOptions.IgnoreCase);
                if (tagMatch.Success)
                {
                    tags.Add(tagMatch.Groups[1].Value);
                    continue;
                }
                if (prev.StartsWith("''", StringComparison.Ordinal)) continue;
                break;
            }

            tests.Add(new TestCase
            {
                Name = name,
                Module = moduleName,
                Line = i + 1,
                Tags = tags.ToArray(),
            });
        }
        return tests;
    }

    internal static ModuleHooks FindModuleHooks(string moduleName, string code)
    {
        var hooks = new ModuleHooks();
        if (string.IsNullOrEmpty(code)) return hooks;

        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            var match = Regex.Match(line,
                @"^(?:Public\s+)?Sub\s+(BeforeAll|AfterAll|BeforeEach|AfterEach)\s*(?:\(\s*\))?\s*(?:'.*)?$",
                RegexOptions.IgnoreCase);
            if (!match.Success) continue;

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

    internal static List<TestCase> SelectTests(List<TestCase> tests, string filter, string moduleFilter, string tagFilter)
    {
        var selected = new List<TestCase>();
        foreach (var test in tests)
        {
            var include = true;
            if (!string.IsNullOrWhiteSpace(filter) && !string.Equals(test.Name, filter, StringComparison.OrdinalIgnoreCase))
            {
                include = false;
            }
            if (!string.IsNullOrWhiteSpace(moduleFilter) && !string.Equals(test.Module, moduleFilter, StringComparison.OrdinalIgnoreCase))
            {
                include = false;
            }
            if (!string.IsNullOrWhiteSpace(tagFilter))
            {
                var tagFound = test.Tags.Any(t => string.Equals(t, tagFilter, StringComparison.OrdinalIgnoreCase));
                if (!tagFound) include = false;
            }
            if (include) selected.Add(test);
        }
        return selected;
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
            if (!hooksByModule.TryGetValue(mod, out var hooks)) continue;

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
            builder.AppendLine(CultureInfo.InvariantCulture, $"        {test.Module}.{test.Name}");
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

    internal static BridgeResponse BuildErrorResponse(
        BridgeRequest request,
        TestCommandArguments args,
        string code,
        string message,
        bool sessionAttached,
        string sessionMode,
        List<TestCase> tests,
        RuntimeInjectionHelper.RuntimeInjectionState? runtimeState,
        bool runtimeInjected)
    {
        var logs = new List<string>();
        var sessionLog = GetSessionUsageLog(sessionMode);
        if (sessionLog is not null) logs.Add(sessionLog);

        var extensions = new Dictionary<string, object?>
        {
            ["workbook"] = new
            {
                path = args.WorkbookPath,
                session = sessionAttached,
                session_mode = sessionMode,
                session_requested = args.UseSession,
                auto_session = sessionAttached && !args.UseSession,
            },
            ["tests"] = tests.Select(t => new { name = t.Name, module = t.Module, line = t.Line, tags = t.Tags }).ToArray(),
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
                Source: "xlflow"),
            Logs = logs,
            Extensions = extensions,
        };
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
        public int Line { get; init; }
        public string[] Tags { get; init; } = [];
        public int Index { get; set; }
    }

    internal sealed class ModuleHooks
    {
        public HookInfo? BeforeAll { get; set; }
        public HookInfo? AfterAll { get; set; }
        public HookInfo? BeforeEach { get; set; }
        public HookInfo? AfterEach { get; set; }
    }

    internal sealed record HookInfo(string Name, string Module, int Line);
}
