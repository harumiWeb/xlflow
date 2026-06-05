using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Text.Json;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Workers;

public sealed record MacroRunWorkerRequest(
    int ExcelProcessId,
    long ExcelHwnd,
    string MacroReference,
    IReadOnlyList<MacroRunWorkerArgument>? Arguments = null,
    string Operation = "run",
    string WorkbookPath = "");

public sealed record MacroRunWorkerArgument(string Type, string Value);

public sealed record MacroRunWorkerError(string Message, string Source, int Number);

public sealed record MacroRunWorkerResult(
    bool Completed,
    bool Ok,
    object? Value,
    MacroRunWorkerError? Error);

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "The worker must normalize COM failures into JSON.")]
public static class MacroRunWorker
{
    internal const string RequestPathEnv = "XLFLOW_WORKER_REQUEST_PATH";
    internal const string ResultPathEnv = "XLFLOW_WORKER_RESULT_PATH";

    public static int Run()
    {
        var requestPath = Environment.GetEnvironmentVariable(RequestPathEnv);
        var resultPath = Environment.GetEnvironmentVariable(ResultPathEnv);
        if (string.IsNullOrWhiteSpace(requestPath) || string.IsNullOrWhiteSpace(resultPath))
        {
            return 2;
        }

        var request = ReadRequest(requestPath);
        if (request is null ||
            request.ExcelProcessId <= 0 ||
            (string.Equals(request.Operation, "run", StringComparison.OrdinalIgnoreCase) &&
             string.IsNullOrWhiteSpace(request.MacroReference)))
        {
            WriteResult(resultPath, new MacroRunWorkerResult(
                Completed: true,
                Ok: false,
                Value: null,
                Error: new MacroRunWorkerError("Invalid macro worker request.", "xlflow-excel-bridge", 0)));
            return 2;
        }

        object? excel = null;
        try
        {
            excel = request.ExcelHwnd != 0
                ? ExcelBridgeSupport.TryGetExcelByHwnd(request.ExcelHwnd)
                : null;
            if (excel is null)
            {
                excel = ExcelBridgeSupport.TryGetExcelByProcessId(request.ExcelProcessId);
            }
            if (excel is null)
            {
                throw new InvalidOperationException("xlflow could not reconnect to the target Excel instance.");
            }

            object? value;
            if (string.Equals(request.Operation, "compile", StringComparison.OrdinalIgnoreCase))
            {
                value = CompileWorkbook(excel, request.WorkbookPath);
            }
            else
            {
                var invokeArgs = new List<object?> { request.MacroReference };
                foreach (var argument in request.Arguments ?? [])
                {
                    invokeArgs.Add(ConvertArgument(argument));
                }
                value = ExcelBridgeSupport.InvokeMethod(excel, "Run", invokeArgs.ToArray());
            }
            WriteResult(resultPath, new MacroRunWorkerResult(true, true, NormalizeValue(value), null));
            return 0;
        }
        catch (Exception ex)
        {
            var detail = ex.InnerException ?? ex;
            WriteResult(resultPath, new MacroRunWorkerResult(
                Completed: true,
                Ok: false,
                Value: null,
                Error: new MacroRunWorkerError(
                    ExcelBridgeSupport.FormatExceptionDetail(ex),
                    detail.Source ?? "xlflow-excel-bridge",
                    detail.HResult)));
            return 1;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static MacroRunWorkerRequest? ReadRequest(string path)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<MacroRunWorkerRequest>(json, JsonOptions.Default);
    }

    private static void WriteResult(string path, MacroRunWorkerResult result)
    {
        var json = JsonSerializer.Serialize(result, JsonOptions.Default);
        File.WriteAllText(path, json);
    }

    private static object ConvertArgument(MacroRunWorkerArgument argument)
    {
        return argument.Type.ToLowerInvariant() switch
        {
            "int" => int.Parse(argument.Value, NumberStyles.Integer, CultureInfo.InvariantCulture),
            "double" => double.Parse(argument.Value, NumberStyles.Float, CultureInfo.InvariantCulture),
            "bool" => bool.Parse(argument.Value),
            _ => argument.Value,
        };
    }

    private static object CompileWorkbook(object excel, string workbookPath)
    {
        object? workbook = null;
        object? vbProject = null;
        object? vbe = null;
        object? commandBars = null;
        object? control = null;
        var stage = "find_workbook";
        try
        {
            workbook = FindWorkbook(excel, workbookPath)
                ?? throw new InvalidOperationException("xlflow could not find the target workbook for VBE Compile.");
            stage = "get_vbproject";
            dynamic workbookObject = workbook;
            vbProject = workbookObject.VBProject;
            stage = "get_vbe";
            dynamic vbProjectObject = vbProject;
            vbe = vbProjectObject.VBE;
            stage = "get_command_bars";
            dynamic vbeObject = vbe;
            commandBars = vbeObject.CommandBars;
            stage = "find_compile_control";
            control = FindCompileControl(commandBars)
                ?? throw new InvalidOperationException("VBE Compile command was not found.");
            stage = "read_compile_enabled";
            dynamic controlObject = control;
            var enabled = Convert.ToBoolean(controlObject.Enabled, CultureInfo.InvariantCulture);
            if (!enabled)
            {
                return new { compiled = false, reason = "compile_command_disabled" };
            }
            stage = "execute_compile";
            controlObject.Execute();
            return new { compiled = true };
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"VBE Compile {stage} failed: {ex.Message}", ex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(control);
            ExcelBridgeSupport.ReleaseComObject(commandBars);
            ExcelBridgeSupport.ReleaseComObject(vbe);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }
    }

    private static object? FindWorkbook(object excel, string workbookPath)
    {
        if (!string.IsNullOrWhiteSpace(workbookPath))
        {
            return ExcelBridgeSupport.GetOpenWorkbook(excel, workbookPath);
        }
        object? workbooks = null;
        try
        {
            dynamic excelObject = excel;
            workbooks = excelObject.Workbooks;
            dynamic workbooksObject = workbooks;
            return workbooksObject.Item(1);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbooks);
        }
    }

    private static object? FindCompileControl(object commandBars)
    {
        const int compileControlId = 578;
        try
        {
            dynamic commandBarsObject = commandBars;
            var byId = commandBarsObject.FindControl(Type.Missing, compileControlId);
            if (byId is not null)
            {
                return byId;
            }
        }
        catch
        {
            // Continue with localized menu lookup.
        }

        foreach (var barName in new[] { "Debug", "デバッグ" })
        {
            object? bar = null;
            object? controls = null;
            try
            {
                dynamic commandBarsObject = commandBars;
                bar = commandBarsObject.Item(barName);
                dynamic barObject = bar;
                controls = barObject.Controls;
                if (controls is null)
                {
                    continue;
                }
                dynamic controlsObject = controls;
                var count = Convert.ToInt32(controlsObject.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    var candidate = controlsObject.Item(index);
                    if (candidate is null)
                    {
                        continue;
                    }
                    dynamic candidateObject = candidate;
                    var id = Convert.ToInt32(candidateObject.Id, CultureInfo.InvariantCulture);
                    var caption = (Convert.ToString(candidateObject.Caption, CultureInfo.InvariantCulture) ?? "")
                        .Replace("&", "", StringComparison.Ordinal);
                    if (id == compileControlId ||
                        caption.Contains("compile", StringComparison.OrdinalIgnoreCase) ||
                        caption.Contains("コンパイル", StringComparison.Ordinal))
                    {
                        return candidate;
                    }
                    ExcelBridgeSupport.ReleaseComObject(candidate);
                }
            }
            catch
            {
                // Try the next localized command bar.
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(controls);
                ExcelBridgeSupport.ReleaseComObject(bar);
            }
        }
        return null;
    }

    private static object? NormalizeValue(object? value)
    {
        if (value is Array array)
        {
            var items = new object?[array.Length];
            for (var i = 0; i < array.Length; i++)
            {
                items[i] = array.GetValue(i);
            }
            return items;
        }
        return value;
    }
}

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Worker process state and cleanup are best-effort lifecycle operations.")]
public sealed class MacroRunWorkerProcess : IDisposable
{
    private readonly string _requestPath;
    private readonly string _resultPath;
    private readonly Process _process;

    private MacroRunWorkerProcess(Process process, string requestPath, string resultPath)
    {
        _process = process;
        _requestPath = requestPath;
        _resultPath = resultPath;
    }

    public int ProcessId => _process.Id;

    public bool HasExited
    {
        get
        {
            try
            {
                return _process.HasExited;
            }
            catch
            {
                return true;
            }
        }
    }

    public static MacroRunWorkerProcess Start(MacroRunWorkerRequest request)
    {
        var requestPath = Path.Combine(Path.GetTempPath(), "xlflow-worker-request-" + Guid.NewGuid().ToString("N") + ".json");
        var resultPath = Path.Combine(Path.GetTempPath(), "xlflow-worker-result-" + Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(requestPath, JsonSerializer.Serialize(request, JsonOptions.Default));

        var startInfo = CreateStartInfo(requestPath, resultPath);
        var process = Process.Start(startInfo) ?? throw new InvalidOperationException("xlflow could not start the .NET macro worker process.");
        return new MacroRunWorkerProcess(process, requestPath, resultPath);
    }

    public MacroRunWorkerResult? WaitForResult(TimeSpan timeout)
    {
        if (!_process.WaitForExit((int)Math.Clamp(timeout.TotalMilliseconds, 1, int.MaxValue)))
        {
            return null;
        }

        if (!File.Exists(_resultPath))
        {
            return null;
        }

        var json = File.ReadAllText(_resultPath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return null;
        }

        return JsonSerializer.Deserialize<MacroRunWorkerResult>(json, JsonOptions.Default);
    }

    public void Stop()
    {
        try
        {
            if (!_process.HasExited)
            {
                _process.Kill(entireProcessTree: false);
            }
        }
        catch
        {
            // best-effort worker cleanup
        }
    }

    public void Dispose()
    {
        Stop();
        _process.Dispose();
        TryDelete(_requestPath);
        TryDelete(_resultPath);
    }

    internal static ProcessStartInfo CreateStartInfo(string requestPath, string resultPath)
    {
        var processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve the bridge executable path.");
        var startInfo = new ProcessStartInfo
        {
            FileName = processPath,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase))
        {
            var assemblyPath = Path.Combine(AppContext.BaseDirectory, $"{Assembly.GetExecutingAssembly().GetName().Name}.dll");
            if (!File.Exists(assemblyPath))
            {
                throw new InvalidOperationException("Unable to resolve the bridge assembly path for dotnet worker startup.");
            }
            startInfo.ArgumentList.Add(assemblyPath);
        }
        startInfo.ArgumentList.Add("--run-worker");
        startInfo.Environment[MacroRunWorker.RequestPathEnv] = requestPath;
        startInfo.Environment[MacroRunWorker.ResultPathEnv] = resultPath;
        return startInfo;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // best-effort temp file cleanup
        }
    }
}
