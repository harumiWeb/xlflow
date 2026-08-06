using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Services;

/// <summary>
/// Executes one isolated, compile-only VBE oracle case.  This service is
/// deliberately not used by any of the production xlflow commands.
/// </summary>
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The oracle is a Windows/Excel-only developer tool.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "COM, worker, and UI failures are returned as oracle infrastructure failures.")]
public sealed class VbeOracleService : IVbeOracleService
{
    private static readonly JsonSerializerOptions PlanJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private static readonly Regex AttributeNamePattern = new(
        "^\\s*Attribute\\s+VB_Name\\s*=\\s*\\\"([^\\\"]+)\\\"",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    public BridgeResponse Execute(BridgeRequest request, VbeOracleCommandArguments args, CancellationToken cancellationToken)
    {
        OraclePlan plan;
        try
        {
            plan = DecodePlan(args.PlanJson64);
            ValidatePlan(plan);
        }
        catch (VbeOracleValidationException ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                "vbe_oracle_plan_invalid",
                ex.Message,
                "vbe-oracle.validate",
                "xlflow-excel-bridge"));
        }
        catch (Exception ex) when (ex is FormatException or JsonException or DecoderFallbackException)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                "vbe_oracle_plan_invalid",
                ex.Message,
                "vbe-oracle.validate",
                "xlflow-excel-bridge"));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException or ArgumentException)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                "vbe_oracle_plan_invalid",
                ex.Message,
                "vbe-oracle.validate",
                "xlflow-excel-bridge"));
        }

        if (!OperatingSystem.IsWindows())
        {
            return InfrastructureResponse(
                request,
                plan,
                "oracle_windows_only",
                "The VBE oracle requires Windows with a locally installed Excel instance.",
                "platform",
                0,
                TimeSpan.Zero,
                new Dictionary<string, object?>());
        }

        var stopwatch = Stopwatch.StartNew();
        object? excel = null;
        object? workbook = null;
        var excelProcessId = 0;
        var excelHwnd = 0L;
        var ownedExcelProcess = OwnedExcelProcess.None;
        var tempDirectory = Path.Combine(Path.GetTempPath(), "xlflow-vbe-oracle-" + Guid.NewGuid().ToString("N"));
        var stage = "startup";
        var cleanupConfirmed = true;
        var tempCleanupConfirmed = true;
        IReadOnlyList<OwnedExcelProcess> baselineExcelProcesses = Array.Empty<OwnedExcelProcess>();
        var ownedExcelProcesses = new List<OwnedExcelProcess>();
        OracleObservation? observation = null;
        Dictionary<string, object?> excelMetadata = new();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            baselineExcelProcesses = ExcelBridgeSupport.CaptureOwnedExcelProcesses();
            Directory.CreateDirectory(tempDirectory);
            stage = "create_excel";
            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Excel.Application COM class is not registered");
            excel = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Failed to create Excel.Application COM instance");
            dynamic app = excel;
            app.Visible = plan.Visible;
            app.DisplayAlerts = false;
            app.EnableEvents = false;

            stage = "create_workbook";
            object? workbooks = null;
            try
            {
                workbooks = ExcelBridgeSupport.Get(excel, "Workbooks")
                    ?? throw new InvalidOperationException("Excel Workbooks collection is unavailable");
                dynamic workbooksObject = workbooks;
                workbook = workbooksObject.Add()
                    ?? throw new InvalidOperationException("Excel could not create a disposable workbook");
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(workbooks);
            }

            excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            excelHwnd = ExcelBridgeSupport.GetExcelMainHwnd(excel);
            if (excelProcessId <= 0 || excelHwnd == 0)
            {
                throw new InvalidOperationException("Could not resolve the dedicated Excel process identity");
            }
            ownedExcelProcess = ExcelBridgeSupport.CaptureOwnedExcelProcess(excelProcessId);
            if (ownedExcelProcess.ProcessId <= 0)
            {
                throw new InvalidOperationException("Could not capture the dedicated Excel process identity");
            }
            if (baselineExcelProcesses.Any(process => SameProcess(process, ownedExcelProcess)))
            {
                throw new InvalidOperationException("Excel.Application did not create a dedicated process instance");
            }
            ownedExcelProcesses.Add(ownedExcelProcess);
            foreach (var process in ExcelBridgeSupport.CaptureOwnedExcelProcesses())
            {
                if (!baselineExcelProcesses.Any(existing => SameProcess(existing, process)) &&
                    !ownedExcelProcesses.Any(existing => SameProcess(existing, process)))
                {
                    ownedExcelProcesses.Add(process);
                }
            }

            excelMetadata = ReadExcelMetadata(excel, excelProcessId);
            stage = "import_modules";
            ImportModules(workbook, plan.Modules ?? [], tempDirectory, cancellationToken);

            stage = "vbe_compile";
            var selectionLocator = new VbeSelectionLocator(
                excelProcessId,
                excelHwnd,
                new VbeSourceMappingOptions("", "", "", "", "sidecar", false, "ignore", false));
            var invocation = ExcelWorkerInvocation.InvokeWithWorker(
                new MacroRunWorkerRequest(
                    excelProcessId,
                    excelHwnd,
                    "",
                    Operation: "compile",
                    WorkbookPath: ""),
                excelHwnd,
                DialogKind.Compile,
                suppressModalErrors: true,
                args.Timeout,
                cancellationToken,
                selectionLocator);
            observation = ClassifyInvocation(invocation, stopwatch.Elapsed, excelMetadata);
            if (observation.Outcome == "accepted")
            {
                var lingeringDialogs = new DialogWatcher().CaptureOracleDialogs(
                    new DialogWatchRequest(
                        excelProcessId,
                        excelHwnd,
                        DialogKind.Any,
                        DialogActionPolicy.ObserveOnly,
                        TimeSpan.Zero,
                        TimeSpan.Zero));
                var lingeringDialog = lingeringDialogs.Count == 0 ? null : lingeringDialogs[0];
                if (lingeringDialog is not null)
                {
                    observation = InfrastructureObservation(
                        "oracle_unknown_modal",
                        "A modal dialog remained after VBE Compile completed.",
                        "vbe_compile",
                        stopwatch.Elapsed,
                        excelMetadata,
                        lingeringDialog);
                }
            }
        }
        catch (OperationCanceledException ex)
        {
            observation = InfrastructureObservation("oracle_cancelled", ex.Message, stage, stopwatch.Elapsed, excelMetadata);
        }
        catch (Exception ex)
        {
            observation = InfrastructureObservation(
                ClassifyInfrastructureCode(stage, ex),
                ExcelBridgeSupport.FormatExceptionDetail(ex),
                stage,
                stopwatch.Elapsed,
                excelMetadata);
        }
        finally
        {
            if (excel is not null || workbook is not null)
            {
                foreach (var process in ExcelBridgeSupport.CaptureOwnedExcelProcesses())
                {
                    if (!baselineExcelProcesses.Any(existing => SameProcess(existing, process)) &&
                        !ownedExcelProcesses.Any(existing => SameProcess(existing, process)))
                    {
                        ownedExcelProcesses.Add(process);
                    }
                }
                cleanupConfirmed = ExcelBridgeSupport.ReleaseOracleExcelAndConfirmExit(
                    workbook,
                    excel,
                    ownedExcelProcesses,
                    baselineExcelProcesses);
                workbook = null;
                excel = null;
            }

            try
            {
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, recursive: true);
                }
            }
            catch
            {
                tempCleanupConfirmed = false;
            }
        }

        observation ??= InfrastructureObservation(
            "oracle_worker_output_invalid",
            "The VBE oracle did not produce an observation.",
            stage,
            stopwatch.Elapsed,
            excelMetadata);

        // A process that did not exit cleanly is never VBA evidence.  This
        // deliberately overrides an otherwise accepted or rejected result.
        if (!cleanupConfirmed || !tempCleanupConfirmed)
        {
            observation = InfrastructureObservation(
                "oracle_cleanup_unconfirmed",
                !cleanupConfirmed
                    ? "The dedicated Excel process did not exit cleanly."
                    : "The oracle temporary directory could not be removed.",
                "excel_cleanup",
                stopwatch.Elapsed,
                excelMetadata);
        }

        return BuildResponse(request, plan, observation, excelProcessId, cleanupConfirmed && tempCleanupConfirmed);
    }

    private static OraclePlan DecodePlan(string planJson64)
    {
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(planJson64);
        }
        catch (FormatException ex)
        {
            throw new VbeOracleValidationException("PlanJson64 is not valid base64", ex);
        }

        var plan = JsonSerializer.Deserialize<OraclePlan>(Encoding.UTF8.GetString(bytes), PlanJsonOptions);
        if (plan is null)
        {
            throw new VbeOracleValidationException("FixtureJSON64 contains an empty oracle fixture");
        }

        if (plan.SchemaVersion != 1)
        {
            throw new VbeOracleValidationException($"unsupported oracle fixture schema_version {plan.SchemaVersion}");
        }

        // The fixture contract uses {id, probe:{mode}, modules:[...]}; the
        // original bridge prototype used {case_id, probe_mode}.  Normalize
        // both spellings at the bridge boundary so the execution path stays
        // deliberately small.
        plan.CaseId = string.IsNullOrWhiteSpace(plan.CaseId) ? plan.Id : plan.CaseId;
        plan.ProbeMode = string.IsNullOrWhiteSpace(plan.ProbeMode) ? plan.Probe?.Mode ?? "" : plan.ProbeMode;
        plan.Modules ??= [];
        foreach (var module in plan.Modules)
        {
            if (string.IsNullOrWhiteSpace(module.SourcePath))
            {
                module.SourcePath = module.Path;
            }
            if (!string.IsNullOrWhiteSpace(module.SourcePath))
            {
                module.SourcePath = Path.GetFullPath(module.SourcePath);
            }
        }
        return plan;
    }

    private static void ValidatePlan(OraclePlan plan)
    {
        if (string.IsNullOrWhiteSpace(plan.CaseId))
        {
            throw new VbeOracleValidationException("oracle plan case_id is required");
        }
        if (!string.Equals(plan.ProbeMode, "compile", StringComparison.OrdinalIgnoreCase))
        {
            throw new VbeOracleValidationException("only probe_mode=compile is supported by oracle schema v1");
        }
        if (plan.Modules is null || plan.Modules.Count == 0)
        {
            throw new VbeOracleValidationException("oracle plan must contain at least one module");
        }
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var module in plan.Modules)
        {
            if (module is null)
            {
                throw new VbeOracleValidationException("oracle plan contains a null module entry");
            }
            if (!string.Equals(module.Kind, "standard", StringComparison.OrdinalIgnoreCase))
            {
                throw new VbeOracleValidationException($"module '{module.Name}' must have kind=standard");
            }
            if (string.IsNullOrWhiteSpace(module.Name) || !names.Add(module.Name))
            {
                throw new VbeOracleValidationException($"module name '{module.Name}' is missing or duplicated");
            }
            if (string.IsNullOrWhiteSpace(module.Source) &&
                (string.IsNullOrWhiteSpace(module.SourcePath) ||
                 !string.Equals(Path.GetExtension(module.SourcePath), ".bas", StringComparison.OrdinalIgnoreCase)))
            {
                throw new VbeOracleValidationException($"module '{module.Name}' must provide a .bas source path or inline source");
            }

            var source = module.Source;
            if (string.IsNullOrWhiteSpace(source))
            {
                var sourcePath = Path.GetFullPath(module.SourcePath);
                if (!File.Exists(sourcePath))
                {
                    throw new VbeOracleValidationException($"module '{module.Name}' source does not exist: {module.SourcePath}");
                }
                source = File.ReadAllText(sourcePath, Encoding.UTF8);
            }
            var match = AttributeNamePattern.Match(source);
            if (!match.Success || !string.Equals(match.Groups[1].Value.Trim(), module.Name, StringComparison.OrdinalIgnoreCase))
            {
                throw new VbeOracleValidationException($"module '{module.Name}' source Attribute VB_Name does not match");
            }
        }
    }

    private static void ImportModules(object workbook, IReadOnlyList<OracleModule> modules, string tempDirectory, CancellationToken cancellationToken)
    {
        object? project = null;
        object? components = null;
        try
        {
            project = ExcelBridgeSupport.Get(workbook, "VBProject")
                ?? throw new InvalidOperationException("VBProject access is denied; enable trusted access to the VBA project object model");
            components = ExcelBridgeSupport.Get(project, "VBComponents")
                ?? throw new InvalidOperationException("VBComponents are unavailable");

            foreach (var module in modules)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var importPath = Path.Combine(tempDirectory, "imports", module.Name + ".bas");
                var sourcePath = module.SourcePath;
                if (!string.IsNullOrWhiteSpace(module.Source))
                {
                    sourcePath = Path.Combine(tempDirectory, "sources", module.Name + ".bas");
                    var sourceParent = Path.GetDirectoryName(sourcePath);
                    if (!string.IsNullOrWhiteSpace(sourceParent))
                    {
                        Directory.CreateDirectory(sourceParent);
                    }
                    File.WriteAllText(sourcePath, module.Source, new UTF8Encoding(false));
                }
                VbaSourceHelper.PrepareSourceForImport(sourcePath, importPath, null, "off");
                object? imported = null;
                try
                {
                    imported = ExcelBridgeSupport.InvokeViaDynamic(components, "Import", importPath)
                        ?? throw new InvalidOperationException($"failed to import VBA module '{module.Name}'");
                    var importedName = ExcelBridgeSupport.GetString(imported, "Name") ?? "";
                    if (!string.Equals(importedName, module.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"imported VBA module name '{importedName}' did not match expected name '{module.Name}'");
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(imported);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
            ExcelBridgeSupport.ReleaseComObject(project);
        }
    }

    private static OracleObservation ClassifyInvocation(WorkerInvocationResult invocation, TimeSpan elapsed, IReadOnlyDictionary<string, object?> excel)
    {
        if (invocation.TimedOut)
        {
            return InfrastructureObservation("oracle_timeout", "VBE Compile timed out.", "vbe_compile", elapsed, excel);
        }

        if (invocation.Dialog is not null)
        {
            if (string.Equals(invocation.Dialog.Kind, "compile", StringComparison.OrdinalIgnoreCase))
            {
                var rejected = new OracleObservation(
                    "rejected",
                    "compile",
                    "vbe_compile",
                    elapsed,
                    true,
                    invocation.Dialog,
                    excel,
                    null);
                rejected.Location = invocation.LocationCapture.Location;
                return rejected;
            }
            var blocked = InfrastructureObservation("oracle_unknown_modal", "A non-compile modal dialog blocked the VBE oracle.", "vbe_compile", elapsed, excel, invocation.Dialog);
            blocked.Location = invocation.LocationCapture.Location;
            return blocked;
        }

        if (invocation.Result is null)
        {
            return InfrastructureObservation("oracle_worker_output_invalid", "The VBE worker exited without a result.", "vbe_compile", elapsed, excel);
        }
        if (!invocation.Result.Completed || !invocation.Result.Ok)
        {
            var message = invocation.Result.Error?.Message ?? "The VBE worker failed without a structured error.";
            return InfrastructureObservation("oracle_worker_failed", message, invocation.Result.Error?.Stage ?? "vbe_compile", elapsed, excel);
        }

        var disabled = TryGetString(invocation.Result.Value, "reason");
        var compileInvoked = !string.Equals(disabled, "compile_command_disabled", StringComparison.OrdinalIgnoreCase);
        var accepted = new OracleObservation(
            "accepted",
            "compile",
            "vbe_compile",
            elapsed,
            compileInvoked,
            null,
            excel,
            null);
        accepted.Location = invocation.LocationCapture.Location;
        return accepted;
    }

    private static Dictionary<string, object?> ReadExcelMetadata(object excel, int excelProcessId)
    {
        var version = ExcelBridgeSupport.GetString(excel, "Version");
        var build = ExcelBridgeSupport.GetString(excel, "Build");
        var locale = CultureInfo.CurrentUICulture.Name;
        object? languageSettings = null;
        try
        {
            languageSettings = ExcelBridgeSupport.Get(excel, "LanguageSettings");
            if (languageSettings is not null)
            {
                dynamic settings = languageSettings;
                var lcid = Convert.ToInt32(settings.LanguageID(2), CultureInfo.InvariantCulture);
                if (lcid > 0)
                {
                    locale = CultureInfo.GetCultureInfo(lcid).Name;
                }
            }
        }
        catch
        {
            // Excel language settings are not exposed by every installation;
            // retain the bridge UI locale as a diagnostic fallback.
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(languageSettings);
        }
        return new Dictionary<string, object?>
        {
            ["version"] = string.IsNullOrWhiteSpace(version) ? null : version,
            ["build"] = string.IsNullOrWhiteSpace(build) ? null : build,
            ["excel_version"] = string.IsNullOrWhiteSpace(version) ? null : version,
            ["excel_build"] = string.IsNullOrWhiteSpace(build) ? null : build,
            ["bitness"] = ExcelBridgeSupport.GetProcessBitness(excelProcessId),
            ["locale"] = locale,
        };
    }

    private static BridgeResponse BuildResponse(BridgeRequest request, OraclePlan plan, OracleObservation observation, int processId, bool cleanupConfirmed)
    {
        var oracle = new Dictionary<string, object?>
        {
            ["schema_version"] = 1,
            ["case_id"] = plan.CaseId,
            ["outcome"] = observation.Outcome,
            ["evidence_phase"] = observation.EvidencePhase,
            ["last_stage"] = observation.LastStage,
            ["duration_ms"] = Math.Max(0L, (long)observation.Duration.TotalMilliseconds),
            ["compile_invoked"] = observation.CompileInvoked,
            ["cleanup_confirmed"] = cleanupConfirmed,
            ["excel_process_id"] = processId > 0 ? processId : null,
            ["excel"] = observation.Excel,
            ["metadata"] = new Dictionary<string, object?>
            {
                ["excel_version"] = observation.Excel.TryGetValue("version", out var version) ? version : null,
                ["excel_build"] = observation.Excel.TryGetValue("build", out var build) ? build : null,
                ["bitness"] = observation.Excel.TryGetValue("bitness", out var bitness) ? bitness : null,
                ["locale"] = observation.Excel.TryGetValue("locale", out var locale) ? locale : null,
            },
        };
        if (observation.Dialog is not null)
        {
            oracle["dialog"] = observation.Dialog;
        }
        if (observation.Location is not null)
        {
            oracle["location"] = observation.Location;
        }
        if (!string.IsNullOrWhiteSpace(observation.ErrorCode))
        {
            oracle["error_code"] = observation.ErrorCode;
            oracle["error"] = observation.ErrorMessage ?? observation.ErrorCode;
        }

        if (string.Equals(observation.Outcome, "infrastructure_failure", StringComparison.OrdinalIgnoreCase))
        {
            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Status = BridgeStatus.Failed,
                Error = new BridgeError(
                    observation.ErrorCode ?? "oracle_infrastructure_failure",
                    observation.ErrorMessage ?? "VBE oracle infrastructure failure.",
                    observation.LastStage,
                    "xlflow-excel-bridge",
                    Details: new Dictionary<string, object?> { ["oracle"] = oracle }),
                Extensions = new Dictionary<string, object?> { ["oracle"] = oracle },
            };
        }

        return BridgeResponse.Ok(request, new Dictionary<string, object?> { ["oracle"] = oracle });
    }

    private static BridgeResponse InfrastructureResponse(BridgeRequest request, OraclePlan plan, string code, string message, string stage, int processId, TimeSpan elapsed, Dictionary<string, object?> excel)
    {
        var observation = InfrastructureObservation(code, message, stage, elapsed, excel);
        return BuildResponse(request, plan, observation, processId, cleanupConfirmed: false);
    }

    private static OracleObservation InfrastructureObservation(string code, string message, string stage, TimeSpan elapsed, IReadOnlyDictionary<string, object?> excel, DialogSnapshot? dialog = null)
    {
        return new OracleObservation("infrastructure_failure", "unknown", stage, elapsed, false, dialog, excel, code, message);
    }

    private static string ClassifyInfrastructureCode(string stage, Exception exception)
    {
        if (exception.Message.Contains("VBProject", StringComparison.OrdinalIgnoreCase))
        {
            return "oracle_vbproject_access_denied";
        }
        if (stage == "create_excel" || stage == "create_workbook")
        {
            return "oracle_excel_startup_failed";
        }
        if (stage == "import_modules")
        {
            return "oracle_import_failed";
        }
        if (stage == "vbe_compile")
        {
            return "oracle_compile_invocation_failed";
        }
        return "oracle_infrastructure_failure";
    }

    private static string? TryGetString(object? value, string property)
    {
        if (value is JsonElement element && element.ValueKind == JsonValueKind.Object && element.TryGetProperty(property, out var propertyValue))
        {
            return propertyValue.ValueKind == JsonValueKind.String ? propertyValue.GetString() : propertyValue.ToString();
        }
        return null;
    }

    private static bool SameProcess(OwnedExcelProcess left, OwnedExcelProcess right)
    {
        return left.ProcessId == right.ProcessId &&
               (left.StartTime is null || right.StartTime is null || left.StartTime.Value == right.StartTime.Value);
    }

    private sealed class VbeOracleValidationException(string message, Exception? inner = null) : Exception(message, inner);

    private sealed record OracleObservation(
        string Outcome,
        string EvidencePhase,
        string LastStage,
        TimeSpan Duration,
        bool CompileInvoked,
        DialogSnapshot? Dialog,
        IReadOnlyDictionary<string, object?> Excel,
        string? ErrorCode,
        string? ErrorMessage = null)
    {
        public object? Location { get; set; }
    }

    private sealed class OraclePlan
    {
        [JsonPropertyName("schema_version")]
        public int SchemaVersion { get; init; }

        [JsonPropertyName("case_id")]
        public string CaseId { get; set; } = "";

        [JsonPropertyName("id")]
        public string Id { get; init; } = "";

        [JsonPropertyName("probe_mode")]
        public string ProbeMode { get; set; } = "";

        [JsonPropertyName("probe")]
        public OracleProbe? Probe { get; init; }

        [JsonPropertyName("visible")]
        public bool Visible { get; init; }

        [JsonPropertyName("timeout_ms")]
        public int TimeoutMilliseconds { get; init; }

        [JsonPropertyName("modules")]
        public List<OracleModule>? Modules { get; set; } = [];
    }

    private sealed class OracleProbe
    {
        [JsonPropertyName("mode")]
        public string Mode { get; init; } = "";
    }

    private sealed class OracleModule
    {
        [JsonPropertyName("name")]
        public string Name { get; init; } = "";

        [JsonPropertyName("kind")]
        public string Kind { get; init; } = "";

        [JsonPropertyName("source_path")]
        public string SourcePath { get; set; } = "";

        [JsonPropertyName("path")]
        public string Path { get; init; } = "";

        [JsonPropertyName("source")]
        public string Source { get; init; } = "";

    }
}
