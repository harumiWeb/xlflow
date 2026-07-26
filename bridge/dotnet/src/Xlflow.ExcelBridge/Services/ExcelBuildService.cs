using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "The bridge converts Excel and filesystem failures into structured build errors.")]
public sealed class ExcelBuildService : IBuildService
{
    private const int ComponentTypeDocument = 100;
    private const int ComponentTypeForm = 3;
    private static readonly JsonSerializerOptions PlanJsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly BuildArtifactPublisher _artifactPublisher;

    public ExcelBuildService() : this(new BuildArtifactPublisher()) { }

    internal ExcelBuildService(BuildArtifactPublisher artifactPublisher)
    {
        _artifactPublisher = artifactPublisher;
    }

    public BridgeResponse Execute(BridgeRequest request, BuildCommandArguments args, CancellationToken cancellationToken)
    {
        object? excel = null;
        object? workbook = null;
        string? stagingDirectory = null;
        BuildArtifactCleanup? stagingCleanup = null;
        var stage = "validate";
        var excelProcessId = 0;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var plan = DecodePlan(args.PlanJson64);
            ValidatePlan(plan, args.ProjectRoot);
            stage = "validate_paths";
            var basePath = ExcelBridgeSupport.NormalizePath(args.BaseWorkbookPath);
            var outputPath = ExcelBridgeSupport.NormalizePath(args.OutputWorkbookPath);
            if (!File.Exists(basePath))
            {
                throw new InvalidOperationException("base workbook does not exist");
            }
            if (string.Equals(basePath, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("base workbook and output workbook must be different files");
            }
            if (HasDirtyMatchingSession(args))
            {
                return BridgeResponse.Failed(request, new BridgeError("build_session_dirty", "A live xlflow session has unsaved changes. Run `xlflow save --session` before building.", "session", "xlflow-excel-bridge"));
            }
            if (HasMatchingSession(args, outputPath))
            {
                return BridgeResponse.Failed(request, new BridgeError("build_output_busy", "The output workbook is owned by a live xlflow session.", "output_busy", "xlflow-excel-bridge"));
            }

            // Stage beside the final output so the final replacement is a
            // same-volume atomic move. Never touch the previous artifact until
            // Excel has saved, closed, and exited cleanly.
            stage = "stage_prepare";
            var artifactStage = BuildArtifactPublisher.CreateStage(outputPath);
            stagingDirectory = artifactStage.Directory;
            var temporaryWorkbook = artifactStage.TemporaryWorkbookPath;

            File.Copy(basePath, temporaryWorkbook);

            stage = "open_workbook";
            var attachment = ExcelBridgeSupport.OpenWorkbookDirect(temporaryWorkbook, args.Visible, disableAutomationMacros: true);
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            stage = "source_applied";
            var applied = Reconstruct(workbook, plan, args.CodeSource, stagingDirectory, cancellationToken);
            var excelHwnd = ExcelBridgeSupport.GetExcelMainHwnd(excel);
            if (excelProcessId <= 0)
            {
                throw new InvalidOperationException("could not resolve Excel process id for VBE compilation");
            }

            stage = "vbe_compile";
            var compile = ExcelWorkerInvocation.InvokeWithWorker(
                new Workers.MacroRunWorkerRequest(excelProcessId, excelHwnd, "", Operation: "compile", WorkbookPath: temporaryWorkbook),
                excelHwnd, Windows.DialogKind.Compile, true, ExcelPushService.ResolveCompileTimeout(request), cancellationToken);
            if (compile.Dialog is not null || compile.TimedOut || compile.Result is null || !compile.Result.Ok)
            {
                var message = compile.Dialog is not null
                    ? "Unexpected VBE compile dialog: " + DialogMessage(compile.Dialog)
                    : compile.TimedOut ? "VBE Compile timed out." : compile.Result?.Error?.Message ?? "VBE Compile failed.";
                var fatalComFailure = compile.Result?.Error is not null && ExcelBridgeSupport.IsFatalComFailure(compile.Result.Error.Number);
                throw new BuildOperationException("build_compile_failed", message, stage, compile.WorkerProcessId, compile.TimedOut || fatalComFailure, compile.LocationCapture.Location);
            }

            stage = "workbook_save";
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
            stage = "workbook_close";
            var cleanupConfirmed = ExcelBridgeSupport.CloseWorkbookAndQuitApplicationAndConfirmExit(workbook, excel, excelProcessId);
            workbook = null;
            excel = null;
            if (!cleanupConfirmed)
            {
                throw new BuildOperationException("build_excel_cleanup_unconfirmed", "Excel did not exit cleanly after saving the staged workbook.", "excel_cleanup", 0, true, null);
            }
            stage = "publish";
            BuildArtifactPublisher.ValidateTemporaryArtifact(temporaryWorkbook);
            var publication = _artifactPublisher.Publish(temporaryWorkbook, outputPath);
            var manifest = CreateManifest(plan, basePath, outputPath, publication, applied);
            var manifestPublication = _artifactPublisher.PublishManifest(stagingDirectory!, outputPath, manifest);
            stagingCleanup = _artifactPublisher.Cleanup(stagingDirectory);
            return BuildSuccess(request, outputPath, publication, manifestPublication, stagingCleanup, applied);
        }
        catch (BuildOperationException failure)
        {
            var cleanupConfirmed = CloseAndConfirm(ref workbook, ref excel, excelProcessId);
            stagingCleanup ??= CleanupStage(stagingDirectory);
            return BuildFailure(request, failure.Code, failure.Message, failure.Stage, excelProcessId, failure.WorkerProcessId, failure.Uncertain || !cleanupConfirmed, failure.Location, stagingCleanup);
        }
        catch (BuildArtifactException failure)
        {
            var cleanupConfirmed = CloseAndConfirm(ref workbook, ref excel, excelProcessId);
            stagingCleanup ??= CleanupStage(stagingDirectory);
            return BuildFailure(request, failure.Code, failure.Message, failure.Stage, excelProcessId, 0, !cleanupConfirmed, null, stagingCleanup);
        }
        catch (Exception ex)
        {
            var cleanupConfirmed = CloseAndConfirm(ref workbook, ref excel, excelProcessId);
            stagingCleanup ??= CleanupStage(stagingDirectory);
            return BuildFailure(request, Classify(ex), ExcelBridgeSupport.FormatExceptionDetail(ex), stage, excelProcessId, 0, !cleanupConfirmed, null, stagingCleanup);
        }
        finally
        {
            CloseDedicated(workbook, excel);
            if (stagingCleanup is null && stagingDirectory is not null)
            {
                _ = CleanupStage(stagingDirectory);
            }
        }
    }

    private BuildArtifactCleanup? CleanupStage(string? stagingDirectory) => stagingDirectory is null ? null : _artifactPublisher.Cleanup(stagingDirectory);

    private static string DialogMessage(Windows.DialogSnapshot dialog)
    {
        var lines = dialog.Text.Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
        return lines.Length > 0 ? string.Join(Environment.NewLine, lines) : dialog.Title;
    }

    private static bool CloseAndConfirm(ref object? workbook, ref object? excel, int excelProcessId)
    {
        if (workbook is null && excel is null)
        {
            return true;
        }

        var confirmed = ExcelBridgeSupport.CloseWorkbookAndQuitApplicationAndConfirmExit(workbook, excel, excelProcessId);
        workbook = null;
        excel = null;
        return confirmed;
    }

    private sealed class BuildOperationException(string code, string message, string stage, int workerProcessId, bool uncertain, object? location) : Exception(message)
    {
        public string Code { get; } = code;
        public string Stage { get; } = stage;
        public int WorkerProcessId { get; } = workerProcessId;
        public bool Uncertain { get; } = uncertain;
        public object? Location { get; } = location;
    }

    private static bool HasDirtyMatchingSession(BuildCommandArguments args)
    {
        if (string.IsNullOrWhiteSpace(args.MetadataPath) || !File.Exists(args.MetadataPath) || string.IsNullOrWhiteSpace(args.SessionWorkbookPath))
        {
            return false;
        }

        object? excel = null;
        object? workbook = null;
        try
        {
            if (!ExcelBridgeSupport.SessionMetadataMatchesWorkbook(args.MetadataPath, args.SessionWorkbookPath))
            {
                return false;
            }

            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(args.SessionWorkbookPath, args.MetadataPath, useSession: false);
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            // If live dirty state cannot be inspected safely, fail closed: the
            // saved base file cannot be trusted as the session's source state.
            return !ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirty) || dirty;
        }
        catch { return true; }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static bool HasMatchingSession(BuildCommandArguments args, string workbookPath)
    {
        return HasLiveMatchingSession(args.MetadataPath, workbookPath, ExcelBridgeSupport.GetSessionExcel, ExcelBridgeSupport.GetOpenWorkbook);
    }

    // Metadata is not ownership proof: Excel can crash after writing it.  Only
    // block publication after reattaching to the recorded Excel session and
    // finding this exact workbook still open.
    internal static bool HasLiveMatchingSession(
        string metadataPath,
        string workbookPath,
        Func<string, object?> getSessionExcel,
        Func<object, string, object> getOpenWorkbook)
    {
        if (string.IsNullOrWhiteSpace(metadataPath) || !File.Exists(metadataPath) ||
            !ExcelBridgeSupport.SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            return false;
        }

        object? excel = null;
        object? workbook = null;
        try
        {
            excel = getSessionExcel(metadataPath);
            if (excel is null)
            {
                return false;
            }

            workbook = getOpenWorkbook(excel, workbookPath);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse BuildSuccess(BridgeRequest request, string outputPath, BuildArtifactPublication publication, BuildManifestPublication manifest, BuildArtifactCleanup cleanup, int applied)
    {
        var extensions = new Dictionary<string, object?>
        {
            ["output"] = new Dictionary<string, object?>
            {
                ["path"] = outputPath,
                ["replaced_existing"] = publication.ReplacedExisting,
                ["publication"] = publication.Publication,
                ["temporary_cleanup"] = CleanupDetails(cleanup),
            },
            ["build"] = new Dictionary<string, object?>
            {
                ["backend"] = "excel",
                ["source_applied"] = true,
                ["components_applied"] = applied,
                ["vbe_compile"] = "passed",
                ["workbook_saved"] = true,
                ["workbook_closed"] = true,
                ["excel_cleanup"] = "clean",
                ["publication"] = new Dictionary<string, object?>
                {
                    ["replaced_existing"] = publication.ReplacedExisting,
                    ["method"] = publication.Publication,
                },
                ["manifest"] = ManifestDetails(manifest),
            },
        };
        var warnings = new List<Dictionary<string, object?>>();
        if (!cleanup.Succeeded)
        {
            warnings.Add(new Dictionary<string, object?>
            {
                ["code"] = "build_temporary_cleanup_failed",
                ["message"] = "The build output was published, but its temporary staging directory could not be removed.",
                ["path"] = cleanup.ResidualPath,
            });
        }
        if (!manifest.Published)
        {
            warnings.Add(new Dictionary<string, object?>
            {
                ["code"] = "build_manifest_publish_failed",
                ["message"] = "The build workbook was published, but its companion manifest could not be published.",
                ["path"] = manifest.Path,
                ["error"] = manifest.Error,
            });
        }
        if (warnings.Count > 0)
        {
            extensions["warnings"] = warnings;
        }
        return BridgeResponse.Ok(request, extensions);
    }

    private static Dictionary<string, object?> ManifestDetails(BuildManifestPublication manifest) => new()
    {
        ["path"] = manifest.Path,
        ["published"] = manifest.Published,
        ["error"] = manifest.Error,
    };

    private static Dictionary<string, object?> CreateManifest(BuildPlanPayload plan, string basePath, string outputPath, BuildArtifactPublication publication, int applied) => new()
    {
        ["schema_version"] = 1,
        ["command"] = "build",
        ["backend"] = "excel",
        ["base"] = basePath,
        ["output"] = outputPath,
        ["included_components"] = plan.Included,
        ["excluded_components"] = plan.Excluded,
        ["validation"] = new Dictionary<string, object?>
        {
            ["source_applied"] = true,
            ["components_applied"] = applied,
            ["vbe_compile"] = "passed",
            ["workbook_saved"] = true,
            ["workbook_closed"] = true,
            ["excel_cleanup"] = "clean",
        },
        ["publication"] = new Dictionary<string, object?>
        {
            ["replaced_existing"] = publication.ReplacedExisting,
            ["method"] = publication.Publication,
        },
    };

    private static BridgeResponse BuildFailure(BridgeRequest request, string code, string message, string stage, int excelPid, int workerPid, bool uncertain, object? location, BuildArtifactCleanup? cleanup)
    {
        var details = new Dictionary<string, object?> { ["stage"] = stage };
        if (location is not null)
        {
            details["location"] = location;
        }
        if (cleanup is not null)
        {
            details["temporary_cleanup"] = CleanupDetails(cleanup);
        }

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(code, message, stage, "xlflow-excel-bridge", Details: details),
            Recovery = uncertain ? new BridgeRecovery { Required = true, Reason = "excel_cleanup_unconfirmed", Operation = "build", ExcelProcessId = excelPid > 0 ? excelPid : null, WorkerProcessId = workerPid > 0 ? workerPid : null, CleanupConfirmed = false } : null,
        };
    }

    private static Dictionary<string, object?> CleanupDetails(BuildArtifactCleanup cleanup) => new()
    {
        ["status"] = cleanup.Succeeded ? "clean" : "failed",
        ["residual_path"] = cleanup.ResidualPath,
        ["error"] = cleanup.Error,
    };

    private static int Reconstruct(object workbook, BuildPlanPayload plan, string codeSource, string tempRoot, CancellationToken cancellationToken)
    {
        ExcelPushService.RemoveNonDocumentComponents(workbook);
        var applied = ImportComponents(workbook, plan.Included.Where(component => component.Type is "standard" or "class" or "form").ToArray(), codeSource, tempRoot, cancellationToken);
        applied += UpdateDocumentModules(workbook, plan.Included.Where(component => component.Type == "document").ToArray());
        return applied;
    }

    private static int ImportComponents(object workbook, IReadOnlyList<BuildComponentPayload> components, string codeSource, string tempRoot, CancellationToken cancellationToken)
    {
        object? project = null;
        object? vbComponents = null;
        try
        {
            project = ExcelBridgeSupport.Get(workbook, "VBProject") ?? throw new InvalidOperationException("VBProject access is denied");
            vbComponents = ExcelBridgeSupport.Get(project, "VBComponents") ?? throw new InvalidOperationException("VBComponents are unavailable");
            var applied = 0;
            // Match push's dependency-friendly import order. In particular,
            // UserForms can reference standard modules or classes while Excel
            // validates their code during later VBE compilation.
            foreach (var component in components
                .OrderBy(component => component.Type switch
                {
                    "standard" => 0,
                    "class" => 1,
                    "form" => 2,
                    _ => 3,
                })
                .ThenBy(component => component.SourcePath, StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Keep the original .frm basename: its designer header names
                // the companion .frx file. Isolation comes from the unique
                // directory, not by renaming the component artifact itself.
                var importPath = Path.Combine(tempRoot, "imports", Guid.NewGuid().ToString("N"), Path.GetFileName(component.SourcePath));
                VbaSourceHelper.PrepareSourceForImport(component.SourcePath, importPath, null, "off");
                if (component.Type == "form")
                {
                    var frx = component.RelatedPaths.FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".frx", StringComparison.OrdinalIgnoreCase));
                    if (frx is not null)
                    {
                        File.Copy(frx, Path.ChangeExtension(importPath, ".frx"), true);
                    }
                }
                object? imported = null;
                try
                {
                    imported = ExcelBridgeSupport.InvokeViaDynamic(vbComponents, "Import", importPath) ?? throw new InvalidOperationException($"failed to import VBA component '{component.Name}'");
                    var name = ExcelBridgeSupport.GetString(imported, "Name") ?? "";
                    if (!string.Equals(name, component.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"imported VBA component name '{name}' did not match expected name '{component.Name}'");
                    }
                }
                finally { ExcelBridgeSupport.ReleaseComObject(imported); }
                if (component.Type == "form" && VbaSourceHelper.IsSidecarMode(codeSource))
                {
                    var codePath = component.RelatedPaths.FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".bas", StringComparison.OrdinalIgnoreCase));
                    if (string.IsNullOrWhiteSpace(codePath))
                    {
                        throw new InvalidOperationException($"UserForm '{component.Name}' has no sidecar code-behind");
                    }

                    ExcelPushService.SyncUserFormCodeBehindFromPath(workbook, component.Name, codePath, false, strict: true);
                }
                applied++;
            }
            return applied;
        }
        finally { ExcelBridgeSupport.ReleaseComObject(vbComponents); ExcelBridgeSupport.ReleaseComObject(project); }
    }

    private static int UpdateDocumentModules(object workbook, IReadOnlyList<BuildComponentPayload> documents)
    {
        var expected = documents.ToDictionary(component => component.Name, StringComparer.OrdinalIgnoreCase);
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        object? project = null;
        object? components = null;
        try
        {
            project = ExcelBridgeSupport.Get(workbook, "VBProject") ?? throw new InvalidOperationException("VBProject access is denied");
            components = ExcelBridgeSupport.Get(project, "VBComponents") ?? throw new InvalidOperationException("VBComponents are unavailable");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(components, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                object? code = null;
                try
                {
                    component = ExcelBridgeSupport.Get(components, "Item", index);
                    if (component is null || ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type")) != ComponentTypeDocument)
                    {
                        continue;
                    }

                    var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
                    code = ExcelBridgeSupport.Get(component, "CodeModule") ?? throw new InvalidOperationException($"document module '{name}' has no code module");
                    if (expected.TryGetValue(name, out var source))
                    {
                        VbaSourceHelper.SetCodeModuleText(code, VbaSourceHelper.NormalizeDocumentModuleContent(File.ReadAllText(source.SourcePath, Encoding.UTF8)));
                        found.Add(name);
                    }
                    else
                    {
                        // Document hosts are workbook-owned and cannot be removed.
                        // Clearing their body is the equivalent of excluding them.
                        VbaSourceHelper.SetCodeModuleText(code, "");
                    }
                }
                finally { ExcelBridgeSupport.ReleaseComObject(code); ExcelBridgeSupport.ReleaseComObject(component); }
            }
        }
        finally { ExcelBridgeSupport.ReleaseComObject(components); ExcelBridgeSupport.ReleaseComObject(project); }
        var missing = expected.Keys.Where(name => !found.Contains(name)).ToArray();
        if (missing.Length > 0)
        {
            throw new InvalidOperationException("document module could not be resolved: " + string.Join(", ", missing));
        }

        return found.Count;
    }

    private static BuildPlanPayload DecodePlan(string json64)
    {
        try
        {
            var plan = JsonSerializer.Deserialize<BuildPlanPayload>(Encoding.UTF8.GetString(Convert.FromBase64String(json64)), PlanJsonOptions);
            return plan ?? throw new InvalidOperationException("build plan is empty");
        }
        catch (FormatException ex) { throw new InvalidOperationException("PlanJson64 is not valid base64", ex); }
        catch (JsonException ex) { throw new InvalidOperationException("PlanJson64 does not contain a valid build plan", ex); }
    }

    private static void ValidatePlan(BuildPlanPayload plan, string projectRoot)
    {
        projectRoot = Path.GetFullPath(projectRoot);
        if (!Directory.Exists(projectRoot))
        {
            throw new InvalidOperationException("project root does not exist");
        }
        foreach (var component in plan.Included)
        {
            if (component.Type is not ("standard" or "class" or "document" or "form"))
            {
                throw new InvalidOperationException($"unsupported build component type '{component.Type}'");
            }

            component.SourcePath = ResolvePlannerPath(projectRoot, component.SourcePath);
            if (string.IsNullOrWhiteSpace(component.Name) || !File.Exists(component.SourcePath))
            {
                throw new InvalidOperationException($"invalid planned component '{component.Name}'");
            }

            for (var index = 0; index < component.RelatedPaths.Count; index++)
            {
                component.RelatedPaths[index] = ResolvePlannerPath(projectRoot, component.RelatedPaths[index]);
                if (!File.Exists(component.RelatedPaths[index]))
                {
                    throw new InvalidOperationException($"missing related artifact for '{component.Name}': {component.RelatedPaths[index]}");
                }
            }
        }
    }

    private static string ResolvePlannerPath(string projectRoot, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }

        var resolved = Path.GetFullPath(Path.IsPathFullyQualified(path)
            ? path
            : Path.Combine(projectRoot, path.Replace('/', Path.DirectorySeparatorChar)));
        if (!IsWithinDirectory(resolved, projectRoot))
        {
            throw new InvalidOperationException($"planned component path is outside the project root: {path}");
        }
        return resolved;
    }

    private static bool IsWithinDirectory(string path, string directory)
    {
        var relative = Path.GetRelativePath(Path.GetFullPath(directory), Path.GetFullPath(path));
        return relative == "." || (!relative.Equals("..", StringComparison.Ordinal) &&
            !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) &&
            !Path.IsPathFullyQualified(relative));
    }

    private static string Classify(Exception ex) => ex.Message.Contains("VBProject", StringComparison.OrdinalIgnoreCase) ? "build_vbproject_access_denied" : ex.Message.Contains("document module", StringComparison.OrdinalIgnoreCase) ? "build_document_module_unresolved" : ex.Message.Contains("UserForm", StringComparison.OrdinalIgnoreCase) ? "build_userform_reconstruct_failed" : "build_reconstruct_failed";

    private static void CloseDedicated(object? workbook, object? excel)
    {
        try { if (workbook is not null) { ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false); } } catch { }
        try { if (excel is not null) { ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit"); } } catch { }
        ExcelBridgeSupport.ReleaseComObject(workbook);
        ExcelBridgeSupport.ReleaseComObject(excel);
    }

    private sealed class BuildPlanPayload
    {
        [JsonPropertyName("base_workbook")]
        public string BaseWorkbook { get; init; } = "";
        [JsonPropertyName("output_path")]
        public string OutputPath { get; init; } = "";
        [JsonPropertyName("included")]
        public List<BuildComponentPayload> Included { get; init; } = [];
        [JsonPropertyName("excluded")]
        public List<BuildComponentPayload> Excluded { get; init; } = [];
    }

    private sealed class BuildComponentPayload
    {
        [JsonPropertyName("source_path")]
        public string SourcePath { get; set; } = "";
        [JsonPropertyName("name")]
        public string Name { get; init; } = "";
        [JsonPropertyName("type")]
        public string Type { get; init; } = "";
        [JsonPropertyName("related_paths")]
        public List<string> RelatedPaths { get; set; } = [];
    }
}
