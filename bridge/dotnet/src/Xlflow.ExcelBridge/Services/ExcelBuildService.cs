using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

// Reconstructs only an isolated workbook copy. Publication, compilation, and
// recovery policy intentionally belong to later build sub-issues.
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "The bridge converts Excel and filesystem failures into structured build errors.")]
public sealed class ExcelBuildService : IBuildService
{
    private const int ComponentTypeDocument = 100;
    private const int ComponentTypeForm = 3;
    private static readonly JsonSerializerOptions PlanJsonOptions = new() { PropertyNameCaseInsensitive = true };

    public BridgeResponse Execute(BridgeRequest request, BuildCommandArguments args, CancellationToken cancellationToken)
    {
        object? excel = null;
        object? workbook = null;
        string? temporaryBuildDirectory = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var plan = DecodePlan(args.PlanJson64);
            ValidatePlan(plan, args.ProjectRoot);
            var basePath = ExcelBridgeSupport.NormalizePath(args.BaseWorkbookPath);
            if (!File.Exists(basePath))
            {
                throw new InvalidOperationException("base workbook does not exist");
            }

            var tempParent = Path.GetFullPath(args.TemporaryDirectory);
            if (IsWithinDirectory(basePath, tempParent))
            {
                throw new InvalidOperationException("temporary directory must not contain the base workbook");
            }

            // TemporaryDirectory belongs to the caller. Own and later remove
            // only this invocation's unique child directory.
            Directory.CreateDirectory(tempParent);
            temporaryBuildDirectory = Path.Combine(tempParent, "xlflow-build-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temporaryBuildDirectory);
            var temporaryWorkbook = Path.Combine(temporaryBuildDirectory, Path.GetFileName(basePath));

            File.Copy(basePath, temporaryWorkbook);

            var attachment = ExcelBridgeSupport.OpenWorkbookDirect(temporaryWorkbook, args.Visible, disableAutomationMacros: true);
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            var applied = Reconstruct(workbook, plan, args.CodeSource, temporaryBuildDirectory, cancellationToken);
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            ExcelBridgeSupport.ReleaseComObject(workbook);
            workbook = null;
            ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit");
            ExcelBridgeSupport.ReleaseComObject(excel);
            excel = null;

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["build"] = new Dictionary<string, object?>
                {
                    ["backend"] = "excel",
                    ["temporary_reconstruction"] = true,
                    ["source_applied"] = true,
                    ["components_applied"] = applied,
                    ["workbook_saved"] = true,
                    ["workbook_closed"] = true,
                },
            });
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: Classify(ex),
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "build_reconstruct",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            CloseDedicated(workbook, excel);
            if (temporaryBuildDirectory is not null && Directory.Exists(temporaryBuildDirectory))
            {
                try { Directory.Delete(temporaryBuildDirectory, true); } catch { }
            }
        }
    }

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
            foreach (var component in components)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var importPath = Path.Combine(tempRoot, "imports", Guid.NewGuid().ToString("N") + Path.GetExtension(component.SourcePath));
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
        if (string.IsNullOrWhiteSpace(path)) return "";
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
        [JsonPropertyName("included")]
        public List<BuildComponentPayload> Included { get; init; } = [];
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
