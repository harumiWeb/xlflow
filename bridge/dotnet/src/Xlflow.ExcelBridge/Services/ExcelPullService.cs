using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Pull export intentionally degrades COM failures into structured bridge responses.")]
public sealed class ExcelPullService : IPullService
{
    private const int ComponentTypeModule = 1;
    private const int ComponentTypeClass = 2;
    private const int ComponentTypeForm = 3;
    private const int ComponentTypeDocument = 100;

    public BridgeResponse Execute(BridgeRequest request, PullCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var stagingRoot = Path.Combine(Path.GetTempPath(), "xlflow-pull-stage-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));

        try
        {
            var (attachment, attached) = ExcelBridgeSupport.RunPhase(
                "attach_or_open_workbook",
                () => AttachOrOpenWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            sessionAttached = attached;

            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var sessionMode = ResolveSessionMode(attached, args.UseSession);
            var modulesDir = args.ModulesDir;
            var classesDir = args.ClassesDir;
            var formsDir = args.FormsDir;
            var workbookDir = args.WorkbookDir;

            var stagedModulesDir = Path.Combine(stagingRoot, "modules");
            var stagedClassesDir = Path.Combine(stagingRoot, "classes");
            var stagedFormsDir = Path.Combine(stagingRoot, "forms");
            var stagedWorkbookDir = Path.Combine(stagingRoot, "workbook");
            EnsureDirectory(stagedModulesDir);
            EnsureDirectory(stagedClassesDir);
            EnsureDirectory(stagedFormsDir);
            EnsureDirectory(stagedWorkbookDir);

            var (exportedCount, exportedFormCodeCount) = ExcelBridgeSupport.RunPhase(
                "export_vba_components",
                () => ExportVbaComponents(
                    workbook, stagedModulesDir, stagedClassesDir, stagedFormsDir, stagedWorkbookDir,
                    args.Folders, args.FolderAnnotation, args.DefaultComponentFolders,
                    args.CodeSource, args.LineNumbersEnabled));

            PromoteStagedSource(stagingRoot, modulesDir, classesDir, formsDir, workbookDir, args.CodeSource);

            var userFormNames = ExcelBridgeSupport.RunPhase(
                "discover_userforms",
                () => GetUserFormNames(workbook));

            var dirty = false;
            var needsSave = false;
            try
            {
                dirty = !ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(workbook, "Saved"));
                needsSave = sessionAttached && dirty;
            }
            catch
            {
                // best-effort save state
            }

            var logs = new List<string>
            {
                string.Format(CultureInfo.InvariantCulture, "exported {0} VBA component(s)", exportedCount),
            };
            if (exportedFormCodeCount > 0)
            {
                logs.Add(string.Format(CultureInfo.InvariantCulture, "exported {0} UserForm code-behind sidecar(s)", exportedFormCodeCount));
            }

            var warnings = new List<Dictionary<string, string>>();
            var hints = new List<Dictionary<string, string>>();

            AddUserFormDiscoveryMessages(warnings, hints, userFormNames);

            if (needsSave)
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live workbook is newer than disk. pull exported from the live workbook rather than the saved workbook file.",
                });
            }

            if (needsSave && userFormNames.Count > 0)
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "userform_unsaved_session_state",
                    ["message"] = "Workbook contains UserForms (" + string.Join(", ", userFormNames) + ") and the live workbook is newer than disk. Run `xlflow save --session` and `xlflow pull` before reviewing `.frm`/`.frx` or disk-backed inspect output.",
                });
            }

            var response = BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?>
                {
                    ["kind"] = sessionAttached ? "live_session" : "file",
                    ["path"] = workbookPath,
                },
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = sessionAttached,
                    ["workbook_path"] = workbookPath,
                    ["dirty"] = needsSave,
                    ["save_required"] = needsSave,
                    ["live_newer_than_disk"] = needsSave,
                    ["mode"] = sessionMode,
                    ["source_of_truth"] = needsSave ? "live_workbook" : "saved_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = workbookPath,
                    ["session"] = sessionAttached,
                    ["session_mode"] = sessionMode,
                    ["session_requested"] = sessionAttached && args.UseSession,
                    ["auto_session"] = sessionAttached && !args.UseSession,
                    ["dirty"] = needsSave,
                    ["needs_save"] = needsSave,
                },
                ["source"] = new Dictionary<string, object?>
                {
                    ["modules_dir"] = modulesDir,
                    ["classes_dir"] = classesDir,
                    ["forms_dir"] = formsDir,
                    ["workbook_dir"] = workbookDir,
                },
                ["logs"] = logs,
            });

            if (warnings.Count > 0)
            {
                response.Extensions["warnings"] = warnings;
            }
            if (hints.Count > 0)
            {
                response.Extensions["hints"] = hints;
            }

            return response;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "session_required",
                Message: ex.Message,
                Phase: "pull",
                Source: "xlflow"));
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("bridge_file_not_openable", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "bridge_file_not_openable",
                Message: ex.Message.Replace("bridge_file_not_openable: ", "", StringComparison.OrdinalIgnoreCase),
                Phase: "pull",
                Source: "xlflow-excel-bridge"));
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("vba_line_number_safety_failed:", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "vba_line_number_safety_failed",
                Message: ex.Message.Replace("vba_line_number_safety_failed: ", "", StringComparison.OrdinalIgnoreCase),
                Phase: "strip_line_numbers",
                Source: "xlflow-excel-bridge"));
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "pull_failed",
                Message: detail,
                Phase: "pull",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            if (Directory.Exists(stagingRoot))
            {
                try { Directory.Delete(stagingRoot, true); }
                catch (IOException) { /* best-effort cleanup */ }
            }
            if (!sessionAttached)
            {
                CloseComInstance(workbook, excel);
            }
            else
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }
        }
    }

    private static (ExcelSessionAttachment Attachment, bool SessionAttached) AttachOrOpenWorkbook(
        string workbookPath, string metadataPath, bool useSession, bool visible)
    {
        try
        {
            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, useSession);
            return (attachment, true);
        }
        catch (InvalidOperationException ex) when (
            ex.Message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("no matching xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            if (useSession)
            {
                throw;
            }

            if (!ExcelBridgeSupport.IsExcelFile(ExcelBridgeSupport.NormalizePath(workbookPath)))
            {
                throw new InvalidOperationException($"bridge_file_not_openable: File does not appear to be an Excel workbook: {workbookPath}");
            }

            var attachment = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
            return (attachment, false);
        }
    }

    private static (int ExportedCount, int ExportedFormCodeCount) ExportVbaComponents(
        object workbook, string modulesDir, string classesDir, string formsDir, string workbookDir,
        bool folders, string folderAnnotation, bool defaultComponentFolders, string? codeSource, bool lineNumbersEnabled)
    {
        object? vbProject = null;
        object? vbComponents = null;
        var exported = 0;
        var exportedFormCode = 0;

        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
            vbComponents = ExcelBridgeSupport.Get(vbProject!, "VBComponents");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(vbComponents!, "Count"));

            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(vbComponents!, "Item", index);
                    ExportComponent(component!, modulesDir, classesDir, formsDir, workbookDir,
                        folders, folderAnnotation, defaultComponentFolders, lineNumbersEnabled);
                    exported++;

                    if (VbaSourceHelper.IsSidecarMode(codeSource) &&
                        ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component!, "Type")) == ComponentTypeForm)
                    {
                        if (ExportUserFormCodeBehind(component!, formsDir, lineNumbersEnabled))
                        {
                            exportedFormCode++;
                        }
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }

        return (exported, exportedFormCode);
    }

    private static void ExportComponent(object component, string modulesDir, string classesDir, string formsDir, string workbookDir,
        bool folders, string folderAnnotation, bool defaultComponentFolders, bool lineNumbersEnabled)
    {
        var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
        var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));

        var targetFile = VbaSourceHelper.GetComponentPath(component, modulesDir, classesDir, formsDir, workbookDir,
            folders, folderAnnotation, defaultComponentFolders);

        if (string.IsNullOrWhiteSpace(targetFile))
        {
            return;
        }

        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-pull-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        try
        {
            Directory.CreateDirectory(tmpDir);
            var tmpFile = Path.Combine(tmpDir, name + (VbaSourceHelper.GetComponentExtension(type) ?? ".bas"));
            ExcelBridgeSupport.InvokeViaDynamic(component, "Export", tmpFile);

            foreach (var exportedFile in Directory.GetFiles(tmpDir))
            {
                var ext = Path.GetExtension(exportedFile).ToLowerInvariant();

                var parentDir = Path.GetDirectoryName(targetFile);
                if (!string.IsNullOrWhiteSpace(parentDir))
                {
                    Directory.CreateDirectory(parentDir);
                }

                if (IsTextFileExtension(ext))
                {
                    var content = VbaSourceHelper.ReadExportedTextAsUtf8(exportedFile);

                    if (type == ComponentTypeDocument)
                    {
                        content = VbaSourceHelper.NormalizeDocumentModuleContent(content);
                        if (lineNumbersEnabled && !ErlLineNumberTransformer.TryRemove(content, out content, out var lineNumberIssue, excelExported: true))
                        {
                            throw new InvalidOperationException($"vba_line_number_safety_failed: {targetFile}:{lineNumberIssue!.Line}: {lineNumberIssue.Message}");
                        }
                        if (!string.IsNullOrWhiteSpace(workbookDir))
                        {
                            var desiredAnnotation = VbaSourceHelper.GetFolderAnnotationForPath(workbookDir, targetFile);
                            content = VbaSourceHelper.UpdateFolderAnnotationText(content, folderAnnotation, desiredAnnotation);
                        }
                    }
                    else if (type == ComponentTypeForm)
                    {
                        if (lineNumbersEnabled && !ErlLineNumberTransformer.TryRemove(content, out content, out var lineNumberIssue, excelExported: true))
                        {
                            throw new InvalidOperationException($"vba_line_number_safety_failed: {targetFile}:{lineNumberIssue!.Line}: {lineNumberIssue.Message}");
                        }
                        content = NormalizeUserFormArtifact(content, component);
                    }
                    else if (lineNumbersEnabled && !ErlLineNumberTransformer.TryRemove(content, out content, out var lineNumberIssue, excelExported: true))
                    {
                        throw new InvalidOperationException($"vba_line_number_safety_failed: {targetFile}:{lineNumberIssue!.Line}: {lineNumberIssue.Message}");
                    }

                    if (type != ComponentTypeDocument && folderAnnotation != "ignore")
                    {
                        var rootDir = VbaSourceHelper.GetComponentRootDir(type, modulesDir, classesDir, formsDir, workbookDir);
                        if (!string.IsNullOrWhiteSpace(rootDir))
                        {
                            var desiredAnnotation = VbaSourceHelper.GetFolderAnnotationForPath(rootDir, targetFile);
                            content = VbaSourceHelper.UpdateFolderAnnotationText(content, folderAnnotation, desiredAnnotation);
                        }
                    }

                    File.WriteAllText(targetFile, content, new UTF8Encoding(false));
                }
                else
                {
                    var binaryTarget = Path.ChangeExtension(targetFile, ext);
                    File.Copy(exportedFile, binaryTarget, true);
                }
            }
        }
        finally
        {
            if (Directory.Exists(tmpDir))
            {
                try { Directory.Delete(tmpDir, true); }
                catch (IOException) { /* best-effort cleanup */ }
            }
        }
    }

    private static bool ExportUserFormCodeBehind(object component, string formsDir, bool lineNumbersEnabled)
    {
        var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
        var codePath = VbaSourceHelper.GetUserFormCodePath(formsDir, name);
        if (string.IsNullOrWhiteSpace(codePath))
        {
            return false;
        }

        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
            var text = VbaSourceHelper.GetCodeModuleText(codeModule!);
            if (lineNumbersEnabled && !ErlLineNumberTransformer.TryRemove(text, out text, out var lineNumberIssue, excelExported: true))
            {
                throw new InvalidOperationException($"vba_line_number_safety_failed: {codePath}:{lineNumberIssue!.Line}: {lineNumberIssue.Message}");
            }
            if (string.IsNullOrWhiteSpace(text))
            {
                if (File.Exists(codePath))
                {
                    try { File.Delete(codePath); }
                    catch (IOException) { /* best-effort */ }
                }
                return false;
            }

            var parentDir = Path.GetDirectoryName(codePath);
            if (!string.IsNullOrWhiteSpace(parentDir))
            {
                Directory.CreateDirectory(parentDir);
            }
            File.WriteAllText(codePath, text, new UTF8Encoding(false));
            return true;
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch
        {
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    private static string NormalizeUserFormArtifact(string content, object component)
    {
        try
        {
            object? designer = null;
            try
            {
                designer = ExcelBridgeSupport.Get(component, "Designer");
                var caption = ExcelBridgeSupport.GetString(designer!, "Caption");
                if (!string.IsNullOrWhiteSpace(caption))
                {
                    content = InjectOrUpdateCaption(content, caption);
                }
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(designer);
            }
        }
        catch
        {
            // best-effort caption normalization
        }
        return content;
    }

    private static string InjectOrUpdateCaption(string content, string caption)
    {
        var lines = new List<string>(content.Split(["\r\n", "\n", "\r"], StringSplitOptions.None));
        var inBeginEnd = false;
        var captionIndex = -1;

        for (var i = 0; i < lines.Count; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed == "Begin")
            {
                inBeginEnd = true;
                continue;
            }
            if (trimmed == "End")
            {
                inBeginEnd = false;
                continue;
            }
            if (inBeginEnd && trimmed.StartsWith("Caption", StringComparison.OrdinalIgnoreCase) && trimmed.Contains('='))
            {
                captionIndex = i;
                break;
            }
        }

        var captionLine = $"   Caption = \"{caption}\"";
        if (captionIndex >= 0)
        {
            lines[captionIndex] = captionLine;
        }
        else if (inBeginEnd)
        {
            for (var i = 0; i < lines.Count; i++)
            {
                if (lines[i].Trim() == "Begin")
                {
                    lines.Insert(i + 1, captionLine);
                    break;
                }
            }
        }

        return string.Join(Environment.NewLine, lines);
    }

    internal static bool IsTextFileExtension(string ext)
    {
        return ext is ".bas" or ".cls" or ".frm";
    }

    private static void ClearExistingSourceFiles(string modulesDir, string classesDir, string formsDir, string workbookDir, string? codeSource)
    {
        var files = VbaSourceHelper.DiscoverSourceFiles(modulesDir, classesDir, formsDir, workbookDir, codeSource);
        foreach (var file in files)
        {
            if (file.Kind == "form_code")
            {
                continue;
            }
            try { File.Delete(file.FullPath); }
            catch (IOException) { /* best-effort */ }
        }
    }

    private static void PromoteStagedSource(string stagingRoot, string modulesDir, string classesDir, string formsDir, string workbookDir, string? codeSource)
    {
        ClearExistingSourceFiles(modulesDir, classesDir, formsDir, workbookDir, codeSource);
        var targets = new[]
        {
            (Stage: Path.Combine(stagingRoot, "modules"), Destination: modulesDir),
            (Stage: Path.Combine(stagingRoot, "classes"), Destination: classesDir),
            (Stage: Path.Combine(stagingRoot, "forms"), Destination: formsDir),
            (Stage: Path.Combine(stagingRoot, "workbook"), Destination: workbookDir),
        };
        foreach (var (stage, destination) in targets)
        {
            if (!Directory.Exists(stage))
            {
                continue;
            }
            foreach (var source in Directory.GetFiles(stage, "*", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(stage, source);
                var target = Path.Combine(destination, relative);
                var parent = Path.GetDirectoryName(target);
                if (!string.IsNullOrWhiteSpace(parent))
                {
                    Directory.CreateDirectory(parent);
                }
                File.Copy(source, target, true);
            }
        }
    }

    private static void EnsureDirectory(string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) && !Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }
    }

    private static void CloseComInstance(object? workbook, object? excel)
    {
        try
        {
            if (workbook is not null)
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            }
        }
        catch
        {
            // best-effort close
        }

        try
        {
            if (excel is not null)
            {
                ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit");
            }
        }
        catch
        {
            // best-effort quit
        }

        ExcelBridgeSupport.ReleaseComObject(workbook);
        ExcelBridgeSupport.ReleaseComObject(excel);
    }

    private static List<string> GetUserFormNames(object workbook)
    {
        var names = new List<string>();
        object? vbProject = null;
        object? vbComponents = null;
        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
            vbComponents = ExcelBridgeSupport.Get(vbProject!, "VBComponents");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(vbComponents!, "Count"));

            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(vbComponents!, "Item", index);
                    var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component!, "Type"));
                    if (type == ComponentTypeForm)
                    {
                        var name = ExcelBridgeSupport.GetString(component!, "Name");
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                }
                catch
                {
                    // best-effort component inspection
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        catch
        {
            // best-effort UserForm detection
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }

        names.Sort(StringComparer.OrdinalIgnoreCase);
        return names;
    }

    private static void AddUserFormDiscoveryMessages(
        List<Dictionary<string, string>> warnings,
        List<Dictionary<string, string>> hints,
        List<string> userFormNames)
    {
        if (userFormNames.Count == 0)
        {
            return;
        }

        var namesList = string.Join(", ", userFormNames);
        warnings.Add(new Dictionary<string, string>
        {
            ["code"] = "userform_state_partial",
            ["message"] = "UserForms detected: " + namesList + ". `.frm` text may not fully represent layout, binary `.frx` state, or VBIDE Designer-backed properties.",
        });
        hints.Add(new Dictionary<string, string>
        {
            ["code"] = "userform_planned_commands",
            ["message"] = "UserForm workflow: `xlflow pull --json`, `xlflow inspect form <name> --designer --json`, `xlflow form snapshot <name> --out src/forms/specs/<name>.yaml`, edit spec/code artifacts, then `xlflow form build src/forms/specs/<name>.yaml --overwrite` and verify with `xlflow form export-image <name> --out <path>`.",
        });
    }

    private static string ResolveSessionMode(bool sessionAttached, bool useSession)
    {
        if (!sessionAttached)
        {
            return "none";
        }
        return useSession ? "explicit" : "auto";
    }
}
