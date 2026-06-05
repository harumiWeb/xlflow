using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Push import intentionally degrades COM failures into structured bridge responses.")]
public sealed class ExcelPushService : IPushService
{
    private const int ComponentTypeModule = 1;
    private const int ComponentTypeClass = 2;
    private const int ComponentTypeForm = 3;
    private const int ComponentTypeDocument = 100;

    private static readonly JsonSerializerOptions IndentedJsonOptions = new() { WriteIndented = true };

    public BridgeResponse Execute(BridgeRequest request, PushCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var tmpImportDir = Path.Combine(Path.GetTempPath(), "xlflow-push-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));

        try
        {
            var sourceFiles = VbaSourceHelper.DiscoverSourceFiles(
                args.ModulesDir, args.ClassesDir, args.FormsDir, args.WorkbookDir, args.CodeSource);
            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var fingerprint = VbaSourceHelper.ComputeFingerprint(
                args.WorkbookPath, args.ModulesDir, args.ClassesDir, args.FormsDir, args.WorkbookDir, args.CodeSource);

            var duplicates = VbaSourceHelper.FindDuplicateModuleNames(sourceFiles);
            if (duplicates.Count > 0)
            {
                var messages = new List<string>();
                foreach (var dup in duplicates)
                {
                    messages.Add(string.Join(", ", dup.Paths));
                }
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "duplicate_module_name",
                    Message: "Duplicate VBA module names detected in source tree. Rename the conflicting files before push.",
                    Phase: "push",
                    Source: "xlflow"));
            }

            if (args.ChangedOnly && VbaSourceHelper.FingerprintMatchesState(fingerprint, args.StatePath))
            {
                var noopSourceUserFormNames = GetSourceUserFormNames(sourceFiles);
                var noopWarnings = new List<Dictionary<string, string>>();
                var noopHints = new List<Dictionary<string, string>>();
                AddUserFormDiscoveryMessages(noopWarnings, noopHints, noopSourceUserFormNames);

                var noopExtensions = new Dictionary<string, object?>
                {
                    ["target"] = new Dictionary<string, object?>
                    {
                        ["kind"] = "file",
                        ["path"] = workbookPath,
                    },
                    ["session"] = new Dictionary<string, object?>
                    {
                        ["active"] = false,
                        ["workbook_path"] = workbookPath,
                        ["dirty"] = false,
                        ["save_required"] = false,
                        ["live_newer_than_disk"] = false,
                        ["mode"] = "none",
                        ["source_of_truth"] = "saved_workbook",
                    },
                    ["workbook"] = new Dictionary<string, object?>
                    {
                        ["path"] = workbookPath,
                        ["session"] = false,
                        ["session_mode"] = "none",
                        ["session_requested"] = false,
                        ["auto_session"] = false,
                        ["saved"] = false,
                        ["dirty"] = false,
                        ["needs_save"] = false,
                    },
                    ["backup"] = new Dictionary<string, object?>
                    {
                        ["path"] = (string?)null,
                        ["mode"] = args.BackupMode,
                    },
                    ["source"] = new Dictionary<string, object?>
                    {
                        ["changed_only"] = true,
                        ["changed"] = false,
                        ["state"] = args.StatePath,
                    },
                    ["logs"] = new List<string> { "source state unchanged; skipped workbook import" },
                };

                if (noopWarnings.Count > 0)
                {
                    noopExtensions["warnings"] = noopWarnings;
                }
                if (noopHints.Count > 0)
                {
                    noopExtensions["hints"] = noopHints;
                }

                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Extensions = noopExtensions,
                };
            }

            var (attachment, attached) = ExcelBridgeSupport.RunPhase(
                "attach_or_open_workbook",
                () => AttachOrOpenWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = attachment.Excel;
            workbook = attachment.Workbook;
            sessionAttached = attached;

            var sessionMode = ResolveSessionMode(attached, args.UseSession);

            string? backupId = null;
            string? backupPath = null;
            const string backupReason = "before-push";

            if (args.BackupMode != "never" && !string.IsNullOrWhiteSpace(args.BackupRoot))
            {
                var (id, path) = ExcelBridgeSupport.RunPhase(
                    "create_backup",
                    () => CreateBackup(workbook, args.BackupRoot, workbookPath));
                backupId = id;
                backupPath = path;
            }

            Directory.CreateDirectory(tmpImportDir);

            ExcelBridgeSupport.RunPhase(
                "remove_non_document_components",
                () => RemoveNonDocumentComponents(workbook));

            var importResult = ExcelBridgeSupport.RunPhase(
                "import_vba_components",
                () => ImportVbaComponents(workbook, args, sourceFiles, tmpImportDir));
            var documentModulesUpdated = ExcelBridgeSupport.RunPhase(
                "update_document_modules",
                () => UpdateDocumentModules(workbook, args));

            var excelProcessId = ExcelBridgeSupport.GetExcelProcessId(excel);
            var excelHwnd = ExcelBridgeSupport.GetExcelMainHwnd(excel);
            if (excelProcessId <= 0)
            {
                throw new InvalidOperationException("compile_vba failed: could not resolve the Excel process id.");
            }

            var compileInvocation = ExcelWorkerInvocation.InvokeWithWorker(
                new MacroRunWorkerRequest(
                    excelProcessId,
                    excelHwnd,
                    "",
                    Operation: "compile",
                    WorkbookPath: args.WorkbookPath),
                excelHwnd,
                DialogKind.Compile,
                suppressModalErrors: true,
                ResolveCompileTimeout(request),
                cancellationToken);
            if (compileInvocation.Dialog is not null ||
                compileInvocation.TimedOut ||
                compileInvocation.Result is null ||
                !compileInvocation.Result.Ok)
            {
                return BuildCompileFailureResponse(
                    request,
                    args,
                    workbookPath,
                    sessionAttached,
                    sessionMode,
                    compileInvocation);
            }

            var saved = false;
            if (!args.NoSave)
            {
                ExcelBridgeSupport.RunPhase(
                    "save_workbook",
                    () => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));
                saved = true;
            }

            VbaSourceHelper.WriteFingerprintState(fingerprint, args.StatePath);

            var needsSave = sessionAttached && !saved;
            var logs = new List<string>
            {
                string.Format(CultureInfo.InvariantCulture, "imported {0} source file(s)", importResult),
                string.Format(CultureInfo.InvariantCulture, "updated {0} workbook module(s)", documentModulesUpdated),
                saved ? "saved workbook in place" : "left workbook unchanged",
            };

            var sourceUserFormNames = GetSourceUserFormNames(sourceFiles);
            var warnings = new List<Dictionary<string, string>>();
            var hints = new List<Dictionary<string, string>>();

            AddUserFormDiscoveryMessages(warnings, hints, sourceUserFormNames);

            if (needsSave)
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "save_required",
                    ["message"] = "Source files were pushed to the live workbook. The live workbook is newer than disk until `xlflow save --session` persists it.",
                });
            }

            if (needsSave && sourceUserFormNames.Count > 0)
            {
                warnings.Add(new Dictionary<string, string>
                {
                    ["code"] = "userform_unsaved_session_state",
                    ["message"] = "Workbook contains UserForms (" + string.Join(", ", sourceUserFormNames) + ") and the live workbook is newer than disk. Run `xlflow save --session` and `xlflow pull` before reviewing `.frm`/`.frx` or disk-backed inspect output.",
                });
            }

            var response = new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Extensions = new Dictionary<string, object?>
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
                        ["saved"] = saved,
                        ["dirty"] = needsSave,
                        ["needs_save"] = needsSave,
                    },
                    ["backup"] = new Dictionary<string, object?>
                    {
                        ["id"] = args.BackupMode == "always" ? backupId : null,
                        ["path"] = args.BackupMode == "always" ? backupPath : null,
                        ["reason"] = args.BackupMode == "always" ? backupReason : null,
                        ["mode"] = args.BackupMode,
                    },
                    ["source"] = new Dictionary<string, object?>
                    {
                        ["changed_only"] = args.ChangedOnly,
                        ["changed"] = true,
                        ["state"] = args.StatePath,
                    },
                    ["logs"] = logs,
                },
            };

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
                Phase: "push",
                Source: "xlflow"));
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("bridge_file_not_openable", StringComparison.OrdinalIgnoreCase))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "bridge_file_not_openable",
                Message: ex.Message.Replace("bridge_file_not_openable: ", "", StringComparison.OrdinalIgnoreCase),
                Phase: "push",
                Source: "xlflow-excel-bridge"));
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "push_failed",
                Message: detail,
                Phase: "push",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
            if (!sessionAttached)
            {
                CloseComInstance(workbook, excel);
            }
            else
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }

            if (Directory.Exists(tmpImportDir))
            {
                try { Directory.Delete(tmpImportDir, true); }
                catch (IOException) { /* best-effort cleanup */ }
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

    private static void RemoveNonDocumentComponents(object workbook)
    {
        object? vbProject = null;
        object? vbComponents = null;
        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
            if (vbProject is null)
            {
                return;
            }

            vbComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (vbComponents is null)
            {
                return;
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(vbComponents, "Count"));

            for (var index = count; index >= 1; index--)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(vbComponents, "Item", index);
                    if (component is null)
                    {
                        continue;
                    }

                    var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
                    if (type != ComponentTypeDocument)
                    {
                        ExcelBridgeSupport.InvokeViaDynamic(vbComponents, "Remove", component);
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        catch
        {
            // best-effort removal; import will replace anyway
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }
    }

    private static int ImportVbaComponents(object workbook, PushCommandArguments args, List<DiscoveredSourceFile> sourceFiles, string tmpImportDir)
    {
        var imported = 0;

        object? vbProject = null;
        object? vbComponents = null;
        try
        {
            vbProject = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"));
            if (vbProject is null)
            {
                return 0;
            }

            vbComponents = ExcelBridgeSupport.RunPhase("get_vbcomponents", () => ExcelBridgeSupport.Get(vbProject, "VBComponents"));
            if (vbComponents is null)
            {
                return 0;
            }

            foreach (var file in sourceFiles)
            {
                if (file.Kind is "document" or "form_code")
                {
                    continue;
                }

                if (file.Extension == ".frx")
                {
                    continue;
                }

                var rootDir = VbaSourceHelper.GetComponentRootDir(
                    file.Kind == "module" ? 1 :
                    file.Kind == "class" ? 2 :
                    file.Kind == "form" ? 3 : 100,
                    args.ModulesDir, args.ClassesDir, args.FormsDir, args.WorkbookDir);

                var importPath = VbaSourceHelper.PrepareSourceForImport(
                    file.FullPath,
                    Path.Combine(tmpImportDir, Path.GetFileName(file.FullPath)),
                    rootDir,
                    args.FolderAnnotation);

                if (string.IsNullOrWhiteSpace(importPath))
                {
                    continue;
                }

                if (file.Extension == ".frm")
                {
                    var frxSource = Path.ChangeExtension(file.FullPath, ".frx");
                    if (File.Exists(frxSource))
                    {
                        var frxDest = Path.ChangeExtension(importPath, ".frx");
                        File.Copy(frxSource, frxDest, true);
                    }
                }

                ExcelBridgeSupport.InvokeViaDynamic(vbComponents, "Import", importPath);
                imported++;

                if (VbaSourceHelper.IsSidecarMode(args.CodeSource) && file.Kind == "form" && file.Extension == ".frm")
                {
                    SyncUserFormCodeBehind(workbook, file.ModuleName, args.FormsDir);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }

        return imported;
    }

    private static void SyncUserFormCodeBehind(object workbook, string formName, string formsDir)
    {
        var codePath = VbaSourceHelper.GetUserFormCodePath(formsDir, formName);
        if (string.IsNullOrWhiteSpace(codePath) || !File.Exists(codePath))
        {
            return;
        }

        var codeText = File.ReadAllText(codePath, Encoding.UTF8);
        if (string.IsNullOrWhiteSpace(codeText))
        {
            return;
        }

        object? vbProject = null;
        object? vbComponents = null;
        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
            if (vbProject is null)
            {
                return;
            }

            vbComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (vbComponents is null)
            {
                return;
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(vbComponents, "Count"));

            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(vbComponents, "Item", index);
                    if (component is null)
                    {
                        continue;
                    }

                    var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
                    var name = (string?)ExcelBridgeSupport.Get(component, "Name");

                    if (type == ComponentTypeForm && string.Equals(name, formName, StringComparison.OrdinalIgnoreCase))
                    {
                        object? codeModule = null;
                        try
                        {
                            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
                            if (codeModule is not null)
                            {
                                VbaSourceHelper.SetCodeModuleText(codeModule, codeText);
                            }
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(codeModule);
                        }
                        return;
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        catch
        {
            // best-effort UserForm code-behind sync
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }
    }

    private static int UpdateDocumentModules(object workbook, PushCommandArguments args)
    {
        if (string.IsNullOrWhiteSpace(args.WorkbookDir) || !Directory.Exists(args.WorkbookDir))
        {
            return 0;
        }

        var updated = 0;
        object? vbProject = null;
        object? vbComponents = null;
        try
        {
            vbProject = ExcelBridgeSupport.Get(workbook, "VBProject");
            if (vbProject is null)
            {
                return 0;
            }

            vbComponents = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (vbComponents is null)
            {
                return 0;
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(vbComponents, "Count"));

            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = ExcelBridgeSupport.Get(vbComponents, "Item", index);
                    if (component is null)
                    {
                        continue;
                    }

                    var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
                    if (type != ComponentTypeDocument)
                    {
                        continue;
                    }

                    var name = (string?)ExcelBridgeSupport.Get(component, "Name");
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    var sourcePath = FindDocumentModuleSource(args.WorkbookDir, name);
                    if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                    {
                        continue;
                    }

                    var sourceContent = File.ReadAllText(sourcePath, Encoding.UTF8);
                    sourceContent = VbaSourceHelper.NormalizeDocumentModuleContent(sourceContent);

                    if (!string.IsNullOrWhiteSpace(args.WorkbookDir))
                    {
                        var desiredAnnotation = VbaSourceHelper.GetFolderAnnotationForPath(args.WorkbookDir, sourcePath);
                        sourceContent = VbaSourceHelper.UpdateFolderAnnotationText(sourceContent, args.FolderAnnotation, desiredAnnotation);
                    }

                    object? codeModule = null;
                    try
                    {
                        codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
                        if (codeModule is null)
                        {
                            continue;
                        }

                        var lineCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(codeModule, "CountOfLines"));
                        if (lineCount > 0)
                        {
                            ExcelBridgeSupport.InvokeViaDynamic(codeModule, "DeleteLines", 1, lineCount);
                        }

                        if (!string.IsNullOrWhiteSpace(sourceContent))
                        {
                            ExcelBridgeSupport.InvokeViaDynamic(codeModule, "InsertLines", 1, sourceContent);
                        }
                    }
                    finally
                    {
                        ExcelBridgeSupport.ReleaseComObject(codeModule);
                    }

                    updated++;
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        catch
        {
            // best-effort document module update
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbComponents);
            ExcelBridgeSupport.ReleaseComObject(vbProject);
        }

        return updated;
    }

    private static string? FindDocumentModuleSource(string workbookDir, string componentName)
    {
        if (!Directory.Exists(workbookDir))
        {
            return null;
        }

        var exactMatch = Path.Combine(workbookDir, componentName + ".bas");
        if (File.Exists(exactMatch))
        {
            return exactMatch;
        }

        foreach (var file in Directory.GetFiles(workbookDir, "*.bas", SearchOption.AllDirectories))
        {
            if (string.Equals(Path.GetFileNameWithoutExtension(file), componentName, StringComparison.OrdinalIgnoreCase))
            {
                return file;
            }
        }

        return null;
    }

    private static (string Id, string Path) CreateBackup(object workbook, string backupRoot, string workbookPath)
    {
        if (!Directory.Exists(backupRoot))
        {
            Directory.CreateDirectory(backupRoot);
        }

        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fff", CultureInfo.InvariantCulture);
        var suffix = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture).Substring(0, 6);
        var backupId = $"{timestamp}-push-{suffix}";
        var backupDir = Path.Combine(backupRoot, backupId);
        Directory.CreateDirectory(backupDir);

        var backupFileName = Path.GetFileName(workbookPath);
        var backupFilePath = Path.Combine(backupDir, backupFileName);

        ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", backupFilePath);

        var metadata = new Dictionary<string, object?>
        {
            ["id"] = backupId,
            ["created_at"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
            ["reason"] = "before-push",
            ["original_workbook_path"] = Path.GetFullPath(workbookPath),
            ["backup_file_path"] = backupFileName,
        };
        var metadataJson = JsonSerializer.Serialize(metadata, IndentedJsonOptions);
        File.WriteAllText(Path.Combine(backupDir, "metadata.json"), metadataJson + "\n",
            new System.Text.UTF8Encoding(false));

        return (backupId, backupFilePath);
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

    private static List<string> GetSourceUserFormNames(List<DiscoveredSourceFile> sourceFiles)
    {
        var names = new List<string>();
        foreach (var file in sourceFiles)
        {
            if (file.Kind == "form" && file.Extension == ".frm")
            {
                if (!string.IsNullOrWhiteSpace(file.ModuleName))
                {
                    names.Add(file.ModuleName);
                }
            }
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

    internal static BridgeResponse BuildCompileFailureResponse(
        BridgeRequest request,
        PushCommandArguments args,
        string workbookPath,
        bool sessionAttached,
        string sessionMode,
        WorkerInvocationResult invocation)
    {
        var message = invocation.Dialog is not null
            ? DialogMessage(invocation.Dialog)
            : invocation.TimedOut
                ? "VBE Compile timed out."
                : invocation.Result?.Error?.Message ?? "VBE Compile failed.";
        var dirty = sessionAttached;
        var extensions = new Dictionary<string, object?>
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
                ["dirty"] = dirty,
                ["save_required"] = dirty,
                ["live_newer_than_disk"] = dirty,
                ["mode"] = sessionMode,
                ["source_of_truth"] = dirty ? "live_workbook" : "saved_workbook",
            },
            ["workbook"] = new Dictionary<string, object?>
            {
                ["path"] = workbookPath,
                ["session"] = sessionAttached,
                ["session_mode"] = sessionMode,
                ["session_requested"] = sessionAttached && args.UseSession,
                ["auto_session"] = sessionAttached && !args.UseSession,
                ["saved"] = false,
                ["dirty"] = dirty,
                ["needs_save"] = dirty,
            },
            ["push_diagnostic"] = new
            {
                kind = invocation.TimedOut ? "timeout" : "compile",
                dialog = invocation.Dialog,
                dialogs = invocation.Dialogs,
                worker = new { pid = invocation.WorkerProcessId, completed = invocation.Result?.Completed ?? false, timed_out = invocation.TimedOut },
            },
        };

        if (dirty)
        {
            extensions["warnings"] = new[]
            {
                new
                {
                    code = "save_required",
                    message = "Source files were imported into the live workbook before VBA compile failed. Inspect or revert the session before saving.",
                },
            };
        }

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
            Logs = ["VBE Compile failed: " + message],
            Extensions = extensions,
        };
    }

    internal static TimeSpan ResolveCompileTimeout(BridgeRequest request)
    {
        if (request.TimeoutMs is > 1000)
        {
            return TimeSpan.FromMilliseconds(request.TimeoutMs.Value - 1000);
        }

        if (request.TimeoutMs is > 0)
        {
            return TimeSpan.FromMilliseconds(Math.Max(1, request.TimeoutMs.Value));
        }

        return TimeSpan.FromMinutes(5);
    }

    private static string DialogMessage(DialogSnapshot dialog)
    {
        var lines = dialog.Text.Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
        return lines.Length > 0 ? string.Join(Environment.NewLine, lines) : dialog.Title;
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
