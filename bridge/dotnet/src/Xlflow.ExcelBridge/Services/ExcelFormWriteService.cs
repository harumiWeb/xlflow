using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize COM failures into structured responses.")]
public sealed class ExcelFormWriteService : IFormWriteService
{
    public BridgeResponse Execute(BridgeRequest request, FormWriteCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            var formName = ReadFormName(args.SpecJson64);
            return args.Phase == "prepare"
                ? PrepareWorkbookCopy(request, args, formName)
                : ImportCompiledForm(request, args, formName, cancellationToken);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            return Failed(request, "session_required", ex.Message, "open_workbook", "xlflow");
        }
        catch (Exception ex)
        {
            return Failed(request,
                args.Action == "apply" ? "form_apply_failed" : "form_build_failed",
                ExcelBridgeSupport.FormatExceptionDetail(ex),
                args.Phase == "prepare" ? "prepare_workbook_copy" : "import_compiled_form",
                "xlflow-excel-bridge");
        }
    }

    internal static void ReplaceFormWithRollback(
        bool hadExistingForm,
        Action removeExisting,
        Action importCompiled,
        Action removePartialImport,
        Action restoreExisting)
    {
        try
        {
            if (hadExistingForm)
            {
                removeExisting();
            }
            importCompiled();
        }
        catch (Exception importError)
        {
            try
            {
                removePartialImport();
                if (hadExistingForm)
                {
                    restoreExisting();
                }
            }
            catch (Exception rollbackError)
            {
                throw new InvalidOperationException(
                    $"form import failed and rollback failed: {ExcelBridgeSupport.FormatExceptionDetail(rollbackError)}",
                    new AggregateException(importError, rollbackError));
            }

            throw;
        }
    }

    private static BridgeResponse PrepareWorkbookCopy(BridgeRequest request, FormWriteCommandArguments args, string formName)
    {
        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";
        try
        {
            (excel, workbook, sessionAttached, sessionMode) = OpenTargetWorkbook(args);
            var parent = Path.GetDirectoryName(args.PreparedWorkbookPath);
            if (!string.IsNullOrWhiteSpace(parent))
            {
                Directory.CreateDirectory(parent);
            }

            ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", args.PreparedWorkbookPath);
            return BridgeResponse.Ok(request, BuildExtensions(args, formName, sessionAttached, sessionMode, saved: false, dirty: ReadDirty(workbook))) with
            {
                Logs =
                [
                    sessionAttached ? $"attached to xlflow session ({sessionMode})" : $"opened workbook {ExcelBridgeSupport.NormalizePath(args.WorkbookPath)}",
                    $"saved a private workbook copy for shared compiler {args.PreparedWorkbookPath}",
                    "left the source workbook unsaved",
                ],
            };
        }
        finally
        {
            ReleaseTarget(excel, workbook, sessionAttached);
        }
    }

    private static BridgeResponse ImportCompiledForm(BridgeRequest request, FormWriteCommandArguments args, string formName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var privateDir = Path.Combine(Path.GetTempPath(), "xlflow-form-write-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(privateDir);
        object? stagedExcel = null;
        object? stagedWorkbook = null;
        object? stagedProject = null;
        object? stagedComponents = null;
        object? stagedComponent = null;
        object? targetExcel = null;
        object? targetWorkbook = null;
        object? targetProject = null;
        object? targetComponents = null;
        var targetSessionAttached = false;
        var targetSessionMode = "none";
        var saved = false;

        try
        {
            var staged = ExcelBridgeSupport.OpenWorkbookDirect(args.CompiledWorkbookPath, visible: false);
            stagedExcel = staged.Excel;
            stagedWorkbook = staged.Workbook;
            stagedProject = ExcelBridgeSupport.Get(stagedWorkbook, "VBProject")
                ?? throw new InvalidOperationException("VBProject access is denied for the compiled workbook.");
            stagedComponents = ExcelBridgeSupport.Get(stagedProject, "VBComponents")
                ?? throw new InvalidOperationException("VBComponents is unavailable in the compiled workbook.");
            stagedComponent = FindComponent(stagedComponents, formName)
                ?? throw new InvalidOperationException($"compiled workbook does not contain UserForm '{formName}'.");
            var compiledFormPath = Path.Combine(privateDir, "compiled.frm");
            ExcelBridgeSupport.InvokeViaDynamic(stagedComponent, "Export", compiledFormPath);
            ExcelBridgeSupport.ReleaseComObject(stagedComponent);
            stagedComponent = null;
            ExcelBridgeSupport.ReleaseComObject(stagedComponents);
            stagedComponents = null;
            ExcelBridgeSupport.ReleaseComObject(stagedProject);
            stagedProject = null;
            ExcelBridgeSupport.CloseWorkbookAndQuitApplication(stagedWorkbook, stagedExcel);
            stagedWorkbook = null;
            stagedExcel = null;

            (targetExcel, targetWorkbook, targetSessionAttached, targetSessionMode) = OpenTargetWorkbook(args);
            targetProject = ExcelBridgeSupport.Get(targetWorkbook, "VBProject")
                ?? throw new InvalidOperationException("VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.");
            targetComponents = ExcelBridgeSupport.Get(targetProject, "VBComponents")
                ?? throw new InvalidOperationException("VBComponents is unavailable.");
            var existing = FindComponent(targetComponents, formName);
            var hadExisting = existing is not null;
            if (args.Action == "apply" && !hadExisting)
            {
                ExcelBridgeSupport.ReleaseComObject(existing);
                throw new InvalidOperationException($"UserForm '{formName}' does not exist; form apply requires an existing form.");
            }
            if (args.Action == "build" && hadExisting && !args.Overwrite)
            {
                ExcelBridgeSupport.ReleaseComObject(existing);
                throw new InvalidOperationException($"UserForm '{formName}' already exists; pass --overwrite to replace it.");
            }
            var backupFormPath = Path.Combine(privateDir, "existing.frm");
            if (existing is not null)
            {
                ExcelBridgeSupport.InvokeViaDynamic(existing, "Export", backupFormPath);
            }
            ExcelBridgeSupport.ReleaseComObject(existing);

            ReplaceFormWithRollback(
                hadExisting,
                removeExisting: () => RemoveComponentIfPresent(targetComponents, formName),
                importCompiled: () =>
                {
                    ImportAndValidate(targetComponents, compiledFormPath, formName);
                    if (!args.NoSave)
                    {
                        ExcelBridgeSupport.InvokeViaDynamic(targetWorkbook, "Save");
                        saved = true;
                    }
                },
                removePartialImport: () => RemoveComponentIfPresent(targetComponents, formName),
                restoreExisting: () => ImportAndValidate(targetComponents, backupFormPath, formName));

            return BridgeResponse.Ok(request, BuildExtensions(
                args,
                formName,
                targetSessionAttached,
                targetSessionMode,
                saved,
                ReadDirty(targetWorkbook))) with
            {
                Logs =
                [
                    targetSessionAttached ? $"attached to xlflow session ({targetSessionMode})" : $"opened workbook {ExcelBridgeSupport.NormalizePath(args.WorkbookPath)}",
                    $"imported compiler-generated UserForm {formName}",
                    args.NoSave ? "left workbook changes unsaved" : "saved workbook after form import",
                    "kept generated form artifacts private",
                ],
            };
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("rollback failed", StringComparison.OrdinalIgnoreCase))
        {
            return FormImportRollbackFailure(request, args, targetSessionAttached, targetSessionMode, ex.Message);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(targetComponents);
            ExcelBridgeSupport.ReleaseComObject(targetProject);
            ReleaseTarget(targetExcel, targetWorkbook, targetSessionAttached);
            ExcelBridgeSupport.ReleaseComObject(stagedComponent);
            ExcelBridgeSupport.ReleaseComObject(stagedComponents);
            ExcelBridgeSupport.ReleaseComObject(stagedProject);
            if (stagedWorkbook is not null || stagedExcel is not null)
            {
                ExcelBridgeSupport.CloseWorkbookAndQuitApplication(stagedWorkbook, stagedExcel);
            }
            try
            {
                Directory.Delete(privateDir, recursive: true);
            }
            catch (IOException)
            {
                // Best-effort cleanup of compiler interchange files.
            }
        }
    }

    internal static BridgeResponse FormImportRollbackFailure(
        BridgeRequest request, FormWriteCommandArguments args, bool sessionAttached, string sessionMode, string message)
    {
        var response = Failed(request, "form_import_rollback_failed", message, "rollback_form_import", "xlflow-excel-bridge");
        if (sessionAttached)
        {
            ExcelBridgeSupport.MarkSessionPoisoned(args.MetadataPath, args.WorkbookPath,
                "UserForm import and rollback failed; the live workbook may be partially modified.", "", "form " + args.Action,
                discardUnsavedChanges: true);
            var session = ExcelBridgeSupport.BuildSessionPayload(args.WorkbookPath, true, sessionMode, true, saveRequired: false);
            session["poisoned"] = true;
            session["discard_required"] = true;
            response.Extensions["session"] = session;
            response.Extensions["warnings"] = new List<Dictionary<string, string>>
            {
                new()
                {
                    ["code"] = "form_import_rollback_failed",
                    ["message"] = string.Equals(sessionMode, "external", StringComparison.OrdinalIgnoreCase)
                        ? "Close the external workbook in Excel without saving, then start a fresh xlflow session."
                        : "Run `xlflow session stop --discard --json`, then start a fresh session before retrying.",
                },
            };
        }
        return response;
    }

    private static Dictionary<string, object?> BuildExtensions(
        FormWriteCommandArguments args,
        string formName,
        bool sessionAttached,
        string sessionMode,
        bool saved,
        bool? dirty)
    {
        var needsSave = sessionAttached && !saved && dirty != false;
        return new Dictionary<string, object?>
        {
            ["target"] = new Dictionary<string, object?>
            {
                ["kind"] = sessionAttached ? "live_session" : "file",
                ["path"] = ExcelBridgeSupport.NormalizePath(args.WorkbookPath),
            },
            ["session"] = ExcelBridgeSupport.BuildSessionPayload(args.WorkbookPath, sessionAttached, sessionMode, dirty, needsSave),
            ["workbook"] = new Dictionary<string, object?>
            {
                ["path"] = ExcelBridgeSupport.NormalizePath(args.WorkbookPath),
                ["session"] = sessionAttached,
                ["session_mode"] = sessionMode,
                ["session_requested"] = sessionAttached && args.UseSession,
                ["auto_session"] = sessionAttached && !args.UseSession,
                ["saved"] = saved,
                ["dirty"] = dirty,
                ["needs_save"] = needsSave,
            },
            ["forms"] = new Dictionary<string, object?>
            {
                ["name"] = formName,
                ["action"] = args.Action,
                ["spec_path"] = args.SpecPath,
                ["overwrite"] = args.Overwrite,
                ["source_synced"] = false,
                ["artifacts_private"] = true,
            },
        };
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenTargetWorkbook(FormWriteCommandArguments args)
    {
        if (args.UseSession || ExcelBridgeSupport.SessionMetadataMatchesWorkbook(args.MetadataPath, args.WorkbookPath))
        {
            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession);
            return (attachment.Excel, attachment.Workbook, true, attachment.SessionMode);
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(args.WorkbookPath, args.Visible);
        return (direct.Excel, direct.Workbook, false, "none");
    }

    private static void ReleaseTarget(object? excel, object? workbook, bool sessionAttached)
    {
        if (sessionAttached)
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
        else
        {
            ExcelBridgeSupport.CloseWorkbookAndQuitApplication(workbook, excel);
        }
    }

    private static object? FindComponent(object components, string name)
    {
        var count = Convert.ToInt32(ExcelBridgeSupport.Get(components, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            object? candidate = null;
            try
            {
                candidate = ExcelBridgeSupport.Get(components, "Item", index);
                var candidateName = candidate is null ? "" : ExcelBridgeSupport.GetString(candidate, "Name") ?? "";
                if (string.Equals(candidateName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }
            catch
            {
                ExcelBridgeSupport.ReleaseComObject(candidate);
                throw;
            }
            ExcelBridgeSupport.ReleaseComObject(candidate);
        }
        return null;
    }

    private static void RemoveComponentIfPresent(object components, string name)
    {
        var component = FindComponent(components, name);
        if (component is null)
        {
            return;
        }
        try
        {
            ExcelBridgeSupport.InvokeViaDynamic(components, "Remove", component);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
        }
    }

    private static void ImportAndValidate(object components, string formPath, string expectedName)
    {
        var imported = ExcelBridgeSupport.InvokeViaDynamic(components, "Import", formPath)
            ?? throw new InvalidOperationException($"Excel failed to import UserForm '{expectedName}'.");
        try
        {
            var actualName = ExcelBridgeSupport.GetString(imported, "Name") ?? "";
            if (!string.Equals(actualName, expectedName, StringComparison.OrdinalIgnoreCase))
            {
                ExcelBridgeSupport.InvokeViaDynamic(components, "Remove", imported);
                throw new InvalidOperationException($"imported UserForm name '{actualName}' did not match '{expectedName}'.");
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(imported);
        }
    }

    private static bool? ReadDirty(object workbook)
    {
        return ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirty) ? dirty : null;
    }

    internal static string ReadFormName(string encoded)
    {
        try
        {
            using var document = JsonDocument.Parse(Convert.FromBase64String(encoded));
            var form = document.RootElement.GetProperty("form");
            var name = form.GetProperty("name").GetString();
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("form.name is required.");
            }
            if (!string.Equals(Path.GetFileName(name), name, StringComparison.Ordinal)
                || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new InvalidOperationException("form.name must be a valid component name.");
            }
            return name;
        }
        catch (Exception ex) when (ex is FormatException or JsonException or KeyNotFoundException or InvalidOperationException)
        {
            throw new InvalidOperationException("invalid form spec payload: " + ex.Message, ex);
        }
    }

    private static BridgeResponse Failed(BridgeRequest request, string code, string message, string phase, string source)
    {
        return BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: phase, Source: source));
    }
}
