using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelFormWriteService : IFormWriteService
{
    private const int ComponentTypeForm = 3;

    private static readonly Dictionary<string, string> TypeProgIdMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["label"] = "Forms.Label.1",
        ["textbox"] = "Forms.TextBox.1",
        ["combobox"] = "Forms.ComboBox.1",
        ["listbox"] = "Forms.ListBox.1",
        ["commandbutton"] = "Forms.CommandButton.1",
        ["checkbox"] = "Forms.CheckBox.1",
        ["optionbutton"] = "Forms.OptionButton.1",
        ["togglebutton"] = "Forms.ToggleButton.1",
        ["spinbutton"] = "Forms.SpinButton.1",
        ["scrollbar"] = "Forms.ScrollBar.1",
        ["image"] = "Forms.Image.1",
        ["frame"] = "Forms.Frame.1",
        ["multipage"] = "Forms.MultiPage.1",
        ["page"] = "Forms.Page.1",
        ["tabstrip"] = "Forms.TabStrip.1",
    };

    public BridgeResponse Execute(BridgeRequest request, FormWriteCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var spec = DecodeSpec(args.SpecJson64);
            var warnings = CollectContractWarnings(spec);

            var openResult = OpenWorkbookForWrite(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible);
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            object vbProject;
            try
            {
                vbProject = ExcelBridgeSupport.Get(workbook, "VBProject")
                    ?? throw new InvalidOperationException("VBProject access is denied.");
            }
            catch
            {
                return Failed(request, "vbproject_access_denied", "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center.", "open_workbook", "Excel");
            }

            object? builtComponent = null;
            try
            {
                builtComponent = args.Action == "build"
                    ? BuildForm(vbProject, workbook, spec, args)
                    : ApplyForm(vbProject, spec, args);

                var saved = false;
                if (sessionAttached)
                {
                    if (!args.NoSave)
                    {
                        ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                        saved = true;
                    }
                }
                else
                {
                    ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                    saved = true;
                }

                var sourceArtifacts = ExportBuiltUserFormArtifacts(builtComponent!, args);
                var dirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
                var dirty = sessionAttached ? dirtyKnown ? dirtyState : true : false;
                var needsSave = sessionAttached && !saved ? dirty : false;

                if (needsSave)
                {
                    warnings.Add(NewMessage("save_required", $"The live workbook is newer than disk after `form {args.Action}`. Run `xlflow save --session` before relying on disk-backed inspect, pull, or source review."));
                }

                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Logs =
                    [
                        sessionAttached ? $"attached to xlflow session ({sessionMode})" : $"opened workbook {workbookPath}",
                        $"{args.Action} form {spec.Form.Name} from {args.SpecPath}",
                        $"synchronized UserForm source artifacts for {spec.Form.Name}",
                    ],
                    Extensions = BuildResponseExtensions(spec, args, sourceArtifacts, workbookPath, sessionAttached, sessionMode, saved, dirty, needsSave, warnings),
                };
            }
            catch (InvalidOperationException ex)
            {
                var (code, message) = ClassifyFormWriteError(args.Action, ex.Message);
                return Failed(request, code, message, "write_designer", "xlflow-excel-bridge");
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(builtComponent);
                ExcelBridgeSupport.ReleaseComObject(vbProject);
            }
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            return Failed(request, "session_required", ex.Message, "open_workbook", "xlflow");
        }
        catch (Exception ex)
        {
            return Failed(request, args.Action == "apply" ? "form_apply_failed" : "form_build_failed", ExcelBridgeSupport.FormatExceptionDetail(ex), "write_designer", "xlflow-excel-bridge");
        }
        finally
        {
            if (sessionAttached)
            {
                ExcelBridgeSupport.ReleaseComObject(workbook);
                ExcelBridgeSupport.ReleaseComObject(excel);
            }
            else
            {
                CloseComInstance(workbook, excel);
            }
        }
    }

    private static Dictionary<string, object?> BuildResponseExtensions(
        FormWriteSpec spec,
        FormWriteCommandArguments args,
        SourceArtifacts sourceArtifacts,
        string workbookPath,
        bool sessionAttached,
        string sessionMode,
        bool saved,
        bool dirty,
        bool needsSave,
        List<Dictionary<string, string>> warnings)
    {
        return new Dictionary<string, object?>
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
                ["dirty"] = dirty,
                ["needs_save"] = needsSave,
            },
            ["forms"] = new Dictionary<string, object?>
            {
                ["name"] = spec.Form.Name,
                ["basis"] = spec.Basis,
                ["action"] = args.Action,
                ["coordinate_system"] = spec.CoordinateSystem ?? "",
                ["control_count"] = spec.Controls?.Count ?? 0,
                ["spec_path"] = args.SpecPath,
                ["overwrite"] = args.Overwrite,
                ["source_synced"] = true,
                ["caption"] = spec.Form.Caption ?? "",
                ["source_artifacts"] = new Dictionary<string, object?>
                {
                    ["form_path"] = sourceArtifacts.FormPath,
                    ["frx_path"] = sourceArtifacts.FrxPath,
                    ["code_path"] = sourceArtifacts.CodePath ?? "",
                },
            },
            ["warnings"] = warnings,
            ["hints"] = new List<Dictionary<string, string>>
            {
                NewMessage("userform_review_commands", $"Review the result with `xlflow inspect form {spec.Form.Name} --designer --json` or `xlflow form export-image {spec.Form.Name} --out <path>`.")
            },
        };
    }

    private static BridgeResponse Failed(BridgeRequest request, string code, string message, string phase, string source)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: code,
            Message: message,
            Phase: phase,
            Source: source));
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForWrite(string workbookPath, string metadataPath, bool useSession, bool visible)
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
                // fall through
            }
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, "none");
    }

    private static FormWriteSpec DecodeSpec(string encoded)
    {
        try
        {
            var bytes = Convert.FromBase64String(encoded);
            var json = Encoding.UTF8.GetString(bytes);
            var spec = JsonSerializer.Deserialize<FormWriteSpec>(json);
            if (spec is null)
            {
                throw new InvalidOperationException("decoded form spec was empty");
            }
            if (spec.Form.Build?.ClientWidth is not null || spec.Form.Build?.ClientHeight is not null)
            {
                throw new InvalidOperationException("form.build.clientWidth/clientHeight are supported only by pure-Go generation, not Excel Designer authoring.");
            }
            ValidateControlSpecs(spec.Controls);
            ValidateDeclaredSelectionIndices(spec);
            return spec;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("invalid form spec payload: " + ex.Message, ex);
        }
    }

    private static List<Dictionary<string, string>> CollectContractWarnings(FormWriteSpec spec)
    {
        var warnings = new List<Dictionary<string, string>>();
        var hasFormSizeExpectation =
            spec.Form.Width is not null ||
            spec.Form.Height is not null ||
            spec.Form.Build?.Width is not null ||
            spec.Form.Build?.Height is not null ||
            spec.Form.Observed?.Width is not null ||
            spec.Form.Observed?.Height is not null;
        if (hasFormSizeExpectation)
        {
            warnings.Add(NewMessage("best_effort_form_size", "Form-level width/height are best-effort in Designer build and may not round-trip through Excel VBIDE Designer APIs. Field scope: form.observed.width, form.observed.height, form.build.width, form.build.height.", "form.observed.width"));
        }

        var listStateControls = spec.Controls
            .Where(control => control.Type is not null &&
                (control.Type.Equals("combobox", StringComparison.OrdinalIgnoreCase) || control.Type.Equals("listbox", StringComparison.OrdinalIgnoreCase)) &&
                ((control.List?.Count ?? 0) > 0 || control.SelectedIndex is not null || (control.Observed?.List?.Count ?? 0) > 0 || control.Observed?.SelectedIndex is not null))
            .Select(control => control.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (listStateControls.Length > 0)
        {
            warnings.Add(NewMessage("best_effort_list_state", "Design-time ComboBox/ListBox list and selectedIndex are best-effort during build and should be treated as observed-only for round-trip expectations. Field scope includes controls[*].list and controls[*].selectedIndex. Controls: " + string.Join(", ", listStateControls) + ".", "controls[*].selectedIndex"));
        }
        return warnings;
    }

    private static object BuildForm(object vbProject, object workbook, FormWriteSpec spec, FormWriteCommandArguments args)
    {
        object? existing = null;
        object? component = null;
        string? existingCodeText = null;
        string? restoreDir = null;
        string? restorePath = null;
        var removedExisting = false;

        try
        {
            existing = GetComponentByName(vbProject, spec.Form.Name);
            if (existing is not null)
            {
                if (ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(existing, "Type")) != ComponentTypeForm)
                {
                    throw new InvalidOperationException($"form_already_exists: component '{spec.Form.Name}' exists but is not a UserForm.");
                }
                if (!args.Overwrite)
                {
                    throw new InvalidOperationException($"form_already_exists: UserForm '{spec.Form.Name}' already exists.");
                }

                restoreDir = NewRestoreDirectory();
                existingCodeText = ReadCodeModuleText(existing);
                restorePath = ExportComponentBackup(existing, restoreDir);
                RemoveComponentByName(vbProject, spec.Form.Name);
                removedExisting = true;
                if (args.NoSave)
                {
                    throw new InvalidOperationException("designer_write_failed: overwrite requires an intermediate workbook save, but save is disabled for this command.");
                }
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
            }

            component = NewUserFormComponent(vbProject, spec.Form.Name);
            PopulateFormComponent(component, spec);
            if (VbaSourceHelper.IsSidecarMode(args.CodeSource))
            {
                SyncUserFormCodeBehind(component, args.FormsDir, existingCodeText);
            }
            else if (!string.IsNullOrWhiteSpace(existingCodeText))
            {
                WriteCodeModuleText(component, existingCodeText);
            }
            return component;
        }
        catch
        {
            if (removedExisting)
            {
                try { RemoveComponentInstance(vbProject, component); } catch { }
                if (!string.IsNullOrWhiteSpace(restorePath))
                {
                    ImportComponentBackup(vbProject, restorePath, spec.Form.Name);
                    ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                }
            }
            throw;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(existing);
            if (!string.IsNullOrWhiteSpace(restoreDir) && Directory.Exists(restoreDir))
            {
                try { Directory.Delete(restoreDir, true); } catch { }
            }
        }
    }

    private static object ApplyForm(object vbProject, FormWriteSpec spec, FormWriteCommandArguments args)
    {
        var component = GetComponentByName(vbProject, spec.Form.Name);
        if (component is null || ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type")) != ComponentTypeForm)
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            throw new InvalidOperationException($"form_not_found: UserForm '{spec.Form.Name}' was not found in the workbook.");
        }

        try
        {
            object? designer = null;
            try
            {
                designer = ExcelBridgeSupport.Get(component, "Designer");
                if (designer is null)
                {
                    throw new InvalidOperationException($"designer_write_failed: failed to access Designer for '{spec.Form.Name}'.");
                }
                PrepareApplySelectionIndices(designer, spec);
                ClearDesignerControls(designer);
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(designer);
            }
            PopulateFormComponent(component, spec);
            if (VbaSourceHelper.IsSidecarMode(args.CodeSource))
            {
                SyncUserFormCodeBehind(component, args.FormsDir, null);
            }
            return component;
        }
        catch
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            throw;
        }
    }

    private static void PopulateFormComponent(object component, FormWriteSpec spec)
    {
        object? designer = null;
        try
        {
            designer = ExcelBridgeSupport.Get(component, "Designer");
            if (designer is null)
            {
                throw new InvalidOperationException($"designer_write_failed: failed to access Designer for '{spec.Form.Name}'.");
            }
            SetDesignerFormProperties(designer, component, spec.Form);
            AddDesignerControls(designer, GetRootControls(spec), spec.Controls, "UserForm");
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(designer);
        }
    }

    private static List<FormWriteControlSpec> GetRootControls(FormWriteSpec spec)
    {
        var roots = spec.Controls.Where(control => string.IsNullOrWhiteSpace(control.ParentId)).ToList();
        if (roots.Count == 0)
        {
            roots = spec.Controls.ToList();
        }
        roots.Sort((left, right) => (left.ZIndex ?? int.MaxValue).CompareTo(right.ZIndex ?? int.MaxValue));
        return roots;
    }

    private static FormWriteControlSpec[] GetChildControls(IEnumerable<FormWriteControlSpec> allControls, FormWriteControlSpec parent)
    {
        var id = string.IsNullOrWhiteSpace(parent.Id) ? parent.Name : parent.Id;
        var explicitChildren = allControls
            .Where(control => !string.IsNullOrWhiteSpace(control.ParentId) && string.Equals(control.ParentId, id, StringComparison.Ordinal))
            .OrderBy(control => control.ZIndex ?? int.MaxValue)
            .ToArray();
        return explicitChildren.Length > 0
            ? explicitChildren
            : (parent.Controls ?? []).OrderBy(control => control.ZIndex ?? int.MaxValue).ToArray();
    }

    private static void AddDesignerControl(
        object parent,
        FormWriteControlSpec controlSpec,
        IReadOnlyList<FormWriteControlSpec> allControls,
        string parentType)
    {
        object? collection = null;
        object? control = null;
        try
        {
            var isPage = IsControlType(controlSpec, "Page");
            if (isPage && !IsControlType(parentType, "MultiPage"))
            {
                throw new InvalidOperationException("designer_write_failed: Page controls must be added through a MultiPage Pages collection.");
            }
            if (!isPage && (IsControlType(parentType, "MultiPage") || IsControlType(parentType, "TabStrip")))
            {
                throw new InvalidOperationException($"designer_write_failed: {parentType} cannot contain '{controlSpec.Type}' through its Controls collection.");
            }

            collection = ExcelBridgeSupport.Get(parent, isPage ? "Pages" : "Controls");
            if (collection is null)
            {
                throw new InvalidOperationException($"designer_write_failed: failed to access {(isPage ? "Pages" : "Controls")} collection.");
            }
            control = isPage
                ? ExcelBridgeSupport.InvokeMethod(collection, "Add", controlSpec.Name)
                : ExcelBridgeSupport.InvokeMethod(collection, "Add", ResolveProgId(controlSpec), controlSpec.Name, true);
            if (control is null)
            {
                throw new InvalidOperationException($"designer_write_failed: failed to add control '{controlSpec.Name}'.");
            }
            VerifyRequiredStringProperty(control, "Name", controlSpec.Name);
            SetDesignerControlProperties(control, controlSpec);

            var children = GetChildControls(allControls, controlSpec);
            if (IsControlType(controlSpec, "MultiPage"))
            {
                if (children.Any(child => !IsControlType(child, "Page")))
                {
                    throw new InvalidOperationException("designer_write_failed: MultiPage children must all be Page controls.");
                }
                ClearMultiPagePages(control);
            }
            else if (IsControlType(controlSpec, "TabStrip") && children.Length > 0)
            {
                throw new InvalidOperationException("designer_write_failed: TabStrip tabs are authored through the tabs array, not child controls.");
            }

            AddDesignerControls(control, children, allControls, controlSpec.Type);

            var tabSpecs = controlSpec.Tabs ?? controlSpec.Observed?.Tabs;
            if (IsControlType(controlSpec, "TabStrip") && tabSpecs is not null)
            {
                ReplaceTabStripTabs(control, tabSpecs);
            }

            ApplySelectedIndex(control, controlSpec, children.Length);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(control);
            ExcelBridgeSupport.ReleaseComObject(collection);
        }
    }

    private static void AddDesignerControls(object parent, IReadOnlyList<FormWriteControlSpec> siblings,
        IReadOnlyList<FormWriteControlSpec> allControls, string parentType)
    {
        foreach (var sibling in siblings)
        {
            AddDesignerControl(parent, sibling, allControls, parentType);
        }
        var tabOrder = siblings.Where(sibling => (sibling.TabIndex ?? sibling.Observed?.TabIndex) is not null)
            .OrderBy(sibling => sibling.TabIndex ?? sibling.Observed?.TabIndex).ToArray();
        if (tabOrder.Length == 0)
        {
            return;
        }
        object? collection = null;
        try
        {
            collection = ExcelBridgeSupport.Get(parent, IsControlType(parentType, "MultiPage") ? "Pages" : "Controls")
                ?? throw new InvalidOperationException("designer_write_failed: tab order collection unavailable.");
            foreach (var sibling in tabOrder)
            {
                object? control = null;
                try
                {
                    control = ExcelBridgeSupport.Get(collection, "Item", sibling.Name)
                        ?? throw new InvalidOperationException($"designer_write_failed: tab order control '{sibling.Name}' unavailable.");
                    SetRequiredMember(control, "TabIndex", (sibling.TabIndex ?? sibling.Observed!.TabIndex)!.Value);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(control);
                }
            }
            foreach (var sibling in tabOrder)
            {
                object? control = null;
                try
                {
                    control = ExcelBridgeSupport.Get(collection, "Item", sibling.Name);
                    var expected = sibling.TabIndex ?? sibling.Observed!.TabIndex;
                    if (control is null || ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(control, "TabIndex")) != expected)
                    {
                        throw new InvalidOperationException($"designer_write_failed: TabIndex did not persist for '{sibling.Name}'.");
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(control);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(collection);
        }
    }

    private static void ValidateControlSpecs(IEnumerable<FormWriteControlSpec> controls)
    {
        foreach (var control in controls)
        {
            if (string.IsNullOrWhiteSpace(control.Name))
            {
                throw new InvalidOperationException("designer_write_failed: every control requires a non-empty name.");
            }
            if (IsControlType(control, "Page"))
            {
                control.ResolvedPageProperties = ResolvePageProperties(control);
            }
            var tabs = control.Tabs ?? control.Observed?.Tabs;
            if (tabs is not null)
            {
                for (var index = 0; index < tabs.Count; index++)
                {
                    if (string.IsNullOrWhiteSpace(tabs[index].Name))
                    {
                        throw new InvalidOperationException($"designer_write_failed: tabs[{index}].name is required.");
                    }
                }
            }
            if (control.Controls is not null)
            {
                ValidateControlSpecs(control.Controls);
            }
        }
    }

    private static void ValidateDeclaredSelectionIndices(FormWriteSpec spec)
    {
        foreach (var control in EnumerateControls(spec.Controls))
        {
            var selectedIndex = control.SelectedIndex ?? control.Observed?.SelectedIndex;
            if (selectedIndex is null)
            {
                continue;
            }
            if (IsControlType(control, "MultiPage"))
            {
                ValidateSelectedIndex(control.Name, "MultiPage", selectedIndex.Value, GetChildControls(spec.Controls, control).Length);
            }
            else if (IsControlType(control, "TabStrip"))
            {
                var tabs = control.Tabs ?? control.Observed?.Tabs;
                ValidateSelectedIndex(control.Name, "TabStrip", selectedIndex.Value, tabs?.Count);
            }
        }
    }

    private static void PrepareApplySelectionIndices(object designer, FormWriteSpec spec)
    {
        foreach (var control in EnumerateControls(spec.Controls))
        {
            if (!IsControlType(control, "TabStrip"))
            {
                continue;
            }

            int? tabCount = control.Tabs?.Count;
            if (control.Tabs is null)
            {
                var existing = FindTabStripState(designer, control.Name);
                if (existing is not null)
                {
                    tabCount = existing.Tabs.Count;
                    control.Observed ??= new FormWriteObservedControlSpec();
                    control.Observed.Tabs = existing.Tabs;
                    if (control.SelectedIndex is null)
                    {
                        control.Observed.SelectedIndex = existing.SelectedIndex;
                    }
                }
                else
                {
                    tabCount = control.Observed?.Tabs?.Count;
                }
            }

            var selectedIndex = control.SelectedIndex ?? control.Observed?.SelectedIndex;
            if (selectedIndex is not null)
            {
                ValidateSelectedIndex(control.Name, "TabStrip", selectedIndex.Value, tabCount);
            }
        }
    }

    private static IEnumerable<FormWriteControlSpec> EnumerateControls(IEnumerable<FormWriteControlSpec> controls)
    {
        var seen = new HashSet<FormWriteControlSpec>();
        var pending = new Stack<FormWriteControlSpec>(controls.Reverse());
        while (pending.Count > 0)
        {
            var control = pending.Pop();
            if (!seen.Add(control))
            {
                continue;
            }
            yield return control;
            if (control.Controls is not null)
            {
                foreach (var child in control.Controls.AsEnumerable().Reverse())
                {
                    pending.Push(child);
                }
            }
        }
    }

    private static void ValidateSelectedIndex(string controlName, string controlType, int selectedIndex, int? itemCount)
    {
        if (selectedIndex < -1 || itemCount is not null &&
            (itemCount.Value == 0 ? selectedIndex != -1 : selectedIndex < 0 || selectedIndex >= itemCount.Value))
        {
            var allowed = itemCount switch
            {
                0 => "-1 for an empty collection",
                > 0 => $"0..{itemCount.Value - 1} for a collection with {itemCount.Value} items",
                _ => "-1 or a valid zero-based index",
            };
            throw new InvalidOperationException($"designer_write_failed: {controlType} '{controlName}' selectedIndex must be {allowed}.");
        }
    }

    private static FormWriteExistingTabStripState? FindTabStripState(object designer, string name)
    {
        object? controls = null;
        try
        {
            controls = ExcelBridgeSupport.Get(designer, "Controls");
            return controls is null ? null : FindTabStripStateInCollection(controls, name);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(controls);
        }
    }

    private static FormWriteExistingTabStripState? FindTabStripStateInCollection(object collection, string name)
    {
        var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(collection, "Count"));
        for (var index = 0; index < count; index++)
        {
            object? control = null;
            object? children = null;
            try
            {
                control = ExcelBridgeSupport.Get(collection, "Item", index);
                if (control is null)
                {
                    continue;
                }
                var controlName = ExcelBridgeSupport.GetString(control, "Name") ?? "";
                var controlType = ExcelFormInspectionService.ResolveControlType(control);
                if (string.Equals(controlName, name, StringComparison.OrdinalIgnoreCase)
                    && IsControlType(controlType, "TabStrip"))
                {
                    return ReadExistingTabStripState(control);
                }

                var childCollectionName = IsControlType(controlType, "MultiPage")
                    ? "Pages"
                    : IsControlType(controlType, "Frame") || IsControlType(controlType, "Page")
                            ? "Controls"
                            : null;
                if (childCollectionName is null)
                {
                    continue;
                }
                children = ExcelBridgeSupport.Get(control, childCollectionName);
                if (children is not null)
                {
                    var match = FindTabStripStateInCollection(children, name);
                    if (match is not null)
                    {
                        return match;
                    }
                }
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(children);
                ExcelBridgeSupport.ReleaseComObject(control);
            }
        }
        return null;
    }

    private static FormWriteExistingTabStripState ReadExistingTabStripState(object tabStrip)
    {
        object? tabs = null;
        try
        {
            tabs = ExcelBridgeSupport.Get(tabStrip, "Tabs")
                ?? throw new InvalidOperationException("designer_write_failed: existing TabStrip Tabs collection is unavailable.");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(tabs, "Count"));
            var tabSpecs = new List<FormWriteTabSpec>(count);
            for (var index = 0; index < count; index++)
            {
                object? tab = null;
                try
                {
                    tab = ExcelBridgeSupport.Get(tabs, "Item", index)
                        ?? throw new InvalidOperationException($"designer_write_failed: failed to read existing TabStrip tab at index {index}.");
                    var tabName = ExcelBridgeSupport.GetString(tab, "Name");
                    if (string.IsNullOrWhiteSpace(tabName))
                    {
                        throw new InvalidOperationException($"designer_write_failed: existing TabStrip tab at index {index} has no name.");
                    }
                    tabSpecs.Add(new FormWriteTabSpec
                    {
                        Name = tabName,
                        Caption = TryGetOptionalString(tab, "Caption"),
                        ControlTipText = TryGetOptionalString(tab, "ControlTipText"),
                        Tag = TryGetOptionalString(tab, "Tag"),
                        Accelerator = TryGetOptionalString(tab, "Accelerator"),
                        Enabled = TryGetOptionalBoolean(tab, "Enabled"),
                        Visible = TryGetOptionalBoolean(tab, "Visible"),
                    });
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(tab);
                }
            }
            var selectedIndex = count == 0
                ? -1
                : Convert.ToInt32(ExcelBridgeSupport.Get(tabStrip, "Value"), CultureInfo.InvariantCulture);
            return new FormWriteExistingTabStripState(tabSpecs, selectedIndex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(tabs);
        }
    }

    private static string? TryGetOptionalString(object target, string propertyName)
    {
        try
        {
            var value = ExcelBridgeSupport.Get(target, propertyName);
            return value is null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryGetOptionalBoolean(object target, string propertyName)
    {
        try
        {
            var value = ExcelBridgeSupport.Get(target, propertyName);
            return value is null ? null : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static FormWritePageProperties ResolvePageProperties(FormWriteControlSpec control)
    {
        var properties = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (control.Properties is not null)
        {
            foreach (var (rawName, rawValue) in control.Properties)
            {
                var name = rawName.Trim();
                if (!IsPageProperty(name))
                {
                    continue;
                }
                if (properties.ContainsKey(name))
                {
                    throw new InvalidOperationException($"designer_write_failed: Page '{control.Name}' has duplicate property aliases for '{name}'.");
                }
                properties.Add(name, ReadPagePropertyValue(control.Name, name, rawValue));
            }
        }

        return new FormWritePageProperties
        {
            Caption = ResolvePageString(control, properties, "caption", control.Caption),
            Enabled = ResolvePageBoolean(control, properties, "enabled", control.Enabled),
            Visible = ResolvePageBoolean(control, properties, "visible", control.Visible),
            Tag = ResolvePageString(control, properties, "tag", control.Tag),
            ControlTipText = ResolvePageString(control, properties, "controltiptext", control.ControlTipText),
            Accelerator = ResolvePageString(control, properties, "accelerator", control.Accelerator),
        };
    }

    private static bool IsPageProperty(string name)
    {
        return name.Equals("caption", StringComparison.OrdinalIgnoreCase)
            || name.Equals("enabled", StringComparison.OrdinalIgnoreCase)
            || name.Equals("visible", StringComparison.OrdinalIgnoreCase)
            || name.Equals("tag", StringComparison.OrdinalIgnoreCase)
            || name.Equals("controltiptext", StringComparison.OrdinalIgnoreCase)
            || name.Equals("accelerator", StringComparison.OrdinalIgnoreCase);
    }

    private static object ReadPagePropertyValue(string controlName, string propertyName, JsonElement value)
    {
        if (propertyName.Equals("enabled", StringComparison.OrdinalIgnoreCase)
            || propertyName.Equals("visible", StringComparison.OrdinalIgnoreCase))
        {
            if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            {
                return value.GetBoolean();
            }
            throw new InvalidOperationException($"designer_write_failed: Page '{controlName}' property '{propertyName}' must be Boolean.");
        }
        if (value.ValueKind == JsonValueKind.String)
        {
            return value.GetString() ?? "";
        }
        throw new InvalidOperationException($"designer_write_failed: Page '{controlName}' property '{propertyName}' must be a string.");
    }

    private static string? ResolvePageString(
        FormWriteControlSpec control,
        Dictionary<string, object?> properties,
        string propertyName,
        string? topLevelValue)
    {
        var bagValue = properties.TryGetValue(propertyName, out var value) ? (string)value! : null;
        if (topLevelValue is not null && bagValue is not null && !string.Equals(topLevelValue, bagValue, StringComparison.Ordinal))
        {
            throw PageAliasConflict(control, propertyName);
        }
        return topLevelValue ?? bagValue;
    }

    private static bool? ResolvePageBoolean(
        FormWriteControlSpec control,
        Dictionary<string, object?> properties,
        string propertyName,
        bool? topLevelValue)
    {
        bool? bagValue = properties.TryGetValue(propertyName, out var value) ? (bool)value! : null;
        if (topLevelValue is not null && bagValue is not null && topLevelValue.Value != bagValue.Value)
        {
            throw PageAliasConflict(control, propertyName);
        }
        return topLevelValue ?? bagValue;
    }

    private static InvalidOperationException PageAliasConflict(FormWriteControlSpec control, string propertyName)
    {
        return new InvalidOperationException($"designer_write_failed: Page '{control.Name}' has conflicting explicit aliases for '{propertyName}'.");
    }

    private static bool IsControlType(FormWriteControlSpec control, string type)
    {
        return IsControlType(control.Type, type);
    }

    private static bool IsControlType(string? actualType, string type)
    {
        return string.Equals(actualType?.Trim(), type, StringComparison.OrdinalIgnoreCase);
    }

    private static void ClearMultiPagePages(object multiPage)
    {
        object? pages = null;
        try
        {
            pages = ExcelBridgeSupport.Get(multiPage, "Pages")
                ?? throw new InvalidOperationException("designer_write_failed: MultiPage Pages collection is unavailable.");
            while (true)
            {
                var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(pages, "Count"));
                if (count == 0)
                {
                    return;
                }

                object? page = null;
                try
                {
                    page = ExcelBridgeSupport.Get(pages, "Item", count - 1)
                        ?? throw new InvalidOperationException("designer_write_failed: failed to read the last MultiPage Page.");
                    var name = ExcelBridgeSupport.GetString(page, "Name");
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        throw new InvalidOperationException("designer_write_failed: MultiPage Page has no removable name.");
                    }
                    // Excel accepts a page name here; numeric indexes have
                    // caused an RPC failure in the supported host environment.
                    ExcelBridgeSupport.InvokeMethod(pages, "Remove", name);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(page);
                }

                var remaining = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(pages, "Count"));
                if (remaining != count - 1)
                {
                    throw new InvalidOperationException("designer_write_failed: MultiPage Page removal did not reduce the collection by one.");
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(pages);
        }
    }

    private static void ReplaceTabStripTabs(object tabStrip, IReadOnlyList<FormWriteTabSpec> tabSpecs)
    {
        object? tabs = null;
        try
        {
            tabs = ExcelBridgeSupport.Get(tabStrip, "Tabs")
                ?? throw new InvalidOperationException("designer_write_failed: TabStrip Tabs collection is unavailable.");
            while (true)
            {
                var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(tabs, "Count"));
                if (count == 0)
                {
                    break;
                }

                // MSForms Tabs.Remove accepts a numeric index. Passing a tab
                // name raises an invalid-argument COM error in Excel.
                ExcelBridgeSupport.InvokeMethod(tabs, "Remove", count - 1);
                var remaining = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(tabs, "Count"));
                if (remaining != count - 1)
                {
                    throw new InvalidOperationException("designer_write_failed: TabStrip tab removal did not reduce the collection by one.");
                }
            }

            foreach (var tabSpec in tabSpecs)
            {
                object? tab = null;
                try
                {
                    tab = ExcelBridgeSupport.InvokeMethod(tabs, "Add", tabSpec.Name)
                        ?? throw new InvalidOperationException($"designer_write_failed: failed to add TabStrip tab '{tabSpec.Name}'.");
                    VerifyRequiredStringProperty(tab, "Name", tabSpec.Name);
                    SetOptionalRequiredStringProperty(tab, "Caption", tabSpec.Caption);
                    SetOptionalRequiredStringProperty(tab, "ControlTipText", tabSpec.ControlTipText);
                    SetOptionalRequiredStringProperty(tab, "Tag", tabSpec.Tag);
                    SetOptionalRequiredStringProperty(tab, "Accelerator", tabSpec.Accelerator);
                    SetOptionalRequiredBooleanProperty(tab, "Enabled", tabSpec.Enabled);
                    SetOptionalRequiredBooleanProperty(tab, "Visible", tabSpec.Visible);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(tab);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(tabs);
        }
    }

    private static void ApplySelectedIndex(object control, FormWriteControlSpec spec, int childCount)
    {
        var selectedIndex = spec.SelectedIndex ?? spec.Observed?.SelectedIndex;
        if (IsControlType(spec, "MultiPage"))
        {
            if (selectedIndex is null)
            {
                return;
            }
            ValidateSelectedIndex(spec.Name, "MultiPage", selectedIndex.Value, childCount);
            if (childCount == 0)
            {
                // Excel can crash when Value=-1 is assigned after the last Page
                // is removed. The empty public selection is normalized on read.
                return;
            }
            SetRequiredMember(control, "Value", selectedIndex.Value);
            return;
        }

        if (IsControlType(spec, "TabStrip"))
        {
            if (selectedIndex is null)
            {
                return;
            }
            object? tabs = null;
            try
            {
                tabs = ExcelBridgeSupport.Get(control, "Tabs")
                    ?? throw new InvalidOperationException("designer_write_failed: TabStrip Tabs collection is unavailable.");
                var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(tabs, "Count"));
                ValidateSelectedIndex(spec.Name, "TabStrip", selectedIndex.Value, count);
                if (count > 0)
                {
                    SetRequiredMember(control, "Value", selectedIndex.Value);
                }
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(tabs);
            }
            return;
        }

        if (selectedIndex is not null)
        {
            TrySetProperty(control, "ListIndex", selectedIndex.Value);
        }
    }

    private static void ClearDesignerControls(object designer)
    {
        object? controls = null;
        try
        {
            controls = ExcelBridgeSupport.Get(designer, "Controls");
            if (controls is null)
            {
                return;
            }
            while (ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(controls, "Count")) > 0)
            {
                var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(controls, "Count"));
                object? control = null;
                try
                {
                    control = ExcelBridgeSupport.Get(controls, "Item", count - 1);
                    var name = ExcelBridgeSupport.GetString(control!, "Name") ?? "";
                    ExcelBridgeSupport.InvokeMethod(controls, "Remove", name);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(control);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("designer_write_failed: failed to clear existing controls. " + ex.Message, ex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(controls);
        }
    }

    private static void SetDesignerFormProperties(object designer, object component, FormWriteFormSpec form)
    {
        var caption = form.Build?.Caption ?? form.Caption ?? form.Observed?.Caption;
        if (caption is not null)
        {
            SetRequiredVBComponentProperty(component, "Caption", caption);
            SetRequiredProperty(designer, "Caption", caption);
        }
        var width = form.Build?.Width ?? form.Width ?? form.Observed?.Width;
        if (width is not null)
        {
            if (!SetVBComponentProperty(component, "Width", width.Value))
            {
                TrySetProperty(designer, "Width", width.Value);
            }
        }
        var height = form.Build?.Height ?? form.Height ?? form.Observed?.Height;
        if (height is not null)
        {
            if (!SetVBComponentProperty(component, "Height", height.Value))
            {
                TrySetProperty(designer, "Height", height.Value);
            }
        }
    }

    private static void SetDesignerControlProperties(object control, FormWriteControlSpec spec)
    {
        var pageProperties = IsControlType(spec, "Page") ? spec.ResolvedPageProperties : null;
        var caption = pageProperties?.Caption ?? spec.Caption ?? spec.Observed?.Caption;
        if (caption is not null)
        {
            SetRequiredProperty(control, "Caption", caption);
        }
        var left = IsControlType(spec, "Page") ? null : spec.Left ?? spec.Observed?.Left;
        if (left is not null)
        {
            SetRequiredMember(control, "Left", left.Value);
        }
        var top = IsControlType(spec, "Page") ? null : spec.Top ?? spec.Observed?.Top;
        if (top is not null)
        {
            SetRequiredMember(control, "Top", top.Value);
        }
        var width = IsControlType(spec, "Page") ? null : spec.Width ?? spec.Observed?.Width;
        if (width is not null)
        {
            SetRequiredMember(control, "Width", width.Value);
        }
        var height = IsControlType(spec, "Page") ? null : spec.Height ?? spec.Observed?.Height;
        if (height is not null)
        {
            SetRequiredMember(control, "Height", height.Value);
        }
        var enabled = pageProperties?.Enabled ?? spec.Enabled ?? spec.Observed?.Enabled;
        if (enabled is not null)
        {
            SetRequiredMember(control, "Enabled", enabled.Value);
        }
        var visible = pageProperties?.Visible ?? spec.Visible ?? spec.Observed?.Visible;
        if (visible is not null)
        {
            SetRequiredMember(control, "Visible", visible.Value);
        }
        SetControlListItems(control, spec.List ?? spec.Observed?.List);
        var hasValueSelection = IsControlType(spec, "MultiPage") || IsControlType(spec, "TabStrip");
        if (!hasValueSelection && spec.Value.ValueKind != JsonValueKind.Undefined)
        {
            var value = ConvertFormValue(spec.Value);
            if (value is not null)
            {
                SetRequiredMember(control, "Value", value);
            }
        }
        else if (!hasValueSelection && spec.Observed is { } observed && observed.Value.ValueKind != JsonValueKind.Undefined)
        {
            var observedValue = ConvertFormValue(observed.Value);
            if (observedValue is not null)
            {
                SetRequiredMember(control, "Value", observedValue);
            }
        }
        var text = spec.Text ?? spec.Observed?.Text;
        if (text is not null)
        {
            SetRequiredProperty(control, "Text", text);
        }
        SetOptionalRequiredStringProperty(control, "Tag", pageProperties?.Tag ?? spec.Tag ?? spec.Observed?.Tag);
        SetOptionalRequiredStringProperty(control, "ControlTipText", pageProperties?.ControlTipText ?? spec.ControlTipText ?? spec.Observed?.ControlTipText);
        SetOptionalRequiredStringProperty(control, "Accelerator", pageProperties?.Accelerator ?? spec.Accelerator ?? spec.Observed?.Accelerator);
    }

    // System.Text.Json deserializes object to JsonElement. COM needs concrete
    // scalar CLR values, especially for Boolean CheckBox/OptionButton values.
    private static object? ConvertFormValue(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.True or JsonValueKind.False => value.GetBoolean(),
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number when value.TryGetInt64(out var integer) => integer,
            JsonValueKind.Number => value.GetDouble(),
            _ => null,
        };
    }

    private static void SetControlListItems(object control, IReadOnlyList<string>? items)
    {
        if (items is null)
        {
            return;
        }
        try
        {
            var listCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(control, "ListCount"));
            if (listCount > 0)
            {
                ExcelBridgeSupport.InvokeMethod(control, "Clear");
            }
        }
        catch
        {
            return;
        }
        foreach (var item in items)
        {
            try
            {
                ExcelBridgeSupport.InvokeMethod(control, "AddItem", item ?? "");
            }
            catch
            {
                return;
            }
        }
    }

    private static bool TrySetProperty(object target, string propertyName, object value)
    {
        try
        {
            ExcelBridgeSupport.Set(target, propertyName, value);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool SetVBComponentProperty(object component, string propertyName, double value)
    {
        object? properties = null;
        try
        {
            properties = ExcelBridgeSupport.Get(component, "Properties");
            if (properties is null)
            {
                return false;
            }
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(properties, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? property = null;
                try
                {
                    property = ExcelBridgeSupport.Get(properties, "Item", index);
                    var name = ExcelBridgeSupport.GetString(property!, "Name") ?? "";
                    if (string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        ExcelBridgeSupport.Set(property!, "Value", value);
                        return true;
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(property);
                }
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(properties);
        }
        return false;
    }

    private static void SetRequiredVBComponentProperty(object component, string propertyName, string expectedValue)
    {
        object? properties = null;
        try
        {
            properties = ExcelBridgeSupport.Get(component, "Properties");
            if (properties is null)
            {
                throw new InvalidOperationException($"designer_write_failed: failed to access VBComponent.Properties for UserForm {propertyName}.");
            }

            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(properties, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? property = null;
                try
                {
                    property = ExcelBridgeSupport.Get(properties, "Item", index);
                    var name = ExcelBridgeSupport.GetString(property!, "Name") ?? "";
                    if (!string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    ExcelBridgeSupport.Set(property!, "Value", expectedValue);
                    var actualValue = ExcelBridgeSupport.GetString(property!, "Value");
                    if (!string.Equals(actualValue, expectedValue, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException($"designer_write_failed: UserForm {propertyName} did not persist through VBComponent.Properties.");
                    }
                    return;
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(property);
                }
            }

            throw new InvalidOperationException($"designer_write_failed: VBComponent.Properties does not expose UserForm {propertyName}.");
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"designer_write_failed: failed to set UserForm {propertyName} through VBComponent.Properties. {ex.Message}", ex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(properties);
        }
    }

    private static void SetRequiredProperty(object target, string propertyName, string expectedValue)
    {
        try
        {
            ExcelBridgeSupport.Set(target, propertyName, expectedValue);
            var actualValue = ExcelBridgeSupport.GetString(target, propertyName);
            if (!string.Equals(actualValue, expectedValue, StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"designer_write_failed: UserForm Designer {propertyName} did not persist.");
            }
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"designer_write_failed: failed to set UserForm Designer {propertyName}. {ex.Message}", ex);
        }
    }

    private static void VerifyRequiredStringProperty(object target, string propertyName, string expectedValue)
    {
        try
        {
            var actualValue = ExcelBridgeSupport.GetString(target, propertyName);
            if (!string.Equals(actualValue, expectedValue, StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"designer_write_failed: UserForm Designer {propertyName} did not persist.");
            }
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"designer_write_failed: failed to read UserForm Designer {propertyName}. {ex.Message}", ex);
        }
    }

    private static void SetOptionalRequiredStringProperty(object target, string propertyName, string? value)
    {
        if (value is not null)
        {
            SetRequiredProperty(target, propertyName, value);
        }
    }

    private static void SetOptionalRequiredBooleanProperty(object target, string propertyName, bool? value)
    {
        if (value is not null)
        {
            SetRequiredMember(target, propertyName, value.Value);
        }
    }

    private static void SetRequiredMember(object target, string propertyName, object expectedValue)
    {
        try
        {
            ExcelBridgeSupport.Set(target, propertyName, expectedValue);
            var actualValue = ExcelBridgeSupport.Get(target, propertyName);
            if (!RequiredValuesEqual(actualValue, expectedValue))
            {
                throw new InvalidOperationException($"designer_write_failed: UserForm Designer {propertyName} did not persist.");
            }
        }
        catch (InvalidOperationException ex) when (ex.Message.StartsWith("designer_write_failed:", StringComparison.Ordinal))
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"designer_write_failed: failed to set UserForm Designer {propertyName}. {ex.Message}", ex);
        }
    }

    private static bool RequiredValuesEqual(object? actualValue, object expectedValue)
    {
        if (actualValue is null)
        {
            return false;
        }
        if (expectedValue is string expectedString)
        {
            return string.Equals(Convert.ToString(actualValue, CultureInfo.InvariantCulture), expectedString, StringComparison.Ordinal);
        }
        if (expectedValue is bool expectedBoolean)
        {
            return Convert.ToBoolean(actualValue, CultureInfo.InvariantCulture) == expectedBoolean;
        }
        if (expectedValue is int expectedInteger)
        {
            return Convert.ToInt32(actualValue, CultureInfo.InvariantCulture) == expectedInteger;
        }
        if (expectedValue is long expectedLong)
        {
            return Convert.ToInt64(actualValue, CultureInfo.InvariantCulture) == expectedLong;
        }
        if (expectedValue is double expectedDouble)
        {
            return Math.Abs(Convert.ToDouble(actualValue, CultureInfo.InvariantCulture) - expectedDouble) <= 0.000001;
        }
        return Equals(actualValue, expectedValue);
    }

    private static object? GetComponentByName(object vbProject, string name)
    {
        object? components = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            return ExcelBridgeSupport.Get(components!, "Item", name);
        }
        catch
        {
            return null;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static object NewUserFormComponent(object vbProject, string name)
    {
        object? components = null;
        object? component = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            component = ExcelBridgeSupport.InvokeMethod(components!, "Add", ComponentTypeForm)
                ?? throw new InvalidOperationException($"designer_write_failed: failed to create UserForm '{name}'.");
            ExcelBridgeSupport.Set(component, "Name", name);
            return component;
        }
        catch (Exception ex)
        {
            if (components is not null && component is not null)
            {
                try { ExcelBridgeSupport.InvokeMethod(components, "Remove", component); } catch { }
            }
            throw new InvalidOperationException($"designer_write_failed: failed to create UserForm '{name}'. {ex.Message}", ex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void RemoveComponentByName(object vbProject, string name)
    {
        object? components = null;
        object? component = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            component = ExcelBridgeSupport.Get(components!, "Item", name);
            if (component is null)
            {
                return;
            }
            ExcelBridgeSupport.InvokeMethod(components!, "Remove", component);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"designer_write_failed: failed to remove existing component '{name}'. {ex.Message}", ex);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static void RemoveComponentInstance(object vbProject, object? component)
    {
        object? components = null;
        try
        {
            if (component is null)
            {
                return;
            }
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is not null)
            {
                ExcelBridgeSupport.InvokeMethod(components, "Remove", component);
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static string NewRestoreDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "xlflow-form-restore-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(path);
        return path;
    }

    private static string ExportComponentBackup(object component, string directory)
    {
        var name = ExcelBridgeSupport.GetString(component, "Name") ?? "UserForm";
        var exportPath = Path.Combine(directory, name + ".frm");
        ExcelBridgeSupport.InvokeViaDynamic(component, "Export", exportPath);
        return exportPath;
    }

    private static void ImportComponentBackup(object vbProject, string exportPath, string expectedName)
    {
        object? components = null;
        object? restored = null;
        try
        {
            components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            restored = ExcelBridgeSupport.InvokeMethod(components!, "Import", exportPath);
            if (restored is not null)
            {
                var restoredName = ExcelBridgeSupport.GetString(restored, "Name") ?? "";
                if (!string.Equals(restoredName, expectedName, StringComparison.OrdinalIgnoreCase))
                {
                    ExcelBridgeSupport.Set(restored, "Name", expectedName);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(restored);
            ExcelBridgeSupport.ReleaseComObject(components);
        }
    }

    private static string? ReadCodeModuleText(object component)
    {
        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
            return codeModule is null ? null : VbaSourceHelper.GetCodeModuleText(codeModule);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    private static void WriteCodeModuleText(object component, string text)
    {
        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
            if (codeModule is not null)
            {
                VbaSourceHelper.SetCodeModuleText(codeModule, text);
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    private static void SyncUserFormCodeBehind(object component, string formsDir, string? fallbackText)
    {
        var formName = ExcelBridgeSupport.GetString(component, "Name") ?? "";
        var codePath = VbaSourceHelper.GetUserFormCodePath(formsDir, formName);
        if (!string.IsNullOrWhiteSpace(codePath) && File.Exists(codePath))
        {
            WriteCodeModuleText(component, File.ReadAllText(codePath, Encoding.UTF8));
            return;
        }
        if (!string.IsNullOrWhiteSpace(fallbackText))
        {
            WriteCodeModuleText(component, fallbackText);
        }
    }

    private static SourceArtifacts ExportBuiltUserFormArtifacts(object component, FormWriteCommandArguments args)
    {
        var exportPath = VbaSourceHelper.GetComponentPath(component, "", "", args.FormsDir, "", args.Folders, args.FolderAnnotation, args.DefaultComponentFolders)
            ?? throw new InvalidOperationException($"designer_write_failed: failed to resolve the source artifact path for built UserForm '{ExcelBridgeSupport.GetString(component, "Name")}'.");
        var parent = Path.GetDirectoryName(exportPath);
        if (!string.IsNullOrWhiteSpace(parent))
        {
            Directory.CreateDirectory(parent);
        }
        RemoveExportArtifacts(exportPath);

        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-form-write-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(tmpDir);
        try
        {
            var name = ExcelBridgeSupport.GetString(component, "Name") ?? "UserForm";
            var tmpFile = Path.Combine(tmpDir, name + ".frm");
            ExcelBridgeSupport.InvokeViaDynamic(component, "Export", tmpFile);
            foreach (var exportedFile in Directory.GetFiles(tmpDir))
            {
                var ext = Path.GetExtension(exportedFile).ToLowerInvariant();
                if (ext is ".bas" or ".cls" or ".frm")
                {
                    var content = VbaSourceHelper.ReadExportedTextAsUtf8(exportedFile);
                    content = NormalizeUserFormArtifact(content, component);
                    if (args.FolderAnnotation == "update")
                    {
                        var desiredAnnotation = VbaSourceHelper.GetFolderAnnotationForPath(args.FormsDir, exportPath);
                        content = VbaSourceHelper.UpdateFolderAnnotationText(content, args.FolderAnnotation, desiredAnnotation);
                    }
                    File.WriteAllText(exportPath, content, new UTF8Encoding(false));
                }
                else
                {
                    File.Copy(exportedFile, Path.ChangeExtension(exportPath, ext), true);
                }
            }

            string? codePath = null;
            if (VbaSourceHelper.IsSidecarMode(args.CodeSource))
            {
                codePath = ExportUserFormCodeBehind(component, args.FormsDir);
            }

            return new SourceArtifacts(exportPath, Path.ChangeExtension(exportPath, ".frx"), codePath);
        }
        finally
        {
            if (Directory.Exists(tmpDir))
            {
                try { Directory.Delete(tmpDir, true); } catch { }
            }
        }
    }

    private static string? ExportUserFormCodeBehind(object component, string formsDir)
    {
        var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
        var codePath = VbaSourceHelper.GetUserFormCodePath(formsDir, name);
        if (string.IsNullOrWhiteSpace(codePath))
        {
            return null;
        }
        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
            if (codeModule is null)
            {
                return null;
            }
            var text = VbaSourceHelper.GetCodeModuleText(codeModule);
            if (string.IsNullOrWhiteSpace(text))
            {
                return File.Exists(codePath) ? codePath : null;
            }
            var dir = Path.GetDirectoryName(codePath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(codePath, text, new UTF8Encoding(false));
            return codePath;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    private static void RemoveExportArtifacts(string exportPath)
    {
        foreach (var path in new[] { exportPath, Path.ChangeExtension(exportPath, ".frx") })
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
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
                var caption = designer is null ? null : ExcelBridgeSupport.GetString(designer, "Caption");
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
        }
        return content;
    }

    private static string InjectOrUpdateCaption(string content, string caption)
    {
        // Multiline captions already live in the exported FRX. Never synthesize
        // additional source lines for them in the single-line FRM assignment.
        if (caption.Contains('\r') || caption.Contains('\n'))
        {
            return content;
        }
        var escapedCaption = caption.Replace("\"", "\"\"", StringComparison.Ordinal);
        var lines = new List<string>(content.Split(["\r\n", "\n", "\r"], StringSplitOptions.None));
        var beginIndex = lines.FindIndex(line => line.Trim() == "Begin");
        var endIndex = beginIndex >= 0 ? lines.FindIndex(beginIndex + 1, line => line.Trim() == "End") : -1;
        if (beginIndex >= 0 && endIndex > beginIndex)
        {
            for (var index = beginIndex + 1; index < endIndex; index++)
            {
                var trimmed = lines[index].Trim();
                if (trimmed.StartsWith("Caption", StringComparison.OrdinalIgnoreCase) && trimmed.Contains('='))
                {
                    lines[index] = $"   Caption = \"{escapedCaption}\"";
                    return string.Join(Environment.NewLine, lines);
                }
            }
            lines.Insert(beginIndex + 1, $"   Caption = \"{escapedCaption}\"");
        }
        return string.Join(Environment.NewLine, lines);
    }

    private static string ResolveProgId(FormWriteControlSpec spec)
    {
        if (!string.IsNullOrWhiteSpace(spec.ProgId))
        {
            return spec.ProgId;
        }
        if (TypeProgIdMap.TryGetValue(spec.Type ?? "", out var progId))
        {
            return progId;
        }
        throw new InvalidOperationException($"unsupported_form_control: unsupported control type '{spec.Type}'.");
    }

    private static Dictionary<string, string> NewMessage(string code, string message, string? fieldPath = null)
    {
        var result = new Dictionary<string, string>
        {
            ["code"] = code,
            ["message"] = message,
        };
        if (!string.IsNullOrWhiteSpace(fieldPath))
        {
            result["field_path"] = fieldPath;
        }
        return result;
    }

    private static (string Code, string Message) ClassifyFormWriteError(string action, string message)
    {
        if (message.StartsWith("form_already_exists: ", StringComparison.OrdinalIgnoreCase))
        {
            return ("form_already_exists", message["form_already_exists: ".Length..]);
        }
        if (message.StartsWith("form_not_found: ", StringComparison.OrdinalIgnoreCase))
        {
            return ("form_not_found", message["form_not_found: ".Length..]);
        }
        if (message.StartsWith("unsupported_form_control: ", StringComparison.OrdinalIgnoreCase))
        {
            return ("unsupported_form_control", message["unsupported_form_control: ".Length..]);
        }
        if (message.StartsWith("designer_write_failed: ", StringComparison.OrdinalIgnoreCase))
        {
            return ("designer_write_failed", message["designer_write_failed: ".Length..]);
        }
        return (action == "apply" ? "form_apply_failed" : "form_build_failed", message);
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
        }
        ExcelBridgeSupport.ReleaseComObject(workbook);
        ExcelBridgeSupport.ReleaseComObject(excel);
    }

    private sealed record SourceArtifacts(string FormPath, string FrxPath, string? CodePath);

    private sealed record FormWriteExistingTabStripState(List<FormWriteTabSpec> Tabs, int SelectedIndex);

    private sealed class FormWriteSpec
    {
        [JsonPropertyName("basis")]
        public string Basis { get; init; } = "";

        [JsonPropertyName("coordinateSystem")]
        public string? CoordinateSystem { get; init; }

        [JsonPropertyName("form")]
        public FormWriteFormSpec Form { get; init; } = new();

        [JsonPropertyName("controls")]
        public List<FormWriteControlSpec> Controls { get; init; } = [];
    }

    private sealed class FormWriteFormSpec
    {
        [JsonPropertyName("name")]
        public string Name { get; init; } = "";

        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("height")]
        public double? Height { get; init; }

        [JsonPropertyName("observed")]
        public FormWriteObservedFormSpec? Observed { get; init; }

        [JsonPropertyName("build")]
        public FormWriteBuildFormSpec? Build { get; init; }
    }

    private sealed class FormWriteObservedFormSpec
    {
        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("height")]
        public double? Height { get; init; }

        [JsonPropertyName("insideWidth")]
        public double? InsideWidth { get; init; }

        [JsonPropertyName("insideHeight")]
        public double? InsideHeight { get; init; }

        [JsonPropertyName("clientWidth")]
        public double? ClientWidth { get; init; }

        [JsonPropertyName("clientHeight")]
        public double? ClientHeight { get; init; }
    }

    private sealed class FormWriteBuildFormSpec
    {
        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("height")]
        public double? Height { get; init; }

        [JsonPropertyName("clientWidth")]
        public double? ClientWidth { get; init; }

        [JsonPropertyName("clientHeight")]
        public double? ClientHeight { get; init; }
    }

    private sealed class FormWriteControlSpec
    {
        [JsonPropertyName("id")]
        public string Id { get; init; } = "";

        [JsonPropertyName("parentId")]
        public string ParentId { get; init; } = "";

        [JsonPropertyName("zIndex")]
        public int? ZIndex { get; init; }

        [JsonPropertyName("type")]
        public string Type { get; init; } = "";

        [JsonPropertyName("name")]
        public string Name { get; init; } = "";

        [JsonPropertyName("progId")]
        public string? ProgId { get; init; }

        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("tag")]
        public string? Tag { get; init; }

        [JsonPropertyName("controlTipText")]
        public string? ControlTipText { get; init; }

        [JsonPropertyName("accelerator")]
        public string? Accelerator { get; init; }

        [JsonPropertyName("text")]
        public string? Text { get; init; }

        [JsonPropertyName("value")]
        public JsonElement Value { get; init; }

        [JsonPropertyName("left")]
        public double? Left { get; init; }

        [JsonPropertyName("top")]
        public double? Top { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("height")]
        public double? Height { get; init; }

        [JsonPropertyName("tabIndex")]
        public int? TabIndex { get; init; }

        [JsonPropertyName("selectedIndex")]
        public int? SelectedIndex { get; init; }

        [JsonPropertyName("enabled")]
        public bool? Enabled { get; init; }

        [JsonPropertyName("visible")]
        public bool? Visible { get; init; }

        [JsonPropertyName("list")]
        public List<string>? List { get; init; }

        [JsonPropertyName("controls")]
        public List<FormWriteControlSpec>? Controls { get; init; }

        [JsonPropertyName("tabs")]
        public List<FormWriteTabSpec>? Tabs { get; init; }

        [JsonPropertyName("properties")]
        public Dictionary<string, JsonElement>? Properties { get; init; }

        [JsonPropertyName("observed")]
        public FormWriteObservedControlSpec? Observed { get; set; }

        [JsonIgnore]
        public FormWritePageProperties? ResolvedPageProperties { get; set; }
    }

    private sealed class FormWritePageProperties
    {
        public string? Caption { get; init; }

        public bool? Enabled { get; init; }

        public bool? Visible { get; init; }

        public string? Tag { get; init; }

        public string? ControlTipText { get; init; }

        public string? Accelerator { get; init; }
    }

    private sealed class FormWriteTabSpec
    {
        [JsonPropertyName("name")]
        public string Name { get; init; } = "";

        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("controlTipText")]
        public string? ControlTipText { get; init; }

        [JsonPropertyName("tag")]
        public string? Tag { get; init; }

        [JsonPropertyName("accelerator")]
        public string? Accelerator { get; init; }

        [JsonPropertyName("enabled")]
        public bool? Enabled { get; init; }

        [JsonPropertyName("visible")]
        public bool? Visible { get; init; }
    }

    private sealed class FormWriteObservedControlSpec
    {
        [JsonPropertyName("caption")]
        public string? Caption { get; init; }

        [JsonPropertyName("tag")]
        public string? Tag { get; init; }

        [JsonPropertyName("controlTipText")]
        public string? ControlTipText { get; init; }

        [JsonPropertyName("accelerator")]
        public string? Accelerator { get; init; }

        [JsonPropertyName("text")]
        public string? Text { get; init; }

        [JsonPropertyName("value")]
        public JsonElement Value { get; init; }

        [JsonPropertyName("left")]
        public double? Left { get; init; }

        [JsonPropertyName("top")]
        public double? Top { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("height")]
        public double? Height { get; init; }

        [JsonPropertyName("tabIndex")]
        public int? TabIndex { get; init; }

        [JsonPropertyName("selectedIndex")]
        public int? SelectedIndex { get; set; }

        [JsonPropertyName("enabled")]
        public bool? Enabled { get; init; }

        [JsonPropertyName("visible")]
        public bool? Visible { get; init; }

        [JsonPropertyName("list")]
        public List<string>? List { get; init; }

        [JsonPropertyName("tabs")]
        public List<FormWriteTabSpec>? Tabs { get; set; }

        [JsonPropertyName("unsupported")]
        public List<string>? Unsupported { get; init; }

        [JsonPropertyName("properties")]
        public Dictionary<string, JsonElement>? Properties { get; init; }
    }
}
