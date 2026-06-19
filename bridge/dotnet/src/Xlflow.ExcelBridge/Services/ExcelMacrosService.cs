using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelMacrosService : IMacrosService
{
    private static readonly string[] EventProcedurePrefixes = ["Workbook_", "Worksheet_"];
    private static readonly string[] EventProcedureNames = ["Auto_Open", "Auto_Close"];

    public BridgeResponse Execute(BridgeRequest request, MacrosCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForMacros(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            var dirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
            var dirty = sessionAttached ? dirtyKnown ? dirtyState : true : false;
            var needsSave = sessionAttached ? dirty : false;

            var macros = DiscoverMacros(workbook, cancellationToken);

            if (args.RunnableOnly)
            {
                macros = macros.Where(m => m.Runnable).ToList();
            }

            var sessionArg = sessionAttached ? " --session" : "";
            foreach (var macro in macros.Where(m => m.Runnable))
            {
                macro.RunCommand = $"xlflow run {macro.QualifiedName}{sessionArg} --json";
            }

            var defaultEntry = "";
            if (!string.IsNullOrWhiteSpace(args.Entry))
            {
                var entryMatch = macros.FirstOrDefault(m => m.QualifiedName == args.Entry && m.Runnable);
                if (entryMatch is not null)
                {
                    defaultEntry = args.Entry;
                }
            }

            var suggestions = new List<object>();
            if (!string.IsNullOrEmpty(defaultEntry))
            {
                suggestions.Add(new { title = "Run the default entrypoint", command = $"xlflow run {defaultEntry}{sessionArg} --json" });
            }
            else
            {
                var firstRunnable = macros.FirstOrDefault(m => m.Runnable);
                if (firstRunnable is not null)
                {
                    suggestions.Add(new { title = "Run the first runnable macro", command = firstRunnable.RunCommand });
                }
            }

            var targetKind = sessionAttached ? "live_session" : "file";
            var logs = new List<string>();
            if (sessionAttached)
            {
                logs.Add($"attached to xlflow session ({sessionMode})");
            }
            logs.Add($"discovered {macros.Count} macro entrypoint(s)");

            var warnings = new List<object>();
            if (needsSave)
            {
                warnings.Add(new
                {
                    code = "save_required",
                    message = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                });
            }

            var hints = new List<object>();
            if (macros.Count == 0)
            {
                hints.Add(new { code = "macros_empty_before_push", message = "If you edited source files, run `xlflow push --session` before `xlflow macros --session`." });
                hints.Add(new { code = "macros_read_from_workbook", message = "`macros` discovers procedures from the workbook, not directly from source files." });
            }

            var extensions = new Dictionary<string, object?>
            {
                ["target"] = new { kind = targetKind, path = args.WorkbookPath },
                ["session"] = new { active = sessionAttached, workbook_path = args.WorkbookPath, dirty, save_required = needsSave, live_newer_than_disk = needsSave, mode = sessionMode, source_of_truth = needsSave ? "live_workbook" : "saved_workbook" },
                ["workbook"] = new { path = args.WorkbookPath, session = sessionAttached, session_mode = sessionMode, session_requested = args.UseSession, auto_session = sessionAttached && !args.UseSession, dirty, needs_save = needsSave },
                ["macros"] = macros.Select(m => m.ToDictionary()).ToArray(),
            };

            if (!string.IsNullOrEmpty(defaultEntry))
            {
                extensions["default_entry"] = defaultEntry;
            }

            if (suggestions.Count > 0)
            {
                extensions["suggestions"] = suggestions;
            }

            if (warnings.Count > 0)
            {
                extensions["warnings"] = warnings;
            }

            if (hints.Count > 0)
            {
                extensions["hints"] = hints;
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = extensions,
            };
        }
        catch (Exception ex)
        {
            var detail = ExcelBridgeSupport.FormatExceptionDetail(ex);
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "macro_discovery_failed",
                Message: detail,
                Phase: "macros",
                Source: "xlflow-excel-bridge"));
        }
        finally
        {
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

    private static List<MacroEntry> DiscoverMacros(object workbook, CancellationToken cancellationToken)
    {
        var macros = new List<MacroEntry>();
        object? project = null;
        object? components = null;

        try
        {
            project = ExcelBridgeSupport.RunPhase("get_vbproject", () => ExcelBridgeSupport.Get(workbook, "VBProject"));
            if (project is null)
            {
                return macros;
            }

            components = ExcelBridgeSupport.RunPhase("get_vbcomponents", () => ExcelBridgeSupport.Get(project, "VBComponents"));
            if (components is null)
            {
                return macros;
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

                    var name = (string?)ExcelBridgeSupport.Get(component, "Name");
                    if (string.IsNullOrWhiteSpace(name) || name.StartsWith("Xlflow", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var componentType = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
                    var componentTypeName = GetComponentTypeName(componentType);

                    var code = GetCodeModuleText(component);
                    if (string.IsNullOrEmpty(code))
                    {
                        continue;
                    }

                    var componentMacros = FindMacroProcedures(name, componentTypeName, code);
                    macros.AddRange(componentMacros);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(component);
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(components);
            ExcelBridgeSupport.ReleaseComObject(project);
        }

        return macros;
    }

    private static string GetComponentTypeName(int componentType)
    {
        return componentType switch
        {
            1 => "standard_module",
            2 => "class_module",
            3 => "userform",
            100 => "document_module",
            _ => "unknown",
        };
    }

    private static string GetCodeModuleText(object component)
    {
        object? codeModule = null;
        try
        {
            codeModule = ExcelBridgeSupport.RunPhase("get_codemodule", () => ExcelBridgeSupport.Get(component, "CodeModule"));
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

    internal static List<MacroEntry> FindMacroProcedures(string moduleName, string componentTypeName, string code)
    {
        var macros = new List<MacroEntry>();
        var lines = code.Split(["\r\n", "\n"], StringSplitOptions.None);

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (line.Length == 0)
            {
                continue;
            }

            if (StartsWithPrivateOrFriend(line))
            {
                continue;
            }

            var match = System.Text.RegularExpressions.Regex.Match(
                line,
                $@"^(?:(Public)\s+)?(Sub|Function)\s+({VbaIdentifierPattern.Identifier})\s*(?:\(([^)]*)\))?",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                continue;
            }

            var name = match.Groups[3].Value;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var visibility = match.Groups[1].Success ? "Public" : "Implicit";
            var argText = match.Groups[4].Value.Trim();
            var macroArgs = new List<string>();
            var hasParams = false;

            if (!string.IsNullOrWhiteSpace(argText))
            {
                macroArgs = argText.Split(',')
                    .Select(a => a.Trim())
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .ToList();
                hasParams = macroArgs.Count > 0;
            }

            string? reason = null;
            var runnable = false;

            if (hasParams)
            {
                reason = "has_parameters";
            }
            else if (IsEventProcedureName(name))
            {
                reason = "event_procedure";
            }
            else if (componentTypeName is "userform" or "document_module" or "unknown")
            {
                reason = "unsupported_component_type";
            }
            else
            {
                runnable = true;
            }

            macros.Add(new MacroEntry
            {
                Module = moduleName,
                Name = name,
                QualifiedName = $"{moduleName}.{name}",
                Kind = match.Groups[2].Value.ToLowerInvariant(),
                Args = macroArgs,
                Line = i + 1,
                ComponentType = componentTypeName,
                Visibility = visibility,
                HasParameters = hasParams,
                Runnable = runnable,
                ReasonNotRunnable = reason,
            });
        }

        return macros;
    }

    private static bool StartsWithPrivateOrFriend(string line)
    {
        return line.StartsWith("PRIVATE ", StringComparison.OrdinalIgnoreCase) ||
               line.StartsWith("FRIEND ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEventProcedureName(string name)
    {
        foreach (var prefix in EventProcedurePrefixes)
        {
            if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        foreach (var eventName in EventProcedureNames)
        {
            if (string.Equals(name, eventName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForMacros(
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
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            }
            catch
            {
                // best-effort close
            }
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }

        if (excel is not null)
        {
            try
            {
                dynamic app = excel;
                app.Quit();
            }
            catch
            {
                // best-effort quit
            }
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    internal sealed class MacroEntry
    {
        public string Module { get; init; } = "";
        public string Name { get; init; } = "";
        public string QualifiedName { get; init; } = "";
        public string Kind { get; init; } = "";
        public List<string> Args { get; init; } = [];
        public int Line { get; init; }
        public string ComponentType { get; init; } = "";
        public string Visibility { get; init; } = "";
        public bool HasParameters { get; init; }
        public bool Runnable { get; init; }
        public string? ReasonNotRunnable { get; init; }
        public string RunCommand { get; set; } = "";

        public Dictionary<string, object?> ToDictionary()
        {
            var dict = new Dictionary<string, object?>
            {
                ["module"] = Module,
                ["name"] = Name,
                ["qualified_name"] = QualifiedName,
                ["kind"] = Kind,
                ["args"] = Args.ToArray(),
                ["line"] = Line,
                ["component_type"] = ComponentType,
                ["visibility"] = Visibility,
                ["has_parameters"] = HasParameters,
                ["runnable"] = Runnable,
            };

            if (ReasonNotRunnable is not null)
            {
                dict["reason_not_runnable"] = ReasonNotRunnable;
            }

            if (!string.IsNullOrEmpty(RunCommand))
            {
                dict["run_command"] = RunCommand;
            }

            return dict;
        }
    }
}
