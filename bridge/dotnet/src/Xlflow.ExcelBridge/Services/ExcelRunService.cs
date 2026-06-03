using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelRunService : IRunService
{
    private static readonly JsonSerializerOptions CachedJsonOptions = new() { PropertyNameCaseInsensitive = true };
    public BridgeResponse Execute(BridgeRequest request, RunCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var sessionAttached = false;
        var sessionMode = "none";

        try
        {
            var openResult = ExcelBridgeSupport.RunPhase("open_workbook", () =>
                OpenWorkbookForRun(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible));
            excel = openResult.Excel;
            workbook = openResult.Workbook;
            sessionAttached = openResult.SessionAttached;
            sessionMode = openResult.SessionMode;

            var dirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
            var dirty = sessionAttached ? dirtyKnown ? dirtyState : true : false;
            var needsSave = sessionAttached ? dirty : false;

            if (!string.IsNullOrWhiteSpace(args.DisplayAlerts.ToString()))
            {
                try
                {
                    dynamic app = excel;
                    app.DisplayAlerts = args.DisplayAlerts;
                }
                catch
                {
                    // best-effort
                }
            }

            var macroArgs = DecodeMacroArgs(args.MacroArgsJSON);
            var macroName = args.MacroName;

            var sw = Stopwatch.StartNew();
            object? runResult = null;
            string? runError = null;
            int? runErrorNumber = null;
            int? runErrorLine = null;

            try
            {
                dynamic app = excel;

                // Clear any prior error state
                try { dynamic err = app.Error; err.Clear(); } catch { /* ignore */ }

                runResult = InvokeMacro(excel, workbook, macroName, macroArgs);
                sw.Stop();
            }
            catch (Exception ex)
            {
                sw.Stop();
                var comEx = ex.InnerException ?? ex;
                runError = comEx.Message;
                runErrorNumber = comEx is System.Runtime.InteropServices.COMException com ? unchecked((int)com.ErrorCode) : ex.HResult;
                runErrorLine = CaptureErrorLine(excel);
            }

            var durationMs = sw.ElapsedMilliseconds;

            // Capture error info from Err object after execution
            if (runError is null)
            {
                var errInfo = CaptureExcelError(excel);
                if (errInfo is not null)
                {
                    runError = errInfo.Value.Message;
                    runErrorNumber = errInfo.Value.Number;
                }
            }

            // Handle save operations
            var saved = false;
            var saveAsCopy = false;
            if (!string.IsNullOrWhiteSpace(args.SaveAsPath))
            {
                var saveAsPath = ExcelBridgeSupport.NormalizePath(args.SaveAsPath);
                AssertSaveAsExtension(args.WorkbookPath, saveAsPath);
                ExcelBridgeSupport.RunPhase("save_as", () =>
                {
                    ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveCopyAs", saveAsPath);
                });
                saveAsCopy = true;
                // SaveCopyAs creates a copy but does NOT save the original workbook.
                // saved stays false to match PowerShell contract.
            }
            else if (args.SaveWorkbook)
            {
                ExcelBridgeSupport.RunPhase("save_workbook", () =>
                {
                    ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");
                });
                saved = true;
            }

            // Re-check dirty state after save
            if (saved)
            {
                needsSave = false;
                dirty = false;
            }
            else if (saveAsCopy && sessionAttached)
            {
                // SaveCopyAs with session: live workbook remains dirty.
                dirty = true;
                needsSave = true;
            }
            else if (sessionAttached)
            {
                // Session-attached no-save path: preserve or set dirty state.
                // Matches PowerShell contract: live session is dirty until explicitly saved.
                dirty = true;
                needsSave = true;
            }
            else if (ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var postDirty))
            {
                dirty = postDirty;
                needsSave = postDirty;
            }

            var targetKind = sessionAttached ? "live_session" : "file";
            var logs = new List<string>();
            if (sessionAttached)
            {
                logs.Add($"attached to xlflow session ({sessionMode})");
            }

            var warnings = new List<object>();
            if (needsSave)
            {
                warnings.Add(new
                {
                    code = "save_required",
                    message = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                });
            }

            var suggestions = new List<object>();
            if (needsSave)
            {
                suggestions.Add(new { code = "save_session", message = "Run xlflow save --session before session stop." });
            }

            var extensions = new Dictionary<string, object?>
            {
                ["target"] = new { kind = targetKind, path = args.WorkbookPath },
                ["session"] = new { active = sessionAttached, workbook_path = args.WorkbookPath, dirty, save_required = needsSave, live_newer_than_disk = needsSave, mode = sessionMode, source_of_truth = needsSave ? "live_workbook" : "saved_workbook" },
            };

            var workbookResult = new Dictionary<string, object?>
            {
                ["path"] = args.WorkbookPath,
                ["session"] = sessionAttached,
                ["session_mode"] = sessionMode,
                ["session_requested"] = args.UseSession,
                ["auto_session"] = sessionAttached && !args.UseSession,
                ["saved"] = saved,
                ["dirty"] = dirty,
                ["needs_save"] = needsSave,
            };

            if (!string.IsNullOrWhiteSpace(args.SaveAsPath))
            {
                workbookResult["save_as"] = ExcelBridgeSupport.NormalizePath(args.SaveAsPath);
            }

            extensions["workbook"] = workbookResult;

            if (runError is not null)
            {
                var errorCode = ClassifyRunError(runError, runErrorNumber);
                logs.Add($"macro execution failed: {runError}");

                extensions["macro"] = new
                {
                    name = macroName,
                    duration_ms = durationMs,
                    arguments = macroArgs.Select(a => new { type = a.Type, value = a.Value }).ToArray(),
                    error = new
                    {
                        message = runError,
                        number = runErrorNumber,
                        line = runErrorLine,
                    },
                };

                if (!string.IsNullOrEmpty(args.RuntimeMode))
                {
                    extensions["runtime"] = new { mode = args.RuntimeMode, source = args.RuntimeSource };
                }

                if (args.Diagnostic)
                {
                    extensions["run_diagnostic"] = new
                    {
                        kind = "runtime",
                        location = new
                        {
                            macro = macroName,
                            line = runErrorLine,
                        },
                    };
                }

                if (suggestions.Count > 0)
                {
                    extensions["suggestions"] = suggestions;
                }

                if (warnings.Count > 0)
                {
                    extensions["warnings"] = warnings;
                }

                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Status = BridgeStatus.Failed,
                    Error = new BridgeError(
                        Code: errorCode,
                        Message: runError,
                        Phase: "invoke_macro",
                        Source: "xlflow-excel-bridge",
                        Number: runErrorNumber),
                    Logs = logs,
                    Extensions = extensions,
                };
            }

            logs.Add($"ran {macroName} in {durationMs}ms");

            if (saveAsCopy)
            {
                logs.Add($"wrote workbook copy to {ExcelBridgeSupport.NormalizePath(args.SaveAsPath)}");
            }

            if (sessionAttached && !saved)
            {
                logs.Add("SAVE REQUIRED: live workbook is newer than disk; run xlflow save before session stop");
            }

            extensions["macro"] = new
            {
                name = macroName,
                duration_ms = durationMs,
                arguments = macroArgs.Select(a => new { type = a.Type, value = a.Value }).ToArray(),
            };

            if (!string.IsNullOrEmpty(args.RuntimeMode))
            {
                extensions["runtime"] = new { mode = args.RuntimeMode, source = args.RuntimeSource, injected = true };
            }

            if (args.Diagnostic)
            {
                extensions["run_diagnostic"] = new
                {
                    kind = "success",
                    location = new
                    {
                        macro = macroName,
                    },
                };
            }

            if (suggestions.Count > 0)
            {
                extensions["suggestions"] = suggestions;
            }

            if (warnings.Count > 0)
            {
                extensions["warnings"] = warnings;
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
                Code: "macro_failed",
                Message: detail,
                Phase: "run",
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

    private static object? InvokeMacro(object excel, object workbook, string macroName, List<MacroArg> args)
    {
        var invokeArgs = new List<object?> { macroName };
        foreach (var arg in args)
        {
            invokeArgs.Add(ConvertArg(arg));
        }

        return ExcelBridgeSupport.RunPhase("invoke_macro", () =>
        {
            return ExcelBridgeSupport.InvokeMethod(excel, "Run", invokeArgs.ToArray());
        });
    }

    private static object? ConvertArg(MacroArg arg)
    {
        return arg.Type.ToLowerInvariant() switch
        {
            "int" or "integer" or "long" => int.TryParse(arg.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : (object)arg.Value,
            "double" or "float" or "number" => double.TryParse(arg.Value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) ? d : (object)arg.Value,
            "bool" or "boolean" => bool.TryParse(arg.Value, out var b) ? b : (object)arg.Value,
            _ => arg.Value,
        };
    }

    private static List<MacroArg> DecodeMacroArgs(string encoded)
    {
        if (string.IsNullOrWhiteSpace(encoded))
        {
            return [];
        }

        try
        {
            var json = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            return JsonSerializer.Deserialize<List<MacroArg>>(json, CachedJsonOptions) ?? [];
        }
        catch
        {
            return [];
        }
    }

    private static int? CaptureErrorLine(object? excel)
    {
        if (excel is null) return null;
        try
        {
            var err = ExcelBridgeSupport.Get(excel, "Err");
            if (err is null) return null;
            dynamic errObj = err;
            var source = errObj.Source as string;
            if (!string.IsNullOrWhiteSpace(source) && int.TryParse(source, out var line))
            {
                return line;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static (string Message, int Number)? CaptureExcelError(object? excel)
    {
        if (excel is null) return null;
        try
        {
            var err = ExcelBridgeSupport.Get(excel, "Err");
            if (err is null) return null;
            dynamic errObj = err;
            var number = ExcelBridgeSupport.ToInt(errObj.Number);
            if (number == 0) return null;
            var description = errObj.Description as string;
            return (description ?? $"Excel error {number}", number);
        }
        catch
        {
            return null;
        }
    }

    internal static void AssertSaveAsExtension(string workbookPath, string saveAsPath)
    {
        var workbookExt = Path.GetExtension(workbookPath);
        var saveAsExt = Path.GetExtension(saveAsPath);
        if (!string.Equals(workbookExt, saveAsExt, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"save-as extension {saveAsExt} does not match workbook extension {workbookExt}");
        }
    }

    internal static string ClassifyRunError(string message, int? number)
    {
        if (IsMacroNotFoundError(message, number))
        {
            return "macro_not_found";
        }

        if (IsMacroDisabledError(message, number))
        {
            return "macro_disabled";
        }

        return "macro_failed";
    }

    private static bool IsMacroNotFoundError(string message, int? number)
    {
        if (string.IsNullOrWhiteSpace(message)) return false;
        var upper = message.ToUpperInvariant();
        if (upper.Contains("CANNOT RUN THE MACRO") ||
            upper.Contains("SUB OR FUNCTION NOT DEFINED") ||
            upper.Contains("MACRO MAY NOT BE AVAILABLE") ||
            upper.Contains("UNABLE TO RUN"))
        {
            return true;
        }
        if (number == 1004 && upper.Contains("MACRO"))
        {
            return true;
        }
        return false;
    }

    private static bool IsMacroDisabledError(string message, int? number)
    {
        if (string.IsNullOrWhiteSpace(message)) return false;
        var upper = message.ToUpperInvariant();
        if (upper.Contains("SECURITY SETTINGS") && upper.Contains("MACRO"))
        {
            return true;
        }
        if (upper.Contains("MACRO") && upper.Contains("DISABLED"))
        {
            return true;
        }
        if (number == 1004 && upper.Contains("SECURITY"))
        {
            return true;
        }
        return false;
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbookForRun(
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

    private sealed class MacroArg
    {
        public string Type { get; set; } = "string";
        public string Value { get; set; } = "";
    }
}
