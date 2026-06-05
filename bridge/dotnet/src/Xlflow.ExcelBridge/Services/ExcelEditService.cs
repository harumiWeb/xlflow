using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelEditService : IEditService
{
    public BridgeResponse Execute(BridgeRequest request, EditCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(args.Action))
        {
            return Failure(request, "edit_args_invalid", "-Action is required.");
        }
        if (args.Action is not ("cell" or "range" or "rows" or "columns"))
        {
            return Failure(request, "edit_args_invalid", "Unsupported edit action: " + args.Action);
        }
        if (!args.UseSession)
        {
            return Failure(request, "session_required", "`xlflow edit` requires an active session. Run `xlflow session start` first.");
        }

        object? excel = null;
        object? workbook = null;
        object? worksheet = null;
        object? targetRange = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
        var sessionMode = "explicit";
        bool? dirty = false;
        var needsSave = false;

        try
        {
            var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, args.MetadataPath, true);
            excel = attached.Excel;
            workbook = attached.Workbook;
            sessionMode = attached.SessionMode;
            UpdateSaveState(workbook, ref dirty, ref needsSave);

            worksheet = GetWorksheet(workbook, args.Sheet);
            if (worksheet is null)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "sheet_not_found", $"Sheet '{args.Sheet}' was not found.", "Excel");
            }

            return args.Action switch
            {
                "cell" => EditCell(request, args, excel, workbook, worksheet, workbookPath, sessionMode, dirty, needsSave),
                "range" => EditRange(request, args, workbook, worksheet, workbookPath, sessionMode, dirty, needsSave),
                "rows" => EditRows(request, args, workbook, worksheet, workbookPath, sessionMode, dirty, needsSave),
                _ => EditColumns(request, args, workbook, worksheet, workbookPath, sessionMode, dirty, needsSave),
            };
        }
        catch (Exception ex)
        {
            return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "edit_failed", ExcelBridgeSupport.FormatExceptionDetail(ex), ex.Source ?? "xlflow-excel-bridge");
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(targetRange);
            ExcelBridgeSupport.ReleaseComObject(worksheet);
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse EditCell(BridgeRequest request, EditCommandArguments args, object excel, object workbook, object worksheet, string workbookPath, string sessionMode, bool? dirty, bool needsSave)
    {
        object? range = null;
        try
        {
            try
            {
                dynamic ws = worksheet;
                range = (object?)ws.Range(args.Cell);
            }
            catch (Exception)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_cell_address", $"Cell address '{args.Cell}' is invalid.", "Excel");
            }

            var cellAddress = GetAddress(range!) ?? args.Cell;
            var extensions = BaseEditExtensions(workbookPath, sessionMode, dirty, needsSave, "cell", GetWorksheetName(worksheet), "cell", cellAddress);

            if (!string.IsNullOrWhiteSpace(args.Fill))
            {
                string normalizedFill;
                try
                {
                    normalizedFill = NormalizeColor(args.Fill);
                }
                catch (InvalidOperationException ex)
                {
                    return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_color", ex.Message, "xlflow");
                }

                var interior = ExcelBridgeSupport.Get(range!, "Interior");
                try
                {
                    var beforeFill = ExcelBridgeSupport.ColorToHex(GetMember(interior!, "Color"));
                    SetMember(interior!, "Pattern", 1);
                    SetMember(interior!, "Color", ToOleColor(normalizedFill));
                    UpdateSaveState(workbook, ref dirty, ref needsSave);
                    extensions["edit"] = MergeMutation(
                        (Dictionary<string, object?>)extensions["edit"]!,
                        new Dictionary<string, object?>
                        {
                            ["mutation"] = new Dictionary<string, object?>
                            {
                                ["style"] = new Dictionary<string, object?>
                                {
                                    ["fill"] = new Dictionary<string, object?> { ["before"] = beforeFill, ["after"] = normalizedFill },
                                    ["changed"] = !string.Equals(beforeFill, normalizedFill, StringComparison.OrdinalIgnoreCase),
                                },
                            },
                        });
                    return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"edited {GetWorksheetName(worksheet)}!{cellAddress} fill in the live Excel session", SaveHint]);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(interior);
                }
            }

            var valueRequested = !string.IsNullOrEmpty(args.Value);
            var formulaRequested = !string.IsNullOrEmpty(args.Formula);
            if (valueRequested == formulaRequested)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "edit_args_invalid", "One of -Value, -Formula, or -Fill is required.", "xlflow");
            }

            var eventMode = NormalizeEventMode(args.Events);
            if (eventMode is null)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "edit_args_invalid", "-Events must be keep, on, or off.", "xlflow");
            }

            bool? enableEventsBefore = null;
            bool? enableEventsAfter = null;
            bool? restored = null;
            try
            {
                enableEventsBefore = Convert.ToBoolean(GetMember(excel, "EnableEvents"), CultureInfo.InvariantCulture);
                if (eventMode == "on")
                {
                    SetMember(excel, "EnableEvents", true);
                }
                else if (eventMode == "off")
                {
                    SetMember(excel, "EnableEvents", false);
                }
            }
            catch (Exception)
            {
            }

            var beforeValue = ValueSnapshot(range!);
            var beforeFormula = FormulaSnapshot(range!);
            try
            {
                if (valueRequested)
                {
                    SetMember(range!, "Value2", args.Value);
                }
                else
                {
                    SetMember(range!, "Formula", args.Formula);
                }
            }
            finally
            {
                try
                {
                    enableEventsAfter = Convert.ToBoolean(GetMember(excel, "EnableEvents"), CultureInfo.InvariantCulture);
                    if (enableEventsBefore.HasValue && eventMode != "keep")
                    {
                        SetMember(excel, "EnableEvents", enableEventsBefore.Value);
                        restored = true;
                    }
                    else
                    {
                        restored = true;
                    }
                }
                catch (Exception)
                {
                    restored = false;
                }
            }

            var mutation = valueRequested
                ? new Dictionary<string, object?> { ["value"] = new Dictionary<string, object?> { ["before"] = beforeValue, ["after"] = ValueSnapshot(range!) } }
                : new Dictionary<string, object?> { ["formula"] = new Dictionary<string, object?> { ["before"] = beforeFormula, ["after"] = FormulaSnapshot(range!) } };
            mutation["events"] = new Dictionary<string, object?>
            {
                ["mode"] = eventMode,
                ["enable_events_before"] = enableEventsBefore,
                ["enable_events_after"] = enableEventsAfter,
                ["restored"] = restored,
            };

            UpdateSaveState(workbook, ref dirty, ref needsSave);
            extensions["edit"] = MergeMutation(
                (Dictionary<string, object?>)extensions["edit"]!,
                new Dictionary<string, object?>
                {
                    ["mutation"] = mutation,
                    ["events"] = mutation["events"],
                });

            var mutationLabel = valueRequested ? "value" : "formula";
            return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"edited {GetWorksheetName(worksheet)}!{cellAddress} {mutationLabel} in the live Excel session", SaveHint]);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(range);
        }
    }

    private static BridgeResponse EditRange(BridgeRequest request, EditCommandArguments args, object workbook, object worksheet, string workbookPath, string sessionMode, bool? dirty, bool needsSave)
    {
        object? range = null;
        try
        {
            try
            {
                dynamic ws = worksheet;
                range = (object?)ws.Range(args.RangeAddress);
            }
            catch (Exception)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_range", $"Range '{args.RangeAddress}' is invalid for sheet '{args.Sheet}'.", "Excel");
            }

            var normalizedRange = GetAddress(range!) ?? args.RangeAddress;
            var affectedCells = RangeCellCount(range!);
            var extensions = BaseEditExtensions(workbookPath, sessionMode, dirty, needsSave, "range", GetWorksheetName(worksheet), "range", normalizedRange);

            if (!string.IsNullOrWhiteSpace(args.Fill))
            {
                string normalizedFill;
                try
                {
                    normalizedFill = NormalizeColor(args.Fill);
                }
                catch (InvalidOperationException ex)
                {
                    return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_color", ex.Message, "xlflow");
                }
                var interior = ExcelBridgeSupport.Get(range!, "Interior");
                try
                {
                    var beforeFill = ExcelBridgeSupport.ColorToHex(GetMember(interior!, "Color"));
                    SetMember(interior!, "Pattern", 1);
                    SetMember(interior!, "Color", ToOleColor(normalizedFill));
                    UpdateSaveState(workbook, ref dirty, ref needsSave);
                    extensions["edit"] = MergeMutation(
                        (Dictionary<string, object?>)extensions["edit"]!,
                        new Dictionary<string, object?>
                        {
                            ["mutation"] = new Dictionary<string, object?>
                            {
                                ["style"] = new Dictionary<string, object?>
                                {
                                    ["fill"] = new Dictionary<string, object?> { ["before"] = beforeFill, ["after"] = normalizedFill },
                                },
                                ["affected_cells"] = affectedCells,
                            },
                        });
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(interior);
                }
                return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"edited {GetWorksheetName(worksheet)}!{normalizedRange} fill in the live Excel session", SaveHint]);
            }

            var clearMode = (args.Clear ?? "").Trim().ToLowerInvariant();
            switch (clearMode)
            {
                case "contents":
                    ExcelBridgeSupport.InvokeMethod(range!, "ClearContents");
                    break;
                case "formats":
                    ExcelBridgeSupport.InvokeMethod(range!, "ClearFormats");
                    break;
                case "all":
                    ExcelBridgeSupport.InvokeMethod(range!, "Clear");
                    break;
                default:
                    return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "edit_args_invalid", "-Clear must be contents, formats, or all.", "xlflow");
            }

            UpdateSaveState(workbook, ref dirty, ref needsSave);
            extensions["edit"] = MergeMutation(
                (Dictionary<string, object?>)extensions["edit"]!,
                new Dictionary<string, object?>
                {
                    ["mutation"] = new Dictionary<string, object?>
                    {
                        ["clear"] = new Dictionary<string, object?> { ["mode"] = clearMode },
                        ["affected_cells"] = affectedCells,
                    },
                });
            return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"cleared {clearMode} on {GetWorksheetName(worksheet)}!{normalizedRange} in the live Excel session", SaveHint]);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(range);
        }
    }

    private static BridgeResponse EditRows(BridgeRequest request, EditCommandArguments args, object workbook, object worksheet, string workbookPath, string sessionMode, bool? dirty, bool needsSave)
    {
        object? range = null;
        try
        {
            try
            {
                dynamic ws = worksheet;
                range = (object?)ws.Rows(args.Rows);
            }
            catch (Exception)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_row_selector", $"Row selector '{args.Rows}' is invalid.", "Excel");
            }

            var before = RowHeightSummary(range!);
            SetMember(range!, "RowHeight", double.Parse(args.Height, CultureInfo.InvariantCulture));
            var after = RowHeightSummary(range!);
            UpdateSaveState(workbook, ref dirty, ref needsSave);
            var extensions = BaseEditExtensions(workbookPath, sessionMode, dirty, needsSave, "rows", GetWorksheetName(worksheet), "rows", args.Rows);
            extensions["edit"] = MergeMutation(
                (Dictionary<string, object?>)extensions["edit"]!,
                new Dictionary<string, object?>
                {
                    ["mutation"] = new Dictionary<string, object?>
                    {
                        ["row_height"] = new Dictionary<string, object?> { ["before"] = before, ["after"] = after },
                        ["affected_rows"] = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(ExcelBridgeSupport.Get(range!, "Rows")!, "Count")),
                    },
                });
            return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"edited row height for {GetWorksheetName(worksheet)}!{args.Rows} in the live Excel session", SaveHint]);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(range);
        }
    }

    private static BridgeResponse EditColumns(BridgeRequest request, EditCommandArguments args, object workbook, object worksheet, string workbookPath, string sessionMode, bool? dirty, bool needsSave)
    {
        object? range = null;
        try
        {
            try
            {
                dynamic ws = worksheet;
                range = (object?)ws.Columns(args.Columns);
            }
            catch (Exception)
            {
                return FailureWithState(request, workbookPath, sessionMode, dirty, needsSave, "invalid_column_selector", $"Column selector '{args.Columns}' is invalid.", "Excel");
            }

            var before = ColumnWidthSummary(range!);
            SetMember(range!, "ColumnWidth", double.Parse(args.Width, CultureInfo.InvariantCulture));
            var after = ColumnWidthSummary(range!);
            UpdateSaveState(workbook, ref dirty, ref needsSave);
            var extensions = BaseEditExtensions(workbookPath, sessionMode, dirty, needsSave, "columns", GetWorksheetName(worksheet), "columns", args.Columns);
            extensions["edit"] = MergeMutation(
                (Dictionary<string, object?>)extensions["edit"]!,
                new Dictionary<string, object?>
                {
                    ["mutation"] = new Dictionary<string, object?>
                    {
                        ["column_width"] = new Dictionary<string, object?> { ["before"] = before, ["after"] = after },
                        ["affected_columns"] = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(ExcelBridgeSupport.Get(range!, "Columns")!, "Count")),
                    },
                });
            return Success(request, workbookPath, sessionMode, dirty, needsSave, extensions, [$"edited column width for {GetWorksheetName(worksheet)}!{args.Columns} in the live Excel session", SaveHint]);
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(range);
        }
    }

    private static Dictionary<string, object?> BaseEditExtensions(string workbookPath, string sessionMode, bool? dirty, bool needsSave, string kind, string sheet, string selectorName, string selectorValue)
    {
        var edit = new Dictionary<string, object?> { ["kind"] = kind, ["sheet"] = sheet };
        edit[selectorName] = selectorValue;
        return new Dictionary<string, object?>
        {
            ["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", workbookPath),
            ["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, true, sessionMode, dirty, needsSave),
            ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, true, sessionMode, true, false, dirty, needsSave),
            ["edit"] = edit,
        };
    }

    private static BridgeResponse Success(BridgeRequest request, string workbookPath, string sessionMode, bool? dirty, bool needsSave, Dictionary<string, object?> extensions, IReadOnlyList<string> logs)
    {
        if (needsSave)
        {
            extensions["warnings"] = new object[]
            {
                new Dictionary<string, object?>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes.",
                },
            };
        }

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = extensions,
        };
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message)
    {
        return BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: "edit", Source: "xlflow"));
    }

    private static BridgeResponse FailureWithState(BridgeRequest request, string workbookPath, string sessionMode, bool? dirty, bool needsSave, string code, string message, string source)
    {
        var response = BridgeResponse.Failed(request, new BridgeError(Code: code, Message: message, Phase: "edit", Source: source));
        response.Extensions["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", workbookPath);
        response.Extensions["session"] = ExcelBridgeSupport.BuildSessionPayload(workbookPath, true, sessionMode, dirty, needsSave);
        response.Extensions["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(workbookPath, true, sessionMode, true, false, dirty, needsSave);
        return response;
    }

    private static object? GetWorksheet(object workbook, string sheet)
    {
        try
        {
            var worksheets = ExcelBridgeSupport.Get(workbook, "Worksheets");
            try
            {
                dynamic sheets = worksheets!;
                return (object?)sheets.Item(sheet);
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(worksheets);
            }
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static void UpdateSaveState(object workbook, ref bool? dirty, ref bool needsSave)
    {
        if (!ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var workbookDirty))
        {
            dirty = null;
            needsSave = true;
            return;
        }
        dirty = workbookDirty;
        needsSave = workbookDirty;
    }

    private static string NormalizeEventMode(string value)
    {
        var mode = (value ?? "").Trim().ToLowerInvariant();
        if (mode.Length == 0)
        {
            mode = "keep";
        }
        return mode is "keep" or "on" or "off" ? mode : null!;
    }

    private static string NormalizeColor(string value)
    {
        var trimmed = (value ?? "").Trim();
        if (!System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$"))
        {
            throw new InvalidOperationException($"Color '{value}' is not supported in MVP. Use #RGB or #RRGGBB.");
        }
        var hex = trimmed[1..].ToUpperInvariant();
        if (hex.Length == 3)
        {
            hex = string.Concat(hex.Select(ch => new string(ch, 2)));
        }
        return "#" + hex;
    }

    private static int ToOleColor(string normalized)
    {
        var hex = normalized[1..];
        var red = Convert.ToInt32(hex[..2], 16);
        var green = Convert.ToInt32(hex.Substring(2, 2), 16);
        var blue = Convert.ToInt32(hex.Substring(4, 2), 16);
        return red + (green * 256) + (blue * 65536);
    }

    private static object? ValueSnapshot(object range)
    {
        var value = GetMember(range, "Value2");
        return value ?? "";
    }

    private static string? FormulaSnapshot(object range)
    {
        var value = Convert.ToString(GetMember(range, "Formula"), CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(value) || !value.StartsWith('=') ? null : value;
    }

    private static int RangeCellCount(object range)
    {
        try
        {
            dynamic dyn = range;
            var rows = (object?)dyn.Rows;
            var cols = (object?)dyn.Columns;
            try
            {
                return ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(rows!, "Count")) * ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(cols!, "Count"));
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(rows);
                ExcelBridgeSupport.ReleaseComObject(cols);
            }
        }
        catch (Exception)
        {
            return ExcelBridgeSupport.ToInt(GetMember(range, "Count"));
        }
    }

    private static object? RowHeightSummary(object rowRange)
    {
        object? rows = null;
        try
        {
            dynamic dyn = rowRange;
            rows = (object?)dyn.Rows;
            var count = ExcelBridgeSupport.ToInt(GetMember(rows!, "Count"));
            double? first = null;
            for (var i = 1; i <= count; i++)
            {
                object? row = null;
                try
                {
                    dynamic rowsDyn = rows!;
                    row = (object?)rowsDyn.Item(i);
                    var current = ExcelBridgeSupport.ToDouble(GetMember(row!, "RowHeight"));
                    if (first is null)
                    {
                        first = current;
                        continue;
                    }
                    if (Math.Abs(first.Value - current) > 0.0001)
                    {
                        return "mixed";
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(row);
                }
            }
            return first;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(rows);
        }
    }

    private static object? ColumnWidthSummary(object columnRange)
    {
        object? columns = null;
        try
        {
            dynamic dyn = columnRange;
            columns = (object?)dyn.Columns;
            var count = ExcelBridgeSupport.ToInt(GetMember(columns!, "Count"));
            double? first = null;
            for (var i = 1; i <= count; i++)
            {
                object? column = null;
                try
                {
                    dynamic colsDyn = columns!;
                    column = (object?)colsDyn.Item(i);
                    var current = ExcelBridgeSupport.ToDouble(GetMember(column!, "ColumnWidth"));
                    if (first is null)
                    {
                        first = current;
                        continue;
                    }
                    if (Math.Abs(first.Value - current) > 0.0001)
                    {
                        return "mixed";
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(column);
                }
            }
            return first;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(columns);
        }
    }

    private static Dictionary<string, object?> MergeMutation(Dictionary<string, object?> edit, Dictionary<string, object?> extra)
    {
        foreach (var pair in extra)
        {
            edit[pair.Key] = pair.Value;
        }
        return edit;
    }

    private const string SaveHint = "Run `xlflow save --session` to persist changes to disk.";

    private static string GetWorksheetName(object worksheet)
    {
        return Convert.ToString(GetMember(worksheet, "Name"), CultureInfo.InvariantCulture) ?? "";
    }

    private static string? GetAddress(object range)
    {
        dynamic dyn = range;
        return Convert.ToString(dyn.Address(false, false), CultureInfo.InvariantCulture);
    }

    private static object? GetMember(object comObject, string memberName)
    {
        dynamic dyn = comObject;
        return memberName switch
        {
            "EnableEvents" => dyn.EnableEvents,
            "Value2" => dyn.Value2,
            "Formula" => dyn.Formula,
            "Color" => dyn.Color,
            "Count" => dyn.Count,
            "RowHeight" => dyn.RowHeight,
            "ColumnWidth" => dyn.ColumnWidth,
            "Name" => dyn.Name,
            _ => ExcelBridgeSupport.Get(comObject, memberName),
        };
    }

    private static void SetMember(object comObject, string memberName, object? value)
    {
        dynamic dyn = comObject;
        switch (memberName)
        {
            case "EnableEvents":
                dyn.EnableEvents = value;
                return;
            case "Value2":
                dyn.Value2 = value;
                return;
            case "Formula":
                dyn.Formula = value;
                return;
            case "Pattern":
                dyn.Pattern = value;
                return;
            case "Color":
                dyn.Color = value;
                return;
            case "RowHeight":
                dyn.RowHeight = value;
                return;
            case "ColumnWidth":
                dyn.ColumnWidth = value;
                return;
            default:
                ExcelBridgeSupport.Set(comObject, memberName, value);
                return;
        }
    }
}
