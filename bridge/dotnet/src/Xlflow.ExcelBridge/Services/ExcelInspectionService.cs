using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Inspect snapshots intentionally degrade best-effort COM reads into structured bridge responses.")]
public sealed class ExcelInspectionService : IInspectService
{
    public BridgeResponse Execute(BridgeRequest request, InspectCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;

        try
        {
            var attachment = ExcelBridgeSupport.AttachToSessionWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession);
            excel = attachment.Excel;
            workbook = attachment.Workbook;

            var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var dirtyKnown = ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out var dirtyState);
            var dirty = attachment.SessionMode == "none" ? false : dirtyKnown ? dirtyState : true;
            var needsSave = attachment.SessionMode == "none" ? false : dirty;

            var target = NewTargetResult("live_session", workbookPath);
            var session = NewSessionResult(workbookPath, attachment.SessionMode, dirty, needsSave);
            var workbookResult = NewWorkbookResult(workbookPath, attachment.SessionMode, dirty, needsSave);

            Dictionary<string, object?> inspect;
            IReadOnlyList<string> logs;

            switch (args.Target)
            {
                case "workbook":
                    {
                        var sheets = ReadSheetSummaries(workbook);
                        var workbookSummary = new Dictionary<string, object?>
                        {
                            ["path"] = workbookPath,
                            ["name"] = Path.GetFileName(workbookPath),
                            ["sheets"] = sheets,
                        };

                        object? activeSheet = null;
                        try
                        {
                            activeSheet = ExcelBridgeSupport.Get(workbook, "ActiveSheet");
                            var activeSheetName = ExcelBridgeSupport.GetString(activeSheet!, "Name");
                            if (!string.IsNullOrWhiteSpace(activeSheetName))
                            {
                                workbookSummary["active_sheet"] = activeSheetName;
                            }
                        }
                        catch
                        {
                            // best-effort active sheet inspection
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(activeSheet);
                        }

                        inspect = new Dictionary<string, object?>
                        {
                            ["target"] = "workbook",
                            ["source"] = "excel_com",
                            ["target_info"] = NewLiveInspectTargetInfo(workbookPath),
                            ["workbook"] = workbookSummary,
                        };
                        logs = [$"inspected live workbook {workbookPath}"];
                        break;
                    }
                case "sheets":
                    {
                        inspect = new Dictionary<string, object?>
                        {
                            ["target"] = "sheets",
                            ["source"] = "excel_com",
                            ["target_info"] = NewLiveInspectTargetInfo(workbookPath),
                            ["sheets"] = ReadSheetSummaries(workbook),
                        };
                        logs = ["inspected live workbook worksheets"];
                        break;
                    }
                case "range":
                    {
                        object? worksheet = null;
                        object? range = null;
                        try
                        {
                            worksheet = GetWorksheetByName(workbook, args.Sheet);
                            range = ExcelBridgeSupport.Get(worksheet!, "Range", args.Address)
                                ?? throw new InvalidOperationException($"failed to resolve range {args.Sheet}!{args.Address}");
                            var normalizedAddress = ExcelBridgeSupport.GetRangeAddress(range);
                            target["sheet"] = args.Sheet;
                            target["range"] = normalizedAddress;

                            inspect = new Dictionary<string, object?>
                            {
                                ["target"] = "range",
                                ["source"] = "excel_com",
                                ["target_info"] = NewLiveInspectTargetInfo(workbookPath),
                                ["range"] = ReadRangeSnapshot(worksheet!, range, normalizedAddress, "", args.MaxRows, args.MaxCols, args.IncludeStyle),
                            };
                            logs = [$"inspected live range {args.Sheet}!{normalizedAddress}"];
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(range);
                            ExcelBridgeSupport.ReleaseComObject(worksheet);
                        }

                        break;
                    }
                case "used-range":
                    {
                        object? worksheet = null;
                        try
                        {
                            worksheet = GetWorksheetByName(workbook, args.Sheet);
                            var usedInfo = GetUsedRangeInfo(worksheet!);
                            target["sheet"] = args.Sheet;
                            if (!string.IsNullOrWhiteSpace(usedInfo.Address))
                            {
                                target["range"] = usedInfo.Address;
                            }

                            Dictionary<string, object?> rangeSnapshot;
                            if (usedInfo.Range is null)
                            {
                                rangeSnapshot = new Dictionary<string, object?>
                                {
                                    ["sheet"] = ExcelBridgeSupport.GetString(worksheet!, "Name") ?? args.Sheet,
                                    ["used_range"] = "",
                                    ["row_count"] = 0,
                                    ["column_count"] = 0,
                                    ["values"] = Array.Empty<object>(),
                                };
                                if (args.IncludeStyle)
                                {
                                    rangeSnapshot["style_included"] = true;
                                    rangeSnapshot["cells"] = Array.Empty<object>();
                                    rangeSnapshot["columns"] = Array.Empty<object>();
                                    rangeSnapshot["rows"] = Array.Empty<object>();
                                    rangeSnapshot["merged_ranges"] = Array.Empty<object>();
                                }
                            }
                            else
                            {
                                rangeSnapshot = ReadRangeSnapshot(worksheet!, usedInfo.Range, "", usedInfo.Address, args.MaxRows, args.MaxCols, args.IncludeStyle);
                            }

                            inspect = new Dictionary<string, object?>
                            {
                                ["target"] = "used-range",
                                ["source"] = "excel_com",
                                ["target_info"] = NewLiveInspectTargetInfo(workbookPath),
                                ["range"] = rangeSnapshot,
                            };
                            logs = [$"inspected live used range for {args.Sheet}"];
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(worksheet);
                        }

                        break;
                    }
                case "cell":
                    {
                        object? worksheet = null;
                        object? cell = null;
                        try
                        {
                            worksheet = GetWorksheetByName(workbook, args.Sheet);
                            cell = ExcelBridgeSupport.Get(worksheet!, "Range", args.Address)
                                ?? throw new InvalidOperationException($"failed to resolve cell {args.Sheet}!{args.Address}");
                            var normalizedAddress = ExcelBridgeSupport.GetRangeAddress(cell);
                            target["sheet"] = args.Sheet;
                            target["range"] = normalizedAddress;

                            var rowCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(ExcelBridgeSupport.Get(cell, "Rows")!, "Count"));
                            var columnCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(ExcelBridgeSupport.Get(cell, "Columns")!, "Count"));
                            if (rowCount * columnCount != 1)
                            {
                                throw new InvalidOperationException($"cell inspect requires a single-cell address, got {normalizedAddress}");
                            }

                            inspect = new Dictionary<string, object?>
                            {
                                ["target"] = "cell",
                                ["source"] = "excel_com",
                                ["target_info"] = NewLiveInspectTargetInfo(workbookPath),
                                ["cell"] = new Dictionary<string, object?>
                                {
                                    ["sheet"] = ExcelBridgeSupport.GetString(worksheet!, "Name") ?? args.Sheet,
                                    ["address"] = normalizedAddress,
                                    ["value"] = ReadCellValue(cell),
                                },
                            };
                            logs = [$"inspected live cell {args.Sheet}!{normalizedAddress}"];
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(cell);
                            ExcelBridgeSupport.ReleaseComObject(worksheet);
                        }

                        break;
                    }
                default:
                    return BridgeResponse.Failed(
                        request,
                        BridgeError.Create("BRIDGE_COMMAND_UNSUPPORTED", $"Inspect target '{args.Target}' is not supported by the .NET bridge.", "bridge.capability"));
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = logs,
                Extensions = new Dictionary<string, object?>
                {
                    ["target"] = target,
                    ["session"] = session,
                    ["workbook"] = workbookResult,
                    ["inspect"] = inspect,
                },
            };
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("xlflow session", StringComparison.OrdinalIgnoreCase))
        {
            return Failure(request, "session_required", ex.Message, "xlflow");
        }
        catch (InvalidOperationException ex) when (ex.Message.StartsWith("sheet '", StringComparison.OrdinalIgnoreCase) && ex.Message.EndsWith("' not found", StringComparison.OrdinalIgnoreCase))
        {
            return Failure(request, "sheet_not_found", ex.Message, "xlflow");
        }
        catch (COMException ex)
        {
            return Failure(request, "inspect_failed", ex.Message, "xlflow-excel-bridge", ex.ErrorCode);
        }
        catch (InvalidOperationException ex) when (ex.Message.StartsWith("cell inspect requires a single-cell address", StringComparison.OrdinalIgnoreCase))
        {
            return Failure(request, "inspect_args_invalid", ex.Message, "xlflow");
        }
        catch (Exception ex)
        {
            return Failure(request, "inspect_failed", ex.Message, "xlflow-excel-bridge");
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message, string source, int? number = null)
    {
        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(code, message, "inspect", source, number),
        };
    }

    private static Dictionary<string, object?> NewTargetResult(string kind, string path)
    {
        return new Dictionary<string, object?>
        {
            ["kind"] = kind,
            ["path"] = path,
            ["description"] = "Workbook currently open through xlflow session",
        };
    }

    private static Dictionary<string, object?> NewSessionResult(string workbookPath, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["active"] = true,
            ["workbook_path"] = workbookPath,
            ["dirty"] = dirty,
            ["save_required"] = needsSave,
            ["live_newer_than_disk"] = needsSave,
            ["mode"] = sessionMode,
            ["source_of_truth"] = needsSave ? "live_workbook" : "saved_workbook",
        };
    }

    private static Dictionary<string, object?> NewWorkbookResult(string workbookPath, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["path"] = workbookPath,
            ["session"] = true,
            ["session_mode"] = sessionMode,
            ["session_requested"] = sessionMode == "explicit",
            ["auto_session"] = sessionMode == "auto",
            ["dirty"] = dirty,
            ["needs_save"] = needsSave,
        };
    }

    private static Dictionary<string, object?> NewLiveInspectTargetInfo(string workbookPath)
    {
        return new Dictionary<string, object?>
        {
            ["kind"] = "live_session",
            ["path"] = workbookPath,
            ["note"] = "This command inspected the live workbook currently open in Excel through xlflow session.",
        };
    }

    private static List<Dictionary<string, object?>> ReadSheetSummaries(object workbook)
    {
        var sheets = new List<Dictionary<string, object?>>();
        dynamic worksheets = workbook;
        var count = Convert.ToInt32(worksheets.Worksheets.Count, CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            object? worksheet = null;
            try
            {
                worksheet = (object?)worksheets.Worksheets.Item(index);
                sheets.Add(NewSheetSummary(worksheet!));
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(worksheet);
            }
        }

        return sheets;
    }

    private static Dictionary<string, object?> NewSheetSummary(object worksheet)
    {
        UsedRangeInfo usedInfo;
        try
        {
            usedInfo = GetUsedRangeInfo(worksheet);
        }
        catch
        {
            usedInfo = new UsedRangeInfo(null, "", 0, 0);
        }
        try
        {
            dynamic sheet = worksheet;
            return new Dictionary<string, object?>
            {
                ["name"] = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "",
                ["index"] = Convert.ToInt32(sheet.Index, CultureInfo.InvariantCulture),
                ["visible"] = ExcelBridgeSupport.VisibleToBool(sheet.Visible),
                ["used_range"] = usedInfo.Address,
                ["row_count"] = usedInfo.RowCount,
                ["column_count"] = usedInfo.ColumnCount,
            };
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(usedInfo.Range);
        }
    }

    private static object GetWorksheetByName(object workbook, string sheetName)
    {
        object? worksheets = null;
        try
        {
            worksheets = ExcelBridgeSupport.Get(workbook, "Worksheets");
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(worksheets!, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? worksheet = null;
                var matched = false;
                try
                {
                    worksheet = ExcelBridgeSupport.Get(worksheets!, "Item", index);
                    var name = ExcelBridgeSupport.GetString(worksheet!, "Name");
                    if (string.Equals(name, sheetName, StringComparison.Ordinal))
                    {
                        matched = true;
                        return worksheet!;
                    }
                }
                finally
                {
                    if (!matched)
                    {
                        ExcelBridgeSupport.ReleaseComObject(worksheet);
                    }
                }
            }
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(worksheets);
        }

        throw new InvalidOperationException($"sheet '{sheetName}' not found");
    }

    private static UsedRangeInfo GetUsedRangeInfo(object worksheet)
    {
        object? usedRange = null;
        object? probe = null;
        try
        {
            dynamic ws = worksheet;
            usedRange = (object?)ws.UsedRange;
            if (usedRange is null)
            {
                return new UsedRangeInfo(null, "", 0, 0);
            }

            dynamic usedRangeDynamic = usedRange;
            try
            {
                var address = ExcelBridgeSupport.GetRangeAddress(usedRange);
                var rowCount = Convert.ToInt32(usedRangeDynamic.Rows.Count, CultureInfo.InvariantCulture);
                var columnCount = Convert.ToInt32(usedRangeDynamic.Columns.Count, CultureInfo.InvariantCulture);
                if (rowCount == 1 && columnCount == 1)
                {
                    probe = (object?)usedRangeDynamic.Cells.Item(1, 1);
                    if (ReadCellValue(probe) is null && ExcelBridgeSupport.TryGetFormula(probe) is null)
                    {
                        ExcelBridgeSupport.ReleaseComObject(usedRange);
                        usedRange = null;
                        return new UsedRangeInfo(null, "", 0, 0);
                    }
                }

                return new UsedRangeInfo(usedRange, address, rowCount, columnCount);
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(probe);
            }
        }
        catch (Exception ex)
        {
            ExcelBridgeSupport.ReleaseComObject(usedRange);
            throw new InvalidOperationException($"failed to inspect used range: {ex.Message}");
        }
    }

    private static Dictionary<string, object?> ReadRangeSnapshot(object worksheet, object range, string rangeAddress, string usedRangeAddress, int maxRows, int maxCols, bool withStyle)
    {
        object? rows = null;
        object? columns = null;
        object? worksheetCells = null;
        try
        {
            rows = ExcelBridgeSupport.Get(range, "Rows");
            columns = ExcelBridgeSupport.Get(range, "Columns");
            worksheetCells = ExcelBridgeSupport.Get(worksheet, "Cells");

            var rowCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(rows!, "Count"));
            var columnCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(columns!, "Count"));
            var returnRows = rowCount;
            var returnCols = columnCount;
            var truncated = false;
            if (maxRows > 0 && rowCount > maxRows)
            {
                returnRows = maxRows;
                truncated = true;
            }

            if (maxCols > 0 && columnCount > maxCols)
            {
                returnCols = maxCols;
                truncated = true;
            }

            var values = new List<List<object?>>();
            var cells = new List<Dictionary<string, object?>>();
            var rowSnapshots = new List<Dictionary<string, object?>>();
            var columnSnapshots = new List<Dictionary<string, object?>>();
            var warnings = new List<string>();
            var mergedRanges = new List<string>();
            var mergedSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var startRow = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(range, "Row"));
            var startCol = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(range, "Column"));

            for (var rowOffset = 0; rowOffset < returnRows; rowOffset++)
            {
                var line = new List<object?>();
                for (var colOffset = 0; colOffset < returnCols; colOffset++)
                {
                    object? cell = null;
                    try
                    {
                        cell = ExcelBridgeSupport.Get(worksheetCells!, "Item", startRow + rowOffset, startCol + colOffset);
                        var cellValue = ReadCellValue(cell);
                        line.Add(cellValue);
                        if (withStyle)
                        {
                            var cellAddress = ExcelBridgeSupport.GetRangeAddress(cell!);
                            string? mergeAddress = null;
                            try
                            {
                                if (ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(cell!, "MergeCells")))
                                {
                                    var mergeArea = ExcelBridgeSupport.Get(cell!, "MergeArea");
                                    try
                                    {
                                        mergeAddress = ExcelBridgeSupport.GetRangeAddress(mergeArea!);
                                        if (!string.IsNullOrWhiteSpace(mergeAddress) && mergedSeen.Add(mergeAddress))
                                        {
                                            mergedRanges.Add(mergeAddress);
                                        }
                                    }
                                    finally
                                    {
                                        ExcelBridgeSupport.ReleaseComObject(mergeArea);
                                    }
                                }
                            }
                            catch
                            {
                                mergeAddress = null;
                            }

                            var style = NewStyleSnapshot(cell!);
                            cells.Add(new Dictionary<string, object?>
                            {
                                ["address"] = cellAddress,
                                ["row"] = startRow + rowOffset,
                                ["column"] = startCol + colOffset,
                                ["value"] = cellValue,
                                ["formula"] = ExcelBridgeSupport.TryGetFormula(cell),
                                ["fill"] = style["fill"],
                                ["font"] = style["font"],
                                ["border"] = style["border"],
                                ["number_format"] = style["number_format"],
                                ["horizontal_alignment"] = style["horizontal_alignment"],
                                ["vertical_alignment"] = style["vertical_alignment"],
                                ["merged"] = mergeAddress is not null,
                                ["merge_range"] = mergeAddress,
                            });
                        }
                    }
                    finally
                    {
                        ExcelBridgeSupport.ReleaseComObject(cell);
                    }
                }

                values.Add(line);
            }

            if (withStyle)
            {
                object? worksheetRows = null;
                object? worksheetColumns = null;
                try
                {
                    worksheetRows = ExcelBridgeSupport.Get(worksheet, "Rows");
                    worksheetColumns = ExcelBridgeSupport.Get(worksheet, "Columns");
                    for (var rowOffset = 0; rowOffset < returnRows; rowOffset++)
                    {
                        object? rowRef = null;
                        try
                        {
                            var rowIndex = startRow + rowOffset;
                            rowRef = ExcelBridgeSupport.Get(worksheetRows!, "Item", rowIndex);
                            rowSnapshots.Add(new Dictionary<string, object?>
                            {
                                ["row"] = rowIndex,
                                ["height"] = Convert.ToDouble(ExcelBridgeSupport.Get(rowRef!, "RowHeight"), CultureInfo.InvariantCulture),
                                ["hidden"] = ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(rowRef!, "Hidden")),
                            });
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(rowRef);
                        }
                    }

                    for (var colOffset = 0; colOffset < returnCols; colOffset++)
                    {
                        object? columnRef = null;
                        try
                        {
                            var colIndex = startCol + colOffset;
                            columnRef = ExcelBridgeSupport.Get(worksheetColumns!, "Item", colIndex);
                            var columnName = ColumnIndexToLetter(colIndex);
                            columnSnapshots.Add(new Dictionary<string, object?>
                            {
                                ["column"] = columnName,
                                ["index"] = colIndex,
                                ["width"] = Convert.ToDouble(ExcelBridgeSupport.Get(columnRef!, "ColumnWidth"), CultureInfo.InvariantCulture),
                                ["hidden"] = ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(columnRef!, "Hidden")),
                            });
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(columnRef);
                        }
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(worksheetRows);
                    ExcelBridgeSupport.ReleaseComObject(worksheetColumns);
                }
            }

            var returnedRange = "";
            if (returnRows > 0 && returnCols > 0)
            {
                object? topLeft = null;
                object? bottomRight = null;
                object? returned = null;
                try
                {
                    topLeft = ExcelBridgeSupport.Get(worksheetCells!, "Item", startRow, startCol);
                    bottomRight = ExcelBridgeSupport.Get(worksheetCells!, "Item", startRow + returnRows - 1, startCol + returnCols - 1);
                    returned = ExcelBridgeSupport.Get(worksheet, "Range", topLeft!, bottomRight!);
                    returnedRange = ExcelBridgeSupport.GetRangeAddress(returned!);
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(returned);
                    ExcelBridgeSupport.ReleaseComObject(topLeft);
                    ExcelBridgeSupport.ReleaseComObject(bottomRight);
                }
            }

            if (truncated)
            {
                warnings.Add($"Output was truncated: selection has {rowCount} row(s) x {columnCount} column(s), returned {returnRows} row(s) x {returnCols} column(s).");
            }

            var snapshot = new Dictionary<string, object?>
            {
                ["sheet"] = ExcelBridgeSupport.GetString(worksheet, "Name") ?? "",
                ["row_count"] = rowCount,
                ["column_count"] = columnCount,
                ["values"] = values,
                ["truncated"] = truncated,
                ["max_rows"] = maxRows,
                ["max_cols"] = maxCols,
                ["returned_range"] = returnedRange,
            };

            if (!string.IsNullOrWhiteSpace(rangeAddress))
            {
                snapshot["range"] = rangeAddress;
            }

            if (!string.IsNullOrWhiteSpace(usedRangeAddress))
            {
                snapshot["used_range"] = usedRangeAddress;
            }

            if (warnings.Count > 0)
            {
                snapshot["warnings"] = warnings;
            }

            if (withStyle)
            {
                snapshot["style_included"] = true;
                snapshot["cells"] = cells;
                snapshot["columns"] = columnSnapshots;
                snapshot["rows"] = rowSnapshots;
                snapshot["merged_ranges"] = mergedRanges;
            }

            return snapshot;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(worksheetCells);
            ExcelBridgeSupport.ReleaseComObject(columns);
            ExcelBridgeSupport.ReleaseComObject(rows);
        }
    }

    private static Dictionary<string, object?> NewStyleSnapshot(object cell)
    {
        object? interior = null;
        object? font = null;
        object? borders = null;
        try
        {
            interior = ExcelBridgeSupport.Get(cell, "Interior");
            font = ExcelBridgeSupport.Get(cell, "Font");
            borders = ExcelBridgeSupport.Get(cell, "Borders");

            var fillType = ExcelBridgeSupport.FillType(ExcelBridgeSupport.Get(interior!, "Pattern"));
            var fillColor = fillType == "none" ? null : ExcelBridgeSupport.ColorToHex(ExcelBridgeSupport.Get(interior!, "Color"));

            Dictionary<string, object?>? fontSnapshot = null;
            try
            {
                fontSnapshot = new Dictionary<string, object?>
                {
                    ["name"] = ExcelBridgeSupport.GetString(font!, "Name") ?? "",
                    ["size"] = Convert.ToDouble(ExcelBridgeSupport.Get(font!, "Size"), CultureInfo.InvariantCulture),
                    ["bold"] = ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(font!, "Bold")),
                    ["italic"] = ExcelBridgeSupport.ToBool(ExcelBridgeSupport.Get(font!, "Italic")),
                    ["color"] = ExcelBridgeSupport.ColorToHex(ExcelBridgeSupport.Get(font!, "Color")),
                };
            }
            catch
            {
                fontSnapshot = null;
            }

            var numberFormat = ExcelBridgeSupport.GetString(cell, "NumberFormat");
            if (string.IsNullOrWhiteSpace(numberFormat))
            {
                numberFormat = null;
            }

            var border = new Dictionary<string, object?>
            {
                ["top"] = NewBorderEdgeSnapshot(borders!, 8),
                ["right"] = NewBorderEdgeSnapshot(borders!, 10),
                ["bottom"] = NewBorderEdgeSnapshot(borders!, 9),
                ["left"] = NewBorderEdgeSnapshot(borders!, 7),
            };

            return new Dictionary<string, object?>
            {
                ["fill"] = fillType == "none" && fillColor is null
                    ? null
                    : new Dictionary<string, object?> { ["type"] = fillType, ["color"] = fillColor },
                ["font"] = fontSnapshot,
                ["border"] = border,
                ["number_format"] = numberFormat,
                ["horizontal_alignment"] = ExcelBridgeSupport.AlignmentName(ExcelBridgeSupport.Get(cell, "HorizontalAlignment"), "horizontal"),
                ["vertical_alignment"] = ExcelBridgeSupport.AlignmentName(ExcelBridgeSupport.Get(cell, "VerticalAlignment"), "vertical"),
            };
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(borders);
            ExcelBridgeSupport.ReleaseComObject(font);
            ExcelBridgeSupport.ReleaseComObject(interior);
        }
    }

    private static Dictionary<string, object?> NewBorderEdgeSnapshot(object borders, int index)
    {
        object? border = null;
        try
        {
            border = ExcelBridgeSupport.Get(borders, "Item", index);
            string? color = null;
            try
            {
                if (ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(border!, "LineStyle")) != -4142)
                {
                    color = ExcelBridgeSupport.ColorToHex(ExcelBridgeSupport.Get(border!, "Color"));
                }
            }
            catch
            {
                color = null;
            }

            return new Dictionary<string, object?>
            {
                ["style"] = ExcelBridgeSupport.BorderStyleName(
                    ExcelBridgeSupport.Get(border!, "LineStyle"),
                    ExcelBridgeSupport.Get(border!, "Weight")),
                ["color"] = color,
            };
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(border);
        }
    }

    private sealed record UsedRangeInfo(object? Range, string Address, int RowCount, int ColumnCount);

    internal static object? ReadCellValue(object? cell)
    {
        var text = ExcelBridgeSupport.TryGetCellText(cell);
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return ExcelBridgeSupport.TryGetCellValue(cell);
    }

    private static string ColumnIndexToLetter(int colIndex)
    {
        var result = "";
        while (colIndex > 0)
        {
            colIndex--;
            result = (char)('A' + colIndex % 26) + result;
            colIndex /= 26;
        }
        return result;
    }
}
