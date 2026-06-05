using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Windows;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services intentionally normalize COM failures into structured bridge responses.")]
public sealed class ExcelExportImageService : IExportImageService
{
    public BridgeResponse Execute(BridgeRequest request, ExportImageCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        object? worksheet = null;
        object? range = null;
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        object? activeSheet = null;
        object? selection = null;
        var phase = "validate_args";
        var createdParentDirs = false;
        var sessionAttached = false;
        var sessionMode = "none";
        var originalVisible = false;
        var restoreVisible = false;
        var savedSheetName = "";
        var savedSelectionAddress = "";
        var outputPath = "";
        var temporaryExportPath = "";
        var dirty = false;
        var needsSave = false;
        var resolvedRangeAddress = "";

        try
        {
            phase = "validate_args";
            var normalizedFormat = NormalizeFormat(args.ImageFormat);
            outputPath = PrepareOutputPath(args.OutputPath, args.Overwrite, out createdParentDirs, out temporaryExportPath);

            phase = "open_workbook";
            var open = OpenWorkbook(args.WorkbookPath, args.MetadataPath, args.UseSession, args.Visible);
            excel = open.Excel;
            workbook = open.Workbook;
            sessionAttached = open.SessionAttached;
            sessionMode = open.SessionMode;

            if (sessionAttached && ExcelBridgeSupport.TryGetWorkbookDirtyState(workbook, out dirty))
            {
                needsSave = dirty;
            }

            originalVisible = ReadExcelVisible(excel);
            if (sessionAttached)
            {
                activeSheet = ExcelBridgeSupport.TryGetActiveSheet(excel);
                if (activeSheet is not null)
                {
                    savedSheetName = ExcelBridgeSupport.GetString(activeSheet, "Name") ?? "";
                }
                selection = ExcelBridgeSupport.TryGetSelection(excel);
                if (selection is not null)
                {
                    savedSelectionAddress = ExcelBridgeSupport.GetRangeAddress(selection);
                }
            }

            phase = "resolve_sheet";
            worksheet = GetWorksheet(workbook, args.Sheet);
            if (worksheet is null)
            {
                return Failure(request, "sheet_not_found", $"Sheet '{args.Sheet}' was not found.", "Excel", phase, args, sessionAttached, sessionMode, dirty, needsSave);
            }

            phase = "resolve_range";
            range = ExcelBridgeSupport.Get(worksheet, "Range", args.RangeAddress)
                ?? throw new InvalidOperationException($"invalid_range: Range '{args.RangeAddress}' is invalid for sheet '{args.Sheet}'.");
            resolvedRangeAddress = ExcelBridgeSupport.GetRangeAddress(range);
            if (string.IsNullOrWhiteSpace(resolvedRangeAddress))
            {
                resolvedRangeAddress = args.RangeAddress;
            }

            phase = "activate_window";
            ActivateWorkbookWindow(excel, worksheet, range, originalVisible, ref restoreVisible);

            phase = "create_chart_host";
            if (!TryExportRangeViaChart(range, worksheet, ref chartObjects, ref chartObject, ref chart, outputPath, restoreVisible, originalVisible, excel))
            {
                throw new InvalidOperationException("export_image_failed: Excel did not create the requested image file.");
            }
            if (!string.IsNullOrWhiteSpace(temporaryExportPath))
            {
                File.Move(outputPath, args.OutputPath, true);
                outputPath = args.OutputPath;
                temporaryExportPath = "";
            }

            var output = new Dictionary<string, object?>
            {
                ["path"] = outputPath,
                ["format"] = normalizedFormat,
                ["default"] = args.OutputIsDefault,
            };
            if (createdParentDirs)
            {
                output["created_parent_dirs"] = true;
            }

            var dimensions = TryGetImageDimensions(outputPath);
            if (dimensions is not null)
            {
                output["width_px"] = dimensions.Value.Width;
                output["height_px"] = dimensions.Value.Height;
            }

            var targetPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            var extensions = new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?>
                {
                    ["kind"] = sessionAttached ? "live_session" : "file",
                    ["path"] = targetPath,
                    ["sheet"] = ExcelBridgeSupport.GetString(worksheet, "Name") ?? args.Sheet,
                    ["range"] = resolvedRangeAddress,
                },
                ["session"] = SessionPayload(targetPath, sessionAttached, sessionMode, dirty, needsSave),
                ["workbook"] = WorkbookPayload(targetPath, sessionAttached, sessionMode, dirty, needsSave),
                ["output"] = output,
            };

            var warnings = new List<Dictionary<string, object?>>();
            if (needsSave)
            {
                warnings.Add(new Dictionary<string, object?>
                {
                    ["code"] = "save_required",
                    ["message"] = "The live workbook is newer than disk. `export-image` used the live workbook state, not the saved workbook file.",
                });
            }
            if (warnings.Count > 0)
            {
                extensions["warnings"] = warnings;
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs =
                [
                    sessionAttached ? $"attached to xlflow session ({sessionMode})" : $"opened workbook {targetPath}",
                    $"exported {(ExcelBridgeSupport.GetString(worksheet, "Name") ?? args.Sheet)}!{resolvedRangeAddress} to {outputPath}",
                ],
                Extensions = extensions,
            };
        }
        catch (InvalidOperationException ex)
        {
            return Failure(request, ClassifyErrorCode(ex.Message), ex.Message, "xlflow-excel-bridge", phase, args, sessionAttached, sessionMode, dirty, needsSave);
        }
        catch (Exception ex)
        {
            return Failure(request, "export_image_failed", ExcelBridgeSupport.FormatExceptionDetail(ex), "xlflow-excel-bridge", phase, args, sessionAttached, sessionMode, dirty, needsSave);
        }
        finally
        {
            if (restoreVisible && excel is not null)
            {
                TrySetVisible(excel, originalVisible);
            }

            if (sessionAttached && workbook is not null && excel is not null)
            {
                RestoreSelection(workbook, savedSheetName, savedSelectionAddress);
            }

            if (chartObject is not null)
            {
                try
                {
                    ExcelBridgeSupport.InvokeMethod(chartObject, "Delete");
                }
                catch
                {
                    // best-effort cleanup
                }
            }

            ExcelBridgeSupport.ReleaseComObject(chart);
            ExcelBridgeSupport.ReleaseComObject(chartObject);
            ExcelBridgeSupport.ReleaseComObject(chartObjects);
            ExcelBridgeSupport.ReleaseComObject(range);
            ExcelBridgeSupport.ReleaseComObject(worksheet);
            ExcelBridgeSupport.ReleaseComObject(selection);
            ExcelBridgeSupport.ReleaseComObject(activeSheet);

            if (!string.IsNullOrWhiteSpace(temporaryExportPath) && File.Exists(temporaryExportPath))
            {
                try
                {
                    File.Delete(temporaryExportPath);
                }
                catch
                {
                    // best-effort cleanup
                }
            }

            CloseWorkbook(excel, workbook, sessionAttached);
        }
    }

    private static string NormalizeFormat(string imageFormat)
    {
        var normalized = string.IsNullOrWhiteSpace(imageFormat) ? "png" : imageFormat.Trim().ToLowerInvariant();
        if (!string.Equals(normalized, "png", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"unsupported_image_format: Image format '{imageFormat}' is not supported. Supported formats: png.");
        }
        return normalized;
    }

    private static string PrepareOutputPath(string outputPath, bool overwrite, out bool createdParentDirs, out string temporaryExportPath)
    {
        createdParentDirs = false;
        temporaryExportPath = "";
        var resolved = Path.GetFullPath(outputPath);
        if (Directory.Exists(resolved))
        {
            throw new InvalidOperationException($"export_image_args_invalid: Output path '{resolved}' is a directory.");
        }

        var extension = Path.GetExtension(resolved);
        if (!string.IsNullOrWhiteSpace(extension) && !string.Equals(extension, ".png", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"unsupported_image_format: Image format '{extension.TrimStart('.')}' is not supported. Supported formats: png.");
        }

        var parent = Path.GetDirectoryName(resolved);
        if (!string.IsNullOrWhiteSpace(parent) && !Directory.Exists(parent))
        {
            Directory.CreateDirectory(parent);
            createdParentDirs = true;
        }

        if (File.Exists(resolved) && !overwrite)
        {
            throw new InvalidOperationException($"output_file_exists: Output file '{resolved}' already exists. Use --overwrite to replace it.");
        }

        if (File.Exists(resolved) && overwrite)
        {
            temporaryExportPath = Path.Combine(parent ?? Path.GetTempPath(), $"xlflow-export-{Guid.NewGuid():N}.png");
            return temporaryExportPath;
        }

        return resolved;
    }

    private static (object Excel, object Workbook, bool SessionAttached, string SessionMode) OpenWorkbook(string workbookPath, string metadataPath, bool useSession, bool visible)
    {
        if (useSession)
        {
            var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, true);
            return (attached.Excel, attached.Workbook, true, attached.SessionMode);
        }
        if (ExcelBridgeSupport.SessionMetadataMatchesWorkbook(metadataPath, workbookPath))
        {
            try
            {
                var attached = ExcelBridgeSupport.AttachToSessionWorkbook(workbookPath, metadataPath, false);
                return (attached.Excel, attached.Workbook, true, attached.SessionMode);
            }
            catch
            {
                // fall through to direct open
            }
        }

        var direct = ExcelBridgeSupport.OpenWorkbookDirect(workbookPath, visible);
        return (direct.Excel, direct.Workbook, false, direct.SessionMode);
    }

    private static object? GetWorksheet(object workbook, string sheet)
    {
        object? worksheets = null;
        try
        {
            worksheets = ExcelBridgeSupport.Get(workbook, "Worksheets");
            return ExcelBridgeSupport.Get(worksheets!, "Item", sheet);
        }
        catch
        {
            return null;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(worksheets);
        }
    }

    private static bool ReadExcelVisible(object excel)
    {
        return ExcelBridgeSupport.TryGetVisible(excel, out var visible) && visible;
    }

    private static void TrySetVisible(object excel, bool visible)
    {
        try
        {
            ExcelBridgeSupport.Set(excel, "Visible", visible);
        }
        catch
        {
            // best-effort visibility restore
        }
    }

    private static void ActivateWorkbookWindow(object excel, object worksheet, object range, bool originalVisible, ref bool restoreVisible)
    {
        var hwnd = ExcelBridgeSupport.GetExcelMainHwnd(excel);
        if (!originalVisible)
        {
            TrySetVisible(excel, true);
            restoreVisible = true;
        }
        try
        {
            ExcelBridgeSupport.Set(excel, "ScreenUpdating", true);
        }
        catch
        {
            // best-effort repaint enablement
        }
        if (hwnd != 0)
        {
            _ = WindowCapture.MoveWindowIntoCaptureBounds(WindowCapture.GetWindowInfo(new IntPtr(hwnd)));
        }
        try
        {
            ExcelBridgeSupport.InvokeMethod(worksheet, "Activate");
        }
        catch
        {
            // best-effort activation
        }
        try
        {
            ExcelBridgeSupport.InvokeMethod(range, "Select");
        }
        catch
        {
            // best-effort selection
        }
        Thread.Sleep(700);
    }

    private static bool TryExportRangeViaChart(object range, object worksheet, ref object? chartObjects, ref object? chartObject, ref object? chart, string outputPath, bool restoreVisible, bool originalVisible, object excel)
    {
        ExcelBridgeSupport.InvokeViaDynamic(range, "CopyPicture", 2, -4147);
        Thread.Sleep(200);

        dynamic ws = worksheet;
        dynamic rg = range;
        dynamic chartObjectsDynamic = ws.ChartObjects();
        chartObjects = (object)chartObjectsDynamic;
        chartObject = (object)chartObjectsDynamic.Add(
            Convert.ToDouble(rg.Left, CultureInfo.InvariantCulture),
            Convert.ToDouble(rg.Top, CultureInfo.InvariantCulture),
            Math.Max(Convert.ToDouble(rg.Width, CultureInfo.InvariantCulture), 1.0),
            Math.Max(Convert.ToDouble(rg.Height, CultureInfo.InvariantCulture), 1.0));
        ((dynamic)chartObject).Name = $"xlflow.export.{Guid.NewGuid():N}";
        chart = (object)((dynamic)chartObject).Chart;
        ExcelBridgeSupport.InvokeViaDynamic(chart, "Paste");
        Thread.Sleep(200);

        if (restoreVisible)
        {
            TrySetVisible(excel, originalVisible);
        }

        var exportOk = ExcelBridgeSupport.InvokeMethod(chart, "Export", outputPath, "PNG");
        return (exportOk is bool exported && exported) || File.Exists(outputPath);
    }

    private static (int Width, int Height)? TryGetImageDimensions(string path)
    {
        try
        {
            using var image = Image.FromFile(path);
            return (image.Width, image.Height);
        }
        catch
        {
            return null;
        }
    }

    private static void RestoreSelection(object workbook, string savedSheetName, string savedSelectionAddress)
    {
        if (string.IsNullOrWhiteSpace(savedSheetName))
        {
            return;
        }

        object? savedSheet = null;
        object? savedSelectionRange = null;
        try
        {
            savedSheet = GetWorksheet(workbook, savedSheetName);
            if (savedSheet is null)
            {
                return;
            }
            ExcelBridgeSupport.InvokeMethod(savedSheet, "Activate");
            if (!string.IsNullOrWhiteSpace(savedSelectionAddress))
            {
                savedSelectionRange = ExcelBridgeSupport.Get(savedSheet, "Range", savedSelectionAddress);
                if (savedSelectionRange is not null)
                {
                    ExcelBridgeSupport.InvokeMethod(savedSelectionRange, "Select");
                }
            }
        }
        catch
        {
            // best-effort selection restore
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(savedSelectionRange);
            ExcelBridgeSupport.ReleaseComObject(savedSheet);
        }
    }

    private static Dictionary<string, object?> SessionPayload(string path, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["active"] = sessionAttached,
            ["workbook_path"] = path,
            ["dirty"] = dirty,
            ["save_required"] = needsSave,
            ["live_newer_than_disk"] = needsSave,
            ["mode"] = sessionMode,
            ["source_of_truth"] = needsSave ? "live_workbook" : "saved_workbook",
        };
    }

    private static Dictionary<string, object?> WorkbookPayload(string path, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        return new Dictionary<string, object?>
        {
            ["path"] = path,
            ["session"] = sessionAttached,
            ["session_mode"] = sessionMode,
            ["session_requested"] = sessionAttached,
            ["auto_session"] = sessionAttached && string.Equals(sessionMode, "auto", StringComparison.OrdinalIgnoreCase),
            ["dirty"] = dirty,
            ["needs_save"] = needsSave,
        };
    }

    private static BridgeResponse Failure(BridgeRequest request, string code, string message, string source, string phase, ExportImageCommandArguments args, bool sessionAttached, string sessionMode, bool dirty, bool needsSave)
    {
        var path = string.IsNullOrWhiteSpace(args.WorkbookPath) ? null : ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
        var extensions = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(path))
        {
            extensions["target"] = new Dictionary<string, object?>
            {
                ["kind"] = sessionAttached ? "live_session" : "file",
                ["path"] = path,
            };
            extensions["session"] = SessionPayload(path, sessionAttached, sessionMode, dirty, needsSave);
            extensions["workbook"] = WorkbookPayload(path, sessionAttached, sessionMode, dirty, needsSave);
        }

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Status = BridgeStatus.Failed,
            Error = new BridgeError(code, message, phase, source),
            Extensions = extensions,
        };
    }

    private static string ClassifyErrorCode(string message)
    {
        if (message.Contains("clipboard_timeout", StringComparison.OrdinalIgnoreCase))
        {
            return "clipboard_timeout";
        }
        if (message.Contains("clipboard_unavailable", StringComparison.OrdinalIgnoreCase))
        {
            return "clipboard_unavailable";
        }
        if (message.Contains("invalid_range", StringComparison.OrdinalIgnoreCase))
        {
            return "invalid_range";
        }
        if (message.Contains("sheet_not_found", StringComparison.OrdinalIgnoreCase))
        {
            return "sheet_not_found";
        }
        if (message.Contains("output_file_exists", StringComparison.OrdinalIgnoreCase))
        {
            return "output_file_exists";
        }
        if (message.Contains("unsupported_image_format", StringComparison.OrdinalIgnoreCase))
        {
            return "unsupported_image_format";
        }
        return "export_image_failed";
    }

    private static void CloseWorkbook(object? excel, object? workbook, bool sessionAttached)
    {
        if (sessionAttached)
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
            return;
        }

        if (workbook is not null)
        {
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false);
            }
            catch
            {
            }
            ExcelBridgeSupport.ReleaseComObject(workbook);
        }
        if (excel is not null)
        {
            try
            {
                ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit");
            }
            catch
            {
            }
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

}
