using System.Diagnostics.CodeAnalysis;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelAttachService : IAttachService
{
    public BridgeResponse Execute(BridgeRequest request, AttachCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!args.Active)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "attach_args_invalid",
                Message: "--active is required for attach in this version.",
                Phase: "attach",
                Source: "xlflow"));
        }

        object? excel = null;
        object? workbook = null;
        var matchedConfiguredWorkbook = false;

        try
        {
            var configuredPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);
            excel = ExcelBridgeSupport.TryGetRunningExcelApplication();
            if (excel is not null)
            {
                workbook = ExcelBridgeSupport.TryGetActiveWorkbook(excel);
                if (workbook is null)
                {
                    try
                    {
                        workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, configuredPath);
                        matchedConfiguredWorkbook = workbook is not null;
                    }
                    catch
                    {
                        workbook = null;
                    }
                }
                if (workbook is null)
                {
                    ExcelBridgeSupport.ReleaseComObject(excel);
                    excel = null;
                }
            }
            if (excel is null)
            {
                excel = ExcelBridgeSupport.TryGetForegroundExcel();
                if (excel is not null)
                {
                    workbook = ExcelBridgeSupport.TryGetActiveWorkbook(excel);
                    if (workbook is null)
                    {
                        try
                        {
                            workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, configuredPath);
                            matchedConfiguredWorkbook = workbook is not null;
                        }
                        catch
                        {
                            workbook = null;
                        }
                    }
                    if (workbook is null)
                    {
                        ExcelBridgeSupport.ReleaseComObject(excel);
                        excel = null;
                    }
                }
            }
            if (excel is null)
            {
                foreach (var process in ExcelBridgeSupport.GetExcelProcesses())
                {
                    excel = ExcelBridgeSupport.TryGetExcelByProcessId(process.ProcessId);
                    if (excel is null)
                    {
                        continue;
                    }

                    workbook = ExcelBridgeSupport.TryGetActiveWorkbook(excel);
                    if (workbook is null)
                    {
                        try
                        {
                            workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, configuredPath);
                            matchedConfiguredWorkbook = workbook is not null;
                        }
                        catch
                        {
                            workbook = null;
                        }
                    }
                    if (workbook is not null)
                    {
                        break;
                    }

                    ExcelBridgeSupport.ReleaseComObject(excel);
                    excel = null;
                }
            }
            if (excel is null)
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "active_workbook_not_found",
                    Message: "No active Excel workbook is available.",
                    Phase: "attach",
                    Source: "Excel"));
            }
            if (workbook is null)
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "active_workbook_not_found",
                    Message: "No active Excel workbook is available.",
                    Phase: "attach",
                    Source: "Excel"));
            }

            var activePath = ExcelBridgeSupport.TryGetWorkbookFullName(workbook) ?? "";
            if (!ExcelBridgeSupport.PathsEqual(configuredPath, activePath))
            {
                if (!matchedConfiguredWorkbook)
                {
                    ExcelBridgeSupport.ReleaseComObject(workbook);
                    workbook = null;

                    try
                    {
                        workbook = ExcelBridgeSupport.GetOpenWorkbook(excel, configuredPath);
                        matchedConfiguredWorkbook = workbook is not null;
                    }
                    catch
                    {
                        workbook = null;
                    }
                }

                if (!matchedConfiguredWorkbook)
                {
                    foreach (var process in ExcelBridgeSupport.GetExcelProcesses())
                    {
                        var candidateExcel = ExcelBridgeSupport.TryGetExcelByProcessId(process.ProcessId);
                        if (candidateExcel is null)
                        {
                            continue;
                        }

                        try
                        {
                            var candidateWorkbook = ExcelBridgeSupport.GetOpenWorkbook(candidateExcel, configuredPath);
                            ExcelBridgeSupport.ReleaseComObject(excel);
                            excel = candidateExcel;
                            workbook = candidateWorkbook;
                            matchedConfiguredWorkbook = true;
                            break;
                        }
                        catch
                        {
                            ExcelBridgeSupport.ReleaseComObject(candidateExcel);
                        }
                    }
                }

                if (!matchedConfiguredWorkbook)
                {
                    return BridgeResponse.Failed(request, new BridgeError(
                        Code: "active_workbook_mismatch",
                        Message: "Active workbook does not match configured workbook: " + activePath,
                        Phase: "attach",
                        Source: "Excel"));
                }
            }

            ExcelBridgeSupport.WriteSessionMetadata(args.MetadataPath, excel, configuredPath);

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = [$"attached xlflow session to active workbook {configuredPath}"],
                Extensions = new Dictionary<string, object?>
                {
                    ["target"] = ExcelBridgeSupport.BuildTargetPayload("live_session", configuredPath),
                    ["session"] = ExcelBridgeSupport.BuildSessionPayload(configuredPath, true, "explicit", false, false),
                    ["workbook"] = ExcelBridgeSupport.BuildWorkbookPayload(configuredPath, true, "explicit", true, false, false, false),
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "active_workbook_not_found",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "attach",
                Source: ex.Source ?? "Excel",
                Number: ex.HResult));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }
}
