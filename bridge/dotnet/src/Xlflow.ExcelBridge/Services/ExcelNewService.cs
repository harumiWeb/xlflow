using System.Diagnostics.CodeAnalysis;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge services normalize Excel COM failures into structured responses.")]
public sealed class ExcelNewService : INewService
{
    public BridgeResponse Execute(BridgeRequest request, NewCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        object? excel = null;
        object? workbook = null;
        var workbookPath = ExcelBridgeSupport.NormalizePath(args.WorkbookPath);

        try
        {
            var parent = Path.GetDirectoryName(workbookPath);
            if (!string.IsNullOrWhiteSpace(parent))
            {
                Directory.CreateDirectory(parent);
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Excel.Application COM class is not registered");
            excel = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Failed to create Excel.Application COM instance");

            dynamic app = excel;
            app.Visible = false;
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Add();
            ExcelBridgeSupport.InvokeViaDynamic(workbook, "SaveAs", workbookPath, FileFormatForWorkbookPath(workbookPath));

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = [$"created workbook {workbookPath}"],
                Extensions = new Dictionary<string, object?>
                {
                    ["workbook"] = new Dictionary<string, object?> { ["path"] = workbookPath },
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "excel_create_failed",
                Message: ExcelBridgeSupport.FormatExceptionDetail(ex),
                Phase: "new",
                Source: ex.Source ?? "xlflow-excel-bridge",
                Number: ex.HResult));
        }
        finally
        {
            if (workbook is not null)
            {
                try { ExcelBridgeSupport.InvokeViaDynamic(workbook, "Close", false); } catch { }
            }
            if (excel is not null)
            {
                try { ExcelBridgeSupport.InvokeViaDynamic(excel, "Quit"); } catch { }
            }
            ExcelBridgeSupport.ReleaseComObject(workbook);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    internal static int FileFormatForWorkbookPath(string workbookPath)
    {
        var extension = Path.GetExtension(workbookPath);
        return extension.ToLowerInvariant() switch
        {
            ".xlsm" => 52,
            ".xlam" => 55,
            _ => throw new InvalidOperationException($"unsupported workbook extension: {extension}"),
        };
    }
}
