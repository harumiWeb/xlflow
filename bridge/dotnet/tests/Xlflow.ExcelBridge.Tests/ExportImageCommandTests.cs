using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExportImageCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeExportImageService((request, args) =>
        {
            Assert.Equal("export-image", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal("Sheet1", args.Sheet);
            Assert.Equal("A1:C10", args.RangeAddress);
            Assert.Equal(@"C:\work\out\sheet.png", args.OutputPath);
            Assert.True(args.OutputIsDefault);
            Assert.Equal("png", args.ImageFormat);
            Assert.True(args.Overwrite);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "live_session", ["path"] = args.WorkbookPath, ["sheet"] = args.Sheet, ["range"] = args.RangeAddress },
                ["session"] = new Dictionary<string, object?> { ["active"] = true, ["workbook_path"] = args.WorkbookPath },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = true },
                ["output"] = new Dictionary<string, object?> { ["path"] = args.OutputPath, ["format"] = "png", ["default"] = args.OutputIsDefault },
            });
        });
        var command = new ExportImageCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-export-image",
            Command = "export-image",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "Sheet": "Sheet1",
                  "RangeAddress": "A1:C10",
                  "OutputPath": "C:\\work\\out\\sheet.png",
                  "OutputIsDefault": "true",
                  "ImageFormat": "png",
                  "Overwrite": "true",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal(@"C:\work\out\sheet.png", json.RootElement.GetProperty("output").GetProperty("path").GetString());
    }

    [Fact]
    public void HandleRejectsMissingRangeAddressBeforeService()
    {
        var command = new ExportImageCommand(new FakeExportImageService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-export-image-invalid",
            Command = "export-image",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "Sheet": "Sheet1",
                  "OutputPath": "C:\\work\\out\\sheet.png"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("export_image_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeExportImageService(Func<BridgeRequest, ExportImageCommandArguments, BridgeResponse> handler) : IExportImageService
    {
        public BridgeResponse Execute(BridgeRequest request, ExportImageCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
