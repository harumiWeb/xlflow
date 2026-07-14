using System.Reflection;
using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormExportImageCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeFormExportImageService((request, args) =>
        {
            Assert.Equal("form-export-image", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal("UserForm1", args.FormName);
            Assert.Equal(@"C:\work\out\UserForm1.png", args.OutputPath);
            Assert.Equal("InitializePreview", args.Initializer);
            Assert.True(args.Overwrite);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "live_session", ["path"] = args.WorkbookPath, ["form"] = args.FormName },
                ["session"] = new Dictionary<string, object?> { ["active"] = true, ["workbook_path"] = args.WorkbookPath },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = true },
                ["forms"] = new Dictionary<string, object?> { ["name"] = args.FormName, ["basis"] = "runtime" },
                ["output"] = new Dictionary<string, object?> { ["path"] = args.OutputPath, ["format"] = "png" },
            });
        });
        var command = new FormExportImageCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-export-image",
            Command = "form-export-image",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "FormName": "UserForm1",
                  "OutputPath": "C:\\work\\out\\UserForm1.png",
                  "Initializer": "InitializePreview",
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
        Assert.Equal("runtime", json.RootElement.GetProperty("forms").GetProperty("basis").GetString());
    }

    [Fact]
    public void HandleRejectsMissingFormNameBeforeService()
    {
        var command = new FormExportImageCommand(new FakeFormExportImageService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-export-image-invalid",
            Command = "form-export-image",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "OutputPath": "C:\\work\\out\\UserForm1.png"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("form_export_image_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void CaptureHelperKeepsTheRuntimeCaptionInsteadOfMaskingItWithDesignerState()
    {
        var method = typeof(ExcelFormExportImageService).GetMethod("BuildHelperCode", BindingFlags.NonPublic | BindingFlags.Static);

        Assert.NotNull(method);
        var helperCode = (string)method.Invoke(null, null)!;

        Assert.Contains("caption = CStr(xlflowCapturedForm.Caption)", helperCode, StringComparison.Ordinal);
        Assert.Contains("xlflowCapturedForm.Caption = caption & \" [xlflow-capture-\"", helperCode, StringComparison.Ordinal);
        Assert.DoesNotContain("xlflowExpectedCaption", helperCode, StringComparison.Ordinal);
        Assert.DoesNotContain("expectedCaption", helperCode, StringComparison.Ordinal);
    }

    private sealed class FakeFormExportImageService(Func<BridgeRequest, FormExportImageCommandArguments, BridgeResponse> handler) : IFormExportImageService
    {
        public BridgeResponse Execute(BridgeRequest request, FormExportImageCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
