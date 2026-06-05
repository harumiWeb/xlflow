using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormWriteCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeFormWriteService((request, args) =>
        {
            Assert.Equal("form-write", request.Command);
            Assert.Equal("build", args.Action);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"src\forms\specs\UserForm1.yaml", args.SpecPath);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.True(args.Folders);
            Assert.Equal("update", args.FolderAnnotation);
            Assert.True(args.DefaultComponentFolders);
            Assert.Equal("eyJmb28iOiJiYXIifQ==", args.SpecJson64);
            Assert.True(args.Overwrite);
            Assert.False(args.NoSave);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "live_session", ["path"] = args.WorkbookPath },
                ["session"] = new Dictionary<string, object?> { ["active"] = true, ["workbook_path"] = args.WorkbookPath },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = true },
                ["forms"] = new Dictionary<string, object?>
                {
                    ["name"] = "UserForm1",
                    ["action"] = args.Action,
                    ["source_synced"] = true,
                },
            });
        });
        var command = new FormWriteCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-write",
            Command = "form-write",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "build",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "SpecPath": "src\\forms\\specs\\UserForm1.yaml",
                  "FormsDir": "C:\\work\\src\\forms",
                  "CodeSource": "sidecar",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "SpecJson64": "eyJmb28iOiJiYXIifQ==",
                  "Overwrite": "true",
                  "NoSave": "false",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("UserForm1", json.RootElement.GetProperty("forms").GetProperty("name").GetString());
        Assert.Equal("build", json.RootElement.GetProperty("forms").GetProperty("action").GetString());
    }

    [Fact]
    public void HandleRejectsInvalidNoSaveUsageBeforeService()
    {
        var command = new FormWriteCommand(new FakeFormWriteService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-write-invalid",
            Command = "form-write",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "build",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "FormsDir": "C:\\work\\src\\forms",
                  "SpecJson64": "eyJmb28iOiJiYXIifQ==",
                  "NoSave": "true",
                  "UseSession": "false"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("form_build_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeFormWriteService(Func<BridgeRequest, FormWriteCommandArguments, BridgeResponse> handler) : IFormWriteService
    {
        public BridgeResponse Execute(BridgeRequest request, FormWriteCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
