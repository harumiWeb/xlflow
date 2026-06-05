using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class PushCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakePushService((request, args) =>
        {
            Assert.Equal("push", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\src\modules", args.ModulesDir);
            Assert.Equal(@"C:\work\src\classes", args.ClassesDir);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal(@"C:\work\src\workbook", args.WorkbookDir);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.Equal(@"C:\work\.xlflow\backups", args.BackupRoot);
            Assert.True(args.Folders);
            Assert.Equal("update", args.FolderAnnotation);
            Assert.True(args.DefaultComponentFolders);
            Assert.Equal(@"C:\work\.xlflow\state\push.json", args.StatePath);
            Assert.False(args.Visible);
            Assert.Equal("always", args.BackupMode);
            Assert.True(args.ChangedOnly);
            Assert.True(args.UseSession);
            Assert.True(args.NoSave);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?>
                {
                    ["kind"] = "live_session",
                    ["path"] = args.WorkbookPath,
                },
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = true,
                    ["workbook_path"] = args.WorkbookPath,
                    ["dirty"] = true,
                    ["save_required"] = true,
                    ["live_newer_than_disk"] = true,
                    ["mode"] = "explicit",
                    ["source_of_truth"] = "live_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = args.WorkbookPath,
                    ["session"] = true,
                    ["session_mode"] = "explicit",
                    ["session_requested"] = true,
                    ["auto_session"] = false,
                    ["saved"] = false,
                    ["dirty"] = true,
                    ["needs_save"] = true,
                },
                ["backup"] = new Dictionary<string, object?>
                {
                    ["id"] = (string?)null,
                    ["path"] = (string?)null,
                    ["reason"] = (string?)null,
                    ["mode"] = args.BackupMode,
                },
                ["source"] = new Dictionary<string, object?>
                {
                    ["changed_only"] = args.ChangedOnly,
                    ["changed"] = true,
                    ["state"] = args.StatePath,
                },
            });
        });
        var command = new PushCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push",
            Command = "push",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "BackupRoot": "C:\\work\\.xlflow\\backups",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "StatePath": "C:\\work\\.xlflow\\state\\push.json",
                  "Visible": "false",
                  "BackupMode": "always",
                  "ChangedOnly": "true",
                  "UseSession": "true",
                  "NoSave": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("always", json.RootElement.GetProperty("backup").GetProperty("mode").GetString());
        Assert.True(json.RootElement.GetProperty("source").GetProperty("changed_only").GetBoolean());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new PushCommand(new FakePushService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-missing-workbook",
            Command = "push",
            Payload = JsonDocument.Parse("""{"ModulesDir":"C:\\work\\src\\modules"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("push_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandlePreservesWarningsAndHintsFromService()
    {
        var service = new FakePushService((request, args) =>
        {
            var response = BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "file", ["path"] = args.WorkbookPath },
                ["session"] = new Dictionary<string, object?> { ["active"] = false, ["workbook_path"] = args.WorkbookPath, ["dirty"] = false, ["save_required"] = false, ["mode"] = "none" },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = false, ["session_mode"] = "none", ["session_requested"] = false, ["auto_session"] = false, ["saved"] = true, ["dirty"] = false, ["needs_save"] = false },
                ["backup"] = new Dictionary<string, object?> { ["id"] = (string?)null, ["path"] = (string?)null, ["reason"] = (string?)null, ["mode"] = args.BackupMode },
                ["source"] = new Dictionary<string, object?> { ["changed_only"] = false, ["changed"] = true, ["state"] = args.StatePath },
                ["logs"] = new List<string> { "imported 2 source file(s)" },
            });
            response.Extensions["warnings"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_state_partial", ["message"] = "UserForms detected: MyForm." },
                new() { ["code"] = "save_required", ["message"] = "Source files were pushed to the live workbook." },
            };
            response.Extensions["hints"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_planned_commands", ["message"] = "UserForm workflow: ..." },
            };
            return response;
        });
        var command = new PushCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-warnings",
            Command = "push",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "BackupRoot": "C:\\work\\.xlflow\\backups",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "StatePath": "C:\\work\\.xlflow\\state\\push.json",
                  "Visible": "false",
                  "BackupMode": "always",
                  "ChangedOnly": "false",
                  "UseSession": "false",
                  "NoSave": "false"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        var warnings = json.RootElement.GetProperty("warnings");
        Assert.Equal(2, warnings.GetArrayLength());
        Assert.Equal("userform_state_partial", warnings[0].GetProperty("code").GetString());
        Assert.Equal("save_required", warnings[1].GetProperty("code").GetString());
        var hints = json.RootElement.GetProperty("hints");
        Assert.Equal(1, hints.GetArrayLength());
        Assert.Equal("userform_planned_commands", hints[0].GetProperty("code").GetString());
    }

    private sealed class FakePushService(Func<BridgeRequest, PushCommandArguments, BridgeResponse> handler) : IPushService
    {
        public BridgeResponse Execute(BridgeRequest request, PushCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
