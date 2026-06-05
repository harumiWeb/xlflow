using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class PullCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakePullService((request, args) =>
        {
            Assert.Equal("pull", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\src\modules", args.ModulesDir);
            Assert.Equal(@"C:\work\src\classes", args.ClassesDir);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal(@"C:\work\src\workbook", args.WorkbookDir);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.True(args.Folders);
            Assert.Equal("update", args.FolderAnnotation);
            Assert.True(args.DefaultComponentFolders);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
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
                    ["dirty"] = false,
                    ["save_required"] = false,
                    ["live_newer_than_disk"] = false,
                    ["mode"] = "explicit",
                    ["source_of_truth"] = "saved_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = args.WorkbookPath,
                    ["session"] = true,
                    ["session_mode"] = "explicit",
                    ["session_requested"] = true,
                    ["auto_session"] = false,
                    ["dirty"] = false,
                    ["needs_save"] = false,
                },
                ["source"] = new Dictionary<string, object?>
                {
                    ["modules_dir"] = args.ModulesDir,
                    ["classes_dir"] = args.ClassesDir,
                    ["forms_dir"] = args.FormsDir,
                    ["workbook_dir"] = args.WorkbookDir,
                },
            });
        });
        var command = new PullCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-pull",
            Command = "pull",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("live_session", json.RootElement.GetProperty("target").GetProperty("kind").GetString());
        Assert.Equal(@"C:\work\src\forms", json.RootElement.GetProperty("source").GetProperty("forms_dir").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new PullCommand(new FakePullService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-pull-missing-workbook",
            Command = "pull",
            Payload = JsonDocument.Parse("""{"ModulesDir":"C:\\work\\src\\modules"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("pull_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandlePreservesWarningsAndHintsFromService()
    {
        var service = new FakePullService((request, args) =>
        {
            var response = BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "file", ["path"] = args.WorkbookPath },
                ["session"] = new Dictionary<string, object?> { ["active"] = false, ["workbook_path"] = args.WorkbookPath, ["dirty"] = false, ["save_required"] = false, ["mode"] = "none" },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = false, ["session_mode"] = "none", ["session_requested"] = false, ["auto_session"] = false, ["dirty"] = false, ["needs_save"] = false },
                ["source"] = new Dictionary<string, object?> { ["modules_dir"] = args.ModulesDir, ["classes_dir"] = args.ClassesDir, ["forms_dir"] = args.FormsDir, ["workbook_dir"] = args.WorkbookDir },
                ["logs"] = new List<string> { "exported 5 VBA component(s)" },
            });
            response.Extensions["warnings"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_state_partial", ["message"] = "UserForms detected: MyForm." },
            };
            response.Extensions["hints"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_planned_commands", ["message"] = "UserForm workflow: ..." },
            };
            return response;
        });
        var command = new PullCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-pull-warnings",
            Command = "pull",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "Visible": "false",
                  "UseSession": "false"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        var warnings = json.RootElement.GetProperty("warnings");
        Assert.Equal(1, warnings.GetArrayLength());
        Assert.Equal("userform_state_partial", warnings[0].GetProperty("code").GetString());
        var hints = json.RootElement.GetProperty("hints");
        Assert.Equal(1, hints.GetArrayLength());
        Assert.Equal("userform_planned_commands", hints[0].GetProperty("code").GetString());
    }

    [Fact]
    public void HandlePreservesUserFormUnsavedSessionStateWarning()
    {
        var service = new FakePullService((request, args) =>
        {
            var response = BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "live_session", ["path"] = args.WorkbookPath },
                ["session"] = new Dictionary<string, object?> { ["active"] = true, ["workbook_path"] = args.WorkbookPath, ["dirty"] = true, ["save_required"] = true, ["mode"] = "explicit" },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = true, ["session_mode"] = "explicit", ["session_requested"] = true, ["auto_session"] = false, ["dirty"] = true, ["needs_save"] = true },
                ["source"] = new Dictionary<string, object?> { ["modules_dir"] = args.ModulesDir, ["classes_dir"] = args.ClassesDir, ["forms_dir"] = args.FormsDir, ["workbook_dir"] = args.WorkbookDir },
                ["logs"] = new List<string> { "exported 3 VBA component(s)" },
            });
            response.Extensions["warnings"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_state_partial", ["message"] = "UserForms detected: Form1." },
                new() { ["code"] = "save_required", ["message"] = "The live workbook is newer than disk." },
                new() { ["code"] = "userform_unsaved_session_state", ["message"] = "Workbook contains UserForms (Form1) and the live workbook is newer than disk." },
            };
            response.Extensions["hints"] = new List<Dictionary<string, string>>
            {
                new() { ["code"] = "userform_planned_commands", ["message"] = "UserForm workflow: ..." },
            };
            return response;
        });
        var command = new PullCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-pull-stale",
            Command = "pull",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        var warnings = json.RootElement.GetProperty("warnings");
        Assert.Equal(3, warnings.GetArrayLength());
        Assert.Equal("userform_state_partial", warnings[0].GetProperty("code").GetString());
        Assert.Equal("save_required", warnings[1].GetProperty("code").GetString());
        Assert.Equal("userform_unsaved_session_state", warnings[2].GetProperty("code").GetString());
    }

    private sealed class FakePullService(Func<BridgeRequest, PullCommandArguments, BridgeResponse> handler) : IPullService
    {
        public BridgeResponse Execute(BridgeRequest request, PullCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
