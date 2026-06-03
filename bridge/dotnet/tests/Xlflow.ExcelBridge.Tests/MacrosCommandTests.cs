using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class MacrosCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeMacrosService((request, args) =>
        {
            Assert.Equal("macros", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);
            Assert.True(args.UseSession);
            Assert.False(args.Visible);
            Assert.True(args.RunnableOnly);

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
                    ["source_of_truth"] = "live_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = args.WorkbookPath,
                    ["session"] = true,
                    ["dirty"] = true,
                    ["needs_save"] = true,
                },
                ["macros"] = Array.Empty<object>(),
            });
        });
        var command = new MacrosCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-macros",
            Command = "macros",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json",
                  "UseSession": "true",
                  "Visible": "false",
                  "RunnableOnly": "true"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("live_session", json.RootElement.GetProperty("target").GetProperty("kind").GetString());
        Assert.True(json.RootElement.GetProperty("session").GetProperty("active").GetBoolean());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new MacrosCommand(new FakeMacrosService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-macros-missing-path",
            Command = "macros",
            Payload = JsonDocument.Parse("""{}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("macros_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void SessionExtensionReflectsDirtyStateWhenSessionAttached()
    {
        var service = new FakeMacrosService((request, args) =>
        {
            Assert.True(args.UseSession);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = true,
                    ["dirty"] = true,
                    ["save_required"] = true,
                    ["source_of_truth"] = "live_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["session"] = true,
                    ["dirty"] = true,
                    ["needs_save"] = true,
                },
                ["macros"] = Array.Empty<object>(),
            });
        });
        var command = new MacrosCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-macros-dirty",
            Command = "macros",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "UseSession": "true"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.True(json.RootElement.GetProperty("session").GetProperty("dirty").GetBoolean());
        Assert.True(json.RootElement.GetProperty("session").GetProperty("save_required").GetBoolean());
        Assert.Equal("live_workbook", json.RootElement.GetProperty("session").GetProperty("source_of_truth").GetString());
        Assert.True(json.RootElement.GetProperty("workbook").GetProperty("dirty").GetBoolean());
        Assert.True(json.RootElement.GetProperty("workbook").GetProperty("needs_save").GetBoolean());
    }

    [Fact]
    public void SessionExtensionShowsCleanStateWhenNotSessionAttached()
    {
        var service = new FakeMacrosService((request, args) =>
        {
            Assert.False(args.UseSession);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = false,
                    ["dirty"] = false,
                    ["save_required"] = false,
                    ["source_of_truth"] = "saved_workbook",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["session"] = false,
                    ["dirty"] = false,
                    ["needs_save"] = false,
                },
                ["macros"] = Array.Empty<object>(),
            });
        });
        var command = new MacrosCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-macros-clean",
            Command = "macros",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.False(json.RootElement.GetProperty("session").GetProperty("dirty").GetBoolean());
        Assert.False(json.RootElement.GetProperty("session").GetProperty("save_required").GetBoolean());
        Assert.Equal("saved_workbook", json.RootElement.GetProperty("session").GetProperty("source_of_truth").GetString());
    }

    private sealed class FakeMacrosService(Func<BridgeRequest, MacrosCommandArguments, BridgeResponse> handler) : IMacrosService
    {
        public BridgeResponse Execute(BridgeRequest request, MacrosCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
