using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ProcessCommandTests
{
    [Fact]
    public void HandleParsesCleanupPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeProcessService((request, args) =>
        {
            Assert.Equal("cleanup", args.Action);
            Assert.Equal(1234, args.TargetPid);
            Assert.False(args.Auto);
            Assert.False(args.All);

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Extensions = new Dictionary<string, object?>
                {
                    ["process"] = new Dictionary<string, object?>
                    {
                        ["action"] = "cleanup",
                        ["mode"] = "pid",
                        ["total"] = 1,
                        ["results"] = new[]
                        {
                            new Dictionary<string, object?>
                            {
                                ["pid"] = 1234,
                                ["terminated"] = true,
                                ["method"] = "graceful",
                            },
                        },
                    },
                },
            };
        });
        var command = new ProcessCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-process-cleanup",
            Command = "process",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "cleanup",
                  "TargetPid": "1234",
                  "Auto": "false",
                  "All": "false"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("cleanup", json.RootElement.GetProperty("process").GetProperty("action").GetString());
        Assert.Equal("pid", json.RootElement.GetProperty("process").GetProperty("mode").GetString());
        Assert.Equal(1, json.RootElement.GetProperty("process").GetProperty("total").GetInt32());
    }

    [Fact]
    public void HandleDefaultsMissingActionToList()
    {
        var service = new FakeProcessService((request, args) =>
        {
            Assert.Equal("list", args.Action);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["process"] = new[]
                {
                    new Dictionary<string, object?>
                    {
                        ["pid"] = 1234,
                        ["has_workbook"] = true,
                    },
                },
            });
        });
        var command = new ProcessCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-process-list",
            Command = "process",
            Payload = JsonDocument.Parse("""{}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal(1234, json.RootElement.GetProperty("process")[0].GetProperty("pid").GetInt32());
        Assert.True(json.RootElement.GetProperty("process")[0].GetProperty("has_workbook").GetBoolean());
    }

    [Fact]
    public void HandleRejectsUnsupportedAction()
    {
        var command = new ProcessCommand(new FakeProcessService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-process-invalid",
            Command = "process",
            Payload = JsonDocument.Parse("""{"Action":"unknown"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("process_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeProcessService(Func<BridgeRequest, ProcessCommandArguments, BridgeResponse> handler) : IProcessService
    {
        public BridgeResponse Execute(BridgeRequest request, ProcessCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
