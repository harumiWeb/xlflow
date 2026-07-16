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
            Assert.Equal([2222, 3333], args.SkipWorkbookProbePids.OrderBy(pid => pid));

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
                  "All": "false",
                  "SkipWorkbookProbePids": [2222, "3333", 0, "invalid"]
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

    [Fact]
    public void WorkbookProbePolicySkipsQuarantinedExcelProcessIds()
    {
        IReadOnlySet<int> skipped = new HashSet<int> { 1234 };

        Assert.False(ExcelBridgeSupport.ShouldProbeWorkbook(1234, skipped));
        Assert.True(ExcelBridgeSupport.ShouldProbeWorkbook(5678, skipped));
        Assert.True(ExcelBridgeSupport.ShouldProbeWorkbook(1234, null));
    }

    [Fact]
    public void ProcessListItemMarksSkippedWorkbookProbeAsRecoveryRequired()
    {
        var item = ExcelProcessService.BuildProcessListItem(
            new ExcelProcessInfo(1234, HasWorkbook: null),
            new HashSet<int> { 1234 });

        Assert.Equal(1234, item["pid"]);
        Assert.Null(item["has_workbook"]);
        Assert.Equal(true, item["workbook_probe_skipped"]);
        Assert.Equal(true, item["recovery_required"]);
    }

    [Fact]
    public void ProcessListItemMarksUnaffectedProcessAsNotRecoveryRequired()
    {
        var item = ExcelProcessService.BuildProcessListItem(
            new ExcelProcessInfo(5678, HasWorkbook: true),
            new HashSet<int> { 1234 });

        Assert.Equal(true, item["has_workbook"]);
        Assert.Equal(false, item["workbook_probe_skipped"]);
        Assert.Equal(false, item["recovery_required"]);
    }

    [Fact]
    public void HandleParsesJsonEncodedSkipWorkbookProbePidArray()
    {
        var service = new FakeProcessService((request, args) =>
        {
            Assert.Equal([2222, 3333], args.SkipWorkbookProbePids.OrderBy(pid => pid));
            return BridgeResponse.Ok(request);
        });
        var command = new ProcessCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-process-list-skip-json",
            Command = "process",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "list",
                  "SkipWorkbookProbePids": "[2222,3333,0]"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);

        Assert.Equal(BridgeStatus.Ok, response.Status);
    }

    [Theory]
    [InlineData("not-json")]
    [InlineData("{}")]
    public void HandleTreatsInvalidJsonEncodedSkipWorkbookProbePidsAsEmpty(string value)
    {
        var service = new FakeProcessService((request, args) =>
        {
            Assert.Empty(args.SkipWorkbookProbePids);
            return BridgeResponse.Ok(request);
        });
        var command = new ProcessCommand(service);
        var payload = JsonSerializer.SerializeToElement(new Dictionary<string, object?>
        {
            ["Action"] = "list",
            ["SkipWorkbookProbePids"] = value,
        });
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-process-list-skip-invalid",
            Command = "process",
            Payload = payload,
        };

        var response = command.Handle(request, CancellationToken.None);

        Assert.Equal(BridgeStatus.Ok, response.Status);
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
