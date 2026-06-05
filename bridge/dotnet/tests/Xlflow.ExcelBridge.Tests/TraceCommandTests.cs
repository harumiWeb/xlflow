using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class TraceCommandTests
{
    [Theory]
    [InlineData("enable")]
    [InlineData("inject")]
    [InlineData("status")]
    [InlineData("disable")]
    public void HandleRejectsMissingWorkbookPathForWorkbookActions(string action)
    {
        var serviceCalled = false;
        var command = new TraceCommand(new FakeTraceService((request, _) =>
        {
            serviceCalled = true;
            return BridgeResponse.Ok(request);
        }));
        var request = Request(action);

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.False(serviceCalled);
        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("trace_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandleAllowsCleanWithoutWorkbookPath()
    {
        var serviceCalled = false;
        var command = new TraceCommand(new FakeTraceService((request, args) =>
        {
            serviceCalled = true;
            Assert.Equal("clean", args.Action);
            return BridgeResponse.Ok(request);
        }));

        var response = command.Handle(Request("clean"), CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.True(serviceCalled);
        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleRejectsUnsupportedActionBeforeCallingService()
    {
        var serviceCalled = false;
        var command = new TraceCommand(new FakeTraceService((request, _) =>
        {
            serviceCalled = true;
            return BridgeResponse.Ok(request);
        }));

        var response = command.Handle(Request("unknown"), CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.False(serviceCalled);
        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("trace_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private static BridgeRequest Request(string action)
    {
        return new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-trace-" + action,
            Command = "trace",
            Payload = JsonDocument.Parse($$"""
                {
                  "Action": "{{action}}"
                }
                """).RootElement.Clone(),
        };
    }

    private sealed class FakeTraceService(Func<BridgeRequest, TraceCommandArguments, BridgeResponse> handler) : ITraceService
    {
        public BridgeResponse Execute(BridgeRequest request, TraceCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
