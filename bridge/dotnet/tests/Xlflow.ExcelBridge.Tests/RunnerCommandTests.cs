using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class RunnerCommandTests
{
    [Fact]
    public void HandleParsesPayload()
    {
        var service = new FakeRunnerService((request, args) =>
        {
            Assert.Equal("install", args.Action);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.False(args.Visible);
            return BridgeResponse.Ok(request);
        });

        var command = new RunnerCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-runner",
            Command = "runner",
            Payload = JsonDocument.Parse("""{"Action":"install","WorkbookPath":"C:\\work\\Book.xlsm","Visible":"false"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    private sealed class FakeRunnerService(Func<BridgeRequest, RunnerCommandArguments, BridgeResponse> handler) : IRunnerService
    {
        public BridgeResponse Execute(BridgeRequest request, RunnerCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
