using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class NewCommandTests
{
    [Fact]
    public void HandleParsesWorkbookPath()
    {
        var service = new FakeNewService((request, args) =>
        {
            Assert.Equal("new", request.Command);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new NewCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-new",
            Command = "new",
            Payload = JsonDocument.Parse("""{"WorkbookPath":"C:\\work\\Book.xlsm"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new NewCommand(new FakeNewService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-new-bad",
            Command = "new",
            Payload = JsonDocument.Parse("""{}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);
        Assert.Equal("new_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeNewService(Func<BridgeRequest, NewCommandArguments, BridgeResponse> handler) : INewService
    {
        public BridgeResponse Execute(BridgeRequest request, NewCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
