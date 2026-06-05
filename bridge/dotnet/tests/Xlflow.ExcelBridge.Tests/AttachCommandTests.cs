using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class AttachCommandTests
{
    [Fact]
    public void HandleParsesPayload()
    {
        var service = new FakeAttachService((request, args) =>
        {
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.True(args.Active);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);
            return BridgeResponse.Ok(request);
        });

        var command = new AttachCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-attach",
            Command = "attach",
            Payload = JsonDocument.Parse("""{"WorkbookPath":"C:\\work\\Book.xlsm","Active":"true","MetadataPath":"C:\\work\\.xlflow\\session.json"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new AttachCommand(new FakeAttachService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-attach-bad",
            Command = "attach",
            Payload = JsonDocument.Parse("""{"Active":"true"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);
        Assert.Equal("attach_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeAttachService(Func<BridgeRequest, AttachCommandArguments, BridgeResponse> handler) : IAttachService
    {
        public BridgeResponse Execute(BridgeRequest request, AttachCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
