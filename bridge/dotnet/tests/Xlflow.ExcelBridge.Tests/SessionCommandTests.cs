using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class SessionCommandTests
{
    [Fact]
    public void HandleParsesPayload()
    {
        var service = new FakeSessionService((request, args) =>
        {
            Assert.Equal("start", args.Action);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);
            Assert.True(args.Visible);
            Assert.False(args.UseSession);
            return BridgeResponse.Ok(request);
        });

        var command = new SessionCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session",
            Command = "session",
            Payload = JsonDocument.Parse("""{"Action":"start","WorkbookPath":"C:\\work\\Book.xlsm","MetadataPath":"C:\\work\\.xlflow\\session.json","Visible":"true","UseSession":"false"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    private sealed class FakeSessionService(Func<BridgeRequest, SessionCommandArguments, BridgeResponse> handler) : ISessionService
    {
        public BridgeResponse Execute(BridgeRequest request, SessionCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
