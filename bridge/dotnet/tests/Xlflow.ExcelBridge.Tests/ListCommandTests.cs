using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ListCommandTests
{
    [Fact]
    public void HandleParsesPayload()
    {
        var service = new FakeListService((request, args) =>
        {
            Assert.Equal("forms", args.Action);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal(@"C:\work", args.ProjectRoot);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);
            return BridgeResponse.Ok(request);
        });

        var command = new ListCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-list",
            Command = "list",
            Payload = JsonDocument.Parse("""{"Action":"forms","WorkbookPath":"C:\\work\\Book.xlsm","FormsDir":"C:\\work\\src\\forms","ProjectRoot":"C:\\work","UseSession":"true","MetadataPath":"C:\\work\\.xlflow\\session.json"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    private sealed class FakeListService(Func<BridgeRequest, ListCommandArguments, BridgeResponse> handler) : IListService
    {
        public BridgeResponse Execute(BridgeRequest request, ListCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
