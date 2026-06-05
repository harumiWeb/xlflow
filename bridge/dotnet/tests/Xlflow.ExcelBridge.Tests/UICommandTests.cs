using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class UICommandTests
{
    [Fact]
    public void HandleParsesButtonAddArguments()
    {
        var service = new FakeUIService((request, args) =>
        {
            Assert.Equal("ui", request.Command);
            Assert.Equal("add", args.Action);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.Equal("Sheet1", args.Sheet);
            Assert.Equal("B2", args.Cell);
            Assert.Equal("Run macro", args.Text);
            Assert.Equal("Module1.RunMacro", args.Macro);
            Assert.Equal("run-macro", args.Id);
            Assert.Equal(180, args.Width);
            Assert.Equal(48, args.Height);
            Assert.True(args.CreateSheet);
            Assert.True(args.VerifyMacro);
            Assert.True(args.UseSession);
            Assert.Equal(@".xlflow\session.json", args.MetadataPath);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new UICommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-ui",
            Command = "ui",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath":"C:\\work\\Book.xlsm",
                  "Action":"add",
                  "Sheet":"Sheet1",
                  "Cell":"B2",
                  "Text":"Run macro",
                  "Macro":"Module1.RunMacro",
                  "Id":"run-macro",
                  "Width":"180",
                  "Height":"48",
                  "CreateSheet":"true",
                  "VerifyMacro":"true",
                  "UseSession":"true",
                  "MetadataPath":".xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new UICommand(new FakeUIService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-ui-bad",
            Command = "ui",
            Payload = JsonDocument.Parse("""{"Action":"list"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);
        Assert.Equal("ui_button_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeUIService(Func<BridgeRequest, UICommandArguments, BridgeResponse> handler) : IUIService
    {
        public BridgeResponse Execute(BridgeRequest request, UICommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
