using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class TestCommandTests
{
    [Fact]
    public void HandlePassesIsolationAndNoSaveOptionsToService()
    {
        var serviceCalled = false;
        var command = new TestCommand(new FakeTestService((_, args) =>
        {
            serviceCalled = true;
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.Equal("module", args.Isolation);
            Assert.True(args.NoSave);
            Assert.True(args.DisableAutoSession);
            Assert.Equal(@"C:\work\Book.xlsm", args.SourceWorkbookPath);
            Assert.Equal(@"C:\work", args.ProjectRoot);
            Assert.Equal(@"C:\work\.xlflow\test-runs", args.TempRunRoot);
            return BridgeResponse.Ok(new BridgeRequest());
        }));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-options",
            Command = "test",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\Book.xlsm",
                  "Isolation": "module",
                  "NoSave": "true",
                  "DisableAutoSession": "true",
                  "SourceWorkbookPath": "C:\\work\\Book.xlsm",
                  "ProjectRoot": "C:\\work",
                  "TempRunRoot": "C:\\work\\.xlflow\\test-runs"
                }
                """).RootElement.Clone(),
        };

        command.Handle(request, CancellationToken.None);

        Assert.True(serviceCalled);
    }

    private sealed class FakeTestService(Func<BridgeRequest, TestCommandArguments, BridgeResponse> handle) : ITestService
    {
        public BridgeResponse Execute(BridgeRequest request, TestCommandArguments args, CancellationToken cancellationToken)
        {
            return handle(request, args);
        }
    }
}
