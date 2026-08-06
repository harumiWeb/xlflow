using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class VbeOracleCommandTests
{
    [Fact]
    public void HandleAcceptsFixtureJson64AndTimeoutMilliseconds()
    {
        var fixtureJson64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("{}"));
        var command = new VbeOracleCommand(new FakeOracleService((request, args) =>
        {
            Assert.Equal("vbe-oracle", request.Command);
            Assert.Equal(fixtureJson64, args.PlanJson64);
            Assert.Equal(TimeSpan.FromSeconds(12), args.Timeout);
            return BridgeResponse.Ok(request);
        }));

        var response = command.Handle(new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-oracle",
            Command = "vbe-oracle",
            Payload = JsonDocument.Parse($$"""{ "FixtureJSON64": "{{fixtureJson64}}", "TimeoutMS": "12000" }""").RootElement.Clone(),
        }, CancellationToken.None);

        Assert.Equal(BridgeStatus.Ok, response.Status);
    }

    [Fact]
    public void HandleRejectsMissingFixture()
    {
        var response = new VbeOracleCommand(new FakeOracleService((request, _) => BridgeResponse.Ok(request)))
            .Handle(new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-oracle-invalid",
                Command = "vbe-oracle",
                Payload = JsonDocument.Parse("{}").RootElement.Clone(),
            }, CancellationToken.None);

        Assert.Equal(BridgeStatus.Failed, response.Status);
        Assert.Equal("vbe_oracle_args_invalid", response.Error?.Code);
    }

    [Fact]
    public void HandleRejectsNonPositiveTimeout()
    {
        var fixtureJson64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("{}"));
        var response = new VbeOracleCommand(new FakeOracleService((request, _) => BridgeResponse.Ok(request)))
            .Handle(new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-oracle-timeout",
                Command = "vbe-oracle",
                Payload = JsonDocument.Parse($$"""{ "FixtureJSON64": "{{fixtureJson64}}", "TimeoutMS": "0" }""").RootElement.Clone(),
            }, CancellationToken.None);

        Assert.Equal(BridgeStatus.Failed, response.Status);
        Assert.Equal("vbe_oracle_args_invalid", response.Error?.Code);
    }

    private sealed class FakeOracleService(Func<BridgeRequest, VbeOracleCommandArguments, BridgeResponse> handler) : IVbeOracleService
    {
        public BridgeResponse Execute(BridgeRequest request, VbeOracleCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
