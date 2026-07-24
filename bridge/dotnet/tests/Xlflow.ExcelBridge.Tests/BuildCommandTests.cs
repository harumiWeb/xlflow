using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class BuildCommandTests
{
    [Fact]
    public void HandlePassesTheAuthoritativePlanWithoutSourceRediscovery()
    {
        var encodedPlan = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"included":[{"source_path":"C:\\work\\src\\modules\\Main.bas","name":"Main","type":"standard"}]}"""));
        var command = new BuildCommand(new FakeBuildService((request, args) =>
        {
            Assert.Equal("build", request.Command);
            Assert.Equal(@"C:\work\Book.xlsm", args.BaseWorkbookPath);
            Assert.Equal(@"C:\work\.xlflow\tmp\build-1", args.TemporaryDirectory);
            Assert.Equal(encodedPlan, args.PlanJson64);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.False(args.Visible);
            return BridgeResponse.Ok(request, new Dictionary<string, object?> { ["build"] = new Dictionary<string, object?> { ["temporary_reconstruction"] = true } });
        }));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-build",
            Command = "build",
            Payload = JsonDocument.Parse($$"""
                { "BaseWorkbookPath": "C:\\work\\Book.xlsm", "TemporaryDirectory": "C:\\work\\.xlflow\\tmp\\build-1", "PlanJson64": "{{encodedPlan}}", "CodeSource": "sidecar", "Visible": "false" }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal(BridgeStatus.Ok, response.Status);
        Assert.True((bool)((Dictionary<string, object?>)response.Extensions["build"]!)["temporary_reconstruction"]!);
    }

    [Fact]
    public void HandleRejectsIncompleteArguments()
    {
        var response = new BuildCommand(new FakeBuildService((request, _) => BridgeResponse.Ok(request))).Handle(new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-build-invalid",
            Command = "build",
            Payload = JsonDocument.Parse("""{"BaseWorkbookPath":"C:\\work\\Book.xlsm"}""").RootElement.Clone(),
        }, CancellationToken.None);

        Assert.Equal(BridgeStatus.Failed, response.Status);
        Assert.Equal("build_args_invalid", response.Error?.Code);
    }

    [Fact]
    public void ServiceRejectsInvalidPlanBeforeCreatingTemporaryWorkspace()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-build-command-test-" + Guid.NewGuid().ToString("N"));
        var request = new BridgeRequest { ProtocolVersion = ProtocolVersion.Current, RequestId = "req-build-invalid-plan", Command = "build" };
        try
        {
            var response = new ExcelBuildService().Execute(request, new BuildCommandArguments(
                Path.Combine(root, "Book.xlsm"),
                Path.Combine(root, "temporary"),
                "not-base64",
                "sidecar",
                false), CancellationToken.None);

            Assert.Equal(BridgeStatus.Failed, response.Status);
            Assert.Equal("build_reconstruct_failed", response.Error?.Code);
            Assert.False(Directory.Exists(Path.Combine(root, "temporary")));
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }

    [Fact]
    public void ServiceReadsSnakeCasePlannerComponentFields()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-build-plan-test-" + Guid.NewGuid().ToString("N"));
        var plan = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"included":[{"source_path":"C:\\missing\\Main.bas","name":"Main","type":"standard","related_paths":[]}]}"""));
        try
        {
            var response = new ExcelBuildService().Execute(new BridgeRequest { ProtocolVersion = ProtocolVersion.Current, RequestId = "req-build-plan", Command = "build" }, new BuildCommandArguments(
                Path.Combine(root, "Book.xlsm"), Path.Combine(root, "temporary"), plan, "sidecar", false), CancellationToken.None);

            Assert.Equal(BridgeStatus.Failed, response.Status);
            Assert.Contains("invalid planned component 'Main'", response.Error?.Message);
            Assert.False(Directory.Exists(Path.Combine(root, "temporary")));
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }

    private sealed class FakeBuildService(Func<BridgeRequest, BuildCommandArguments, BridgeResponse> handler) : IBuildService
    {
        public BridgeResponse Execute(BridgeRequest request, BuildCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
