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
            Assert.Equal(@"C:\work", args.ProjectRoot);
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
                { "ProjectRoot": "C:\\work", "BaseWorkbookPath": "C:\\work\\Book.xlsm", "TemporaryDirectory": "C:\\work\\.xlflow\\tmp\\build-1", "PlanJson64": "{{encodedPlan}}", "CodeSource": "sidecar", "Visible": "false" }
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
                root,
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
        var sourcePath = Path.Combine(root, "src", "modules", "Main.bas");
        var plan = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"included":[{"source_path":"src/modules/Main.bas","name":"Main","type":"standard","related_paths":[]}]}"""));
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(sourcePath)!);
            File.WriteAllText(sourcePath, "Attribute VB_Name = \"Main\"");
            var response = new ExcelBuildService().Execute(new BridgeRequest { ProtocolVersion = ProtocolVersion.Current, RequestId = "req-build-plan", Command = "build" }, new BuildCommandArguments(
                root, Path.Combine(root, "Book.xlsm"), Path.Combine(root, "temporary"), plan, "sidecar", false), CancellationToken.None);

            Assert.Equal(BridgeStatus.Failed, response.Status);
            Assert.Contains("base workbook does not exist", response.Error?.Message);
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
    public void ServiceKeepsCallerOwnedTemporaryParentWhenCopyOrOpenFails()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-build-temp-owner-test-" + Guid.NewGuid().ToString("N"));
        var sourcePath = Path.Combine(root, "src", "modules", "Main.bas");
        var tempParent = Path.Combine(root, "caller-temp");
        var sentinel = Path.Combine(tempParent, "keep.txt");
        var baseWorkbook = Path.Combine(root, "Book.txt");
        var plan = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"included":[{"source_path":"src/modules/Main.bas","name":"Main","type":"standard"}]}"""));
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(sourcePath)!);
            Directory.CreateDirectory(tempParent);
            File.WriteAllText(sourcePath, "Attribute VB_Name = \"Main\"");
            File.WriteAllText(baseWorkbook, "not a workbook");
            File.WriteAllText(sentinel, "caller-owned");

            var response = new ExcelBuildService().Execute(new BridgeRequest { ProtocolVersion = ProtocolVersion.Current, RequestId = "req-build-temp-owner", Command = "build" }, new BuildCommandArguments(
                root, baseWorkbook, tempParent, plan, "sidecar", false), CancellationToken.None);

            Assert.Equal(BridgeStatus.Failed, response.Status);
            Assert.True(File.Exists(sentinel));
            Assert.Empty(Directory.GetDirectories(tempParent, "xlflow-build-*"));
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, true);
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
