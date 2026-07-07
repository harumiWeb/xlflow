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
            Assert.False(args.Active);
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

    [Fact]
    public void HandleParsesNonAsciiPayload()
    {
        const string workbookPath = @"C:\work\てすと\sources\帳票.xlsm";
        const string metadataPath = @"C:\work\てすと\.xlflow\session.json";

        var service = new FakeSessionService((request, args) =>
        {
            Assert.Equal("status", args.Action);
            Assert.Equal(workbookPath, args.WorkbookPath);
            Assert.Equal(metadataPath, args.MetadataPath);
            return BridgeResponse.Ok(request);
        });

        var command = new SessionCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session-nonascii",
            Command = "session",
            Payload = JsonDocument.Parse("""{"Action":"status","WorkbookPath":"C:\\work\\てすと\\sources\\帳票.xlsm","MetadataPath":"C:\\work\\てすと\\.xlflow\\session.json","Visible":"false","UseSession":"false"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void ReadSessionMetadataKeepsCompatibilityWithLegacyShape()
    {
        var dir = Path.Combine(Path.GetTempPath(), "xlflow-session-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "session.json");
            File.WriteAllText(path, """{"hwnd":123,"pid":456,"workbook_path":"C:\\work\\Book.xlsm"}""");

            var metadata = ExcelBridgeSupport.ReadSessionMetadata(path);

            Assert.NotNull(metadata);
            Assert.Equal(123, metadata!.Hwnd);
            Assert.Equal(456, metadata.Pid);
            Assert.False(metadata.Poisoned);
            Assert.Equal("managed", metadata.Owner);
            Assert.Equal("", metadata.HResult);
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
    }

    [Fact]
    public void ReadSessionMetadataPreservesNonAsciiWorkbookPath()
    {
        var dir = Path.Combine(Path.GetTempPath(), "xlflow-session-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "session.json");
            var workbookPath = Path.Combine(dir, "てすと", "sources", "帳票.xlsm");
            Directory.CreateDirectory(Path.GetDirectoryName(workbookPath)!);
            File.WriteAllText(path, $$"""{"hwnd":123,"pid":456,"workbook_path":"{{workbookPath.Replace("\\", "\\\\", StringComparison.Ordinal)}}"}""");

            var metadata = ExcelBridgeSupport.ReadSessionMetadata(path);

            Assert.NotNull(metadata);
            Assert.Equal(workbookPath, metadata!.WorkbookPath);
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
    }

    [Fact]
    public void ReadSessionMetadataPreservesExternalOwner()
    {
        var dir = Path.Combine(Path.GetTempPath(), "xlflow-session-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "session.json");
            File.WriteAllText(path, """{"hwnd":123,"pid":456,"workbook_path":"C:\\work\\Book.xlsm","owner":"external"}""");

            var metadata = ExcelBridgeSupport.ReadSessionMetadata(path);

            Assert.NotNull(metadata);
            Assert.Equal("external", metadata!.Owner);
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
    }

    [Fact]
    public void MarkSessionPoisonedBlocksSessionReuse()
    {
        var dir = Path.Combine(Path.GetTempPath(), "xlflow-session-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "session.json");
            var workbook = Path.Combine(dir, "Book.xlsm");
            File.WriteAllText(path, $$"""{"hwnd":123,"pid":456,"workbook_path":"{{workbook.Replace("\\", "\\\\", StringComparison.Ordinal)}}"}""");

            ExcelBridgeSupport.MarkSessionPoisoned(path, workbook, "RPC failed", "0x800706BE", "run");
            var metadata = ExcelBridgeSupport.ReadSessionMetadata(path);

            Assert.NotNull(metadata);
            Assert.True(metadata!.Poisoned);
            Assert.Equal("0x800706BE", metadata.HResult);
            var ex = Assert.Throws<SessionPoisonedException>(() =>
                ExcelBridgeSupport.ThrowIfSessionPoisoned(path, workbook));
            Assert.Equal("run", ex.Metadata.LastCommand);
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
    }

    [Fact]
    public void RunPhaseDoesNotWrapSessionPoisonedException()
    {
        var metadata = new SessionMetadata(
            Hwnd: 123,
            Pid: 456,
            WorkbookPath: @"C:\work\Book.xlsm",
            Owner: "managed",
            Poisoned: true,
            PoisonedAt: "2026-06-07T00:00:00.0000000Z",
            PoisonReason: "RPC failed",
            HResult: "0x800706BE",
            LastCommand: "run");

        var ex = Assert.Throws<SessionPoisonedException>(() =>
            ExcelBridgeSupport.RunPhase("open_workbook", () => throw new SessionPoisonedException(metadata)));

        Assert.Equal("0x800706BE", ex.Metadata.HResult);
    }

    private sealed class FakeSessionService(Func<BridgeRequest, SessionCommandArguments, BridgeResponse> handler) : ISessionService
    {
        public BridgeResponse Execute(BridgeRequest request, SessionCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
