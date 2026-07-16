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
            Assert.True(args.Discard);
            return BridgeResponse.Ok(request);
        });

        var command = new SessionCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session",
            Command = "session",
            Payload = JsonDocument.Parse("""{"Action":"start","WorkbookPath":"C:\\work\\Book.xlsm","MetadataPath":"C:\\work\\.xlflow\\session.json","Visible":"true","UseSession":"false","Discard":"true"}""").RootElement.Clone(),
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
            Assert.False(metadata.DiscardUnsavedChanges);
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
    public void ReplacementPoisonRequestsDiscardOnStop()
    {
        var dir = Path.Combine(Path.GetTempPath(), "xlflow-session-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "session.json");
            var workbook = Path.Combine(dir, "Book.xlsm");
            File.WriteAllText(path, $$"""{"hwnd":123,"pid":456,"workbook_path":"{{workbook.Replace("\\", "\\\\", StringComparison.Ordinal)}}"}""");

            ExcelBridgeSupport.MarkSessionPoisoned(
                path,
                workbook,
                "partial VBA replacement",
                "0x800A03EC",
                "push",
                discardUnsavedChanges: true);
            var metadata = Assert.IsType<SessionMetadata>(ExcelBridgeSupport.ReadSessionMetadata(path));

            Assert.True(metadata.Poisoned);
            Assert.True(metadata.DiscardUnsavedChanges);
            Assert.True(ExcelSessionService.ShouldDiscardUnsavedChanges(metadata));
            Assert.True(ExcelSessionService.RequiresExplicitDiscard(metadata, discardRequested: false));
            Assert.False(ExcelSessionService.RequiresExplicitDiscard(metadata, discardRequested: true));
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
    public void StatusDoesNotRequestSaveForDiscardPoisonedSession()
    {
        var metadata = new SessionMetadata(
            Hwnd: 123,
            Pid: 456,
            WorkbookPath: @"C:\work\Book.xlsm",
            Owner: "managed",
            Poisoned: true,
            DiscardUnsavedChanges: true);

        Assert.False(ExcelSessionService.NeedsSaveForStatus(
            running: true,
            open: true,
            dirty: true,
            metadata));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StatusPreservesSaveRequiredForNonDiscardDirtySession(bool poisoned)
    {
        var metadata = new SessionMetadata(
            Hwnd: 123,
            Pid: 456,
            WorkbookPath: @"C:\work\Book.xlsm",
            Owner: "managed",
            Poisoned: poisoned,
            DiscardUnsavedChanges: false);

        Assert.True(ExcelSessionService.NeedsSaveForStatus(
            running: true,
            open: true,
            dirty: true,
            metadata));
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

    [Fact]
    public void ManagedDiscardStopReportsConfirmedRecoveryAndExcelExit()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session-stop-discard",
            Command = "session",
        };
        var recovery = ExcelSessionService.BuildRecovery(
            required: false,
            reason: "managed_session_discarded",
            operation: "session stop",
            excelProcessId: 4321,
            cleanupConfirmed: true,
            sessionActive: false,
            sessionOwner: "managed");

        var response = ExcelSessionService.BuildStopResponse(
            request,
            @"C:\work\Book.xlsm",
            autoSavedOnStop: false,
            removedStaleMetadata: false,
            discardedUnsavedChanges: true,
            excelProcessId: 4321,
            excelProcessExited: true,
            recovery: recovery);
        using var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.False(json.RootElement.GetProperty("recovery").GetProperty("required").GetBoolean());
        Assert.True(json.RootElement.GetProperty("recovery").GetProperty("cleanup_confirmed").GetBoolean());
        Assert.Equal(4321, json.RootElement.GetProperty("excel_process").GetProperty("pid").GetInt32());
        Assert.True(json.RootElement.GetProperty("excel_process").GetProperty("exited").GetBoolean());
        Assert.True(json.RootElement.GetProperty("workbook").GetProperty("discarded_unsaved_changes").GetBoolean());
    }

    [Fact]
    public void ManagedDiscardWithLiveUnreachableProcessFailsAndKeepsRecovery()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session-stop-unconfirmed",
            Command = "session",
        };
        var response = ExcelSessionService.BuildCleanupUnconfirmedResponse(
            request,
            @"C:\work\Book.xlsm",
            excelProcessId: 4321,
            discardedUnsavedChanges: false);
        using var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal(BridgeStatus.Failed, response.Status);
        Assert.Equal("session_stop_cleanup_unconfirmed", response.Error?.Code);
        Assert.True(json.RootElement.GetProperty("recovery").GetProperty("required").GetBoolean());
        Assert.False(json.RootElement.GetProperty("recovery").GetProperty("cleanup_confirmed").GetBoolean());
        Assert.Equal(4321, json.RootElement.GetProperty("recovery").GetProperty("excel_pid").GetInt32());
        Assert.False(json.RootElement.GetProperty("error").GetProperty("details").GetProperty("discarded_unsaved_changes").GetBoolean());
    }

    [Fact]
    public void ExternalDiscardDetachKeepsRecoveryRequired()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-session-stop-external",
            Command = "session",
        };
        var recovery = ExcelSessionService.BuildRecovery(
            required: true,
            reason: "external_session_detached",
            operation: "session stop",
            excelProcessId: 4321,
            cleanupConfirmed: false,
            sessionActive: false,
            sessionOwner: "external");

        var response = ExcelSessionService.BuildStopResponse(
            request,
            @"C:\work\Book.xlsm",
            autoSavedOnStop: false,
            removedStaleMetadata: false,
            detachedExternal: true,
            detachedDirty: true,
            unsafeChangesNotDiscarded: true,
            recovery: recovery);
        using var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.True(json.RootElement.GetProperty("recovery").GetProperty("required").GetBoolean());
        Assert.False(json.RootElement.GetProperty("recovery").GetProperty("cleanup_confirmed").GetBoolean());
        Assert.Equal("external", json.RootElement.GetProperty("recovery").GetProperty("session").GetProperty("owner").GetString());
        Assert.True(json.RootElement.GetProperty("workbook").GetProperty("unsafe_changes_not_discarded").GetBoolean());
    }

    private sealed class FakeSessionService(Func<BridgeRequest, SessionCommandArguments, BridgeResponse> handler) : ISessionService
    {
        public BridgeResponse Execute(BridgeRequest request, SessionCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
