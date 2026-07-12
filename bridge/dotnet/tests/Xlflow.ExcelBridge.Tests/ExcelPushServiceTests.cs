using System.Reflection;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;
using Xlflow.ExcelBridge.Windows;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelPushServiceTests
{
    [Fact]
    public void Execute_ChangedOnlyNoOpSkipsExcelAndKeepsSessionStateClean()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-push-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            var modulesDir = Path.Combine(root, "src", "modules");
            var classesDir = Path.Combine(root, "src", "classes");
            var formsDir = Path.Combine(root, "src", "forms");
            var workbookDir = Path.Combine(root, "src", "workbook");
            Directory.CreateDirectory(modulesDir);
            Directory.CreateDirectory(classesDir);
            Directory.CreateDirectory(formsDir);
            Directory.CreateDirectory(workbookDir);

            File.WriteAllText(
                Path.Combine(modulesDir, "Module1.bas"),
                "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\nPublic Sub Main()\r\nEnd Sub\r\n");

            var workbookPath = Path.Combine(root, "Book.txt");
            var statePath = Path.Combine(root, ".xlflow", "state", "push.json");
            var fingerprint = VbaSourceHelper.ComputeFingerprint(
                workbookPath,
                modulesDir,
                classesDir,
                formsDir,
                workbookDir,
                "");
            VbaSourceHelper.WriteFingerprintState(fingerprint, statePath);

            var service = new ExcelPushService();
            var request = new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-push-noop",
                Command = "push",
            };
            var args = new PushCommandArguments(
                WorkbookPath: workbookPath,
                ModulesDir: modulesDir,
                ClassesDir: classesDir,
                FormsDir: formsDir,
                WorkbookDir: workbookDir,
                CodeSource: "",
                BackupRoot: Path.Combine(root, ".xlflow", "backups"),
                Folders: false,
                FolderAnnotation: "ignore",
                DefaultComponentFolders: false,
                StatePath: statePath,
                Visible: false,
                BackupMode: "never",
                ChangedOnly: true,
                UseSession: false,
                NoSave: false,
                MetadataPath: Path.Combine(root, ".xlflow", "session.json"));

            var response = service.Execute(request, args, CancellationToken.None);

            Assert.Equal("ok", response.Status);

            var target = Assert.IsType<Dictionary<string, object?>>(response.Extensions["target"]);
            Assert.Equal("file", target["kind"]);

            var session = Assert.IsType<Dictionary<string, object?>>(response.Extensions["session"]);
            Assert.Equal(false, session["active"]);
            Assert.Equal(false, session["save_required"]);
            Assert.Equal("none", session["mode"]);

            var workbook = Assert.IsType<Dictionary<string, object?>>(response.Extensions["workbook"]);
            Assert.Equal(false, workbook["session"]);
            Assert.Equal(false, workbook["needs_save"]);

            var source = Assert.IsType<Dictionary<string, object?>>(response.Extensions["source"]);
            Assert.Equal(true, source["changed_only"]);
            Assert.Equal(false, source["changed"]);
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
    public void ImportVbaComponents_UsesVbComponentsImportInsteadOfVbProjectImport()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-push-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            var modulesDir = Path.Combine(root, "src", "modules");
            Directory.CreateDirectory(modulesDir);
            var modulePath = Path.Combine(modulesDir, "Module1.bas");
            File.WriteAllText(modulePath, "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n");

            var workbook = new FakeWorkbook();
            var args = new PushCommandArguments(
                WorkbookPath: Path.Combine(root, "Book.xlsm"),
                ModulesDir: modulesDir,
                ClassesDir: Path.Combine(root, "src", "classes"),
                FormsDir: Path.Combine(root, "src", "forms"),
                WorkbookDir: Path.Combine(root, "src", "workbook"),
                CodeSource: "",
                BackupRoot: Path.Combine(root, ".xlflow", "backups"),
                Folders: false,
                FolderAnnotation: "ignore",
                DefaultComponentFolders: false,
                StatePath: Path.Combine(root, ".xlflow", "state", "push.json"),
                Visible: false,
                BackupMode: "never",
                ChangedOnly: false,
                UseSession: false,
                NoSave: false,
                MetadataPath: Path.Combine(root, ".xlflow", "session.json"));
            var sourceFiles = VbaSourceHelper.DiscoverSourceFiles(
                args.ModulesDir, args.ClassesDir, args.FormsDir, args.WorkbookDir, args.CodeSource);
            var tmpImportDir = Path.Combine(root, ".tmp-import");

            var method = typeof(ExcelPushService).GetMethod(
                "ImportVbaComponents",
                BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(method);

            var imported = (int)method!.Invoke(null, [workbook, args, sourceFiles, tmpImportDir])!;

            Assert.Equal(1, imported);
            Assert.False(workbook.VBProject.ImportCalled);
            Assert.Single(workbook.VBProject.VBComponents.ImportedPaths);
            var importedPath = workbook.VBProject.VBComponents.ImportedPaths[0];
            Assert.Equal(Path.Combine(tmpImportDir, "Module1.bas"), importedPath);
            Assert.True(File.Exists(importedPath));
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
    public void BuildCompileFailureResponseMarksPushCompileFailureBeforeSave()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-compile-failed",
            Command = "push",
        };
        var args = new PushCommandArguments(
            WorkbookPath: @"C:\work\Book.xlsm",
            ModulesDir: @"C:\work\src\modules",
            ClassesDir: @"C:\work\src\classes",
            FormsDir: @"C:\work\src\forms",
            WorkbookDir: @"C:\work\src\workbook",
            CodeSource: "",
            BackupRoot: @"C:\work\.xlflow\backups",
            Folders: false,
            FolderAnnotation: "ignore",
            DefaultComponentFolders: false,
            StatePath: @"C:\work\.xlflow\state\push.json",
            Visible: false,
            BackupMode: "never",
            ChangedOnly: false,
            UseSession: true,
            NoSave: true,
            MetadataPath: @"C:\work\.xlflow\session.json");
        var dialog = new DialogSnapshot
        {
            Kind = "compile",
            Title = "Microsoft Visual Basic for Applications",
            Text = ["Compile error:", "Expected: expression"],
            Action = "compile_close",
            ActionSucceeded = true,
        };
        var invocation = new WorkerInvocationResult(
            Result: null,
            Dialog: dialog,
            Dialogs: [dialog],
            LocationCapture: new VbeSelectionCapture(
                new ErrorLocation(
                    "high",
                    "vbe.selection",
                    "src/modules/Main.bas",
                    "Main",
                    "module",
                    "CompileError",
                    4,
                    3,
                    4,
                    6,
                    "  x ="),
                [new VbeSelectionCaptureAttempt("before_dialog_action", true)]),
            TimedOut: false,
            WorkerProcessId: 1234);

        var response = ExcelPushService.BuildCompileFailureResponse(
            request,
            args,
            @"C:\work\Book.xlsm",
            sessionAttached: true,
            sessionMode: "explicit",
            invocation);

        Assert.Equal("failed", response.Status);
        Assert.NotNull(response.Error);
        Assert.Equal("vba_compile_failed", response.Error.Code);
        Assert.Equal("compile_vba", response.Error.Phase);
        Assert.Contains("Expected: expression", response.Error.Message);

        var session = Assert.IsType<Dictionary<string, object?>>(response.Extensions["session"]);
        Assert.Equal(true, session["active"]);
        Assert.Equal(true, session["dirty"]);
        Assert.Equal(true, session["save_required"]);

        var workbook = Assert.IsType<Dictionary<string, object?>>(response.Extensions["workbook"]);
        Assert.Equal(false, workbook["saved"]);
        Assert.Equal(true, workbook["needs_save"]);

        var diagnostic = Assert.IsType<Dictionary<string, object?>>(response.Extensions["push_diagnostic"]);
        var location = Assert.IsType<ErrorLocation>(diagnostic["location"]);
        Assert.Equal("src/modules/Main.bas", location.SourcePath);
        Assert.Equal(4, location.Line);
        Assert.Equal("  x =", location.Text);
        Assert.True(diagnostic.ContainsKey("location_capture"));
    }

    [Fact]
    public void BuildCompileFailureResponseIncludesCreatedBackup()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-compile-failed-backup",
            Command = "push",
        };
        var args = PushArgs(backupMode: "always");
        var invocation = new WorkerInvocationResult(
            Result: null,
            Dialog: null,
            Dialogs: [],
            LocationCapture: new VbeSelectionCapture(null, []),
            TimedOut: true,
            WorkerProcessId: 1234);
        var backup = new ExcelPushService.BackupRef(
            "20260712-134201-123-push-a1b2c3",
            @"C:\work\.xlflow\backups\20260712-134201-123-push-a1b2c3\Book.xlsm",
            "before-push",
            "always");

        var response = ExcelPushService.BuildCompileFailureResponse(
            request,
            args,
            @"C:\work\Book.xlsm",
            sessionAttached: false,
            sessionMode: "none",
            backup,
            invocation);

        var payload = Assert.IsType<Dictionary<string, object?>>(response.Extensions["backup"]);
        Assert.Equal(backup.Id, payload["id"]);
        Assert.Equal(backup.Path, payload["path"]);
        Assert.Equal("before-push", payload["reason"]);
        Assert.Equal("always", payload["mode"]);
    }

    [Fact]
    public void WithBackupAddsBackupToGenericFailureResponse()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-failed-backup",
            Command = "push",
        };
        var response = BridgeResponse.Failed(request, new BridgeError(
            Code: "push_failed",
            Message: "import failed",
            Phase: "push",
            Source: "xlflow-excel-bridge"));
        var backup = new ExcelPushService.BackupRef("backup-id", @"C:\work\.xlflow\backups\backup-id\Book.xlsm", "before-push", "always");

        var enriched = ExcelPushService.WithBackup(response, backup);

        Assert.Equal("push_failed", enriched.Error?.Code);
        var payload = Assert.IsType<Dictionary<string, object?>>(enriched.Extensions["backup"]);
        Assert.Equal("backup-id", payload["id"]);
        Assert.Equal(@"C:\work\.xlflow\backups\backup-id\Book.xlsm", payload["path"]);
    }

    [Fact]
    public void CreateBackupRemovesDirectoryWhenSaveCopyAsFails()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-backup-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var workbook = new FakeBackupWorkbook(_ => throw new InvalidOperationException("save copy failed"));

            var ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelPushService.CreateBackup(workbook, Path.Combine(root, ".xlflow", "backups"), Path.Combine(root, "Book.xlsm")));

            Assert.Contains("save copy failed", ex.Message);
            var backupRoot = Path.Combine(root, ".xlflow", "backups");
            Assert.Empty(Directory.Exists(backupRoot) ? Directory.GetDirectories(backupRoot) : []);
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
    public void CreateBackupRemovesDirectoryWhenMetadataWriteFails()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-backup-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var workbook = new FakeBackupWorkbook(path => Directory.CreateDirectory(path));

            Assert.ThrowsAny<Exception>(() =>
                ExcelPushService.CreateBackup(workbook, Path.Combine(root, ".xlflow", "backups"), Path.Combine(root, "metadata.json")));

            var backupRoot = Path.Combine(root, ".xlflow", "backups");
            Assert.Empty(Directory.Exists(backupRoot) ? Directory.GetDirectories(backupRoot) : []);
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
    public void CreateBackupKeepsCompleteBackupAfterSuccess()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-backup-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var workbook = new FakeBackupWorkbook(path => File.WriteAllText(path, "book"));

            var (id, path) = ExcelPushService.CreateBackup(
                workbook,
                Path.Combine(root, ".xlflow", "backups"),
                Path.Combine(root, "Book.xlsm"));

            Assert.True(File.Exists(path));
            Assert.True(File.Exists(Path.Combine(root, ".xlflow", "backups", id, "metadata.json")));
            using var metadata = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, ".xlflow", "backups", id, "metadata.json")));
            Assert.Equal(id, metadata.RootElement.GetProperty("id").GetString());
            Assert.Equal("Book.xlsm", metadata.RootElement.GetProperty("backup_file_path").GetString());
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }

    [Theory]
    [InlineData(2500, 1500)]
    [InlineData(1000, 1000)]
    [InlineData(500, 500)]
    [InlineData(1, 1)]
    public void ResolveCompileTimeoutHonorsProvidedRequestTimeout(int requestTimeoutMs, int expectedTimeoutMs)
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-timeout",
            Command = "push",
            TimeoutMs = requestTimeoutMs,
        };

        var timeout = ExcelPushService.ResolveCompileTimeout(request);

        Assert.Equal(TimeSpan.FromMilliseconds(expectedTimeoutMs), timeout);
    }

    [Fact]
    public void ResolveCompileTimeoutFallsBackToDefaultWhenUnset()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-push-timeout-default",
            Command = "push",
        };

        var timeout = ExcelPushService.ResolveCompileTimeout(request);

        Assert.Equal(TimeSpan.FromMinutes(5), timeout);
    }

    public sealed class FakeWorkbook
    {
        public FakeVBProject VBProject { get; } = new();
    }

    public sealed class FakeVBProject
    {
        public FakeVBComponents VBComponents { get; } = new();

        public bool ImportCalled { get; private set; }

        public void Import(string path)
        {
            ImportCalled = true;
            throw new InvalidOperationException("VBProject.Import should not be used for push imports.");
        }
    }

    public sealed class FakeVBComponents
    {
        public List<string> ImportedPaths { get; } = [];

        public object Import(object path)
        {
            ImportedPaths.Add((string)path);
            return new object();
        }
    }

    public sealed class FakeBackupWorkbook(Action<string> saveCopyAs)
    {
        public object? SaveCopyAs(object path)
        {
            saveCopyAs((string)path);
            return null;
        }
    }

    private static PushCommandArguments PushArgs(string backupMode)
    {
        return new PushCommandArguments(
            WorkbookPath: @"C:\work\Book.xlsm",
            ModulesDir: @"C:\work\src\modules",
            ClassesDir: @"C:\work\src\classes",
            FormsDir: @"C:\work\src\forms",
            WorkbookDir: @"C:\work\src\workbook",
            CodeSource: "",
            BackupRoot: @"C:\work\.xlflow\backups",
            Folders: false,
            FolderAnnotation: "ignore",
            DefaultComponentFolders: false,
            StatePath: @"C:\work\.xlflow\state\push.json",
            Visible: false,
            BackupMode: backupMode,
            ChangedOnly: false,
            UseSession: false,
            NoSave: false,
            MetadataPath: @"C:\work\.xlflow\session.json");
    }
}
