using System.Reflection;
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

        Assert.True(response.Extensions.ContainsKey("push_diagnostic"));
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
}
