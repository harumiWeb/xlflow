using System.Reflection;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

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
