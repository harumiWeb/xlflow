using System.Reflection;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelPushServiceTests
{
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
