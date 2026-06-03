using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelPullServiceTests
{
    [Theory]
    [InlineData(".bas", true)]
    [InlineData(".cls", true)]
    [InlineData(".frm", true)]
    [InlineData(".frx", false)]
    public void IsTextFileExtension_ClassifiesCorrectly(string ext, bool expected)
    {
        Assert.Equal(expected, ExcelPullService.IsTextFileExtension(ext));
    }

    [Fact]
    public void Execute_ReturnsBridgeFileNotOpenableForNonExcelFile()
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmpDir);
            var filePath = Path.Combine(tmpDir, "test.txt");
            File.WriteAllBytes(filePath, [0]);

            var service = new ExcelPullService();
            var request = new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-test",
                Command = "pull",
            };
            var args = new PullCommandArguments(
                WorkbookPath: filePath,
                ModulesDir: Path.Combine(tmpDir, "modules"),
                ClassesDir: Path.Combine(tmpDir, "classes"),
                FormsDir: Path.Combine(tmpDir, "forms"),
                WorkbookDir: Path.Combine(tmpDir, "workbook"),
                CodeSource: "sidecar",
                Folders: false,
                FolderAnnotation: "update",
                DefaultComponentFolders: false,
                Visible: false,
                UseSession: false,
                MetadataPath: "");

            var response = service.Execute(request, args, CancellationToken.None);

            Assert.Equal("failed", response.Status);
            Assert.NotNull(response.Error);
            Assert.Equal("bridge_file_not_openable", response.Error.Code);
        }
        finally
        {
            if (Directory.Exists(tmpDir))
            {
                try { Directory.Delete(tmpDir, true); }
                catch (IOException) { /* best-effort */ }
            }
        }
    }
}
