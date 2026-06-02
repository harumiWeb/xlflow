using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelBridgeSupportTests
{
    [Theory]
    [InlineData(".xlsm", true)]
    [InlineData(".xlsx", true)]
    [InlineData(".xls", true)]
    [InlineData(".xlt", true)]
    [InlineData(".xla", true)]
    [InlineData(".xlam", true)]
    [InlineData(".xltx", true)]
    [InlineData(".xltm", true)]
    [InlineData(".txt", false)]
    [InlineData(".csv", false)]
    [InlineData(".json", false)]
    [InlineData("", false)]
    public void IsExcelFile_ClassifiesExtensionsCorrectly(string ext, bool expected)
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmpDir);
            var filePath = Path.Combine(tmpDir, "test" + ext);
            File.WriteAllBytes(filePath, [0]);

            Assert.Equal(expected, ExcelBridgeSupport.IsExcelFile(filePath));
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

    [Fact]
    public void IsExcelFile_ReturnsFalseForMissingFile()
    {
        var fakePath = Path.Combine(Path.GetTempPath(), "nonexistent-" + Guid.NewGuid().ToString("N"), "book.xlsm");
        Assert.False(ExcelBridgeSupport.IsExcelFile(fakePath));
    }

    [Fact]
    public void OpenWorkbookDirect_ThrowsBridgeFileNotOpenableForNonExcelFile()
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmpDir);
            var filePath = Path.Combine(tmpDir, "test.txt");
            File.WriteAllBytes(filePath, [0]);

            var ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelBridgeSupport.OpenWorkbookDirect(filePath, false));
            Assert.Contains("bridge_file_not_openable", ex.Message);
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

    [Fact]
    public void OpenWorkbookDirect_ThrowsBridgeFileNotOpenableForMissingFile()
    {
        var fakePath = Path.Combine(Path.GetTempPath(), "nonexistent-" + Guid.NewGuid().ToString("N"), "book.xlsm");

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelBridgeSupport.OpenWorkbookDirect(fakePath, false));
        Assert.Contains("bridge_file_not_openable", ex.Message);
    }

    [Fact]
    public void GetOpenWorkbook_ResolvesWorkbookThroughLateBoundCollection()
    {
        var targetPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "Book.xlsm"));
        var expected = new FakeWorkbook(targetPath);
        var excel = new FakeExcel(new FakeWorkbook(Path.Combine(Path.GetTempPath(), "Other.xlsm")), expected);

        var actual = ExcelBridgeSupport.GetOpenWorkbook(excel, targetPath);

        Assert.Same(expected, actual);
        Assert.Equal(2, excel.Workbooks.ItemCalls);
    }

    public sealed class FakeExcel(params FakeWorkbook[] workbooks)
    {
        public FakeWorkbookCollection Workbooks { get; } = new(workbooks);
    }

    public sealed class FakeWorkbookCollection(params FakeWorkbook[] workbooks)
    {
        public int Count => workbooks.Length;

        public int ItemCalls { get; private set; }

        public FakeWorkbook Item(int index)
        {
            ItemCalls++;
            return workbooks[index - 1];
        }
    }

    public sealed class FakeWorkbook(string fullName)
    {
        public string FullName { get; } = fullName;
    }
}
