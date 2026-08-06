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
    [InlineData(".xlsb", true)]
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
        Assert.Equal(2, excel.Workbooks.IntegerItemCalls);
        Assert.Equal(0, excel.Workbooks.StringItemCalls);
    }

    [Fact]
    public void GetOpenWorkbook_ResolvesXlsbThroughNormalWorkbookEnumeration()
    {
        var targetPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "Model.xlsb"));
        var expected = new FakeWorkbook(targetPath);
        var collection = new FakeWorkbookCollection(
            [new FakeWorkbook(Path.Combine(Path.GetTempPath(), "Other.xlsm")), expected],
            new Dictionary<string, FakeWorkbook>(StringComparer.OrdinalIgnoreCase)
            {
                ["Model.xlsb"] = expected,
            });
        var excel = new FakeExcel(collection);

        var actual = ExcelBridgeSupport.GetOpenWorkbook(excel, targetPath);

        Assert.Same(expected, actual);
        Assert.Equal(2, excel.Workbooks.IntegerItemCalls);
        Assert.Equal(0, excel.Workbooks.StringItemCalls);
    }

    [Fact]
    public void GetOpenWorkbook_ResolvesHiddenAddInThroughDirectFilenameLookup()
    {
        var targetPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "selfmacros.xlam"));
        var expected = new FakeWorkbook(targetPath);
        var collection = new FakeWorkbookCollection([], new Dictionary<string, FakeWorkbook>(StringComparer.OrdinalIgnoreCase)
        {
            ["selfmacros.xlam"] = expected,
        });
        var excel = new FakeExcel(collection);

        var actual = ExcelBridgeSupport.GetOpenWorkbook(excel, targetPath);

        Assert.Same(expected, actual);
        Assert.Equal(0, excel.Workbooks.IntegerItemCalls);
        Assert.Equal(1, excel.Workbooks.StringItemCalls);
    }

    [Fact]
    public void GetOpenWorkbook_RejectsDirectFilenameLookupWhenFullPathDiffers()
    {
        var targetPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "project-a", "selfmacros.xlam"));
        var otherPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "project-b", "selfmacros.xlam"));
        var collection = new FakeWorkbookCollection([], new Dictionary<string, FakeWorkbook>(StringComparer.OrdinalIgnoreCase)
        {
            ["selfmacros.xlam"] = new(otherPath),
        });
        var excel = new FakeExcel(collection);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelBridgeSupport.GetOpenWorkbook(excel, targetPath));

        Assert.Contains("xlflow session workbook is not open", ex.Message);
        Assert.Equal(0, excel.Workbooks.IntegerItemCalls);
        Assert.Equal(1, excel.Workbooks.StringItemCalls);
    }

    [Fact]
    public void GetOpenWorkbook_PreservesNotOpenErrorWhenDirectFilenameLookupFails()
    {
        var targetPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "selfmacros.xlam"));
        var excel = new FakeExcel(new FakeWorkbookCollection([]));

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelBridgeSupport.GetOpenWorkbook(excel, targetPath));

        Assert.Contains("xlflow session workbook is not open", ex.Message);
        Assert.Equal(0, excel.Workbooks.IntegerItemCalls);
        Assert.Equal(1, excel.Workbooks.StringItemCalls);
    }

    [Fact]
    public void InvokeViaDynamic_SupportsCopyPictureAndPaste()
    {
        var range = new FakeRange();
        var chart = new FakeChart();

        ExcelBridgeSupport.InvokeViaDynamic(range, "CopyPicture", 2, -4147);
        ExcelBridgeSupport.InvokeViaDynamic(chart, "Paste");

        Assert.Equal((2, -4147), range.LastCopyPictureArgs);
        Assert.True(chart.PasteCalled);
    }

    [Fact]
    public void ConfirmOracleProcessCleanup_AllowsDelayedOwnedExit()
    {
        var owned = new OwnedExcelProcess(1234, new DateTime(2026, 8, 6, 10, 0, 0, DateTimeKind.Utc));
        var discoverCalls = 0;
        var exitCalls = 0;
        var result = ExcelBridgeSupport.ConfirmOracleProcessCleanup(
            new List<OwnedExcelProcess> { owned },
            Array.Empty<OwnedExcelProcess>(),
            () => ++discoverCalls == 1
                ? new[] { owned }
                : Array.Empty<OwnedExcelProcess>(),
            _ => ++exitCalls >= 2,
            _ => { });

        Assert.True(result.Confirmed);
        Assert.Equal(2, result.DrainAttempts);
        Assert.Equal(2, exitCalls);
        Assert.Single(result.OwnedProcesses);
        Assert.True(result.OwnedProcesses[0].ExitConfirmed);
        Assert.Equal(0, result.RemainingProcesses);
        Assert.Null(result.FailureStage);
    }

    [Fact]
    public void ConfirmOracleProcessCleanup_RemainsFailedAfterUnexpectedProcess()
    {
        var owned = new OwnedExcelProcess(1234, new DateTime(2026, 8, 6, 10, 0, 0, DateTimeKind.Utc));
        var unrelated = new OwnedExcelProcess(5678, new DateTime(2026, 8, 6, 10, 1, 0, DateTimeKind.Utc));
        var discoverCalls = 0;
        var result = ExcelBridgeSupport.ConfirmOracleProcessCleanup(
            new List<OwnedExcelProcess> { owned },
            Array.Empty<OwnedExcelProcess>(),
            () => ++discoverCalls == 1
                ? new[] { owned, unrelated }
                : Array.Empty<OwnedExcelProcess>(),
            (_, _) => true,
            ownershipPredicate: candidate => ExcelBridgeSupport.SameOwnedProcess(candidate, owned),
            delay: _ => { });

        Assert.False(result.Confirmed);
        Assert.Equal("unexpected-process", result.FailureStage);
        Assert.Single(result.OwnedProcesses);
        Assert.Equal(0, result.RemainingProcesses);
    }

    [Fact]
    public void SameOwnedProcess_DoesNotMatchKnownStartTimeWithUnknownStartTime()
    {
        var known = new OwnedExcelProcess(1234, new DateTime(2026, 8, 6, 10, 0, 0, DateTimeKind.Utc));
        var unknown = new OwnedExcelProcess(1234);

        Assert.False(ExcelBridgeSupport.SameOwnedProcess(known, unknown));
    }

    public sealed class FakeExcel
    {
        public FakeExcel(params FakeWorkbook[] workbooks)
            : this(new FakeWorkbookCollection(workbooks))
        {
        }

        public FakeExcel(FakeWorkbookCollection workbooks)
        {
            Workbooks = workbooks;
        }

        public FakeWorkbookCollection Workbooks { get; }
    }

    public sealed class FakeWorkbookCollection(
        IReadOnlyList<FakeWorkbook> workbooks,
        IReadOnlyDictionary<string, FakeWorkbook>? directWorkbooks = null)
    {
        public int Count => workbooks.Count;

        public int IntegerItemCalls { get; private set; }

        public int StringItemCalls { get; private set; }

        public FakeWorkbook Item(int index)
        {
            IntegerItemCalls++;
            return workbooks[index - 1];
        }

        public FakeWorkbook Item(string name)
        {
            StringItemCalls++;
            if (directWorkbooks is not null && directWorkbooks.TryGetValue(name, out var workbook))
            {
                return workbook;
            }

            throw new InvalidOperationException("workbook not found");
        }
    }

    public sealed class FakeWorkbook(string fullName)
    {
        public string FullName { get; } = fullName;
    }

    public sealed class FakeRange
    {
        public (object? Appearance, object? Format)? LastCopyPictureArgs { get; private set; }

        public void CopyPicture(object? appearance, object? format)
        {
            LastCopyPictureArgs = (appearance, format);
        }
    }

    public sealed class FakeChart
    {
        public bool PasteCalled { get; private set; }

        public void Paste()
        {
            PasteCalled = true;
        }
    }
}
