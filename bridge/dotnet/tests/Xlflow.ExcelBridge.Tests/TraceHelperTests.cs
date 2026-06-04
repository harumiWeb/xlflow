using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class TraceHelperTests
{
    [Fact]
    public void BuildTraceModuleCodeContainsOptionExplicit()
    {
        var code = TraceHelper.BuildTraceModuleCode();
        Assert.Contains("Option Explicit", code);
    }

    [Fact]
    public void BuildTraceModuleCodeContainsXlflowSetTraceFile()
    {
        var code = TraceHelper.BuildTraceModuleCode();
        Assert.Contains("Public Sub XlflowSetTraceFile(ByVal path As String)", code);
        Assert.Contains("mTraceFile = path", code);
    }

    [Fact]
    public void BuildTraceModuleCodeContainsXlflowLog()
    {
        var code = TraceHelper.BuildTraceModuleCode();
        Assert.Contains("Public Sub XlflowLog(ByVal message As String)", code);
        Assert.Contains("Format$(Now, \"yyyy-mm-dd hh:nn:ss\")", code);
        Assert.Contains("vbTab", code);
    }

    [Fact]
    public void BuildTraceModuleCodeContainsErrorHandler()
    {
        var code = TraceHelper.BuildTraceModuleCode();
        Assert.Contains("On Error GoTo Handler", code);
        Assert.Contains("Handler:", code);
        Assert.Contains("Err.Raise", code);
    }

    [Fact]
    public void GetTraceModuleSourceTextStartsWithAttributeVBName()
    {
        var source = TraceHelper.GetTraceModuleSourceText();
        Assert.StartsWith("Attribute VB_Name = \"XlflowTrace\"", source);
    }

    [Fact]
    public void GetTraceModuleSourceTextContainsModuleCode()
    {
        var source = TraceHelper.GetTraceModuleSourceText();
        Assert.Contains("Public Sub XlflowLog(ByVal message As String)", source);
        Assert.Contains("Public Sub XlflowSetTraceFile(ByVal path As String)", source);
    }

    [Fact]
    public void ReadTraceEventsParsesTabDelimitedLines()
    {
        var tempFile = Path.GetTempFileName();
        try
        {
            File.WriteAllLines(tempFile, new[]
            {
                "2024-01-15 10:30:00\thello world",
                "2024-01-15 10:30:01\tsecond event",
            });

            var events = TraceHelper.ReadTraceEvents(tempFile);

            Assert.Equal(2, events.Count);
            Assert.Equal("2024-01-15 10:30:00", events[0].Timestamp);
            Assert.Equal("hello world", events[0].Message);
            Assert.Equal("2024-01-15 10:30:00\thello world", events[0].Raw);
            Assert.Equal("2024-01-15 10:30:01", events[1].Timestamp);
            Assert.Equal("second event", events[1].Message);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadTraceEventsHandlesLinesWithoutTabs()
    {
        var tempFile = Path.GetTempFileName();
        try
        {
            File.WriteAllLines(tempFile, new[]
            {
                "no tab here",
            });

            var events = TraceHelper.ReadTraceEvents(tempFile);

            Assert.Single(events);
            Assert.Equal("", events[0].Timestamp);
            Assert.Equal("no tab here", events[0].Message);
            Assert.Equal("no tab here", events[0].Raw);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadTraceEventsSkipsEmptyLines()
    {
        var tempFile = Path.GetTempFileName();
        try
        {
            File.WriteAllLines(tempFile, new[]
            {
                "2024-01-15 10:30:00\tevent",
                "",
                "   ",
                "2024-01-15 10:30:01\tevent2",
            });

            var events = TraceHelper.ReadTraceEvents(tempFile);

            Assert.Equal(2, events.Count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void ReadTraceEventsReturnsEmptyForNullPath()
    {
        Assert.Empty(TraceHelper.ReadTraceEvents(null!));
    }

    [Fact]
    public void ReadTraceEventsReturnsEmptyForWhitespacePath()
    {
        Assert.Empty(TraceHelper.ReadTraceEvents("   "));
    }

    [Fact]
    public void ReadTraceEventsReturnsEmptyForNonExistentFile()
    {
        Assert.Empty(TraceHelper.ReadTraceEvents(@"C:\nonexistent\path\trace.log"));
    }

    [Fact]
    public void TraceModuleSourceMatchesReturnsTrueForMatchingSource()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            Directory.CreateDirectory(tempDir);
            var path = Path.Combine(tempDir, "XlflowTrace.bas");
            File.WriteAllText(path, TraceHelper.GetTraceModuleSourceText(), new System.Text.UTF8Encoding(false));

            Assert.True(TraceHelper.TraceModuleSourceMatches(tempDir));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void TraceModuleSourceMatchesReturnsFalseForModifiedSource()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            Directory.CreateDirectory(tempDir);
            var path = Path.Combine(tempDir, "XlflowTrace.bas");
            File.WriteAllText(path, "modified content", new System.Text.UTF8Encoding(false));

            Assert.False(TraceHelper.TraceModuleSourceMatches(tempDir));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void TraceModuleSourceMatchesReturnsFalseForMissingFile()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            Directory.CreateDirectory(tempDir);
            Assert.False(TraceHelper.TraceModuleSourceMatches(tempDir));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void TraceModuleSourceMatchesReturnsFalseForEmptyDir()
    {
        Assert.False(TraceHelper.TraceModuleSourceMatches(""));
        Assert.False(TraceHelper.TraceModuleSourceMatches(null!));
    }

    [Fact]
    public void WriteTraceModuleSourceCreatesFileWithCorrectContent()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "xlflow-test-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var result = TraceHelper.WriteTraceModuleSource(tempDir);

            Assert.NotNull(result);
            Assert.Equal(Path.Combine(tempDir, "XlflowTrace.bas"), result);
            Assert.True(File.Exists(result));

            var content = File.ReadAllText(result);
            Assert.Equal(TraceHelper.GetTraceModuleSourceText(), content);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void WriteTraceModuleSourceReturnsNullForEmptyDir()
    {
        Assert.Null(TraceHelper.WriteTraceModuleSource(""));
        Assert.Null(TraceHelper.WriteTraceModuleSource(null!));
    }
}
