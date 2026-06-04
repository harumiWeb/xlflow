using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelTraceServiceTests
{
    [Fact]
    public void CanRemoveTraceSourceRejectsModifiedSourceWithoutForce()
    {
        var modulesDir = Directory.CreateTempSubdirectory("xlflow-trace-modified-").FullName;
        File.WriteAllText(Path.Combine(modulesDir, "XlflowTrace.bas"), "modified");

        Assert.False(ExcelTraceService.CanRemoveTraceSource(modulesDir, force: false));
    }

    [Fact]
    public void CanRemoveTraceSourceAllowsModifiedSourceWithForce()
    {
        var modulesDir = Directory.CreateTempSubdirectory("xlflow-trace-force-").FullName;
        File.WriteAllText(Path.Combine(modulesDir, "XlflowTrace.bas"), "modified");

        Assert.True(ExcelTraceService.CanRemoveTraceSource(modulesDir, force: true));
    }
}
